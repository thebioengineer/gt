as_ooxml_tbl <- function(
    ooxml_type,
    data,
    align = "center",
    split = FALSE,
    keep_with_next = TRUE,
    embedded_heading = FALSE
) {

  # Perform input object validation
  stop_if_not_gt_tbl(data = data)

  tbl_properties <- create_table_properties_ooxml(ooxml_type, data = data, align = align)

  # <a:tblGrid> is not optional in pptx, so create_table_grid must set it
  #
  # things are different in word where we can have w:tblLayoutType="autofit" and then
  # not have a <w:tblGrid> node
  tbl_grid <- create_table_grid_ooxml(ooxml_type, data = data)

  tbl_spanner_rows <- create_spanner_rows_ooxml(ooxml_type, data = data)
  tbl_header_rows <- create_column_header_rows_ooxml(ooxml_type, data = data)
  tbl_body_rows <- create_body_rows_ooxml(ooxml_type, data = data)

  ooxml_tbl(ooxml_type,
    properties = tbl_properties,
    grid       = tbl_grid,
    !!!tbl_spanner_rows,
    !!!tbl_header_rows,
    !!!tbl_body_rows
  )
}


# table properties --------------------------------------------------------

create_table_properties_ooxml <- function(ooxml_type, data , align = c("center", "start", "end"), look = c("first row")) {
  # TODO: set layout as autofit when dt_boxhead_get()$column_width
  #       are all NULL and figure out equivalent in pptx
  ooxml_tbl_properties(ooxml_type,
    justify = align,
    # margins = list(left = 60, right = 60),
    width   = "100%",
    look    = look
  )
}


# table grid --------------------------------------------------------------

create_table_grid_ooxml <- function(ooxml_type, data) {
  boxh <- dt_boxhead_get(data = data)

  widths <- boxh[boxh$type %in% c("default", "stub"), , drop = FALSE]
  # returns vector of column widths where `stub` is first
  widths <- dplyr::arrange(widths, dplyr::desc(type))$column_width

  # widths may be NULL, pct(), px() ...

  ooxml_tblGrid(ooxml_type, !!!widths)
}


# spanner rows ------------------------------------------------------------

create_spanner_rows_ooxml <- function(ooxml_type, data) {
  if (dt_options_get_value(data = data, option = "column_labels_hidden") || !dt_spanners_exists(data = data)) {
    return(NULL)
  }

  # Determine the finalized number of spanner rows
  spanner_row_count <- dt_spanners_matrix_height(data = data, omit_columns_row = TRUE)

  spanner_rows <- lapply(seq_len(spanner_row_count),
    create_spanner_row_cells_ooxml,
    ooxml_type = ooxml_type, data = data
  )

  spanner_rows
}

create_spanner_row_cells_ooxml <- function(ooxml_type, data, span_row_idx) {
  styles_tbl <- dt_styles_get(data = data)
  column_labels_vlines_color <- dt_options_get_value(data = data, option = "column_labels_vlines_color")
  column_labels_border_top_color <- dt_options_get_value(data = data, option = "column_labels_border_top_color")
  column_labels_border_bottom_color <- dt_options_get_value(data = data, option = "column_labels_border_bottom_color")

  spanners <- dt_spanners_print_matrix(data, include_hidden = FALSE)
  spanner_ids <- dt_spanners_print_matrix(data, include_hidden = FALSE, ids = TRUE)

  spanner_row_values <- spanners[span_row_idx,]
  spanner_row_ids <- spanner_ids[span_row_idx,]

  spanners_rle <- rle(spanner_row_ids)
  sig_cells <- c(1, utils::head(cumsum(spanners_rle$lengths) + 1, -1))
  colspans <- ifelse(
    seq_along(spanner_row_values) %in% sig_cells,
    spanners_rle$lengths[match(seq_along(spanner_row_ids), sig_cells)],
    0
  )

  # there are spanners, so the spanners row for the stub are empty cells that continue merge
  stub_cell <- if (dt_stub_df_exists(data = data)) {

    if (span_row_idx == 1) {
      cell_style <- styles_tbl[styles_tbl$locname %in% "stubhead", "styles", drop = TRUE]
      cell_style <- cell_style[1][[1]]

      borders <- list(
        top    = list(color = column_labels_border_top_color),
        bottom = list(size = 8, color = column_labels_border_bottom_color),
        left   = list(color = column_labels_vlines_color),
        right  = list(color = column_labels_vlines_color)
      )

      ooxml_tbl_cell(ooxml_type,
        ooxml_paragraph(ooxml_type,
          ooxml_run(ooxml_type,
            ooxml_text(ooxml_type,
              stubh$label,
              space = cell_style[["cell_text"]][["whitespace"]] %||% "default"
            ),
            properties = ooxml_run_properties(ooxml_type, cell_style = cell_style)
          )
        ),
        properties = ooxml_tbl_cell_properties(ooxml_type,
          borders  = borders,
          fill     = cell_style[["cell_fill"]][["color"]],
          v_align  = cell_style[["cell_text"]][["v_align"]],
          col_span = colspans[i]
        )
      )
    } else {
      borders <- list(
        left   = list(color = column_labels_vlines_color),
        right  = list(color = column_labels_vlines_color),
        bottom = if (span_row_idx == nrow(spanners)) list(size = 8, color = column_labels_border_bottom_color)
      )
      ooxml_tbl_cell(ooxml_type,
        row_span = "continue",
        properties = ooxml_tbl_cell_properties(ooxml_type, borders = borders)
      )
    }


  }

  cells <- lapply(seq_along(spanner_row_values), \(i) {
    if (is.na(spanner_row_ids[i])) {
      borders <- list(
        left = if (i == 1L) { list(color = column_labels_vlines_color) },
        right = if (i == length(spanner_row_values)) { list(color = column_labels_vlines_color) },
        top = if (span_row_idx == 1) { list(color = column_labels_border_top_color) }
      )
      return(ooxml_tbl_cell(ooxml_type,
        properties = ooxml_tbl_cell_properties(ooxml_type,
          borders  = borders
        )
      ))
    }

    if (colspans[i] == 0) {
      return(NULL)
    }

    cell_style <- vctrs::vec_slice(styles_tbl, styles_tbl$locname %in% c("columns_groups") & styles_tbl$grpname %in% spanner_row_ids[i])
    cell_style <- cell_style$styles[1][[1]]

    borders <- list(
      left = if (i == 1) { list(color = column_labels_vlines_color) },
      right = if (i == (length(spanner_row_values) + 1 - colspans[i] )) { list(color = column_labels_vlines_color) },
      bottom = list(size = 8, color = column_labels_border_bottom_color),
      top = if (span_row_idx == 1) { list(size = 8, color = column_labels_border_top_color) }
    )

    ooxml_tbl_cell(ooxml_type,
      ooxml_paragraph(ooxml_type,
        ooxml_run(ooxml_type,
          ooxml_text(ooxml_type,
            spanner_row_values[i],
            space = cell_style[["cell_text"]][["whitespace"]] %||% "default"
          ),
          properties = ooxml_run_properties(ooxml_type, cell_style = cell_style)
        )
      ),
      properties = ooxml_tbl_cell_properties(ooxml_type,
        borders  = borders,
        fill     = cell_style[["cell_fill"]][["color"]],
        v_align  = cell_style[["cell_text"]][["v_align"]],
        col_span = colspans[i]
      )
    )

  })

  ooxml_tbl_row(ooxml_type, stub_cell, !!!cells, is_header = TRUE)
}

# column header rows ------------------------------------------------------

create_column_header_rows_ooxml <- function(ooxml_type, data) {
  NULL
}

# data rows ---------------------------------------------------------------

create_body_rows_ooxml <- function(ooxml_type, data) {
  NULL
}

