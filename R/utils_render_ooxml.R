as_ooxml_tbl_body <- function(
    ooxml_type,
    data,
    align = "center",
    split = FALSE,
    keep_with_next = TRUE,
    embedded_heading = FALSE
) {

  # Perform input object validation
  stop_if_not_gt_tbl(data = data)

  tbl_properties <- create_table_props_component_ooxml(ooxml_type = ooxml_type, data = data, align = align)

  # tbl_grid       <- create_table_grid(ooxml_type = ooxml_type, data = data)
  tbl_grid <- NULL

  tbl_properties
}

create_table_props_component_ooxml <- function(data, ooxml_type, align = c("center", "start", "end"), look = c("first row")) {
  ooxml_tbl_properties(ooxml_type,
    justify = align,
    margins = list(left = 60, right = 60),
    width   = NULL,
    look    = look
  )
}

create_table_grid <- function(data, ooxml_type) {
  boxh <- dt_boxhead_get(data = data)

  widths <- boxh[boxh$type %in% c("default", "stub"), , drop = FALSE]
  # returns vector of column widths where `stub` is first
  widths <- dplyr::arrange(widths, dplyr::desc(type))$column_width

  # widths may be NULL, pct(), px() ...

  ooxml_tblGrid(ooxml_type, widths)
}
