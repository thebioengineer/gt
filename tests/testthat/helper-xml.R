#' @title Add gt table into a Word document
#' @description Add a gt into a Word document.
#' @param x `rdocx` object
#' @param value `gt` object
#' @param align left, center (default) or right.
#' @param caption_location top (default), bottom, or embed Indicating if the title and subtitle should be listed above, below, or be embedded in the table
#' @param caption_align left (default), center, or right. Alignment of caption (title and subtitle). Used when `caption_location` is not "embed".
#' @param split set to TRUE if you want to activate Word
#' option 'Allow row to break across pages'.
#' @param keep_with_next Word option 'keep rows together' can be
#' activated when TRUE. It avoids page break within tables.
#' @param pos where to add the gt table relative to the cursor,
#' one of "after" (default), "before", "on" (end of line).
#'
#' @seealso \link[flextable]{body_add_flextable}
#'
#' @examples
#'
#' library(officer)
#' library(gt)
#' gt_tbl <- gt( head( exibble ) )
#'
#' doc <- read_docx()
#' doc <- body_add_gt(doc, value = gt_tbl)
#' fileout <- tempfile(fileext = ".docx")
#' print(doc, target = fileout)
#' @noRd
body_add_gt <- function(
    x,
    value,
    align = "center",
    pos = c("after","before","on"),
    caption_location = c("top","bottom","embed"),
    caption_align = "left",
    split = FALSE,
    keep_with_next = TRUE
) {

  ## check that officer is available
  rlang::check_installed("officer", "to add gt tables to word documents.")

  ## check that inputs are an officer rdocx and gt tbl
  stopifnot(inherits(x, "rdocx"))
  stopifnot(inherits(value, "gt_tbl"))

  pos <- rlang::arg_match(pos)
  caption_location <- rlang::arg_match(caption_location)

  # Build all table data objects through a common pipeline
  value <- build_data(data = value, context = "word")

  ## Create and add table caption if it is to come before the table
  if (caption_location %in% c("top")) {
    header_xml <-
      as_word_tbl_header_caption(
        data = value,
        align = caption_align,
        split = split,
        keep_with_next = keep_with_next
      )

    if (!identical(header_xml,c(""))) {
      for (header_component in header_xml) {
        x <- officer::body_add_xml(x, str = header_component, pos) |>
          suppressWarnings()
      }
    }
  }

  ## Create and add the table to the docxr. If the
  tbl_xml <-
    as_word_tbl_body(
      data = value,
      align = align,
      split = split,
      keep_with_next = keep_with_next,
      embedded_heading = identical(caption_location, "embed")
    ) |>
    as.character()

  x <-
    officer::body_add_xml(x, str = tbl_xml, pos) |>
    suppressWarnings()

  ## Create and add table caption if it is to come after the table
  if (caption_location %in% c("bottom")) {

    ## set keep_with_next to false here to prevent it trying to keep with non-table content
    header_xml <-
      as_word_tbl_header_caption(
        data = value,
        align = caption_align,
        split = split,
        keep_with_next = FALSE
      )

    if (!identical(header_xml,c(""))) {
      for (header_component in header_xml) {
        x <-
          officer::body_add_xml(x, str = header_component, pos) |>
          suppressWarnings()
      }
    }
  }

  x
}
