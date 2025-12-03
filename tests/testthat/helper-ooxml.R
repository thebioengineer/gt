ooxml_body_add_gt <- function(
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

  pos <- rlang::arg_match(pos)

  xml <- as_word_ooxml(value,
    align = align,
    caption_location = caption_location,
    caption_align = caption_align,
    split = split,
    keep_with_next = keep_with_next
  )
  suppressWarnings({
    for (i in seq_along(xml)) {
      x <- officer::body_add_xml(x, xml[[i]], pos)
    }
  })
  x
}

read_xml_word_nodes <- function(x) {
  xml2::xml_children(suppressWarnings(xml2::read_xml(paste0(
    '<w:wrapper xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">',
    paste(x, collapse = ""),
    "</w:wrapper>"
  ))))
}

read_pptx_word_nodes <- function(x) {
  xml2::xml_children(suppressWarnings(xml2::read_xml(paste0(
    '<w:wrapper xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    paste(x, collapse = ""),
    "</w:wrapper>"
  ))))
}

expect_xml_snapshot <- function(xml) {
  expect_snapshot(writeLines(as.character(xml)))
}
