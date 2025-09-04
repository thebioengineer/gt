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

  # Create the table properties component
  table_props_component <-
    create_table_props_component_ooxml(ooxml_type = ooxml_type, data = data, align = align)

  table_props_component
}

create_table_props_component_ooxml <- function(data, ooxml_type, align = c("center", "start", "end"), look = c("first row")) {
  ooxml_tbl_properties(ooxml_type,
    justify = align,
    margins = list(left = 60, right = 60),
    width   = NULL,
    look    = look
  )
}
