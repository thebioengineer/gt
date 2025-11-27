skip_on_cran()

check_suggests <- function() {
  skip_if_not_installed("officer")
}

test_that("word ooxml can be generated from gt object", {

  # Create a one-row table for these tests
  exibble_min <- exibble[1, ]

  ## basic table
  xml <- read_xml_word_nodes(as_word_ooxml(gt(exibble_min)))
  expect_equal(length(xml), 1)
  expect_equal(xml_name(xml), "tbl")
  expect_equal(length(xml_find_all(xml, "//w:keepNext")), 18)

  expect_xml_snapshot(xml)
  expect_equal(
    xml_attr(xml_find_all(xml, "(//w:tr)[1]//w:pPr/w:jc"), "val"),
    c("end", "start", "center", "end", "end", "end", "end", "start", "start")
  )
  expect_equal(
    xml_attr(xml_find_all(xml, "(//w:tr)[2]//w:pPr/w:jc"), "val"),
    c("end", "start", "center", "end", "end", "end", "end", "start", "start")
  )

  ## basic table with title
  gt_tbl_1 <-
    exibble_min |>
    gt() |>
    tab_header(
      title = "TABLE TITLE",
      subtitle = "table subtitle"
    )
  xml_top <- read_xml_word_nodes(as_word_ooxml(gt_tbl_1))
  expect_equal(length(xml_top), 3)
  expect_equal(xml_name(xml_top), c("p", "p", "tbl"))
  expect_xml_snapshot(xml_top[[1]])
  expect_xml_snapshot(xml_top[[2]])
  expect_equal(length(xml_find_all(xml_top, "//w:keepNext")), 20)
  # check autonum nodes
  expect_equal(length(xml_find_all(xml_top, "//w:tr")), 2)
  expect_equal(length(xml_find_all(xml_top, "//w:instrText")), 1)
  expect_equal(length(xml_find_all(xml_top, "//w:fldChar")), 3)

  ## basic table with title added below table
  xml_bottom <- read_xml_word_nodes(as_word_ooxml(gt_tbl_1, caption_location = "bottom"))
  expect_equal(length(xml_bottom), 3)
  expect_equal(xml_name(xml_bottom), c("tbl", "p", "p"))
  expect_equal(xml_top[[3]], xml_bottom[[1]])
  expect_equal(length(xml_find_all(xml_bottom, "//w:tr")), 2)
  expect_equal(length(xml_find_all(xml_bottom, "//w:keepNext")), 18)
  # expect_snapshot_ooxml_word(gt_tbl_1, caption_location = "bottom")

  ## basic table with title embedded on the top of table
  xml_embed <- read_xml_word_nodes(as_word_ooxml(gt_tbl_1, caption_location = "embed"))
  expect_equal(length(xml_embed), 1)
  expect_equal(xml_name(xml_embed), c("tbl"))
  expect_equal(length(xml_find_all(xml_embed, "//w:tr")), 3)
  expect_equal(length(xml_find_all(xml_embed, "//w:keepNext")), 20)

  expect_equal(
    xml_find_all(xml_top, "(//w:tr)[1]")[[1]],
    xml_find_all(xml_embed, "(//w:tr)[2]")[[1]]
  )
  expect_equal(
    xml_find_all(xml_top, "(//w:tr)[2]")[[1]],
    xml_find_all(xml_embed, "(//w:tr)[3]")[[1]]
  )

  ## basic table with split enabled
  xml_split <- read_xml_word_nodes(as_word_ooxml(gt_tbl_1, split = FALSE))
  expect_equal(
    purrr::map_lgl(xml_find_all(xml_split, "//w:trPr"), \(x) {
      length(xml_find_all(x, ".//w:cantSplit")) == 1
    }),
    c(TRUE, TRUE)
  )
  expect_equal(length(xml_find_all(xml_split, "//w:keepNext")), 20)


  ## basic table with autonum disabled
  xml_autonum_false <- read_xml_word_nodes(as_word_ooxml(gt_tbl_1, autonum = FALSE))
  expect_equal(xml_name(xml_autonum_false), c("p", "p", "tbl"))
  expect_xml_snapshot(xml_autonum_false[[1]])
  expect_xml_snapshot(xml_autonum_false[[2]])
  expect_equal(length(xml_find_all(xml_autonum_false, "//w:instrText")), 0)
  expect_equal(length(xml_find_all(xml_autonum_false, "//w:fldChar")), 0)
  expect_equal(length(xml_find_all(xml_autonum_false, "//w:keepNext")), 20)

  ## basic table with keep_with_next disabled
  xml_keep_next_false <- read_xml_word_nodes(as_word_ooxml(gt_tbl_1, keep_with_next = FALSE))
  expect_equal(length(xml_find_all(xml_keep_next_false, "//w:keepNext")), 0)
})

test_that("word ooxml can be generated from gt object with cell styling", {
  ## Table with cell styling
  gt_tbl_2 <-
    exibble[1:4, ] |>
    gt(rowname_col = "char") |>
    tab_row_group("My Row Group 1", c(1:2)) |>
    tab_row_group("My Row Group 2", c(3:4)) |>
    tab_style(
      style = cell_fill(color = "orange"),
      locations = cells_body(columns = c(num, fctr, time, currency, group))
    ) |>
    tab_style(
      style = cell_fill(color = "orange"),
      locations = cells_body(columns = c(num, fctr, time, currency, group))
    ) |>
    tab_style(
      style = cell_text(
        color = "green",
        font = "Biome",
        style = "italic",
        weight = "bold"
      ),
      locations = cells_stub()
    ) |>
    tab_style(
      style = cell_text(color = "blue"),
      locations = cells_row_groups()
    )
  xml <- read_xml_word_nodes(as_word_ooxml(gt_tbl_2, keep_with_next = FALSE))

  # row groups
  expect_equal(xml_attr(xml_find_all(xml, '//w:t[text() = "My Row Group 2"]/../w:rPr/w:color'), "val"), "0000FF")
  expect_equal(xml_attr(xml_find_all(xml, '//w:t[text() = "My Row Group 2"]/../w:rPr/w:sz'), "val"), "20")
  expect_equal(xml_attr(xml_find_all(xml, '//w:t[text() = "My Row Group 1"]/../w:rPr/w:color'), "val"), "0000FF")
  expect_equal(xml_attr(xml_find_all(xml, '//w:t[text() = "My Row Group 1"]/../w:rPr/w:sz'), "val"), "20")

  # body rows
  xml_body <- xml_find_all(xml, "//w:tr")[c(3, 4, 6, 7)]
  purrr::walk(xml_find_all(xml_body, "(.//w:tc)[1]//w:rPr"), \(node) {
    expect_equal(xml_attr(xml_find_all(node, ".//w:rFonts"), "ascii"), "Biome")
    expect_equal(xml_attr(xml_find_all(node, ".//w:sz"), "val"), "20")
    expect_equal(length(xml_find_all(node, ".//w:i")), 1)
    expect_equal(xml_attr(xml_find_all(node, ".//w:color"), "val"), "00FF00")
    expect_equal(length(xml_find_all(node, ".//w:b")), 1)
  })

  # orange cells
  purrr::walk(c(2, 3, 5, 7, 9), \(i) {
    shd <- xml_find_all(xml_body, paste0("(.//w:tc)[", i, "]/w:tcPr/w:shd"))
    expect_equal(xml_attr(shd, "fill"), rep("FFA500", 4))
    expect_equal(xml_attr(shd, "val"), rep("clear", 4))
    expect_equal(xml_attr(shd, "color"), rep("auto", 4))
  })

  # regular cells
  purrr::walk(c(4, 6, 8), \(i) {
    expect_equal(length(xml_find_all(xml_body, paste0("(.//w:tc)[", i, "]/w:tcPr/w:shd"))), 0)
  })

  # ## table with column and span styling
  gt_exibble_min <-
    exibble[1, ] |>
    gt() |>
    tab_spanner("My Span Label", columns = 1:5) |>
    tab_spanner("My Span Label top", columns = 2:4, level = 2) |>
    tab_style(
      style = cell_text(color = "purple"),
      locations = cells_column_labels()
    ) |>
    tab_style(
      style = cell_fill(color = "green"),
      locations = cells_column_labels()
    ) |>
    tab_style(
      style = cell_fill(color = "orange"),
      locations = cells_column_spanners("My Span Label")
    ) |>
    tab_style(
      style = cell_fill(color = "red"),
      locations = cells_column_spanners("My Span Label top")
    )

  xml <- read_xml_word_nodes(as_word_ooxml(gt_exibble_min))

  # level 2 span
  for (j in c(1, 3:7)) {
    expect_equal(xml_text(xml_find_all(xml, paste0("//w:tr[1]/w:tc[", j, "]//w:t"))), "")
  }
  xml_top_span <- xml_find_all(xml, "//w:tr[1]/w:tc[2]")
  expect_equal(xml_attr(xml_find_all(xml_top_span, "./w:tcPr/w:gridSpan"), "val"), "3")
  expect_equal(xml_attr(xml_find_all(xml_top_span, "./w:tcPr/w:shd"), "fill"), "FF0000")
  expect_equal(xml_text(xml_find_all(xml_top_span, ".//w:t")), "My Span Label top")

  # level 1 span
  xml_bottom_span <- xml_find_all(xml, "//w:tr[2]/w:tc[1]")
  expect_equal(xml_attr(xml_find_all(xml_bottom_span, "./w:tcPr/w:gridSpan"), "val"), "5")
  expect_equal(xml_attr(xml_find_all(xml_bottom_span, "./w:tcPr/w:shd"), "fill"), "FFA500")
  expect_equal(xml_text(xml_find_all(xml_bottom_span, ".//w:t")), "My Span Label")
  for (j in c(2:5)) {
    expect_equal(xml_text(xml_find_all(xml, paste0("//w:tr[2]/w:tc[", j, "]//w:t"))), "")
  }

  # columns
  expect_equal(
    xml_attr(xml_find_all(xml, "//w:tr[3]/w:tc/w:tcPr/w:shd"), "fill"),
    rep("00FF00", 9)
  )
  expect_equal(
    xml_attr(xml_find_all(xml, "//w:tr[3]//w:color"), "val"),
    rep("A020F0", 9)
  )

})

test_that("process_text() handles ooxml/word", {
  expect_equal(process_text("simple", context = "ooxml/word"), "simple")
  expect_equal(
    process_text(md("simple <br> markdown"), context = "ooxml/word"),
    process_text(md("simple <br> markdown"), context = "word")
  )
  expect_equal(
    process_text(html("simple <br> html"), context = "ooxml/word"),
    process_text(html("simple <br> html"), context = "word")
  )
})

test_that("word ooxml handles md() and html()", {

  # Create a one-row table for these tests
  exibble_min <- exibble[1, ]

  ## basic table with linebreak in title
  gt_tbl_linebreaks_md <-
    exibble_min |>
    gt() |>
    tab_header(
      title = md("TABLE <br> TITLE"),
      subtitle = md("table <br> subtitle")
    )
  xml <- read_xml_word_nodes(as_word_ooxml(gt_tbl_linebreaks_md))
  expect_equal(
    xml_text(xml_find_all(xml[[1]], "(.//w:r)[last()]//w:t")),
    "TABLE"
  )
  expect_equal(xml_text(xml_find_all(xml[[2]], ".//w:r//w:t")), "TITLE")
  expect_equal(xml_text(xml_find_all(xml[[3]], ".//w:r//w:t")), "table")
  expect_equal(xml_text(xml_find_all(xml[[4]], ".//w:r//w:t")), "subtitle")

  ## basic table with linebreak in title
  gt_tbl_linebreaks_html <-
    exibble_min |>
    gt() |>
    tab_header(
      title = html("TABLE <br> TITLE"),
      subtitle = html("table <br> subtitle")
    )

  xml <- read_xml_word_nodes(as_word_ooxml(gt_tbl_linebreaks_md))
  expect_equal(
    xml_text(xml_find_all(xml[[1]], "(.//w:r)[last()]//w:t")),
    "TABLE"
  )
  expect_equal(xml_text(xml_find_all(xml[[2]], ".//w:r//w:t")), "TITLE")
  expect_equal(xml_text(xml_find_all(xml[[3]], ".//w:r//w:t")), "table")
  expect_equal(xml_text(xml_find_all(xml[[4]], ".//w:r//w:t")), "subtitle")
})

test_that("word ooxml escapes special characters in gt object", {
  df <- data.frame(special_characters = "><&\n\r\"'", stringsAsFactors = FALSE)
  xml <- read_xml_word_nodes(as_word_ooxml(gt(df)))

  expect_snapshot(
    xml_find_all(xml, "(//w:t)[last()]/text()")
  )
})

test_that("word ooxml escapes special characters in gt object footer", {

  gt_tbl <- gt(data.frame(num = 1)) |>
    tab_footnote(footnote = "p < .05, ><&\n\r\"'")

  xml <- read_xml_word_nodes(as_word_ooxml(gt_tbl))
  expect_snapshot(xml_find_all(xml, "//w:tr[last()]//w:t"))
})

test_that("multicolumn stub are supported", {
  test_data <- dplyr::tibble(
    mfr = c("Ford", "Ford", "BMW", "BMW", "Audi"),
    model = c("GT", "F-150", "X5", "X3", "A4"),
    trim = c("Base", "XLT", "xDrive35i", "sDrive28i", "Premium"),
    year = c(2017, 2018, 2019, 2020, 2021),
    hp = c(647, 450, 300, 228, 261),
    msrp = c(447000, 28000, 57000, 34000, 37000)
  )

  # Three-column stub
  triple_stub <- gt(test_data, rowname_col = c("mfr", "model", "trim"))

  # The merge cells on the first column
  xml <- read_xml(as_word_ooxml(triple_stub))
  nodes_Ford <- xml_find_all(xml, ".//w:t[. = 'Ford']")
  expect_equal(xml_attr(xml_find_all(nodes_Ford[[1]], "../../..//w:vMerge"), "val"), "restart")
  expect_equal(xml_attr(xml_find_all(nodes_Ford[[2]], "../../..//w:vMerge"), "val"), "continue")

  nodes_BMW <- xml_find_all(xml, ".//w:t[. = 'BMW']")
  expect_equal(xml_attr(xml_find_all(nodes_BMW[[1]], "../../..//w:vMerge"), "val"), "restart")
  expect_equal(xml_attr(xml_find_all(nodes_BMW[[2]], "../../..//w:vMerge"), "val"), "continue")

  nodes_Audi <- xml_find_all(xml, ".//w:t[. = 'Audi']")
  expect_equal(xml_length(xml_find_all(nodes_Audi[[1]], "../../..//w:vMerge")), 0)

  # no other merge cells
  expect_equal(length(xml_find_all(xml, ".//w:vMerge")), 4)

  # no stub head, i.e. empty text
  expect_equal(
    xml_text(xml_find_all(xml, "(.//w:tr)[1]//w:t")),
    c("", "year", "hp", "msrp")
  )
  tcPr <- xml_find_all(xml, "(.//w:tr)[1]/w:tc/w:tcPr")
  expect_equal(xml_attr(xml_find_all(tcPr[[1]], ".//w:gridSpan"), "val"), "3")
  for (i in 2:4) {
    expect_equal(length(xml_find_all(tcPr[[i]], ".//w:gridSpan")), 0)
  }

  # one label: merged
  xml <- test_data |>
    gt(rowname_col = c("mfr", "model", "trim")) |>
    tab_stubhead("one") |>
    as_word() %>%
    read_xml()
  tcPr <- xml_find_all(xml, "(.//w:tr)[1]/w:tc/w:tcPr")
  expect_equal(xml_attr(xml_find_all(tcPr[[1]], ".//w:gridSpan"), "val"), "3")
  for (i in 2:4) {
    expect_equal(length(xml_find_all(tcPr[[i]], ".//w:gridSpan")), 0)
  }

  expect_equal(
    xml_text(xml_find_all(xml, "(.//w:tr)[1]//w:t")),
    c("one", "year", "hp", "msrp")
  )

  # 3 labels
  xml <- test_data |>
    gt(rowname_col = c("mfr", "model", "trim")) |>
    tab_stubhead(c("one", "two", "three")) |>
    as_word() %>%
    read_xml()

  expect_equal(
    xml_text(xml_find_all(xml, "(.//w:tr)[1]//w:t")),
    c("one", "two", "three", "year", "hp", "msrp")
  )

  # add spanner
  xml <- test_data |>
    gt(rowname_col = c("mfr", "model", "trim")) |>
    tab_stubhead(c("one", "two", "three")) |>
    tab_spanner(label = "span", columns = c(hp, msrp)) |>
    as_word() %>%
    read_xml()

  expect_equal(
    xml_text(xml_find_all(xml, "(.//w:tr)[1]//w:t")),
    c("one", "two", "three", "", "span")
  )
  # first row
  tcPr <- xml_find_all(xml, "(.//w:tr)[1]/w:tc/w:tcPr")
  for (i in 1:3) {
    expect_equal(xml_attr(xml_find_all(tcPr[[i]], ".//w:vMerge"), "val"), "restart")
  }
  expect_equal(xml_attr(xml_find_first(tcPr[[5]], ".//w:gridSpan"), "val"), "2")

  # second row
  tcPr <- xml_find_all(xml, "(.//w:tr)[2]/w:tc/w:tcPr")
  for (i in 1:3) {
    expect_equal(xml_attr(xml_find_all(tcPr[[i]], ".//w:vMerge"), "val"), "continue")
  }

  # spanner - one label
  xml <- test_data |>
    gt(rowname_col = c("mfr", "model", "trim")) |>
    tab_stubhead(c("one")) |>
    tab_spanner(label = "span", columns = c(hp, msrp)) |>
    as_word() %>%
    read_xml()

  expect_equal(
    xml_text(xml_find_all(xml, "(.//w:tr)[1]//w:t")),
    c("one", "", "span")
  )

  # first row
  tcPr <- xml_find_all(xml, "(.//w:tr)[1]/w:tc/w:tcPr")
  expect_equal(xml_attr(xml_find_all(tcPr[[1]], ".//w:vMerge"), "val"), "restart")
  expect_equal(xml_attr(xml_find_all(tcPr[[1]], ".//w:gridSpan"), "val"), "3")
  expect_equal(xml_attr(xml_find_first(tcPr[[3]], ".//w:gridSpan"), "val"), "2")

  # second row
  tcPr <- xml_find_all(xml, "(.//w:tr)[2]/w:tc/w:tcPr")
  expect_equal(xml_attr(xml_find_all(tcPr[[1]], ".//w:vMerge"), "val"), "continue")
  expect_equal(xml_attr(xml_find_all(tcPr[[1]], ".//w:gridSpan"), "val"), "3")

})

test_that("tables can be added to a word doc", {
  check_suggests()

  ## simple table
  gt_exibble_min <-
    exibble[1:2, ] |>
    gt() |>
    tab_header(
      title = "table title",
      subtitle = "table subtitle"
    )

  ## Add table to empty word document
  word_doc <- officer::read_docx() |>
    ooxml_body_add_gt(
      gt_exibble_min,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <- xml2::xml_children(xml2::xml_children(docx$doc_obj$get()))

  ## extract table caption
  docx_table_caption_text <- xml2::xml_text(docx_contents[1:2])

  ## extract table contents
  docx_table_body_header <-
    docx_contents[3] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[3] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: table title", "table subtitle")
  )

  expect_equal(
    xml2::xml_text(xml2::xml_find_all(docx_table_body_header, ".//w:p")),
    c(
      "num", "char", "fctr", "date", "time",
      "datetime", "currency", "row", "group"
    )
  )

  expect_equal(
    lapply(
      docx_table_body_contents,
      FUN = function(x) xml2::xml_text(xml2::xml_find_all(x, ".//w:p"))
    ),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "2.2220",
        "banana",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      )
    )
  )
})

test_that("tables with special characters can be added to a word doc", {

  skip_on_ci()
  check_suggests()

  ## simple table
  gt_exibble_min <-
    exibble[1, ] |>
    dplyr::mutate(special_characters = "><&\"'") |>
    gt() |>
    tab_header(
      title = "table title",
      subtitle = "table subtitle"
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(
      gt_exibble_min,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <- xml2::xml_children(xml2::xml_children(docx$doc_obj$get()))

  ## extract table caption
  docx_table_caption_text <- xml2::xml_text(docx_contents[1:2])

  ## extract table contents
  docx_table_body_header <-
    docx_contents[3] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[3] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: table title", "table subtitle")
  )

  expect_equal(
    xml2::xml_text(xml2::xml_find_all(docx_table_body_header, ".//w:p")),
    c(
      "num", "char", "fctr", "date", "time",
      "datetime", "currency", "row", "group",
      "special_characters"
    )
  )

  expect_equal(
    lapply(
      docx_table_body_contents,
      FUN = function(x) xml2::xml_text(xml2::xml_find_all(x, ".//w:p"))
    ),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a",
        "><&\"'"
      )
    )
  )
})

test_that("tables with embedded titles can be added to a word doc", {

  check_suggests()

  ## simple table
  gt_exibble_min <-
    exibble[1:2, ] |>
    gt() |>
    tab_header(
      title = "table title",
      subtitle = "table subtitle"
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(
      gt_exibble_min,
      caption_location = "embed",
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc, target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <- xml2::xml_children(xml2::xml_children(docx$doc_obj$get()))

  ## extract table contents
  docx_table_body_header <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:t") |>
      xml2::xml_text(),
    c(
      "table title", "table subtitle", "num", "char", "fctr",
      "date", "time", "datetime", "currency", "row", "group"
    )
  )

  expect_equal(
    lapply(docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "2.2220",
        "banana",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      )
    )
  )
})

test_that("tables with spans can be added to a word doc", {
  check_suggests()

  ## simple table
  gt_exibble_min <-
    exibble[1:2, ] |>
    gt() |>
    tab_header(
      title = "table title",
      subtitle = "table subtitle"
    ) |>
    ## add spanner across columns 1:5
    tab_spanner(
      "My Column Span",
      columns = 3:5
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(
      gt_exibble_min,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table caption
  docx_table_caption_text <-
    docx_contents[1:2] |>
    xml2::xml_text()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[3] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[3] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: table title", "table subtitle")
  )

  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c(
      "", "", "My Column Span", "", "", "", "", "num", "char", "fctr",
      "date", "time", "datetime", "currency", "row", "group"
    )
  )

  expect_equal(
    lapply(docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "2.2220",
        "banana",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      )
    )
  )
})

test_that("tables with multi-level spans can be added to a word doc", {

  check_suggests()

  ## simple table
  gt_exibble_min <-
    exibble[1:2, ] |>
    gt() |>
    tab_header(
      title = "table title",
      subtitle = "table subtitle"
    ) |>
    ## add spanner across columns 1:5
    tab_spanner(
      "My 1st Column Span L1",
      columns = 1:5
    ) |>
    tab_spanner(
      "My Column Span L2",
      columns = 2:5,level = 2
    ) |>
    tab_spanner(
      "My 2nd Column Span L1",
      columns = 8:9
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(
      gt_exibble_min,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table caption
  docx_table_caption_text <-
    docx_contents[1:2] |>
    xml2::xml_text()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[3] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[3] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: table title", "table subtitle")
  )

  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c(
      "", "My Column Span L2", "", "", "", "", "My 1st Column Span L1", "", "",
      "My 2nd Column Span L1", "num", "char", "fctr", "date", "time",
      "datetime", "currency", "row", "group"
    )
  )

  expect_equal(
    lapply(
      docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "2.2220",
        "banana",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      )
    )
  )
})

test_that("tables with summaries can be added to a word doc", {
  check_suggests()
  skip("summaries not implemented yet")

  ## simple table
  gt_exibble_min <-
    exibble |>
    dplyr::select(-c(fctr, date, time, datetime)) |>
    gt(rowname_col = "row", groupname_col = "group") |>
    summary_rows(
      groups = everything(),
      columns = num,
      fns = list(
        avg = ~mean(., na.rm = TRUE),
        total = ~sum(., na.rm = TRUE),
        s.d. = ~sd(., na.rm = TRUE)
      ),
      fmt = list(~ fmt_number(.))
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(
      gt_exibble_min,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  ## "" at beginning for stubheader
  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c("", "num", "char", "currency")
  )

  expect_equal(
    lapply(
      docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()),
    list(
      "grp_a",
      c("row_1", "1.111e-01", "apricot", "49.950"),
      c("row_2", "2.222e+00", "banana", "17.950"),
      c("row_3", "3.333e+01", "coconut", "1.390"),
      c("row_4", "4.444e+02", "durian", "65100.000"),
      c("avg", "120.02", "—", "—"),
      c("total", "480.06", "—", "—"),
      c("s.d.", "216.79", "—", "—"),
      "grp_b",
      c("row_5", "5.550e+03", "NA", "1325.810"),
      c("row_6", "NA", "fig", "13.255"),
      c("row_7", "7.770e+05", "grapefruit", "NA"),
      c("row_8", "8.880e+06", "honeydew", "0.440"),
      c("avg", "3,220,850.00", "—", "—"),
      c("total", "9,662,550.00", "—", "—"),
      c("s.d.", "4,916,123.25", "—", "—")
    )
  )

  ## Now place the summary on the top

  ## simple table
  gt_exibble_min_top <-
    exibble |>
    dplyr::select(-c(fctr, date, time, datetime)) |>
    gt(rowname_col = "row", groupname_col = "group") |>
    summary_rows(
      groups = everything(),
      columns = num,
      fns = list(
        avg = ~mean(., na.rm = TRUE),
        total = ~sum(., na.rm = TRUE),
        s.d. = ~sd(., na.rm = TRUE)
      ),
      fmt = list(~ fmt_number(.)),
      side = "top"
    )

  ## Add table to empty word document
  word_doc_top <-
    officer::read_docx() |>
    body_add_gt(
      gt_exibble_min_top,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file_top <- tempfile(fileext = ".docx")
  print(word_doc_top,target = temp_word_file_top)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file_top)
  }

  ## Programmatic Review
  docx_top <- officer::read_docx(temp_word_file_top)

  ## get docx table contents
  docx_contents_top <-
    docx_top$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header_top <-
    docx_contents_top[1] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents_top <-
    docx_contents_top[1] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header_top)

  ## "" at beginning for stub header
  expect_equal(
    docx_table_body_header_top |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c( "", "num", "char", "currency")
  )

  expect_equal(
    lapply(
      docx_table_body_contents_top, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()
    ),
    list(
      "grp_a",
      c("avg", "120.02", "—", "—"),
      c("total", "480.06", "—", "—"),
      c("s.d.", "216.79", "—", "—"),
      c("row_1", "1.111e-01", "apricot", "49.950"),
      c("row_2", "2.222e+00", "banana", "17.950"),
      c("row_3", "3.333e+01", "coconut", "1.390"),
      c("row_4", "4.444e+02", "durian", "65100.000"),
      "grp_b",
      c("avg", "3,220,850.00", "—", "—"),
      c("total", "9,662,550.00", "—", "—"),
      c("s.d.", "4,916,123.25", "—", "—"),
      c("row_5", "5.550e+03", "NA", "1325.810"),
      c("row_6", "NA", "fig", "13.255"),
      c("row_7", "7.770e+05", "grapefruit", "NA"),
      c("row_8", "8.880e+06", "honeydew", "0.440")
    )
  )
})

test_that("tables with grand summaries but no rownames can be added to a word doc", {
  check_suggests()
  skip("grand summaries not yet implemented")

  ## simple table
  gt_exibble_min <-
    exibble |>
    dplyr::select(-c(fctr, date, time, datetime, row, group)) |>
    dplyr::slice(1:3) |>
    gt() |>
    grand_summary_rows(
      c(everything(), -char),
      fns = c("Total" = ~length(.))
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(
      gt_exibble_min,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  ## "" at beginning for stubheader
  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c( "", "num", "char", "currency")
  )

  expect_equal(
    lapply(docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()),
    list(
      c("", "0.1111", "apricot", "49.95"),
      c("", "2.2220", "banana", "17.95"),
      c("", "33.3300", "coconut", "1.39"),
      c("Total", "3", "—", "3")
    )
  )

  ## simple table
  gt_exibble_min_top <-
    exibble |>
    dplyr::select(-c(fctr, date, time, datetime, row, group)) |>
    dplyr::slice(1:3) |>
    gt() |>
    grand_summary_rows(
      c(everything(), -char),
      fns = c("Total" = ~length(.)),
      side = "top"
    )

  ## Add table to empty word document
  word_doc_top <-
    officer::read_docx() |>
    body_add_gt(
      gt_exibble_min_top,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file_top <- tempfile(fileext = ".docx")
  print(word_doc_top, target = temp_word_file_top)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file_top)
  }

  ## Programmatic Review
  docx_top <- officer::read_docx(temp_word_file_top)

  ## get docx table contents
  docx_contents_top <-
    docx_top$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header_top <-
    docx_contents_top[1] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents_top <-
    docx_contents_top[1] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header_top)

  ## "" at beginning for stubheader
  expect_equal(
    docx_table_body_header_top |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c("", "num", "char", "currency")
  )

  expect_equal(
    lapply(
      docx_table_body_contents_top, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()
    ),
    list(
      c("Total", "3","—", "3"),
      c("", "0.1111", "apricot", "49.95"),
      c("", "2.2220", "banana","17.95"),
      c("", "33.3300", "coconut", "1.39")
    )
  )
})

test_that("tables with footnotes can be added to a word doc", {
  check_suggests()
  skip("footnote marks not yet implemented")

  ## simple table
  gt_exibble_min <-
    exibble[1:2, ] |>
    gt() |>
    tab_footnote(
      footnote = md("this is a footer example"),
      locations = cells_column_labels(columns = num )
    ) |>
    tab_footnote(
      footnote = md("this is a second footer example"),
      locations = cells_column_labels(columns = char )
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(
      gt_exibble_min,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc, target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  ## superscripts will display as "true#false" due to
  ## xml being:
  ## <w:vertAlign w:val="superscript"/><w:i>true</w:i><w:t xml:space="default">1</w:t><w:i>false</w:i>,
  ## and being converted to TRUE due to italic being true, then the superscript, then turning off italics
  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c(
      "num1", "char2", "fctr", "date", "time",
      "datetime", "currency", "row", "group"
    )
  )

  ## superscripts will display as "true##" due to
  ## xml being:
  ## <w:vertAlign w:val="superscript"/><w:i>true</w:i><w:t xml:space="default">1</w:t>,
  ## and being converted to TRUE due to italic being true, then the superscript,
  expect_equal(
    lapply(
      docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()
    ),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "2.2220",
        "banana",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      ),
      c("1this is a footer example"),
      c("2this is a second footer example")
    )
  )
})

test_that("tables with source notes can be added to a word doc", {
  check_suggests()
  skip("source notes not yet implemented")

  ## simple table
  gt_exibble_min <-
    exibble[1:2, ] |>
    gt() |>
    tab_source_note(source_note = "this is a source note example")

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(gt_exibble_min, align = "center")

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc, target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c(
      "num",
      "char",
      "fctr",
      "date",
      "time",
      "datetime",
      "currency",
      "row",
      "group"
    )
  )

  expect_equal(
    lapply(
      docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()
    ),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "2.2220",
        "banana",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      ),
      c("this is a source note example")
    )
  )
})

test_that("long tables can be added to a word doc", {
  check_suggests()

  ## simple table
  gt_letters <-
    dplyr::tibble(
      upper_case = c(LETTERS,LETTERS),
      lower_case = c(letters,letters)
    ) |>
    gt() |>
    tab_header(title = "LETTERS")

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(gt_letters, align = "center")

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table caption
  docx_table_caption_text <-
    docx_contents[1] |>
    xml2::xml_text()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[2] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[2] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: LETTERS")
  )

  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c("upper_case", "lower_case")
  )

  expect_equal(
    lapply(
      docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()
    ),
    lapply(c(1:26,1:26), function(i) c(LETTERS[i], letters[i]))
  )
})

test_that("long tables with spans can be added to a word doc", {
  check_suggests()

  ## simple table
  gt_letters <-
    dplyr::tibble(
      upper_case = c(LETTERS,LETTERS),
      lower_case = c(letters,letters)
    ) |>
    gt() |>
    tab_header(title = "LETTERS") |>
    tab_spanner(
      "LETTERS",
      columns = 1:2
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(gt_letters, align = "center")

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table caption
  docx_table_caption_text <-
    docx_contents[1] |>
    xml2::xml_text()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[2] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[2] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: LETTERS")
  )

  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c("LETTERS", "upper_case", "lower_case")
  )

  expect_equal(
    lapply(
      docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()
    ),
    lapply(c(1:26,1:26), function(i) c(LETTERS[i], letters[i]))
  )
})

test_that("tables with cell & text coloring can be added to a word doc - no spanner", {

  check_suggests()

  ## simple table
  gt_exibble_min <-
    exibble[1:4, ] |>
    gt(rowname_col = "char") |>
    tab_row_group("My Row Group 1", c(1:2)) |>
    tab_row_group("My Row Group 2", c(3:4)) |>
    tab_style(
      style = cell_fill(color = "orange"),
      locations = cells_body(columns = c(num, fctr, time, currency, group))
    ) |>
    tab_style(
      style = cell_text(
        color = "green",
        font = "Biome"
      ),
      locations = cells_stub()
    ) |>
    tab_style(
      style = cell_text(size = 25, v_align = "middle"),
      locations = cells_body(columns = c(num, fctr, time, currency, group))
    ) |>
    tab_style(
      style = cell_text(
        color = "blue",
        stretch = "extra-expanded"
      ),
      locations = cells_row_groups()
    ) |>
    tab_style(
      style = cell_text(color = "teal"),
      locations = cells_column_labels()
    ) |>
    tab_style(
      style = cell_fill(color = "green"),
      locations = cells_column_labels()
    ) |>
    tab_style(
      style = cell_fill(color = "pink"),
      locations = cells_stubhead()
    )

  if (!testthat::is_testing() && interactive()) {
    print(gt_exibble_min)
  }

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(gt_exibble_min, align = "center")

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tr") |>
    setdiff(docx_table_body_header)

  ## header
  expect_equal(
    docx_table_body_header |>
      xml2::xml_find_all(".//w:p") |>
      xml2::xml_text(),
    c("", "num", "fctr", "date", "time", "datetime", "currency", "row", "group")
  )
  expect_equal(
    lapply(
      docx_table_body_header, function(x) x |> xml2::xml_find_all(".//w:shd") |> xml2::xml_attr(attr = "fill")
    ),
    list(c("FFC0CB", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00"))
  )

  expect_equal(
    lapply(
      docx_table_body_header, function(x) x |> xml2::xml_find_all(".//w:color") |> xml2::xml_attr(attr = "val")
    ),
    list(c("008080", "008080", "008080", "008080", "008080", "008080", "008080", "008080"))
  )

  ## cell background styling
  expect_equal(
    lapply(
      docx_table_body_contents,
      function(x) {
        x |> xml2::xml_find_all(".//w:tc") |> lapply(function(y) {
          y |> xml2::xml_find_all(".//w:shd") |> xml2::xml_attr(attr = "fill")
        })
      }
    ),
    list(
      list(character()),
      list(character(), "FFA500", "FFA500", character(), "FFA500", character(), "FFA500", character(), "FFA500"),
      list(character(), "FFA500", "FFA500", character(), "FFA500", character(), "FFA500", character(), "FFA500"),
      list(character()),
      list(character(), "FFA500", "FFA500", character(), "FFA500", character(),"FFA500", character(), "FFA500"),
      list(character(), "FFA500", "FFA500", character(), "FFA500", character(),"FFA500", character(), "FFA500")
    )
  )

  ## cell text styling
  expect_equal(
    lapply(
      docx_table_body_contents,
      function(x) {
        x |> xml2::xml_find_all(".//w:tc") |> lapply(function(y) {
          y |> xml2::xml_find_all(".//w:color") |> xml2::xml_attr(attr = "val")
        })
      }
    ),
    list(
      list("0000FF"),
      list("00FF00",character(), character(), character(), character(), character(), character(), character(), character()),
      list("00FF00",character(), character(), character(), character(), character(), character(), character(), character()),
      list("0000FF"),
      list("00FF00", character(), character(), character(), character(), character(), character(), character(), character()),
      list("00FF00", character(), character(), character(), character(), character(), character(), character(), character())
    )
  )

  expect_equal(
    lapply(
      docx_table_body_contents, function(x)
      x |> xml2::xml_find_all(".//w:p") |> xml2::xml_text()
    ),
    list(
      "My Row Group 2",
      c(
        "coconut",
        "33.3300",
        "three",
        "2015-03-15",
        "15:45",
        "2018-03-03 03:44",
        "1.39",
        "row_3",
        "grp_a"
      ),
      c(
        "durian",
        "444.4000",
        "four",
        "2015-04-15",
        "16:50",
        "2018-04-04 15:55",
        "65100.00",
        "row_4",
        "grp_a"
      ),
      "My Row Group 1",
      c(
        "apricot",
        "0.1111",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "banana",
        "2.2220",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      )
    )
  )
})

test_that("tables with cell & text coloring can be added to a word doc - with spanners", {

  check_suggests()

  ## simple table
  gt_exibble_min <-
    exibble[1:4, ] |>
    gt(rowname_col = "char") |>
    tab_row_group("My Row Group 1", c(1:2)) |>
    tab_row_group("My Row Group 2", c(3:4)) |>
    tab_spanner("My Span Label", columns = 1:5) |>
    tab_spanner("My Span Label top", columns = 2:4, level = 2) |>
    tab_style(
      style = cell_text(color = "purple"),
      locations = cells_column_labels()
    ) |>
    tab_style(
      style = cell_fill(color = "green"),
      locations = cells_column_labels()
    ) |>
    tab_style(
      style = cell_fill(color = "orange"),
      locations = cells_column_spanners("My Span Label")
    ) |>
    tab_style(
      style = cell_fill(color = "red"),
      locations = cells_column_spanners("My Span Label top")
    ) |>
    tab_style(
      style = cell_fill(color = "pink"),
      locations = cells_stubhead()
    )

  if (!testthat::is_testing() && interactive()) {
    print(gt_exibble_min)
  }

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() |>
    ooxml_body_add_gt(gt_exibble_min, align = "center")

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() && interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() |>
    xml2::xml_children() |>
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[1] |>
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  ## header
  expect_equal(
    docx_table_body_header |> xml2::xml_find_all(".//w:p") |> xml2::xml_text(),
    c("", "", "My Span Label top", "", "", "", "", "",
      "", "My Span Label", "", "", "", "",
      "", "num", "fctr", "date", "time", "datetime", "currency", "row", "group")
  )

  expect_equal(
    lapply(docx_table_body_header, function(x) {
      x |> xml2::xml_find_all(".//w:tc") |> lapply(function(y) {
        y |> xml2::xml_find_all(".//w:shd") |> xml2::xml_attr(attr = "fill")
      })}),
    list(
      list("FFC0CB", character(0L), "FF0000", character(0L), character(0L), character(0L), character(0L), character(0L)),
      list(character(0L), "FFA500", character(0L), character(0L), character(0L), character(0L)),
      list(character(0L), "00FF00", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00")
      )
  )

  expect_equal(
    lapply(docx_table_body_header, function(x) {
      x |> xml2::xml_find_all(".//w:tc") |> lapply(function(y) {
        y |> xml2::xml_find_all(".//w:color") |> xml2::xml_attr(attr = "val")
      })}),
    list(
      list(character(0L), character(0L), character(0L), character(0L),character(0L), character(0L), character(0L), character(0L)),
      list(character(0L), character(0L), character(0L), character(0L),character(0L), character(0L)),
      list(character(0L), "A020F0", "A020F0", "A020F0", "A020F0", "A020F0", "A020F0", "A020F0", "A020F0")
      )
  )
})

