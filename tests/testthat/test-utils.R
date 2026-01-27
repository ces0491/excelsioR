# Tests for utility functions in R/utils.R

test_that("col_to_num handles single letters", {
  expect_equal(col_to_num("A"), 1L)
  expect_equal(col_to_num("B"), 2L)
  expect_equal(col_to_num("Z"), 26L)
})

test_that("col_to_num handles multi-letter columns", {
  expect_equal(col_to_num("AA"), 27L)
  expect_equal(col_to_num("AB"), 28L)
  expect_equal(col_to_num("AZ"), 52L)
  expect_equal(col_to_num("BA"), 53L)
  expect_equal(col_to_num("ZZ"), 702L)
  expect_equal(col_to_num("AAA"), 703L)
  expect_equal(col_to_num("XFD"), 16384L)  # Max Excel column
})

test_that("col_to_num is case insensitive", {
  expect_equal(col_to_num("a"), 1L)
  expect_equal(col_to_num("aa"), 27L)
  expect_equal(col_to_num("Aa"), 27L)
})

test_that("col_to_num handles whitespace", {
  expect_equal(col_to_num(" A "), 1L)
  expect_equal(col_to_num(" AA "), 27L)
})

test_that("col_to_num errors on invalid input", {
  expect_error(col_to_num("1"), "only letters")
  expect_error(col_to_num("A1"), "only letters")
  expect_error(col_to_num(""), "only letters")
})

test_that("num_to_col handles single-letter columns", {
  expect_equal(num_to_col(1), "A")
  expect_equal(num_to_col(2), "B")
  expect_equal(num_to_col(26), "Z")
})

test_that("num_to_col handles multi-letter columns", {
  expect_equal(num_to_col(27), "AA")
  expect_equal(num_to_col(28), "AB")
  expect_equal(num_to_col(52), "AZ")
  expect_equal(num_to_col(53), "BA")
  expect_equal(num_to_col(702), "ZZ")
  expect_equal(num_to_col(703), "AAA")
  expect_equal(num_to_col(16384), "XFD")
})

test_that("num_to_col errors on invalid input", {
  expect_error(num_to_col(0), "positive integer")
  expect_error(num_to_col(-1), "positive integer")
  expect_error(num_to_col("A"), "positive integer")
})

test_that("num_to_col and col_to_num are inverses", {
  for (n in c(1, 26, 27, 52, 100, 256, 702, 703, 1000, 16384)) {
    expect_equal(col_to_num(num_to_col(n)), n)
  }
})

test_that("parse_cell_reference works for simple references", {
  result <- parse_cell_reference("A1")
  expect_equal(result$col, 1L)
  expect_equal(result$row, 1L)

  result <- parse_cell_reference("Z99")
  expect_equal(result$col, 26L)
  expect_equal(result$row, 99L)
})

test_that("parse_cell_reference works for multi-letter columns", {
  result <- parse_cell_reference("AA1")
  expect_equal(result$col, 27L)
  expect_equal(result$row, 1L)

  result <- parse_cell_reference("AA100")
  expect_equal(result$col, 27L)
  expect_equal(result$row, 100L)

  result <- parse_cell_reference("XFD1048576")
  expect_equal(result$col, 16384L)
  expect_equal(result$row, 1048576L)
})

test_that("parse_cell_reference is case insensitive", {
  expect_equal(parse_cell_reference("a1"), parse_cell_reference("A1"))
  expect_equal(parse_cell_reference("aa100"), parse_cell_reference("AA100"))
})

test_that("parse_cell_reference handles whitespace", {
  result <- parse_cell_reference(" A1 ")
  expect_equal(result$col, 1L)
  expect_equal(result$row, 1L)
})

test_that("parse_cell_reference errors on invalid input", {
  expect_error(parse_cell_reference(""), "Invalid cell reference")
  expect_error(parse_cell_reference("1"), "missing column")
  expect_error(parse_cell_reference("A"), "missing row")
  expect_error(parse_cell_reference("A0"), "invalid row")
  expect_error(parse_cell_reference("A-1"), "invalid row")
})

test_that("validate_workbook_path accepts valid paths", {
  expect_silent(validate_workbook_path("test.xlsx"))
  expect_silent(validate_workbook_path("path/to/file.xlsx"))
  expect_silent(validate_workbook_path("file.xls"))
  expect_silent(validate_workbook_path("file.xlsm"))
})

test_that("validate_workbook_path rejects invalid paths", {
  expect_error(validate_workbook_path(NULL), "cannot be NULL")
  expect_error(validate_workbook_path(c("a.xlsx", "b.xlsx")), "single character")
  expect_error(validate_workbook_path(123), "single character")
  expect_error(validate_workbook_path("file.csv"), "Excel file extension")
  expect_error(validate_workbook_path("file.txt"), "Excel file extension")
})

test_that("validate_workbook_path checks file existence when required", {
  expect_error(
    validate_workbook_path("nonexistent_file.xlsx", must_exist = TRUE),
    "does not exist"
  )
})
