# Tests for read_excel functions

test_that("read_excel validates input columns", {
  bad_df <- data.frame(path = "test.xlsx")

  expect_error(read_excel(bad_df, "Sheet1"), "file_path")
})

test_that("read_excel_single checks file existence", {
  expect_error(
    read_excel_single("nonexistent_file.xlsx", NA),
    "doesn't exist"
  )
})

test_that("read_excel works with test data", {
  test_file <- "mtcars_test.xlsx"
  expected_file <- "mtcars_expected.rds"
  src_dir <- system.file("testdata", package = "excelsioR")

  skip_if(!file.exists(file.path(src_dir, test_file)), "Test file not found")
  skip_if(!file.exists(file.path(src_dir, expected_file)), "Expected file not found")

  src_df <- data.frame(
    file_path = file.path(src_dir, test_file),
    file_name = test_file
  )

  test_tbl <- read_excel(src_df, "Sheet1")

  expect_true("raw_excel_data" %in% names(test_tbl))
  expect_true("file_name" %in% names(test_tbl))

  test_list <- test_tbl$raw_excel_data[[1]]
  test_df <- test_list[[1]]

  expected_df <- readRDS(file.path(src_dir, expected_file))

  # Compare only common columns (tidyxl versions may have different columns)
  common_cols <- intersect(names(test_df), names(expected_df))
  expect_true(length(common_cols) > 0)
  expect_equal(test_df[common_cols], expected_df[common_cols])
})

test_that("read_excel returns tibble with correct structure", {
  test_file <- "mtcars_test.xlsx"
  src_dir <- system.file("testdata", package = "excelsioR")

  skip_if(!file.exists(file.path(src_dir, test_file)), "Test file not found")

  src_df <- data.frame(
    file_path = file.path(src_dir, test_file),
    file_name = "mtcars"
  )

  result <- read_excel(src_df, "Sheet1")

  expect_s3_class(result, "tbl_df")
  expect_true("file_name" %in% names(result))
  expect_true("raw_excel_data" %in% names(result))
})

test_that("read_excel_single reads all sheets when reqd_sheets is NA", {
  test_file <- "mtcars_test.xlsx"
  src_dir <- system.file("testdata", package = "excelsioR")
  file_path <- file.path(src_dir, test_file)

  skip_if(!file.exists(file_path), "Test file not found")

  result <- read_excel_single(file_path, NA)

  expect_true(is.list(result))
  expect_true(length(result) >= 1)
})
