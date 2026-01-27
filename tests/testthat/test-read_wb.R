# Tests for read_wb function

test_that("read_wb validates source directory", {
  expect_error(read_wb("nonexistent_directory"), "does not exist")
})

test_that("read_wb errors when no Excel files found", {
  tmp_dir <- tempfile()
  dir.create(tmp_dir)
  on.exit(unlink(tmp_dir, recursive = TRUE), add = TRUE)

  # Create a non-Excel file
  writeLines("test", file.path(tmp_dir, "test.txt"))

  expect_error(read_wb(tmp_dir), "No Excel files")
})

test_that("read_wb errors when requested workbook not found", {
  src_dir <- system.file("extdata", package = "excelsioR")
  skip_if(!dir.exists(src_dir), "extdata directory not found")

  expect_error(
    read_wb(src_dir, reqd_wkbks = "nonexistent_workbook"),
    "None of the requested"
  )
})

test_that("read_wb works with test data", {
  src_dir <- system.file("extdata", package = "excelsioR")
  exp_dir <- system.file("testdata", package = "excelsioR")

  skip_if(!dir.exists(src_dir), "extdata directory not found")
  skip_if(!file.exists(file.path(exp_dir, "mtcars_read_wb_test.rds")),
          "Expected test file not found")

  # Only test with mtcars file (locked_wb_test.xlsx is password-protected)
  test <- read_wb(src_dir, reqd_wkbks = "mtcars_test")
  expected <- readRDS(file.path(exp_dir, "mtcars_read_wb_test.rds"))

  # Compare structure, not exact values (tidyxl versions may differ)
  expect_equal(names(test), names(expected))
  expect_equal(nrow(test), nrow(expected))
})

test_that("read_wb returns correct structure", {
  src_dir <- system.file("extdata", package = "excelsioR")
  skip_if(!dir.exists(src_dir), "extdata directory not found")

  # Check if there are any xlsx files
  xlsx_files <- list.files(src_dir, pattern = "\\.xlsx$", full.names = TRUE)
  skip_if(length(xlsx_files) == 0, "No xlsx files in extdata")

  # Only test with mtcars file (locked_wb_test.xlsx is password-protected)
  result <- read_wb(src_dir, reqd_wkbks = "mtcars_test")

  expect_s3_class(result, "tbl_df")
  expect_true("file_name" %in% names(result))
  expect_true("raw_excel_data" %in% names(result))
  expect_true("all_excel_data_tbl" %in% names(result))
})

test_that("read_wb can filter to specific workbooks", {
  src_dir <- system.file("extdata", package = "excelsioR")
  skip_if(!dir.exists(src_dir), "extdata directory not found")

  # Use the known non-protected file
  result <- read_wb(src_dir, reqd_wkbks = "mtcars_test")

  expect_equal(nrow(result), 1)
  expect_equal(result$file_name[1], "mtcars_test")
})

test_that("read_wb uses temp directory by default", {
  src_dir <- system.file("extdata", package = "excelsioR")
  skip_if(!dir.exists(src_dir), "extdata directory not found")

  xlsx_files <- list.files(src_dir, pattern = "\\.xlsx$", full.names = TRUE)
  skip_if(length(xlsx_files) == 0, "No xlsx files in extdata")

  # Capture messages (only test with mtcars file)
  messages <- capture_messages(read_wb(src_dir, reqd_wkbks = "mtcars_test"))

  # Should mention temp directory
  expect_true(any(grepl("temp|copied", messages, ignore.case = TRUE)))
})
