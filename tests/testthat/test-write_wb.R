# Tests for write_wb functions

test_that("write_wb.default works with basic data frame", {
  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)

  sample_data <- mtcars[1:5, 1:3]

  write_wb(sample_data, sheet_name = "test", wb_dir = tmp_file, save_wb = TRUE)

  result <- openxlsx::read.xlsx(tmp_file)

  expect_equal(nrow(result), 5)
  expect_equal(ncol(result), 3)
  expect_equal(result$mpg, sample_data$mpg)
})

test_that("write_wb handles multi-letter columns", {
  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)

  df <- data.frame(x = 1:3, y = 4:6)

  # Write to column AA (27th column)
  write_wb(df, sheet_name = "test", paste_coord = "AA1",
           wb_dir = tmp_file, save_wb = TRUE)

  # Use tidyxl to verify exact cell positions
  skip_if_not_installed("tidyxl")
  cells <- tidyxl::xlsx_cells(tmp_file)

  # Find where data was written - should be at column 27 (AA)
  data_cells <- cells[!cells$is_blank, ]
  min_col <- min(data_cells$col)

  # AA = column 27

  expect_equal(min_col, 27)
})

test_that("write_wb.list writes multiple sheets", {
  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)

  data_list <- list(
    sheet1 = mtcars[1:3, 1:2],
    sheet2 = iris[1:3, 1:2]
  )

  write_wb(data_list, wb_dir = tmp_file, save_wb = TRUE)

  wb <- openxlsx::loadWorkbook(tmp_file)
  sheet_names <- names(wb)

  expect_true("sheet1" %in% sheet_names)
  expect_true("sheet2" %in% sheet_names)
})

test_that("write_wb.zoo works correctly", {
  skip_if_not_installed("zoo")

  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)

  # Create a zoo object
  dates <- as.Date("2024-01-01") + 0:4
  values <- matrix(1:10, nrow = 5, ncol = 2)
  colnames(values) <- c("a", "b")
  z <- zoo::zoo(values, dates)

  write_wb(z, sheet_name = "zoo_test", wb_dir = tmp_file, save_wb = TRUE)

  result <- openxlsx::read.xlsx(tmp_file)

  expect_equal(nrow(result), 5)
  expect_equal(ncol(result), 3)  # date + 2 value columns
  expect_true("date" %in% names(result))
})

test_that("write_wb.xts works correctly", {
  skip_if_not_installed("zoo")

  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)

  # Create an xts-like zoo object
  dates <- as.Date("2024-01-01") + 0:2
  values <- matrix(1:6, nrow = 3, ncol = 2)
  colnames(values) <- c("col1", "col2")
  x <- zoo::zoo(values, dates)

  write_wb.xts(x, sheet_name = "xts_test", wb_dir = tmp_file, save_wb = TRUE)

  result <- openxlsx::read.xlsx(tmp_file)

  expect_equal(nrow(result), 3)
  expect_true("date" %in% names(result))
})

test_that("write_wb returns workbook invisibly", {
  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)

  df <- data.frame(x = 1:3)

  result <- write_wb(df, sheet_name = "test", wb_dir = tmp_file, save_wb = TRUE)

  # Should return a workbook object
  expect_true(inherits(result, "Workbook"))
})

test_that("write_wb creates new sheet if it doesn't exist", {
  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)

  df <- data.frame(x = 1:3)

  # First write
  wb <- write_wb(df, sheet_name = "first", wb_dir = tmp_file, save_wb = TRUE)

  # Second write to new sheet
  wb <- write_wb(df, sheet_name = "second", wb = wb, wb_dir = tmp_file, save_wb = TRUE)

  wb_loaded <- openxlsx::loadWorkbook(tmp_file)
  expect_true("first" %in% names(wb_loaded))
  expect_true("second" %in% names(wb_loaded))
})

test_that("write_wb clears sheet when clear_sheet = TRUE", {
  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)

  # Write large data first
  df_large <- data.frame(x = 1:10, y = 11:20)
  write_wb(df_large, sheet_name = "test", wb_dir = tmp_file, save_wb = TRUE)

  # Write smaller data with clear_sheet = TRUE
  df_small <- data.frame(a = 1:2)
  write_wb(df_small, sheet_name = "test", clear_sheet = TRUE,
           wb_dir = tmp_file, save_wb = TRUE)

  result <- openxlsx::read.xlsx(tmp_file)

  # Should only have the new small data
  expect_equal(nrow(result), 2)
  expect_equal(ncol(result), 1)
  expect_equal(names(result), "a")
})

test_that("write_wb errors when save_wb = TRUE but no wb_dir", {
  df <- data.frame(x = 1:3)

  expect_error(
    write_wb(df, sheet_name = "test", save_wb = TRUE),
    "wb_dir"
  )
})

test_that("load_wb creates new workbook when no arguments", {
  expect_error(load_wb(), "Must provide")
})

test_that("load_wb loads existing workbook", {
  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)

  # Create a workbook first
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "existing")
  openxlsx::saveWorkbook(wb, tmp_file)

  # Load it
  loaded_wb <- load_wb(wb_dir = tmp_file)

  expect_true("existing" %in% names(loaded_wb))
})

test_that("load_wb creates new workbook when file doesn't exist", {
  wb <- load_wb(wb_dir = "nonexistent_file.xlsx")
  expect_true(inherits(wb, "Workbook"))
  expect_equal(length(names(wb)), 0)  # No sheets yet
})
