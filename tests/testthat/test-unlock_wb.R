# Tests for unlock_wb function

test_that("unlock_wb requires a backend package", {
  skip_if(
    requireNamespace("rpxl", quietly = TRUE) ||
      requireNamespace("XLConnect", quietly = TRUE),
    "A backend is installed, skipping missing package test"
  )

  file_dir_df <- data.frame(
    file_path = "test.xlsx",
    file_name = "test"
  )

  expect_error(unlock_wb(file_dir_df, wb_password = "test"), "No backend available")
})

test_that("unlock_wb validates file_dir input", {
  skip_if(
    !requireNamespace("rpxl", quietly = TRUE) &&
      !requireNamespace("XLConnect", quietly = TRUE),
    "No backend installed"
  )

  # Missing required columns
  bad_df <- data.frame(path = "test.xlsx")

  expect_error(unlock_wb(bad_df, wb_password = "test"), "file_path")
})

test_that("unlock_wb requires password in non-RStudio environment", {
  skip_if(
    !requireNamespace("rpxl", quietly = TRUE) &&
      !requireNamespace("XLConnect", quietly = TRUE),
    "No backend installed"
  )
  skip_if(rstudioapi::isAvailable(), "Running in RStudio")

  file_dir_df <- data.frame(
    file_path = "test.xlsx",
    file_name = "test"
  )

  expect_error(unlock_wb(file_dir_df), "Password required")
})

test_that("unlock_wb rejects empty password", {
  skip_if(
    !requireNamespace("rpxl", quietly = TRUE) &&
      !requireNamespace("XLConnect", quietly = TRUE),
    "No backend installed"
  )

  file_dir_df <- data.frame(
    file_path = "test.xlsx",
    file_name = "test"
  )

  expect_error(unlock_wb(file_dir_df, wb_password = ""), "cannot be empty")
})

test_that("unlock_wb checks file existence", {
  skip_if(
    !requireNamespace("rpxl", quietly = TRUE) &&
      !requireNamespace("XLConnect", quietly = TRUE),
    "No backend installed"
  )

  file_dir_df <- data.frame(
    file_path = "nonexistent_file.xlsx",
    file_name = "nonexistent"
  )

  expect_error(
    unlock_wb(file_dir_df, wb_password = "test"),
    "does not exist"
  )
})

test_that("unlock_wb backend selection works", {
  # Test explicit backend selection errors
  file_dir_df <- data.frame(
    file_path = "test.xlsx",
    file_name = "test"
  )

  if (!requireNamespace("rpxl", quietly = TRUE)) {
    expect_error(
      unlock_wb(file_dir_df, wb_password = "test", backend = "rpxl"),
      "rpxl backend requested but not installed"
    )
  }

  if (!requireNamespace("XLConnect", quietly = TRUE)) {
    expect_error(
      unlock_wb(file_dir_df, wb_password = "test", backend = "xlconnect"),
      "XLConnect backend requested but not installed"
    )
  }
})

# Integration test - only run manually with a real password-protected file
test_that("unlock_wb works on password-protected file", {
  skip("Skipping integration test - requires password-protected test file")

  test_file <- system.file("extdata", "locked_wb_test.xlsx", package = "excelsioR")
  skip_if(!file.exists(test_file), "Test file not found")

  # Create a temp copy to test on
  tmp_file <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp_file), add = TRUE)
  file.copy(test_file, tmp_file)

  file_dir_df <- data.frame(
    file_path = tmp_file,
    file_name = "test"
  )

  # This would need the actual password
 # unlock_wb(file_dir_df, wb_password = "123")
})
