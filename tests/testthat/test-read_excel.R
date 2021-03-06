test_that("read_excel works as expected", {

  test_file <- "mtcars_test.xlsx"
  expected_file <- "mtcars_expected.rds"
  src_dir <- system.file("testdata", package = "excelsioR")

  src_df <- data.frame(file_path = paste0(src_dir, "/", test_file), file_name = test_file)

  test_tbl <- read_excel(src_df, "Sheet1")

  test_list <- test_tbl$raw_excel_data[[1]]

  test_df <- test_list[[1]]

  expected_df <- readRDS(paste0(src_dir, "/", expected_file))

  expect_equal(test_df, expected_df)

})
