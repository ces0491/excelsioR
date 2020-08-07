test_that("read_excel works as expected", {

  test_file <- "mtcars_test"
  src_dir <- system.file("extdata", package = "excelsioR")

  src_df <- data.frame(file_path = paste0(src_dir, "/", test_file, ".xlsx"), file_name = test_file)

  test_tbl <- read_excel(src_df, "Sheet1")

  test_list <- test_tbl$raw_excel_data[[1]]

  test_df <- test_list[[1]]

})
