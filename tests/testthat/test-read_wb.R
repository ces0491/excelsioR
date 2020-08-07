test_that("read_wb works", {

  test_file <- "mtcars_test.xlsx"
  src_dir <- system.file("extdata", package = "excelsioR")

  file_path = paste0(src_dir, "/", test_file)

  test <- read_wb(src_dir)

})
