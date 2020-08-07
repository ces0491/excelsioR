test_that("write_wb works", {

  dest_base <- system.file("extdata", package = "excelsioR")
  dest_path <- paste0(dest_base, "/mtcars_test.xlsx")

  sample_data <- mtcars %>%
    dplyr::mutate(model = row.names(.)) %>%
    dplyr::relocate(model, before = mpg)

  write_wb(sample_data, sheet_name = "Sheet1", wb_dir = dest_path, save_wb = TRUE)

  test <- openxlsx::read.xlsx(dest_path)

  expected <- sample_data

  testthat::expect_equal(test, expected)
})
