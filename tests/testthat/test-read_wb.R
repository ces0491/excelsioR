test_that("read_wb works", {

  src_dir <- system.file("extdata", package = "excelsioR")
  test <- read_wb(src_dir)

  exp_dir <- system.file("testdata", package = "excelsioR")
  expected <- readRDS(paste0(exp_dir, "/mtcars_read_wb_test.rds"))

  expect_equal(test, expected)
})
