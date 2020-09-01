#' unlock a protected excel workbook
#'
#' @param file_dir \code{data.frame} of file paths and names for files you wish to unlock
#' @param options.java.params optional string argument specifying java parameters
#'
#' @return NULL
#' @export
#'
unlock_wb <- function(file_dir, options.java.params = NULL) {

  options(java.parameters = options.java.params)

  assertR::assert_present(names(file_dir), c("file_path", "file_name"))

  wb_password <- rstudioapi::askForPassword("Please enter password to unlock workbook") # enter password via rstudio api rather than a string for security

  # you may unlock multiple files in a directory provided they all have the same password
  for(n in 1:nrow(file_dir)) {

    reqd_file <- file_dir[n, ]

    file.exists <- base::file.exists(reqd_file$file_path)
    assertR::assert_true(file.exists, "the file you're trying to unlock doesn't exist")

    p <- which(file_dir$file_path == reqd_file$file_path)
    progress <- round(p/length(file_dir$file_path), 2) * 100
    print(glue::glue("{reqd_file$file_name} unlocked - {progress}% complete"))

    wb <- XLConnect::loadWorkbook(reqd_file$file_path, password = wb_password)
    XLConnect::saveWorkbook(wb)

    rm(wb)
    XLConnect::xlcFreeMemory()  # free Java Virtual Machine memory
  }
}
