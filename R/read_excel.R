#' read a single excel workbook
#'
#' @param file_path string indicating the file path to the workbook of interest
#' @param reqd_sheets the sheets you'd like to copy from the workbook
#'
#' @return a named list containing raw excel data in a tbl_df where the name equals the reqd_sheet
#'
read_excel_single <- function(file_path, reqd_sheets) {

  is_file_open <- function(file_name) {
    con <- file(description = file_name)
    isOpen(con)
  }

  if(is_file_open(file_path)) {
    message(glue::glue("{file_path} is open, please close before proceding"))
  }

  assertR::assert_true(length(file_path) == 1, "logic error")
  assertR::assert_true(file.exists(file_path), paste0("File doesn't exist:", file_path))

  # if reqd sheets is NA, all available sheets in the workbook are read in
  if (any(is.na(reqd_sheets))) {
    raw_excel <- suppressWarnings(tidyxl::xlsx_cells(file_path, sheets = reqd_sheets))
    assertR::assert_present(names(raw_excel), "sheet")

    wsheet_list <- raw_excel %>%
      dplyr::group_split(sheet)

    names(wsheet_list) <- dplyr::group_keys(raw_excel, sheet)[[1]]

  } else {

    wsheet_list <- list()
    for (wsheet in reqd_sheets) {
      raw_excel <- suppressWarnings(tidyxl::xlsx_cells(file_path, sheets = wsheet))
      wsheet_list[[wsheet]] <- raw_excel
    }
  }

  wsheet_list
}

#' read excel workbooks in to R
#'
#' @param file_names_df \code{data.frame} containing the directories of files you'd like to read
#' @param reqd_sheets string vector with the names of the worksheets to read in
#'
#' @return \code{tbl_df} with file_name and raw_excel_data as a nested named list where the list is indexed by the sheet name specified in reqd_sheets
#'
read_excel <- function(file_names_df, reqd_sheets) {

  assertR::assert_present(names(file_names_df), c("file_path", "file_name"))

  raw_data <- file_names_df %>%
    dplyr::group_by(file_name) %>%
    dplyr::do(raw_excel_data = read_excel_single(.$file_path, reqd_sheets)) %>% # get list of wsheets from each wbook
    dplyr::ungroup()

  raw_data

}
