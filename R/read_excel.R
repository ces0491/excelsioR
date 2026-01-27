#' Read a single Excel workbook
#'
#' Reads raw cell data from an Excel workbook using tidyxl.
#'
#' @param file_path string indicating the file path to the workbook
#' @param reqd_sheets the sheets to read from the workbook. Use NA to read all sheets.
#'
#' @return a named list containing raw Excel data in a tbl_df, indexed by sheet name
#'
read_excel_single <- function(file_path, reqd_sheets) {

  # Check if file is open by another process
  if (is_file_open(file_path)) {
    stop(
      glue::glue("{file_path} appears to be open by another process. Please close it before proceeding."),
      call. = FALSE
    )
  }

  assert_true(length(file_path) == 1, "file_path must be a single path")
  assert_true(file.exists(file_path), paste0("File doesn't exist: ", file_path))

  # If reqd_sheets is NA, read all available sheets in the workbook
  if (any(is.na(reqd_sheets))) {
    raw_excel <- suppressWarnings(tidyxl::xlsx_cells(file_path, sheets = reqd_sheets))
    assert_present(names(raw_excel), "sheet")

    grouped_excel <- raw_excel %>%
      dplyr::group_by(sheet)

    wsheet_list <- grouped_excel %>%
      dplyr::group_split()

    names(wsheet_list) <- dplyr::group_keys(grouped_excel)[[1]]

  } else {
    wsheet_list <- list()
    for (wsheet in reqd_sheets) {
      raw_excel <- suppressWarnings(tidyxl::xlsx_cells(file_path, sheets = wsheet))
      wsheet_list[[wsheet]] <- raw_excel
    }
  }

  wsheet_list
}

#' Read Excel workbooks into R
#'
#' Reads multiple Excel workbooks and returns raw cell data in a tidy format.
#'
#' @param file_names_df data.frame containing columns `file_path` and `file_name`
#' @param reqd_sheets string vector with the names of the worksheets to read. Use NA for all sheets.
#'
#' @return tbl_df with columns `file_name` and `raw_excel_data` (a nested named list indexed by sheet name)
#'
read_excel <- function(file_names_df, reqd_sheets) {

  assert_present(names(file_names_df), c("file_path", "file_name"))

  # Use purrr::map instead of deprecated dplyr::do
  raw_data <- file_names_df %>%
    dplyr::group_by(file_name) %>%
    dplyr::summarise(
      raw_excel_data = list(read_excel_single(file_path[[1]], reqd_sheets)),
      .groups = "drop"
    )

  raw_data
}
