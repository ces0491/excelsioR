#' convert list of tbl_dfs extracted from excel to a tibble containing all extracted data
#'
#' @param raw_excel_data \code{list} of \code{data.frame}'s containing raw excel data where each list object contains data from a sheet required from the workbook
#'
#' @return \code{tbl_df} with columns sheet, row, col, numeric, character and date
#'
get_excel_data_tbl_single <- function(raw_excel_data) {

  assertR::assert_true(is.list(raw_excel_data), "logic error")

  raw_excel_df <- raw_excel_data %>%
    tibble::enframe() %>%
    tidyr::unnest(value) %>%
    dplyr::select(-name)

  assertR::assert_present(names(raw_excel_df), c("sheet", "row", "col", "is_blank", "numeric", "date", "character"))

  num_df <- raw_excel_df %>%
    dplyr::filter(!is_blank) %>%
    dplyr::select(sheet, row, col, numeric) %>%
    tidyr::drop_na()

  char_df <- raw_excel_df %>%
    dplyr::filter(!is_blank) %>%
    dplyr::mutate(character = trimws(character, "both")) %>%
    dplyr::group_by(row) %>%
    dplyr::filter(col == min(col)) %>% # we assume that the left most column contains the characters of interest i.e. variable names
    dplyr::select(sheet, row, character) %>%
    tidyr::drop_na()

  date_df <- raw_excel_df %>%
    dplyr::filter(!is_blank) %>%
    dplyr::select(sheet, col, date) %>%
    dplyr::mutate(date = as.Date(date)) %>% # tidyxl reads dates as POSIXct so convert to Date
    tidyr::drop_na() %>%
    dplyr::distinct()

  if(is.data.frame(date_df) && nrow(date_df) == 0) {
    tidy_df <- num_df %>%
      dplyr::left_join(char_df, by = c("sheet", "row")) %>%
      tidyr::drop_na() %>%
      dplyr::mutate(date = as.Date(NA)) # if no dates are detected, create a date column of NAs

  } else {

    tidy_df <- num_df %>%
      dplyr::left_join(char_df, by = c("sheet", "row")) %>%
      dplyr::left_join(date_df, by = c("sheet", "col")) %>%
      tidyr::drop_na()
  }

  all_excel_data_tbl <- tidy_df %>%
    dplyr::distinct()

  all_excel_data_tbl
}

#' get single tibbles per file for all required data given lists of tbl_dfs extracted from excel
#'
#' @param raw_excel_data_tbl tbl with file_name and nested list of raw excel data
#'
#' @return tbl_df with col file_name, nested raw excel data containing named lists of tibbles per sheet, nested tibbles containing combined data
#'
get_excel_data_tbl <- function(raw_excel_data_tbl) {

  assertR::assert_present(names(raw_excel_data_tbl), c("file_name", "raw_excel_data"))

  tidy_excel_data_tbl <- raw_excel_data_tbl %>%
    dplyr::group_by(file_name) %>%
    dplyr::mutate(all_excel_data_tbl = purrr::map(raw_excel_data, get_excel_data_tbl_single)) %>%
    dplyr::ungroup()

  tidy_excel_data_tbl
}
