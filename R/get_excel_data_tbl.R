#' Convert raw Excel data to a tidy tibble
#'
#' Processes raw Excel cell data extracted by tidyxl and converts it to a tidy
#' tibble format with separate columns for numeric, character, and date values.
#'
#' @param raw_excel_data list of data.frames containing raw Excel data,
#'   where each list element contains data from one worksheet
#'
#' @return tbl_df with columns:
#'   \itemize{
#'     \item `sheet`: worksheet name
#'     \item `row`: row number
#'     \item `col`: column number
#'     \item `numeric`: numeric value (if present)
#'     \item `character`: character value from the leftmost column of each row
#'     \item `date`: date value (if present)
#'   }
#'
#' @section Assumptions:
#' This function makes the following assumptions about the spreadsheet structure:
#' \itemize{
#'   \item The leftmost non-blank column in each row contains row labels or variable names
#'   \item Dates are stored in a dedicated date column and will be matched by column position
#'   \item Blank cells are excluded from the output
#' }
#'
#' These assumptions work well for financial/time-series data where the first column
#' contains variable names and subsequent columns contain values. For other layouts,
#' you may need to post-process the output.
#'
get_excel_data_tbl_single <- function(raw_excel_data) {

  assert_true(is.list(raw_excel_data), "raw_excel_data must be a list")

  raw_excel_df <- raw_excel_data %>%
    tibble::enframe() %>%
    tidyr::unnest(value) %>%
    dplyr::select(-name)

  assert_present(
    names(raw_excel_df),
    c("sheet", "row", "col", "is_blank", "numeric", "date", "character")
  )

  # Extract numeric values (excluding blank cells)
  num_df <- raw_excel_df %>%
    dplyr::filter(!is_blank) %>%
    dplyr::select(sheet, row, col, numeric) %>%
    tidyr::drop_na()

  # Extract character values from the leftmost column of each row
  # Assumption: leftmost column contains variable names/row labels
  char_df <- raw_excel_df %>%
    dplyr::filter(!is_blank) %>%
    dplyr::mutate(character = trimws(character, "both")) %>%
    dplyr::group_by(row) %>%
    dplyr::filter(col == min(col)) %>%
    dplyr::ungroup() %>%
    dplyr::select(sheet, row, character) %>%
    tidyr::drop_na()

  # Extract date values (convert POSIXct to Date)
  date_df <- raw_excel_df %>%
    dplyr::filter(!is_blank) %>%
    dplyr::select(sheet, col, date) %>%
    dplyr::mutate(date = as.Date(date)) %>%
    tidyr::drop_na() %>%
    dplyr::distinct()

  # Join the data frames
  if (nrow(date_df) == 0) {
    # No dates detected - create empty date column
    tidy_df <- num_df %>%
      dplyr::left_join(char_df, by = c("sheet", "row")) %>%
      tidyr::drop_na() %>%
      dplyr::mutate(date = as.Date(NA))
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

#' Process multiple workbooks to tidy tibbles
#'
#' Takes a tibble containing raw Excel data from multiple files and converts
#' each file's data to a tidy format.
#'
#' @param raw_excel_data_tbl tbl_df with columns `file_name` and `raw_excel_data`
#'   (a nested list of raw Excel data per file)
#'
#' @return tbl_df with columns:
#'   \itemize{
#'     \item `file_name`: name of the source file
#'     \item `raw_excel_data`: original raw data (nested list)
#'     \item `all_excel_data_tbl`: processed tidy data (nested tibble)
#'   }
#'
#' @seealso [get_excel_data_tbl_single()] for the assumptions made during processing
#'
get_excel_data_tbl <- function(raw_excel_data_tbl) {

  assert_present(names(raw_excel_data_tbl), c("file_name", "raw_excel_data"))

  tidy_excel_data_tbl <- raw_excel_data_tbl %>%
    dplyr::group_by(file_name) %>%
    dplyr::mutate(all_excel_data_tbl = purrr::map(raw_excel_data, get_excel_data_tbl_single)) %>%
    dplyr::ungroup()

  tidy_excel_data_tbl
}
