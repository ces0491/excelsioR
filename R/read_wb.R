#' read data from workbooks
#'
#' Data is read from a copy of the specified workbooks rather than the original to prevent potential file corruption of the original
#'
#' @param source_data_dir string indicating source directory
#' @param reqd_wkbks character vector indicating the names of the files you wish to extract from source_dir. default NA extracts all workbooks in the dir
#' @param reqd_sheets character vector containing names of sheets to extract data from. default NA extracts from all sheets in workbook
#' @param password_protected logical indicating whether the excel files are password protected. Password must be the same for multiple file reads
#' @param overwrite logical indicating whether files in the destination directory should be overwritten. FALSE results in no copy rather than duplicates
#' @param dest_data_dir optional string indicating destination directory - default NULL will use temp directory
#'
#' @return \code{tbl_df} with columns file_name, raw_excel_data and all_excel_data_tbl
#' @export
#'
read_wb <- function(source_data_dir, reqd_wkbks = NA, reqd_sheets = NA, password_protected = FALSE, overwrite = TRUE, dest_data_dir = NULL) {

  # get file path and names for the source data
  original_file_dir <- get_file_names(source_data_dir)

  # filter out the specific workbooks you require from the source dir. NB: file_name contains no leading numerics, special characters or spaces
  if(any(!is.na(reqd_wkbks))) {
    original_file_dir <- dplyr::filter(original_file_dir, file_name %in% reqd_wkbks)
  }

  # create temp folder to copy source data to else use user specified dest_data_dir
  if (is.null(dest_data_dir)) {
    dest_data_dir <- glue::glue("{tempdir()}/read_wb_data")
    message(glue::glue("no dest_data_dir specified. source data will be copied to {dest_data_dir}. this copied data will be cleared when your session restarts"))
  } else {
    dest_data_dir <- dest_data_dir
  }

  # create a timestamped folder in the dest_data_dir
  dest_dir <- glue::glue("{dest_data_dir}/source_data_copy_{format(Sys.Date(), '%Y%m%d')}")

  # if the destination directory exists and overwrite is TRUE, clear everything in it
  if (file.exists(dest_dir) && overwrite) {
    message("destination directory already exists. clearing contents")
    do.call(file.remove, list(list.files(dest_dir, full.names = TRUE)))
  }

  # here, duplicate files are not copied. only files that exist in source but not destination are copied
  if (file.exists(dest_dir) && !overwrite) {
    message("destination directory already exists. duplicate files already in the destination directory will not be copied from source")
  }

  # copy files to a separate directory so they can be used without risk of contaminating the original data
  # if the destination directory does not exist, it will be created
  fileR::copy_files(original_file_dir, dest_dir, overwrite)

  # get the file names from the directory containing the copied source files
  copied_files_dir <- get_file_names(dest_dir)

  # unlock protected wbs
  if (password_protected) {
    unlock_wb(copied_files_dir)
  }

  # read excel data
  raw_excel_data_tbl <- read_excel(copied_files_dir, reqd_sheets)

  # clean raw excel data to return a tidy tbl
  excel_data <- get_excel_data_tbl(raw_excel_data_tbl)

  excel_data

}
