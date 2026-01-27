#' Get Excel file names from a directory
#'
#' Scans a directory for Excel files and returns a data frame with file paths
#' and cleaned file names suitable for use as R variable names.
#'
#' @param data_dir string indicating the directory to scan
#' @param recursive logical indicating whether to search subdirectories. Default is FALSE.
#' @param include_dirs logical indicating whether directory names should be included. Default is FALSE.
#' @param extensions character vector of file extensions to include (without the dot).
#'   Default is `c("xlsx", "xlsm", "xls")`.
#' @param clean_names logical indicating whether to clean file names (remove leading numbers,
#'   replace non-alphanumeric characters). Default is TRUE.
#'
#' @return data.frame with columns:
#'   \itemize{
#'     \item `file_path`: full path to the file
#'     \item `file_name`: cleaned file name without extension
#'   }
#'
#' @examples
#' \dontrun{
#' # Get all Excel files in a directory
#' files <- get_file_names("path/to/excel/files")
#' files
#' #>                         file_path    file_name
#' #> 1 path/to/excel/files/sales.xlsx        sales
#' #> 2 path/to/excel/files/01 Data.xlsx       Data
#'
#' # Search subdirectories
#' files <- get_file_names("path/to/files", recursive = TRUE)
#'
#' # Only xlsx files (exclude xls, xlsm)
#' files <- get_file_names("path/to/files", extensions = "xlsx")
#'
#' # Keep original file names without cleaning
#' files <- get_file_names("path/to/files", clean_names = FALSE)
#' }
#'
#' @export
#'
get_file_names <- function(data_dir, recursive = FALSE, include_dirs = FALSE,
                            extensions = c("xlsx", "xlsm", "xls"),
                            clean_names = TRUE) {

  if (!dir.exists(data_dir)) {
    stop("Directory does not exist: ", data_dir, call. = FALSE)
  }

  # Build regex pattern for extensions
  ext_pattern <- paste0("\\.(", paste(extensions, collapse = "|"), ")$")

  # Get all files
  all_files <- list.files(data_dir, full.names = TRUE, recursive = recursive,
                          include.dirs = include_dirs)

  # Filter to Excel files only
  file_names_df <- data.frame(file_path = all_files, stringsAsFactors = FALSE) %>%
    dplyr::filter(stringr::str_detect(tolower(file_path), ext_pattern)) %>%
    dplyr::mutate(
      file_name = tools::file_path_sans_ext(basename(file_path))
    )

  if (clean_names && nrow(file_names_df) > 0) {
    file_names_df <- file_names_df %>%
      dplyr::mutate(
        file_name = trimws(file_name, "both"),
        file_name = gsub("^\\d+", "", file_name),
        file_name = gsub("[^[:alnum:]]", "_", file_name),
        file_name = gsub("^_+|_+$", "", file_name),
        file_name = gsub("_+", "_", file_name)
      )
  }

  if (nrow(file_names_df) == 0) {
    message("No Excel files found in: ", data_dir)
  }

  file_names_df
}
