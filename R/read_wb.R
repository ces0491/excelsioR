#' Read data from Excel workbooks
#'
#' Reads data from Excel workbooks in a directory and returns it in a tidy format.
#' For safety, data is read from a timestamped copy of the workbooks rather than
#' the originals to prevent potential file corruption.
#'
#' @param source_data_dir string indicating the source directory containing Excel files
#' @param reqd_wkbks character vector of workbook names to extract (without extension).
#'   Default NA extracts all workbooks in the directory.
#' @param reqd_sheets character vector of sheet names to extract.
#'   Default NA extracts all sheets from each workbook.
#' @param password_protected logical indicating whether the Excel files are password protected.
#'   If TRUE, you will be prompted for the password. All files must share the same password.
#' @param overwrite logical indicating whether to overwrite existing files in the destination
#'   directory. FALSE preserves existing files and skips duplicates.
#' @param dest_data_dir optional string indicating destination directory for file copies.
#'   Default NULL uses a temp directory that is cleared when your R session ends.
#'
#' @return tbl_df with columns:
#'   \itemize{
#'     \item `file_name`: cleaned name of the source file
#'     \item `raw_excel_data`: nested list of raw cell data per sheet (from tidyxl)
#'     \item `all_excel_data_tbl`: nested tibble of processed tidy data
#'   }
#'
#' @section Workflow:
#' 1. Scans source directory for Excel files
#' 2. Creates timestamped copy of files in destination directory
#' 3. Unlocks password-protected files if needed
#' 4. Reads raw cell data using tidyxl
#' 5. Processes data into tidy format
#'
#' @seealso [get_excel_data_tbl()] for details on how data is tidied
#'
#' @examples
#' \donttest{
#' # Create test files in a temp directory
#' test_dir <- file.path(tempdir(), "read_wb_example")
#' dir.create(test_dir, showWarnings = FALSE)
#' openxlsx::write.xlsx(mtcars[1:5, ], file.path(test_dir, "cars.xlsx"))
#' openxlsx::write.xlsx(iris[1:5, ], file.path(test_dir, "flowers.xlsx"))
#'
#' # Read all Excel files from the directory
#' data <- read_wb(test_dir)
#' data
#'
#' # Read a specific workbook
#' data <- read_wb(test_dir, reqd_wkbks = "cars")
#'
#' # Access raw tidyxl data for complex layouts
#' raw <- data$raw_excel_data[[1]][[1]]
#' head(raw)
#' }
#'
#' @export
#'
read_wb <- function(source_data_dir, reqd_wkbks = NA, reqd_sheets = NA,
                    password_protected = FALSE, overwrite = TRUE, dest_data_dir = NULL) {

  # Validate source directory
  if (!dir.exists(source_data_dir)) {
    stop("Source directory does not exist: ", source_data_dir, call. = FALSE)
  }

  # Get file paths and names for the source data
  original_file_dir <- get_file_names(source_data_dir)

  if (nrow(original_file_dir) == 0) {
    stop("No Excel files found in: ", source_data_dir, call. = FALSE)
  }

  # Filter to specific workbooks if requested
  # Note: file_name contains no leading numerics, special characters, or spaces
  if (any(!is.na(reqd_wkbks))) {
    original_file_dir <- dplyr::filter(original_file_dir, file_name %in% reqd_wkbks)

    if (nrow(original_file_dir) == 0) {
      stop("None of the requested workbooks found: ",
           paste(reqd_wkbks, collapse = ", "), call. = FALSE)
    }
  }

  message(sprintf("Found %d workbook(s) to process", nrow(original_file_dir)))

  # Create temp folder or use user-specified destination
  if (is.null(dest_data_dir)) {
    dest_data_dir <- file.path(tempdir(), "read_wb_data")
    message(
      "No dest_data_dir specified. Source data will be copied to:\n",
      "  ", dest_data_dir, "\n",
      "This copied data will be cleared when your session restarts."
    )
  }

  # Create a timestamped folder in the destination directory
  dest_dir <- file.path(dest_data_dir,
                         paste0("source_data_copy_", format(Sys.Date(), "%Y%m%d")))

  # Handle existing destination directory
  if (dir.exists(dest_dir) && overwrite) {
    message("Destination directory exists. Clearing contents...")
    existing_files <- list.files(dest_dir, full.names = TRUE)
    if (length(existing_files) > 0) {
      file.remove(existing_files)
    }
  } else if (dir.exists(dest_dir) && !overwrite) {
    message("Destination directory exists. Duplicate files will not be copied.")
  }

  # Copy files to destination (creates directory if needed)
  copy_files(original_file_dir, dest_dir, overwrite)

  # Get file names from the copied directory
  copied_files_dir <- get_file_names(dest_dir)

  # Unlock password-protected workbooks if needed
  if (password_protected) {
    unlock_wb(copied_files_dir)
  }

  # Read raw Excel data
  message("Reading Excel data...")
  raw_excel_data_tbl <- read_excel(copied_files_dir, reqd_sheets)

  # Clean raw data to return a tidy tibble
  message("Processing data...")
  excel_data <- get_excel_data_tbl(raw_excel_data_tbl)

  message("Done!")
  excel_data
}
