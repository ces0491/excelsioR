# Global variables used in NSE (dplyr/tidyr pipelines) to avoid R CMD check NOTEs
utils::globalVariables(c(
  "file_name", "file_path", "raw_excel_data", "value", "name",
  "is_blank", "sheet", ".data"
))

#' Pipe operator
#'
#' @name %>%
#' @rdname pipe
#' @keywords internal
#' @importFrom dplyr %>%
#' @export
#' @usage lhs \%>\% rhs
NULL

#' Convert Excel column letters to numeric index
#'
#' Supports multi-letter columns (A=1, Z=26, AA=27, AB=28, etc.)
#'
#' @param col_letter Character string representing Excel column (e.g., "A", "AA", "XFD")
#' @return Integer representing the column number
#'
#' @keywords internal
#'
col_to_num <- function(col_letter) {
  col_letter <- toupper(trimws(col_letter))

  if (!grepl("^[A-Z]+$", col_letter)) {
    stop("Column must contain only letters A-Z: ", col_letter, call. = FALSE)
  }

  letters_vec <- strsplit(col_letter, "")[[1]]
  result <- sum(vapply(seq_along(letters_vec), function(i) {
    (match(letters_vec[i], LETTERS)) * (26L ^ (length(letters_vec) - i))
  }, FUN.VALUE = numeric(1)))

  as.integer(result)
}

#' Convert numeric index to Excel column letters
#'
#' @param num Integer representing column number (1-16384)
#' @return Character string representing Excel column
#'
#' @keywords internal
#'
num_to_col <- function(num) {
  if (!is.numeric(num) || length(num) != 1 || num < 1) {
    stop("Column number must be a positive integer", call. = FALSE)
  }

  num <- as.integer(num)
  result <- ""

  while (num > 0) {
    num <- num - 1L
    result <- paste0(LETTERS[(num %% 26L) + 1L], result)
    num <- num %/% 26L
  }

  result
}

#' Parse Excel cell reference into column and row components
#'
#' @param cell_ref Character string like "A1", "AA100", "XFD1048576"
#' @return Named list with `col` (integer) and `row` (integer)
#'
#' @keywords internal
#'
parse_cell_reference <- function(cell_ref) {
  cell_ref <- toupper(trimws(cell_ref))

  col_part <- gsub("[0-9]", "", cell_ref)
  row_part <- gsub("[A-Z]", "", cell_ref)

  if (nchar(col_part) == 0) {
    stop("Invalid cell reference (missing column): ", cell_ref, call. = FALSE)
  }

  if (nchar(row_part) == 0) {
    stop("Invalid cell reference (missing row): ", cell_ref, call. = FALSE)
  }

  row_num <- suppressWarnings(as.integer(row_part))
  if (is.na(row_num) || row_num < 1) {
    stop("Invalid cell reference (invalid row): ", cell_ref, call. = FALSE)
  }

  list(
    col = col_to_num(col_part),
    row = row_num
  )
}

#' Get dimensions of data in a worksheet
#'
#' Reads the worksheet to determine actual data extent.
#'
#' @param wb Workbook object
#' @param sheet Sheet name or index
#' @return Named list with max_row and max_col
#'
#' @keywords internal
#'
get_sheet_dimensions <- function(wb, sheet) {
  sheet_data <- tryCatch(
    suppressWarnings(
      openxlsx::readWorkbook(wb, sheet = sheet, colNames = FALSE, skipEmptyRows = FALSE)
    ),
    error = function(e) NULL
  )

  if (is.null(sheet_data) || nrow(sheet_data) == 0) {
    return(list(max_row = 0L, max_col = 0L))
  }

  list(
    max_row = nrow(sheet_data),
    max_col = ncol(sheet_data)
  )
}

#' Validate workbook path
#'
#' @param wb_dir Path to workbook file
#' @param must_exist Logical, whether file must already exist
#' @return Normalized path (invisibly)
#'
#' @keywords internal
#'
validate_workbook_path <- function(wb_dir, must_exist = FALSE) {
  if (is.null(wb_dir)) {
    stop("`wb_dir` cannot be NULL", call. = FALSE)
  }

  if (!is.character(wb_dir) || length(wb_dir) != 1) {
    stop("`wb_dir` must be a single character string", call. = FALSE)
  }

  if (must_exist && !file.exists(wb_dir)) {
    stop("File does not exist: ", wb_dir, call. = FALSE)
  }

  if (!grepl("\\.(xlsx|xlsm|xls)$", wb_dir, ignore.case = TRUE)) {
    stop("`wb_dir` must have an Excel file extension (.xlsx, .xlsm, or .xls)", call. = FALSE)
  }

  invisible(normalizePath(wb_dir, mustWork = FALSE))
}

#' Check if a file is currently open by another process
#'
#' @param file_path Path to file to check
#' @return Logical, TRUE if file is open/locked
#'
#' @keywords internal
#'
is_file_open <- function(file_path) {
  if (!file.exists(file_path)) {
    return(FALSE)
  }

  tryCatch({
    con <- file(file_path, open = "r+b")
    close(con)
    FALSE
  }, error = function(e) {
    TRUE
  })
}

# ============================================================================
# Assertion helpers (previously from assertR)
# ============================================================================

#' Assert that an expression is TRUE
#'
#' @param expr Expression to evaluate
#' @param msg Error message if assertion fails
#' @return NULL invisibly; stops with error if assertion fails
#'
#' @keywords internal
#'
assert_true <- function(expr, msg = "Assertion failed") {
  if (is.null(expr) || length(expr) == 0 || is.na(expr[1]) || !isTRUE(as.logical(expr[1]))) {
    stop(msg, call. = FALSE)
  }
  invisible(NULL)
}

#' Assert that required values are present in a vector
#'
#' @param available Vector of available values
#' @param required Vector of required values that must be present
#' @return NULL invisibly; stops with error if any required values are missing
#'
#' @keywords internal
#'
assert_present <- function(available, required) {
  missing <- required[!required %in% available]
  if (length(missing) > 0) {
    stop(
      "Missing required elements: ", paste(missing, collapse = ", "), "\n",
      "Available: ", paste(available, collapse = ", "),
      call. = FALSE
    )
  }
  invisible(NULL)
}

# ============================================================================
# File helpers (previously from fileR)
# ============================================================================

#' Copy files to a destination directory
#'
#' @param source_files data.frame with columns `file_path` and `file_name`
#' @param dest_dir Destination directory (created if it doesn't exist)
#' @param overwrite Logical, whether to overwrite existing files
#' @param timestamp Logical, whether to add timestamp prefix to copied files
#' @return NULL invisibly
#'
#' @keywords internal
#'
copy_files <- function(source_files, dest_dir, overwrite = FALSE, timestamp = FALSE) {
  assert_present(names(source_files), c("file_path", "file_name"))

  # Create destination directory if needed
 if (!dir.exists(dest_dir)) {
    dir.create(dest_dir, recursive = TRUE)
  }

  n_files <- nrow(source_files)

  for (i in seq_len(n_files)) {
    src_path <- source_files$file_path[i]
    file_name <- basename(src_path)

    if (!file.exists(src_path)) {
      warning("Source file does not exist: ", src_path)
      next
    }

    # Add timestamp prefix if requested
    if (timestamp) {
      file_name <- paste0(format(Sys.time(), "%H%M%S"), "_", file_name)
    }

    dest_path <- file.path(dest_dir, file_name)

    # Skip if file exists and overwrite is FALSE
    if (file.exists(dest_path) && !overwrite) {
      message("Skipping (exists): ", file_name)
      next
    }

    file.copy(src_path, dest_path, overwrite = overwrite, copy.date = TRUE)

    progress <- round(i / n_files * 100)
    message(sprintf("[%d%%] Copied: %s", progress, file_name))
  }

  invisible(NULL)
}
