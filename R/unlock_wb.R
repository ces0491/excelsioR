#' Unlock a protected Excel workbook
#'
#' Removes password protection from Excel workbooks. Supports two backends:
#' \itemize{
#'   \item \strong{rpxl} (recommended): Uses Python's msoffcrypto-tool via reticulate. Lighter weight.
#'   \item \strong{XLConnect}: Uses Java. Heavier but doesn't require Python.
#' }
#'
#' The function automatically selects the available backend, preferring rpxl.
#'
#' @param file_dir data.frame with columns `file_path` and `file_name` for files to unlock
#' @param wb_password optional string argument for the workbook password.
#'   If NULL, will prompt via RStudio API (recommended for security).
#' @param backend character string specifying which backend to use: "auto" (default),
#'   "rpxl", or "xlconnect". With "auto", rpxl is preferred if available.
#' @param options.java.params optional string argument specifying Java parameters
#'   (e.g., "-Xmx4g"). Only used with XLConnect backend.
#'
#' @return NULL (invisibly). Files are decrypted in place.
#'
#' @section Backend Installation:
#' \strong{rpxl (recommended):}
#' \preformatted{
#' remotes::install_github("epicentre-msf/rpxl")
#' }
#' Requires Python with msoffcrypto-tool. rpxl will guide you through setup.
#'
#' \strong{XLConnect:}
#' \preformatted{
#' install.packages("XLConnect")
#' }
#' Requires Java to be installed on your system.
#'
#' @examples
#' # Example requires rpxl or XLConnect and password-protected files
#' \dontrun{
#' # Prepare file info
#' files_to_unlock <- data.frame(
#'   file_path = c("protected1.xlsx", "protected2.xlsx"),
#'   file_name = c("protected1", "protected2")
#' )
#'
#' # Unlock - will prompt for password in RStudio
#' unlock_wb(files_to_unlock)
#'
#' # Or provide password directly
#' unlock_wb(files_to_unlock, wb_password = "secret123")
#'
#' # Force specific backend
#' unlock_wb(files_to_unlock, wb_password = "secret", backend = "rpxl")
#' }
#'
#' @export
#'
unlock_wb <- function(file_dir, wb_password = NULL, backend = "auto",
                      options.java.params = NULL) {

 assert_present(names(file_dir), c("file_path", "file_name"))

  # Determine which backend to use
  backend <- match.arg(backend, c("auto", "rpxl", "xlconnect"))

  has_rpxl <- requireNamespace("rpxl", quietly = TRUE)
  has_xlconnect <- requireNamespace("XLConnect", quietly = TRUE)

  if (backend == "auto") {
    if (has_rpxl) {
      backend <- "rpxl"
    } else if (has_xlconnect) {
      backend <- "xlconnect"
    } else {
      stop(
        "No backend available for unlocking password-protected files.\n",
        "Install one of:\n",
        "  - rpxl (recommended): remotes::install_github('epicentre-msf/rpxl')\n",
        "  - XLConnect: install.packages('XLConnect') (requires Java)",
        call. = FALSE
      )
    }
  } else if (backend == "rpxl" && !has_rpxl) {
    stop(
      "rpxl backend requested but not installed.\n",
      "Install with: remotes::install_github('epicentre-msf/rpxl')",
      call. = FALSE
    )
  } else if (backend == "xlconnect" && !has_xlconnect) {
    stop(
      "XLConnect backend requested but not installed.\n",
      "Install with: install.packages('XLConnect')\n",
      "Note: XLConnect requires Java.",
      call. = FALSE
    )
  }

  # Get password securely
  if (is.null(wb_password)) {
    if (rstudioapi::isAvailable()) {
      wb_password <- rstudioapi::askForPassword("Enter password to unlock workbook")
    } else {
      stop(
        "Password required. Either:\n",
        "  1. Run in RStudio to use the secure password prompt, or\n",
        "  2. Provide password via the `wb_password` parameter",
        call. = FALSE
      )
    }
  }

  if (is.null(wb_password) || nchar(wb_password) == 0) {
    stop("Password cannot be empty", call. = FALSE)
  }

  # Dispatch to appropriate backend
  if (backend == "rpxl") {
    unlock_with_rpxl(file_dir, wb_password)
  } else {
    unlock_with_xlconnect(file_dir, wb_password, options.java.params)
  }

  invisible(NULL)
}

#' Unlock files using rpxl (Python/msoffcrypto backend)
#' @keywords internal
#' @noRd
unlock_with_rpxl <- function(file_dir, wb_password) {

  n_files <- nrow(file_dir)
  message(sprintf("Unlocking %d file(s) using rpxl (Python)...", n_files))

  for (n in seq_len(n_files)) {
    reqd_file <- file_dir[n, ]

    assert_true(
      file.exists(reqd_file$file_path),
      paste0("File does not exist: ", reqd_file$file_path)
    )

    # rpxl::decrypt_xlsx returns path to decrypted file
    # We decrypt to a temp file, then replace the original
    decrypted_path <- rpxl::decrypt_xlsx(
      path = reqd_file$file_path,
      password = wb_password
    )

    # Replace original with decrypted version
    file.copy(decrypted_path, reqd_file$file_path, overwrite = TRUE)
    unlink(decrypted_path)

    progress <- round(n / n_files * 100)
    message(sprintf("[%d/%d] %s unlocked (%d%% complete)",
                    n, n_files, reqd_file$file_name, progress))
  }

  message("All files unlocked successfully (rpxl)")
}

#' Unlock files using XLConnect (Java backend)
#' @keywords internal
#' @noRd
unlock_with_xlconnect <- function(file_dir, wb_password, options.java.params) {

  # Set Java parameters if provided
  if (!is.null(options.java.params)) {
    options(java.parameters = options.java.params)
  }

  n_files <- nrow(file_dir)
  message(sprintf("Unlocking %d file(s) using XLConnect (Java)...", n_files))

  for (n in seq_len(n_files)) {
    reqd_file <- file_dir[n, ]

    assert_true(
      file.exists(reqd_file$file_path),
      paste0("File does not exist: ", reqd_file$file_path)
    )

    wb <- XLConnect::loadWorkbook(reqd_file$file_path, password = wb_password)
    XLConnect::saveWorkbook(wb)

    progress <- round(n / n_files * 100)
    message(sprintf("[%d/%d] %s unlocked (%d%% complete)",
                    n, n_files, reqd_file$file_name, progress))

    rm(wb)
    XLConnect::xlcFreeMemory()
  }

  message("All files unlocked successfully (XLConnect)")
}
