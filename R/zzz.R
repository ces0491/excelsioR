# Package load and unload hooks

.onLoad <- function(libname, pkgname) {
  # Optional dependency checks are done at function call time, not load time
  invisible()
}

.onAttach <- function(libname, pkgname) {
  # Check for optional dependencies and inform user
  if (!requireNamespace("XLConnect", quietly = TRUE)) {
    packageStartupMessage(
      "Note: XLConnect is not installed. ",
      "Password-protected workbook features (unlock_wb) will be unavailable.\n",
      "Install with: install.packages('XLConnect') (requires Java)"
    )
  }

  if (!requireNamespace("zoo", quietly = TRUE)) {
    packageStartupMessage(
      "Note: zoo is not installed. ",
      "Time series export (write_wb.zoo) will be unavailable.\n",
      "Install with: install.packages('zoo')"
    )
  }
}
