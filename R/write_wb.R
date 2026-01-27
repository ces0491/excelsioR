#' Load a workbook object from file or use the current workbook if it exists
#'
#' @param wb optional workbook object
#' @param wb_dir optional path to a file to load; if file doesn't exist it will create a new workbook
#'
#' @return workbook object. See \link[openxlsx]{loadWorkbook}
#'
load_wb <- function(wb = NULL, wb_dir = NULL) {

  assert_true(
    (!is.null(wb)) || (!is.null(wb_dir)),
    "Must provide either a workbook object (wb) or a file path (wb_dir)"
  )

  if (is.null(wb)) {

    if (!is.null(wb_dir) && file.exists(wb_dir)) {
      wb <- suppressWarnings(openxlsx::loadWorkbook(wb_dir))
    } else {
      wb <- openxlsx::createWorkbook()
      message("New workbook created")
    }

  }

  wb
}

#' Write from R to Excel
#'
#' Generic function to write R objects to an Excel workbook. Supports data frames,
#' lists (multiple sheets), and zoo/xts time series objects.
#'
#' @param r_data an object to be written to an xlsx workbook
#' @param ... arguments for other methods
#'
#' @examples
#' \dontrun{
#' # Write a data frame to Excel
#' write_wb(mtcars, sheet_name = "cars", wb_dir = "output.xlsx", save_wb = TRUE)
#'
#' # Write a named list (each element becomes a sheet)
#' sheets <- list(cars = mtcars[1:5, ], flowers = iris[1:5, ])
#' write_wb(sheets, wb_dir = "multi_sheet.xlsx", save_wb = TRUE)
#' }
#'
#' @export
#'
write_wb <- function(r_data, ...) UseMethod("write_wb")

#' Default method for writing to Excel
#'
#' @param r_data an object to be written to an xlsx workbook
#' @param sheet_name a string indicating the name of the sheet to write to
#' @param paste_coord string indicating the coordinates of the top left cell to write data to.
#'   Supports multi-letter columns (e.g., "A1", "AA100", "XFD1"). Default is "A1".
#' @param clear_sheet logical indicating whether to clear the contents of the destination sheet
#'   before writing. Default is TRUE.
#' @param wb workbook object. Default is NULL (creates or loads from wb_dir).
#' @param wb_dir string indicating the workbook's file path
#' @param save_wb logical indicating whether workbook should be saved. Default is FALSE.
#' @param ... arguments for other methods
#'
#' @return The workbook object (invisibly)
#'
#' @examples
#' \dontrun{
#' # Basic usage - write data frame to Excel
#' write_wb(mtcars, sheet_name = "cars", wb_dir = "output.xlsx", save_wb = TRUE)
#'
#' # Write starting at a specific cell
#' write_wb(iris[1:10, ], sheet_name = "flowers", paste_coord = "B5",
#'          wb_dir = "positioned.xlsx", save_wb = TRUE)
#'
#' # Write to multi-letter column (e.g., column AA = 27th column)
#' write_wb(data.frame(x = 1:5), sheet_name = "Sheet1", paste_coord = "AA1",
#'          wb_dir = "wide.xlsx", save_wb = TRUE)
#'
#' # Build workbook incrementally
#' wb <- write_wb(mtcars[1:5, ], sheet_name = "Sheet1", wb_dir = "report.xlsx")
#' wb <- write_wb(iris[1:5, ], sheet_name = "Sheet2", wb = wb)
#' openxlsx::saveWorkbook(wb, "report.xlsx", overwrite = TRUE)
#' }
#'
#' @export
#'
write_wb.default <- function(r_data, sheet_name, paste_coord = "A1", clear_sheet = TRUE,
                              wb = NULL, wb_dir = NULL, save_wb = FALSE, ...) {

  # Parse cell reference using utility function (supports multi-letter columns)
  parsed <- parse_cell_reference(paste_coord)
  start_col <- parsed$col
  start_row <- parsed$row

  # Ensure that we have a workbook object

  wb <- load_wb(wb, wb_dir)

  # If the existing workbook doesn't have sheet_name, create it
  if (!(sheet_name %in% names(wb))) {

    openxlsx::addWorksheet(wb = wb, sheetName = sheet_name)
    message(paste("Sheet", sheet_name, "added to workbook"))

  } else if (clear_sheet) {
    # If sheet exists and clear_sheet is TRUE, clear contents dynamically
    dims <- get_sheet_dimensions(wb, sheet_name)

    if (dims$max_row > 0 && dims$max_col > 0) {
      openxlsx::deleteData(
        wb = wb,
        sheet = sheet_name,
        cols = seq_len(dims$max_col),
        rows = seq_len(dims$max_row),
        gridExpand = TRUE
      )
    }
  }

  # Write data to sheet_name in existing workbook
 suppressWarnings(openxlsx::writeData(
    wb = wb,
    sheet = sheet_name,
    x = r_data,
    startCol = start_col,
    startRow = start_row,
    keepNA = TRUE
  ))

  # Save the workbook with the data written to it
  if (save_wb) {
    assert_true(!is.null(wb_dir),
                "Cannot save: `wb_dir` must be specified when `save_wb = TRUE`")
    openxlsx::saveWorkbook(wb = wb, file = wb_dir, overwrite = TRUE)
  }

  invisible(wb)
}

#' Write an object of class zoo or xts to Excel
#'
#' Converts the time series to a data frame with a date column and writes to Excel.
#'
#' @param r_data an object of class \code{zoo} or \code{xts}
#' @param sheet_name a string indicating the name of the sheet to write to
#' @param paste_coord string indicating the coordinates of the top left cell to write data to.
#'   Supports multi-letter columns. Default is "A1".
#' @param clear_sheet logical indicating whether to clear the contents of the destination sheet
#'   before writing. Default is TRUE.
#' @param wb workbook object. Default is NULL.
#' @param wb_dir string indicating the workbook's file path
#' @param save_wb logical indicating whether workbook should be saved. Default is FALSE.
#' @param ... arguments for other methods
#'
#' @return The workbook object (invisibly)
#'
#' @examples
#' \dontrun{
#' library(zoo)
#'
#' # Create a zoo time series
#' dates <- seq(as.Date("2024-01-01"), by = "day", length.out = 10)
#' values <- data.frame(price = cumsum(rnorm(10)), volume = rpois(10, 100))
#' ts_data <- zoo(values, dates)
#'
#' # Write to Excel - date column is added automatically
#' write_wb(ts_data, sheet_name = "daily", wb_dir = "timeseries.xlsx", save_wb = TRUE)
#' }
#'
#' @export
#'
write_wb.zoo <- function(r_data, sheet_name, paste_coord = "A1", clear_sheet = TRUE,
                          wb = NULL, wb_dir = NULL, save_wb = FALSE, ...) {

  if (!requireNamespace("zoo", quietly = TRUE)) {
    stop("Package 'zoo' is required. Install with: install.packages('zoo')", call. = FALSE)
  }

  # Convert zoo to data.frame so that date column is written
  r_data_df <- data.frame(date = zoo::index(r_data), zoo::coredata(r_data))
  colnames(r_data_df) <- c("date", colnames(r_data))

  write_wb.default(
    r_data = r_data_df,
    sheet_name = sheet_name,
    paste_coord = paste_coord,
    clear_sheet = clear_sheet,
    wb = wb,
    wb_dir = wb_dir,
    save_wb = save_wb
  )
}

#' @rdname write_wb.zoo
#' @export
write_wb.xts <- write_wb.zoo

#' Write an object of class list to Excel
#'
#' Each named element of the list is written to a separate sheet.
#'
#' @param r_data a named list of objects to write. Names become sheet names.
#' @param clear_sheet logical indicating whether to clear sheet contents before writing.
#'   Default is TRUE.
#' @param wb workbook object. Default is NULL.
#' @param wb_dir string indicating the workbook's file path
#' @param save_wb logical indicating whether workbook should be saved. Default is FALSE.
#' @param ... arguments for other methods
#'
#' @return The workbook object (invisibly)
#'
#' @examples
#' \dontrun{
#' # Create a named list of data frames
#' my_sheets <- list(
#'   summary = data.frame(metric = c("Total", "Average"), value = c(100, 50)),
#'   details = mtcars[1:10, ],
#'   metadata = data.frame(created = Sys.Date(), author = "Me")
#' )
#'
#' # Write all sheets at once
#' write_wb(my_sheets, wb_dir = "multi_sheet_report.xlsx", save_wb = TRUE)
#' }
#'
#' @export
#'
write_wb.list <- function(r_data, clear_sheet = TRUE, wb = NULL, wb_dir = NULL,
                           save_wb = FALSE, ...) {

  wb <- load_wb(wb, wb_dir)

  for (nm in names(r_data)) {
    list_elem <- r_data[[nm]]
    write_wb(list_elem, sheet_name = nm, wb = wb, clear_sheet = clear_sheet, save_wb = FALSE, ...)
  }

  if (save_wb) {
    assert_true(!is.null(wb_dir),
                "Cannot save: `wb_dir` must be specified when `save_wb = TRUE`")
    openxlsx::saveWorkbook(wb = wb, file = wb_dir, overwrite = TRUE)
  }

  invisible(wb)
}

#' Write nested tibbles to multiple workbooks
#'
#' Takes a tibble with columns specifying file names, sheet names, and data,
#' and writes each data element to the appropriate file and sheet.
#'
#' @param write_tbl a tibble with columns:
#'   \itemize{
#'     \item \code{name}: file name (without extension)
#'     \item \code{object}: sheet name
#'     \item \code{data}: data to write (list column)
#'   }
#' @param write_dir string indicating the directory to write files to
#' @param ... additional arguments passed to write_wb
#'
#' @examples
#' \dontrun{
#' # Create a tibble defining multiple workbooks and sheets
#' write_tbl <- tibble::tibble(
#'   name = c("sales", "sales", "inventory"),
#'   object = c("Q1", "Q2", "stock"),
#'   data = list(
#'     data.frame(month = 1:3, revenue = c(100, 120, 150)),
#'     data.frame(month = 4:6, revenue = c(140, 160, 180)),
#'     data.frame(item = c("A", "B"), qty = c(50, 30))
#'   )
#' )
#'
#' # Creates sales.xlsx (with Q1 and Q2 sheets) and inventory.xlsx (with stock sheet)
#' write_wb_multi(write_tbl, write_dir = "output")
#' }
#'
#' @export
#'
write_wb_multi <- function(write_tbl, write_dir, ...) {

  assert_present(names(write_tbl), c("name", "object", "data"))

  # Ensure write directory exists
  if (!dir.exists(write_dir)) {
    dir.create(write_dir, recursive = TRUE)
  }

  file_names <- unique(write_tbl$name)

  # Write one file per file_name containing separate sheets per object
  for (fname in file_names) {
    write_tbl_single <- dplyr::filter(write_tbl, .data$name == fname)
    wb_path <- file.path(write_dir, paste0(fname, ".xlsx"))

    # Create workbook for this file
    wb <- NULL

    for (n in seq_len(nrow(write_tbl_single))) {
      sheet <- write_tbl_single$object[[n]]
      data <- write_tbl_single$data[[n]]

      # Load existing workbook or create new one
      wb <- load_wb(wb, wb_path)

      write_wb(
        r_data = data,
        sheet_name = sheet,
        clear_sheet = TRUE,
        wb = wb,
        wb_dir = wb_path,
        save_wb = FALSE,
        ...
      )
    }

    # Save workbook after all sheets are written
    if (!is.null(wb)) {
      openxlsx::saveWorkbook(wb = wb, file = wb_path, overwrite = TRUE)
      message(paste("Workbook saved:", wb_path))
    }
  }

  invisible(NULL)
}
