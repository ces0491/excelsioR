#' load a workbook object from file or use the current workbook if it exists
#'
#' @param wb optional workbook object
#' @param wb_dir optional path to a file to load if file doesn't exist it will just save the file
#'
#' @return workbook object. See \link[openxlsx]{loadWorkbook}
#'
load_wb <- function(wb = NULL, wb_dir = NULL) {

  fractalAssert::assert_true( (!is.null(wb)) | (!is.null(wb_dir)), "Can't have a NULL workbook and wb_dir")

  if (is.null(wb)) {

    if (!is.null(wb_dir) && file.exists(wb_dir)) {

      wb <- suppressWarnings(openxlsx::loadWorkbook(wb_dir)) # if wb null and wb_dir exists, load it

    } else {

      wb <- openxlsx::createWorkbook() # if wb null and wb_dir null, create new workbook
      message("New workbook created")
    }

  }

  wb

}

#' Write from R to Excel
#'
#' @param r_data an object to be written to an xlsx workbook when r_data is a list it will use the names of the list as sheet_names
#' @param ... arguments for other methods
#'
#' @export
#'
write_wb <- function(r_data, ...) UseMethod("write_wb")

#' Default method for writing to excel
#'
#' @param r_data an object to be written to an xlsx workbook when r_data is a list it will use the names of the list as sheet_names
#' @param sheet_name a string indicating the name of the sheet to write to
#' @param paste_coord string indicating the coordinates of the top left cell you wish to write data to - default "A1"
#' @param clear_sheet logical indicating whether you'd like to clear the contents of the destination sheet before writing - default TRUE
#' @param wb workbook object - default NULL
#' @param wb_dir string indicating the workbook's directory
#' @param save_wb logical indicating whether workbook should be saved
#' @param ... arguments for other methods
#'
#' @export
#'
write_wb.default <- function(r_data, sheet_name, paste_coord = "A1", clear_sheet = TRUE, wb = NULL, wb_dir = NULL, save_wb = FALSE, ...) {

  col <- toupper(gsub("[0-9]", "", paste_coord))
  start_col <- which(LETTERS %in% col)
  assertR::assert_true(length(start_col) > 0, "only start columns between A and Z are currently supported")
  start_row <- gsub("[a-z A-Z]", "", paste_coord)

  # ensure that we have a workbook object
  wb <- load_wb(wb, wb_dir)

  # if the existing workbook doesn't have sheet_name, create it
  if (!(sheet_name %in% names(wb))) {

    openxlsx::addWorksheet(wb = wb,
                           sheetName = sheet_name)
    message(paste("sheet", sheet_name, "added to workbook", sep = " "))

  } else {
    # if sheet does exist, clear contents
    assertR::assert_true((sheet_name %in% names(wb)) && clear_sheet, "sheet must exist in order to clear it")

    openxlsx::deleteData(wb = wb,
                         sheet = sheet_name,
                         cols = 1:100,
                         rows = 1:20000,
                         gridExpand = TRUE) # all available cells in excel uses too much memory so we make this assumption
  }

  # write data to sheet_name in existing workbook
  suppressWarnings(openxlsx::writeData(wb = wb,
                                       sheet = sheet_name,
                                       x = r_data,
                                       startCol = start_col,
                                       startRow = start_row,
                                       keepNA = TRUE))

  # save the workbook with the data written to it
  # if save is TRUE wb_dir must be a string
  if (save_wb) {
    assertR::assert_true(!is.null(wb_dir), "no wb_dir specified to save to")
    openxlsx::saveWorkbook(wb = wb,
                           file =  wb_dir,
                           overwrite = TRUE)
  }

  wb

}

#' Write an object of class \code{zoo} or \code{xts} to excel
#'
#' @param r_data an object of class \code{zoo}
#' @param sheet_name a string indicating the name of the sheet to write to
#' @param paste_coord string indicating the coordinates of the top left cell you wish to write data to - default "A1"
#' @param clear_sheet logical indicating whether you'd like to clear the contents of the destination sheet before writing - default TRUE
#' @param wb workbook object - default NULL
#' @param wb_dir string indicating the workbook's directory
#' @param save_wb logical indicating whether workbook should be saved
#' @param ... arguments for other methods
#'
#' @export
#'
write_wb.zoo <- function(r_data, sheet_name, paste_coord = "A1", clear_sheet = TRUE, wb = NULL, wb_dir = NULL, save_wb = FALSE, ...) {

  #convert zoo to data.frame so that date column is written
  r_data_df <- data.frame(date = zoo::index(r_data), zoo::coredata(r_data))
  colnames(r_data_df) <- c("date", colnames(r_data))

  write_to_wb.default(r_data = r_data.df,
                      sheet_name = sheet_name,
                      paste_coord = paste_coord,
                      clear_sheet = clear_sheet,
                      wb = wb,
                      wb_dir = wb_dir,
                      save_wb = save_wb)

}

#' Write an object of class \code{list} to excel
#'
#' @param r_data an object of class \code{list}
#' @param clear_sheet logical indicating whether you'd like to clear the contents of the destination sheet before writing - default TRUE
#' @param wb workbook object - default NULL
#' @param wb_dir string indicating the workbook's directory
#' @param save_wb logical indicating whether workbook should be saved
#' @param ... arguments for other methods
#'
#' @export
#'
write_wb.list <- function(r_data, clear_sheet = TRUE, wb = NULL, wb_dir = NULL, save_wb = FALSE, ...) {

  wb <- load_wb(wb, wb_dir)

  for (nm in names(r_data)) {

    list_elem <- r_data[[nm]]
    write_wb(list_elem, sheet_name = nm, wb = wb, save_wb = FALSE, keepNA = TRUE)

  }

  if (save_wb) {

    assertR::assert_true(!is.null(wb_dir), "no wb_dir specified to save to")
    openxlsx::saveWorkbook(wb = wb,
                           file =  wb_dir,
                           overwrite = TRUE)

  }

  wb

}
