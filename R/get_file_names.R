#' get excel file names from directory
#'
#' @param data_dir string indicating the directory of interest
#' @param recursive logical indicating whether to return file names recursively
#' @param include_dirs logical indicating whether names of sub directories should also be included
#'
#' @return \code{data.frame} with columns file_path and file_name
#'
get_file_names <- function(data_dir, recursive = FALSE, include_dirs = FALSE) {

  file_name <- list.files(data_dir, full.names = TRUE, recursive = recursive, include.dirs = include_dirs)

  file_names_df <- file_name %>%
    data.frame(file_path = .) %>%
    dplyr::filter(stringr::str_detect(file_path, ".xls")) %>% # return only excel file names xls, xlsx, xlsm
    dplyr::mutate(file_name = gsub(".*[//]([^.]+)[.].*", "\\1", file_path)) %>% # get string between fwd slash and dot - filename without the extension
    dplyr::mutate(file_name = trimws(file_name, "both")) %>% # trim the whitespace on either end
    dplyr::mutate(file_name = gsub("^\\d+", "", file_name)) %>% # remove leading numerics
    dplyr::mutate(file_name = gsub("[^[:alnum:]]", "_", file_name)) # replace all non alphanums with _

  file_names_df
}
