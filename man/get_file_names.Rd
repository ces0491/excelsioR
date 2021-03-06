% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/get_file_names.R
\name{get_file_names}
\alias{get_file_names}
\title{get excel file names from directory}
\usage{
get_file_names(data_dir, recursive = FALSE, include_dirs = FALSE)
}
\arguments{
\item{data_dir}{string indicating the directory of interest}

\item{recursive}{logical indicating whether to return file names recursively}

\item{include_dirs}{logical indicating whether names of sub directories should also be included}
}
\value{
\code{data.frame} with columns file_path and file_name
}
\description{
get excel file names from directory
}
