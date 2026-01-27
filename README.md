# excelsioR

Tidy retrieval of often untidy spreadsheet data. Read and write Excel files
with a focus on handling messy, real-world spreadsheets.

## Installation

```r
# Install from GitHub
devtools::install_github("ces0491/excelsioR")

# Or using pak
pak::pak("ces0491/excelsioR")
```

## Features

- **Safe reading**: Reads from copies of files to prevent corruption
- **Untidy data handling**: Extracts structured data from messy spreadsheets
  using tidyxl
- **Multi-letter column support**: Write to columns beyond Z (AA, AB, etc.)
- **Password protection**: Unlock protected workbooks (requires rpxl or
  XLConnect)
- **Time series support**: Native support for zoo/xts objects
- **Multiple sheets**: Read/write multiple sheets and workbooks at once

## Quick Start

### Reading Excel Files

```r
library(excelsioR)

# Read all Excel files from a directory
data <- read_wb("path/to/excel/files")

# Read specific workbooks and sheets
data <- read_wb(
  source_data_dir = "path/to/files",
  reqd_wkbks = c("sales_2024", "inventory"),
  reqd_sheets = c("Summary", "Details")
)

# The result is a tibble with nested data
data$file_name          # File names
data$raw_excel_data     # Raw cell data (from tidyxl)
data$all_excel_data_tbl # Processed tidy data
```

### Writing Excel Files

```r
# Write a data frame to Excel
write_wb(
  mtcars,
  sheet_name = "cars",
  wb_dir = "output.xlsx",
  save_wb = TRUE
)

# Write to a specific cell (supports multi-letter columns)
write_wb(
  my_data,
  sheet_name = "Sheet1",
  paste_coord = "AA10",
  wb_dir = "output.xlsx",
  save_wb = TRUE
)

# Write multiple sheets from a named list
sheets <- list(
  cars = mtcars,
  flowers = iris
)
write_wb(sheets, wb_dir = "multi_sheet.xlsx", save_wb = TRUE)

# Write zoo/xts time series (date column added automatically
library(zoo)
ts_data <- zoo(matrix(rnorm(20), 10, 2), Sys.Date() + 1:10)
write_wb(ts_data, sheet_name = "timeseries", wb_dir = "ts.xlsx", save_wb = TRUE)
```

### Password-Protected Files

```r
# Read password-protected files (will prompt for password)
data <- read_wb("path/to/files", password_protected = TRUE)

# Or unlock files directly
file_info <- data.frame(
  file_path = "protected.xlsx",
  file_name = "protected"
)
unlock_wb(file_info, wb_password = "secret")
```

#### Backend Options for Password Protection

excelsioR supports two backends for unlocking protected files:

| Backend       | Pros             | Cons          | Install                 |
|---------------|------------------|---------------|-------------------------|
| **rpxl**      | No Java required | Needs Python  | `install_github("...")` |
| **XLConnect** | Battle-tested    | Requires Java | `install.packages(...)` |

The function automatically selects the available backend, preferring rpxl.

## How It Works

### Reading Workflow

1. Scans source directory for Excel files (.xlsx, .xlsm, .xls)
2. Creates timestamped copies to prevent file corruption
3. Unlocks password-protected files if needed
4. Reads raw cell data using tidyxl
5. Processes into tidy format with columns: sheet, row, col, numeric,
   character, date

### Data Structure Assumptions

The `read_wb()` function assumes:

- The leftmost non-blank column contains row labels/variable names
- Dates are in dedicated columns
- This works well for financial/time-series data

For other layouts, use the `raw_excel_data` column which contains the
unprocessed tidyxl output.

## Dependencies

**Required:**

- dplyr, tidyr, purrr, tibble, stringr, glue
- openxlsx (writing)
- tidyxl (reading untidy data)
- rstudioapi (password prompts)

**Optional:**

- zoo (time series support)
- rpxl or XLConnect (password-protected files)

## License

GPL-3
