# excelsioR 0.1.0

Initial CRAN release.
## Features

### Reading Excel Files

- `read_wb()`: Read Excel files from a directory with automatic file copying
  for safety
- Support for filtering by workbook name and sheet name
- Returns both raw tidyxl cell data and processed tidy data
- Handles password-protected files via `unlock_wb()`

### Writing Excel Files

- `write_wb()`: Generic function with methods for data frames, lists, and
  zoo/xts time series
- Support for multi-letter column references (AA, AB, etc.)
- `write_wb_multi()`: Batch write multiple workbooks from a nested tibble

### Password Protection

- `unlock_wb()`: Remove password protection from Excel files
- Dual backend support: rpxl (Python-based) and XLConnect (Java-based)
- Automatic backend selection with preference for rpxl

### Utilities

- `get_file_names()`: Scan directories for Excel files with name cleaning
- Multi-letter column conversion (A=1, Z=26, AA=27, etc.)
- Dynamic sheet dimension detection for clearing data
