#' ============================================================================
#' excelsioR Demo: Taming Messy Spreadsheets
#' ============================================================================
#'
#' This script demonstrates excelsioR's core value proposition:
#' Reading spreadsheets that look nice to humans but are hard to analyze.
#'
#' Run this entire script to see excelsioR in action!
#' ============================================================================

library(excelsioR)
library(dplyr)
library(tidyr)
library(openxlsx)

cat("excelsioR Demo: Taming Messy Spreadsheets\n")
cat("==========================================\n\n")

# Create a temporary directory for our demo
demo_dir <- file.path(tempdir(), "excelsioR_demo")
dir.create(demo_dir, showWarnings = FALSE)
cat("Demo directory:", demo_dir, "\n\n")

# ============================================================================
# PART 1: Create "Human-Friendly" but Analysis-Hostile Spreadsheets
# ============================================================================

cat("PART 1: Creating messy spreadsheets...\n")
cat("--------------------------------------\n")

# These represent common patterns: titles, footnotes, inconsistent layouts

# File 1: Sales report with title row and footnotes
create_sales_report <- function(path) {
  wb <- createWorkbook()
  addWorksheet(wb, "Q1 Sales")
  addWorksheet(wb, "Q2 Sales")

  for (quarter in c("Q1", "Q2")) {
    sheet <- paste(quarter, "Sales")
    # Row 1: Title (NOT headers!)
    writeData(wb, sheet, x = paste("Regional Sales -", quarter, "2024"),
              startRow = 1, startCol = 1)

    # Row 2: Empty (visual spacing)

    # Row 3: Headers
    headers <- data.frame(Region = "", Jan = "", Feb = "", Mar = "")
    if (quarter == "Q2") headers <- data.frame(Region = "", Apr = "", May = "", Jun = "")
    writeData(wb, sheet, x = headers, startRow = 3)

    # Rows 4-7: Actual data
    if (quarter == "Q1") {
      data <- data.frame(
        Region = c("North", "South", "East", "West"),
        Jan = c(12450, 9800, 11300, 10200),
        Feb = c(13200, 10100, 11800, 10500),
        Mar = c(14100, 11200, 12400, 10900)
      )
    } else {
      data <- data.frame(
        Region = c("North", "South", "East", "West"),
        Apr = c(14500, 11500, 12800, 11200),
        May = c(15200, 12100, 13400, 11800),
        Jun = c(15800, 12600, 13900, 12300)
      )
    }
    writeData(wb, sheet, x = data, startRow = 4, colNames = FALSE)

    # Row 9: Footnote
    writeData(wb, sheet, x = "* Figures in USD thousands. Excludes inter-company.",
              startRow = 9, startCol = 1)
  }

  saveWorkbook(wb, path, overwrite = TRUE)
}

# File 2: Inventory report with different structure
create_inventory_report <- function(path) {
  wb <- createWorkbook()
  addWorksheet(wb, "Inventory")

  # This one has headers in row 2 (different from sales report!)
  writeData(wb, "Inventory", x = "Warehouse Inventory Snapshot",
            startRow = 1, startCol = 1)

  data <- data.frame(
    SKU = c("A001", "A002", "B001", "B002", "C001"),
    Description = c("Widget Alpha", "Widget Beta", "Gadget Pro", "Gadget Lite", "Gizmo X"),
    Quantity = c(1250, 890, 2100, 1560, 780),
    Reorder_Point = c(500, 400, 800, 600, 300),
    Status = c("OK", "OK", "OK", "OK", "LOW")
  )
  writeData(wb, "Inventory", x = data, startRow = 2)

  saveWorkbook(wb, path, overwrite = TRUE)
}

# File 3: Multiple tables in one sheet (a classic headache)
create_multi_table_report <- function(path) {
  wb <- createWorkbook()
  addWorksheet(wb, "Dashboard")

  # Table 1: Top left
  writeData(wb, "Dashboard", x = "Revenue by Channel", startRow = 1, startCol = 1)
  writeData(wb, "Dashboard",
            x = data.frame(Channel = c("Online", "Retail", "Wholesale"),
                           Revenue = c(45000, 32000, 28000)),
            startRow = 2, startCol = 1)

  # Table 2: Top right (starts at column E)
  writeData(wb, "Dashboard", x = "Costs by Category", startRow = 1, startCol = 5)
  writeData(wb, "Dashboard",
            x = data.frame(Category = c("Materials", "Labor", "Overhead"),
                           Amount = c(22000, 18000, 12000)),
            startRow = 2, startCol = 5)

  # Table 3: Below (starts at row 8)
  writeData(wb, "Dashboard", x = "Monthly Trend", startRow = 8, startCol = 1)
  writeData(wb, "Dashboard",
            x = data.frame(Month = month.abb[1:6],
                           Value = c(12, 14, 13, 16, 18, 17)),
            startRow = 9, startCol = 1)

  saveWorkbook(wb, path, overwrite = TRUE)
}

# Create all the files
create_sales_report(file.path(demo_dir, "sales_report.xlsx"))
create_inventory_report(file.path(demo_dir, "inventory.xlsx"))
create_multi_table_report(file.path(demo_dir, "dashboard.xlsx"))

cat("Created 3 messy spreadsheets:\n")
cat("  - sales_report.xlsx (title rows, footnotes, multiple sheets)\n")
cat("  - inventory.xlsx (headers in row 2)\n")
cat("  - dashboard.xlsx (multiple tables in one sheet)\n\n")

# ============================================================================
# PART 2: Read All Files with One Command
# ============================================================================

cat("PART 2: Reading all files with read_wb()...\n")
cat("-------------------------------------------\n")

all_data <- read_wb(source_data_dir = demo_dir)

cat("\nResult structure:\n")
print(all_data)

cat("\nFiles found:", paste(all_data$file_name, collapse = ", "), "\n")

# ============================================================================
# PART 3: Examine Raw Cell Data
# ============================================================================

cat("\n\nPART 3: Examining raw cell data...\n")
cat("-----------------------------------\n")

# Look at the sales report's raw cells
# raw_excel_data is a list of tibbles, one per sheet
sales_raw_list <- all_data %>%
  filter(file_name == "sales_report") %>%
  pull(raw_excel_data) %>%
  .[[1]]

# Get the first sheet (Q1 Sales)
sales_raw <- sales_raw_list[[1]]

cat("\nFirst 15 cells from sales_report (first sheet):\n")
sales_raw %>%
  select(row, col, data_type, numeric, character) %>%
  head(15) %>%
  print()

cat("\nNotice how we can see:\n")
cat("  - Row 1, Col 1: Title text (not a header)\n")
cat("  - Row 3: Column headers\n")
cat("  - Rows 4-7: Actual data\n")
cat("  - Row 9: Footnote\n")

# ============================================================================
# PART 4: Extract Clean Data
# ============================================================================

cat("\n\nPART 4: Extracting clean data...\n")
cat("---------------------------------\n")

# Extract Q1 sales data properly using raw cell access
q1_clean <- sales_raw %>%
  filter(row >= 3, row <= 7, !is_blank) %>%  # Headers + data rows only
  select(row, col, numeric, character) %>%
  mutate(value = coalesce(as.character(numeric), character)) %>%
  select(row, col, value) %>%
  pivot_wider(names_from = col, values_from = value)

# First row is headers, rest is data
headers <- as.character(q1_clean[1, -1])
q1_data <- q1_clean[-1, -1]
names(q1_data) <- headers

# Convert to proper types
q1_data <- q1_data %>%
  mutate(across(where(function(x) all(!is.na(suppressWarnings(as.numeric(x))))), as.numeric))

cat("\nClean Q1 Sales data extracted from messy spreadsheet:\n")
print(q1_data)

# ============================================================================
# PART 5: Default Behavior Without Structure Hints
# ============================================================================

cat("\n\nPART 5: Default read without structure hints...\n")
cat("------------------------------------------------\n")

messy_file <- file.path(demo_dir, "sales_report.xlsx")
cat("\nDefault readxl::read_excel() output:\n")
readxl_result <- readxl::read_excel(messy_file)
print(readxl_result)

cat("\nWithout hints, the first row becomes column names.\n")
cat("For a single file, you'd specify skip=2, n_max=4.\n")
cat("For many files with varying layouts, cell-level access helps.\n")

# ============================================================================
# PART 6: Write Results Back to Excel
# ============================================================================

cat("\n\nPART 6: Writing results to Excel...\n")
cat("------------------------------------\n")

# Create a summary
summary_data <- data.frame(
  Region = c("North", "South", "East", "West"),
  Q1_Total = c(39750, 31100, 35500, 31600),
  Avg_Monthly = c(13250, 10367, 11833, 10533)
)

# Write as multi-sheet workbook using a named list
output_data <- list(
  "Summary" = summary_data,
  "Q1 Detail" = q1_data,
  "Processing Log" = data.frame(
    Step = c("Files Read", "Sheets Processed", "Records Extracted"),
    Count = c(3, 4, 8)
  )
)

output_file <- file.path(demo_dir, "analysis_output.xlsx")

write_wb(
  r_data = output_data,
  wb_dir = output_file,
  save_wb = TRUE
)

cat("Wrote multi-sheet report to:", output_file, "\n")

# ============================================================================
# Summary
# ============================================================================

cat("\n")
cat("============================================================\n")
cat("DEMO COMPLETE\n")
cat("============================================================\n\n")

cat("What we demonstrated:\n")
cat("  1. Created 3 'messy' spreadsheets with common issues:\n")
cat("     - Title rows (headers not in row 1)\n")
cat("     - Footnotes (extra rows after data)\n")
cat("     - Multiple sheets\n")
cat("     - Multiple tables in one sheet\n\n")

cat("  2. Read ALL files with a single read_wb() call\n\n")

cat("  3. Accessed raw cell data to understand spreadsheet structure\n\n")

cat("  4. Extracted clean data by filtering on row/col positions\n\n")

cat("  5. Wrote results back as a multi-sheet Excel file\n\n")

cat("The key insight:\n")
cat("  Presentation spreadsheets (titles, logos, footnotes) need cell-level\n")
cat("  access to parse reliably. tidyxl provides this brilliantly.\n\n")

cat("  excelsioR builds on tidyxl and openxlsx to provide a complete workflow:\n")
cat("  - Batch processing: read folders of files at once\n")
cat("  - Structured output: nested tibbles for easy filtering\n")
cat("  - Write capabilities: send results back to Excel\n\n")

cat("Output files in:", demo_dir, "\n")
cat("  - analysis_output.xlsx (your results)\n")
