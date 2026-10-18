# Example: monthly numerator/denominator data (no raw records). The data and the
# indicator sheet are two sheets in one workbook, so only the file is needed.
# p-chart limits suit monthly totals with denominators in the hundreds.
source("R/kpi_functions.R")
make_kpi_deck("examples/kpi_monthly.xlsx", title = "Monthly KPI Report", year = 2026,
              limits_method = "p", output = c("pptx", "pdf"))
