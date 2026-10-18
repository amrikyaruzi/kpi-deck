# =============================================================================
# KPI deck: one function call per dataset.
# =============================================================================

# 1. Load the code (once per R session). Point this at the folder you unzipped.
kpi_home <- if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
  dirname(rstudioapi::getSourceEditorContext()$path)
} else getwd()
source(file.path(kpi_home, "R", "kpi_functions.R"))
#   or simply:  source("C:/Users/you/Documents/kpi-deck/R/kpi_functions.R")


# 2. NEW DATASET, first time only: create its indicator sheet, then fill it in Excel.
#    It finds the Yes/No columns and saves <name>_indicators.xlsx next to the data.
# make_indicator_sheet("C:/Audits/hand_hygiene_2026.xlsx")


# 3. Build the deck. The PPTX is saved in an "output" folder next to the data file.
make_kpi_deck(
  data       = file.path(kpi_home, "data", "raw_data.xlsx"),
  indicators = file.path(kpi_home, "data", "indicators.xlsx"),  # omit if <name>_indicators.xlsx sits next to the data
  title      = "Medical Records Documentation Audit",
  year       = 2026

  # Optional settings (defaults shown):
  # date_col       = NULL,       # NULL = first date column found
  # group_col      = NULL,       # e.g. "unit" -> one deck per unit
  # limits_method  = "sd",       # mean +/- 3 SD as in the Excel sheet
  # quarter_method = "pooled",   # or "mean" (mean of the monthly rates)
  # output         = "pptx",     # or "pdf", or c("pptx", "pdf") (PDF needs Typst)
  # facility       = "Mji Mpya Referral Hospital, Dodoma",     # name, location (title slide box)
  # period         = NULL,       # e.g. "Q3 2026"; NULL = worked out from the data
  # indicator_list = TRUE,       # slide 2: list of indicators
  # subtitle       = ""
)
