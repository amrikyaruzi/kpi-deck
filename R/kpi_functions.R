# =============================================================================
# KPI slide-deck pipeline: Excel -> R (calculations) -> PowerPoint and/or PDF
# Main entry point: make_kpi_deck()  (see the bottom of this file)
# =============================================================================

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(purrr)
})

`%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || all(is.na(a))) b else a

# Folder this code lives in (the one holding R/ and templates/), found when sourced
KPI_HOME <- local({
  home <- getOption("kpi.home")
  if (is.null(home)) for (i in rev(seq_len(sys.nframe()))) {
    f <- sys.frame(i)$ofile
    if (!is.null(f) && grepl("kpi_functions\\.R$", f)) { home <- dirname(dirname(normalizePath(f))); break }
  }
  normalizePath(home %||% getwd(), winslash = "/")
})

MONTHS <- month.abb

# ---- Small helpers ----------------------------------------------------------


slug <- function(x) {
  x <- gsub("[^A-Za-z0-9]+", "_", trimws(x))
  gsub("^_|_$", "", x)
}

blank <- function(x) is.null(x) || length(x) == 0 || is.na(x) || !nzchar(trimws(as.character(x)))

fmt_num <- function(x, digits = 1) {
  ifelse(is.na(x), "", formatC(round(x, digits), format = "f", digits = digits))
}

fmt_int <- function(x) ifelse(is.na(x), "", format(x, big.mark = ",", trim = TRUE))

# Splits a multi-line Excel cell ("1. foo\n2. bar") into clean list items
split_lines <- function(x) {
  if (blank(x)) return(character(0))
  out <- trimws(unlist(strsplit(as.character(x), "\r?\n")))
  out <- sub("^([0-9]+[.)]|[-*\u2022])\\s*", "", out)
  out[nzchar(out)]
}

# Accepts real dates, month numbers (1-12) or month names ("Jan", "July", "Sept")
parse_month <- function(x) {
  if (inherits(x, c("Date", "POSIXct", "POSIXt"))) return(as.integer(format(x, "%m")))
  if (is.numeric(x)) return(as.integer(x))
  x <- tolower(substr(trimws(as.character(x)), 1, 3))
  match(x, tolower(MONTHS))
}

parse_year <- function(x) {
  if (inherits(x, c("Date", "POSIXct", "POSIXt"))) return(as.integer(format(x, "%Y")))
  rep(NA_integer_, length(x))
}

# Converts Yes/No style answers to 1 / 0; anything else (N/A, blank) -> NA
to_flag <- function(x, yes_values, no_values) {
  v <- tolower(trimws(as.character(x)))
  out <- rep(NA_real_, length(v))
  out[v %in% tolower(yes_values)] <- 1
  out[v %in% tolower(no_values)]  <- 0
  out
}

# ---- Reading inputs ---------------------------------------------------------

read_indicators <- function(cfg) {
  meta <- read_excel(cfg$meta_file, sheet = cfg$meta_sheet %||% 1)
  names(meta) <- tolower(trimws(names(meta)))
  need <- c("indicator", "title")
  miss <- setdiff(need, names(meta))
  if (length(miss)) stop("Indicator sheet is missing column(s): ", paste(miss, collapse = ", "))
  meta <- meta |> filter(!is.na(indicator))
  if ("include" %in% names(meta)) {
    meta <- meta |> filter(is.na(include) | toupper(as.character(include)) %in% c("Y", "YES", "TRUE", "1"))
  }
  # Optional columns: create them empty so later code can rely on them
  opt <- c("measure", "target", "reporting_schedule", "definition", "numerator", "denominator",
           "rule", "source_column", "criteria", "data_provider", "responsible", "analysis",
           "action", "chart_title")
  for (o in setdiff(opt, names(meta))) meta[[o]] <- NA_character_
  if ("order" %in% names(meta)) meta <- arrange(meta, order)
  meta
}

# Raw data: one row per audited record/opportunity, one column per criterion
read_raw <- function(cfg) {
  df <- read_excel(cfg$data_file, sheet = cfg$data_sheet %||% 1)
  if (blank(cfg$date_col)) {
    is_date <- vapply(df, function(x) inherits(x, c("Date", "POSIXct")), logical(1))
    cand <- c(names(df)[is_date], intersect(names(df), c("month", "Month", "MONTH", "date", "Date")))
    if (!length(cand)) stop("No date or month column found. Set date_col = \"<column name>\".")
    cfg$date_col <- cand[1]
    message("Using '", cfg$date_col, "' as the date column.")
  }
  if (!cfg$date_col %in% names(df)) stop("Date/month column '", cfg$date_col, "' not found in raw data.")
  df$.month <- parse_month(df[[cfg$date_col]])
  df$.year  <- parse_year(df[[cfg$date_col]])
  if (!is.null(cfg$year)) df <- df |> filter(is.na(.year) | .year == cfg$year)
  bad <- sum(is.na(df$.month))
  if (bad) warning(bad, " row(s) dropped: month could not be read from '", cfg$date_col, "'.")
  filter(df, !is.na(.month))
}

# Aggregate data: columns indicator, month, numerator, denominator (+ optional year, group)
read_aggregate <- function(cfg) {
  df <- read_excel(cfg$data_file, sheet = cfg$data_sheet %||% 1)
  names(df) <- tolower(trimws(names(df)))
  df$.month <- parse_month(df$month)
  if (!is.null(cfg$year) && "year" %in% names(df)) df <- filter(df, year == cfg$year)
  filter(df, !is.na(.month))
}

# ---- Monthly numerator / denominator ---------------------------------------

empty_months <- function() tibble(month = 1:12, num = NA_real_, den = NA_real_)

# rule = "column": numerator = Yes, denominator = Yes + No in that column
# rule = "all":    record counts as compliant only if every assessed criterion is Yes
# rule = "pooled": every criterion answer pooled together (total Yes / total assessed)
counts_from_raw <- function(df, ind, criteria_cols, cfg) {
  rule <- tolower(ind$rule %||% "column")
  cols <- if (rule == "column") {
    ind$source_column %||% ind$indicator
  } else if (!blank(ind$criteria)) {
    trimws(strsplit(ind$criteria, ",")[[1]])
  } else criteria_cols

  miss <- setdiff(cols, names(df))
  if (length(miss)) stop("Indicator '", ind$indicator, "': column(s) not in raw data: ", paste(miss, collapse = ", "))

  flags <- as.data.frame(lapply(df[cols], to_flag, cfg$yes_values, cfg$no_values))

  per_row <- switch(rule,
    column = tibble(num = flags[[1]] == 1, den = !is.na(flags[[1]])),
    all = {
      assessed <- rowSums(!is.na(flags)) > 0
      tibble(num = assessed & rowSums(flags == 0, na.rm = TRUE) == 0, den = assessed)
    },
    pooled = tibble(num = rowSums(flags == 1, na.rm = TRUE), den = rowSums(!is.na(flags))),
    stop("Unknown rule '", rule, "' for indicator '", ind$indicator, "'. Use column, all or pooled.")
  )

  per_row |>
    mutate(month = df$.month, num = as.numeric(coalesce(num, FALSE)), den = as.numeric(den)) |>
    group_by(month) |>
    summarise(num = sum(num), den = sum(den), .groups = "drop") |>
    right_join(tibble(month = 1:12), by = "month") |>
    arrange(month) |>
    mutate(across(c(num, den), ~ ifelse(is.na(den) | den == 0, NA, .x)))
}

counts_from_aggregate <- function(df, ind) {
  df |>
    filter(indicator == ind$indicator) |>
    group_by(month = .month) |>
    summarise(num = sum(numerator, na.rm = TRUE), den = sum(denominator, na.rm = TRUE), .groups = "drop") |>
    right_join(tibble(month = 1:12), by = "month") |>
    arrange(month) |>
    mutate(across(c(num, den), ~ ifelse(is.na(den) | den == 0, NA, .x)))
}

# ---- Rates, quarters, control limits ---------------------------------------

# method "pooled": sum(num) / sum(den) across the quarter (weights by volume)
# method "mean":   simple mean of the monthly rates (how the original slides did it)
quarter_rates <- function(m, method = "pooled") {
  m |>
    mutate(q = (month - 1) %/% 3 + 1) |>
    group_by(q) |>
    summarise(
      n_months = sum(!is.na(den)),
      # Only report a quarter once all three of its months have data
      rate = if (n_months < 3) NA_real_ else if (method == "mean") mean(rate, na.rm = TRUE)
             else 100 * sum(num, na.rm = TRUE) / sum(den, na.rm = TRUE),
      .groups = "drop"
    )
}

# method "p":  p-chart; centre = pooled rate, limits = p +/- 3*sqrt(p(1-p)/n_month)
# method "laney": p-chart adjusted for over-dispersion (see below)
# method "sd": mean of monthly rates +/- 3 * population SD (the original Excel sheet)
control_limits <- function(m, method = "sd", cap = FALSE) {
  ok <- !is.na(m$den)
  m$cl <- m$ucl <- m$lcl <- NA_real_
  if (sum(ok) < 2) return(m)
  if (method == "p") {
    p <- sum(m$num[ok]) / sum(m$den[ok])
    s <- sqrt(p * (1 - p) / m$den)
    m$cl  <- 100 * p
    m$ucl <- ifelse(ok, 100 * pmin(1, p + 3 * s), NA)
    m$lcl <- ifelse(ok, 100 * pmax(0, p - 3 * s), NA)
  } else if (method == "laney") {
    # Laney p' chart: p-chart limits widened by the observed between-month variation
    # (sigma_z). Recommended when denominators are large (hundreds+), where a plain
    # p-chart flags almost every month.
    p <- sum(m$num[ok]) / sum(m$den[ok])
    s <- sqrt(p * (1 - p) / m$den)
    z <- (m$num / m$den - p) / s
    sz <- if (p %in% c(0, 1)) 1 else mean(abs(diff(z[ok]))) / 1.128
    m$cl  <- 100 * p
    m$ucl <- ifelse(ok, 100 * pmin(1, p + 3 * s * sz), NA)
    m$lcl <- ifelse(ok, 100 * pmax(0, p - 3 * s * sz), NA)
  } else {
    # Same as the Excel sheet: AVERAGE and STDEV.P of the monthly rates,
    # UCL = mean + 3 SD, LCL = mean - 3 SD, constant across the year
    r <- m$rate[ok]
    mu <- mean(r); sdp <- sqrt(mean((r - mu)^2))
    m$cl  <- mu
    m$ucl <- if (cap) min(100, mu + 3 * sdp) else mu + 3 * sdp
    m$lcl <- if (cap) max(0, mu - 3 * sdp) else mu - 3 * sdp
  }
  m$signal <- ok & (m$rate > m$ucl + 1e-9 | m$rate < m$lcl - 1e-9)
  m
}

compute_indicator <- function(counts, cfg) {
  m <- counts |> mutate(rate = 100 * num / den)
  m <- control_limits(m, cfg$limits_method, cap = cfg$cap_limits %||% FALSE)
  if (is.null(m$signal)) m$signal <- FALSE
  m$signal[is.na(m$signal)] <- FALSE
  list(months = m, quarters = quarter_rates(m, cfg$quarter_method))
}

# ---- Auto-drafted analysis (used only when the analysis cell is empty) -----

draft_analysis <- function(st, target, cfg) {
  m <- st$months |> filter(!is.na(den))
  if (nrow(m) == 0) return("No data recorded for this period.")
  d <- cfg$rate_digits
  last <- tail(m, 1)
  out <- character(0)

  s1 <- sprintf("%s compliance was %s%% (%s/%s)", MONTHS[last$month], fmt_num(last$rate, d),
                fmt_int(last$num), fmt_int(last$den))
  if (!is.na(target)) s1 <- paste0(s1, if (last$rate >= target) ", meeting" else ", below", " the ", target, "% target")
  if (nrow(m) > 1) {
    prev <- m[nrow(m) - 1, ]
    diff <- round(last$rate - prev$rate, d)
    trend <- if (diff > 0) "up from" else if (diff < 0) "down from" else "unchanged from"
    s1 <- sprintf("%s; %s %s%% in %s", s1, trend, fmt_num(prev$rate, d), MONTHS[prev$month])
  }
  out <- c(out, paste0(s1, "."))

  if (!is.na(target)) {
    met <- sum(m$rate >= target)
    out <- c(out, sprintf("Target met in %d of %d month%s reported.", met, nrow(m), if (nrow(m) == 1) "" else "s"))
  }

  q <- st$quarters |> filter(!is.na(rate))
  if (nrow(q) >= 2) {
    a <- q[nrow(q) - 1, ]; b <- q[nrow(q), ]
    out <- c(out, sprintf("Q%d rate %s%% vs Q%d %s%%.", b$q, fmt_num(b$rate, d), a$q, fmt_num(a$rate, d)))
  }

  if (nrow(m) >= 2 && !all(is.na(m$ucl))) {
    sig <- m$month[m$signal]
    out <- c(out, if (length(sig))
      sprintf("%s outside the control limits: special-cause variation, review what changed.",
              paste(MONTHS[sig], collapse = ", "))
    else "All months within the control limits: common-cause variation only.")
  }
  out
}

# ---- Title-slide period ------------------------------------------------------
# Whole quarters -> "Q3 2026" or "Q1 - Q2 2026"; otherwise "January to June 2026"
period_label <- function(first, last, year) {
  q1 <- (first - 1) %/% 3 + 1; q2 <- (last - 1) %/% 3 + 1
  if (first %% 3 == 1 && last %% 3 == 0)
    return(trimws(paste(if (q1 == q2) paste0("Q", q1) else paste0("Q", q1, " - Q", q2), year)))
  trimws(paste(month.name[first], "to", month.name[last], year))
}

# ---- Y-axis range -------------------------------------------------------------
# mode "range" (default): pad by 20% of the spread between the lowest and highest
#   plotted values, so small differences (e.g. 98-100%) stay visible.
# mode "value": highest x 1.2 and lowest x 0.8 (x 1.2 if negative).
# Both are then rounded outward to a gridline. The axis only goes below 0 if a line does.
axis_limits <- function(values, pad = 0.2, mode = "range") {
  v <- values[is.finite(values)]
  if (!length(v)) return(list(lims = c(0, 100), step = 20))
  vmax <- max(v); vmin <- min(v)
  if (mode == "value") {
    hi <- if (vmax > 0) vmax * (1 + pad) else vmax * (1 - pad)
    lo <- if (vmin > 0) vmin * (1 - pad) else vmin * (1 + pad)
  } else {
    spread <- max(vmax - vmin, 2)
    hi <- vmax + pad * spread
    lo <- vmin - pad * spread
  }
  if (vmin >= 0) lo <- max(0, lo)
  if (hi - lo < 1) { hi <- hi + 1; lo <- lo - 1 }
  br <- pretty(c(lo, hi), n = 5)
  step <- diff(br)[1]
  list(lims = c(floor(lo / step) * step, ceiling(hi / step) * step), step = step)
}

# ---- Chart ------------------------------------------------------------------

# Use the sigma symbol only where R runs in UTF-8 (default on Windows from R 4.2)
SIG <- if (isTRUE(l10n_info()[["UTF-8"]])) "\u03c3" else " SD"
UCL_LAB <- paste0("UCL (3", SIG, ")")
LCL_LAB <- paste0("LCL (3", SIG, ")")
chart_colours <- setNames(c("#1F4E79", "#ED7D31", "#548235", "#5B9BD5", "#7030A0"),
                          c("Compliance (%)", "Target", "Average", UCL_LAB, LCL_LAB))

make_chart <- function(st, title, target, file, cfg) {
  m <- st$months
  lines <- bind_rows(
    tibble(month = m$month, value = m$rate, series = "Compliance (%)"),
    tibble(month = 1:12, value = target, series = "Target"),
    tibble(month = m$month, value = if (cfg$limits_method %in% c("p", "laney")) ifelse(is.na(m$den), NA, m$cl) else m$cl, series = "Average"),
    tibble(month = m$month, value = m$ucl, series = UCL_LAB),
    tibble(month = m$month, value = m$lcl, series = LCL_LAB)
  ) |>
    filter(!is.na(value)) |>
    mutate(series = factor(series, levels = names(chart_colours)))

  # Keep the y-axis tight around the data so small differences stay visible
  ax <- axis_limits(c(lines$value, target), cfg$axis_pad %||% 0.2, cfg$axis_pad_mode %||% "range")
  lims <- ax$lims

  pts <- filter(m, !is.na(rate))

  g <- ggplot(lines, aes(month, value, colour = series)) +
    {
      lim_lines <- filter(lines, series %in% c(UCL_LAB, LCL_LAB))
      if (cfg$limits_method %in% c("p", "laney"))
        geom_step(data = lim_lines, direction = "mid", linewidth = 0.47)
      else geom_line(data = lim_lines, linewidth = 0.47)
    } +
    geom_line(data = filter(lines, series %in% c("Target", "Average")), linewidth = 0.47) +
    geom_line(data = filter(lines, series == "Compliance (%)"), linewidth = 0.59) +   # = 1.25 pt
    geom_point(data = filter(lines, series == "Compliance (%)"), size = 2.4) +
    geom_point(data = filter(pts, signal), aes(month, rate), inherit.aes = FALSE,
               colour = "#C00000", size = 3.4, shape = 21, stroke = 1.3, fill = NA) +
    scale_x_continuous(breaks = 1:12, labels = MONTHS, limits = c(0.7, 12.3), expand = c(0, 0)) +
    scale_y_continuous(limits = lims, breaks = seq(lims[1], lims[2], by = ax$step), expand = c(0, 0)) +
    scale_colour_manual(values = chart_colours, drop = FALSE) +
    labs(title = title, x = NULL, y = NULL, colour = NULL) +
    theme_minimal(base_size = 11, base_family = cfg$chart_font) +
    theme(
      plot.title = element_text(hjust = 0.5, colour = "#595959", size = 13),
      panel.grid.major.x = element_blank(),
      panel.grid.minor = element_blank(),
      panel.grid.major.y = element_line(colour = "#D9D9D9", linewidth = 0.4),
      axis.text = element_text(colour = "#595959", size = 9),
      legend.position = "bottom",
      legend.text = element_text(colour = "#595959", size = 9),
      legend.key.width = unit(1.2, "lines"),
      plot.margin = margin(4, 10, 2, 4)
    ) +
    guides(colour = guide_legend(nrow = 1, override.aes = list(linewidth = 1)))

  ggsave(file, g, width = cfg$chart_size[1], height = cfg$chart_size[2], units = "in")
  file
}

# ---- Slide records for Typst -----------------------------------------------

as_target <- function(x) {
  v <- suppressWarnings(as.numeric(gsub("%", "", as.character(x))))
  if (!is.na(v) && v <= 1 && grepl("^0?\\.", as.character(x))) v <- v * 100  # 0.95 -> 95
  v
}

slide_record <- function(ind, st, chart_path, cfg) {
  d <- cfg$rate_digits
  target <- as_target(ind$target)
  m <- st$months
  analysis <- split_lines(ind$analysis)
  auto <- length(analysis) == 0 && isTRUE(cfg$auto_analysis)
  if (auto) analysis <- draft_analysis(st, target, cfg)
  q <- st$quarters

  list(
    title            = ind$title,
    measure          = ind$measure %||% ind$title,
    target           = if (is.na(target)) "" else paste0(fmt_num(target, if (target %% 1 == 0) 0 else 1), "%"),
    schedule         = ind$reporting_schedule %||% "Monthly",
    definition       = ind$definition %||% "",
    numerator_text   = ind$numerator %||% "",
    denominator_text = ind$denominator %||% "",
    months = map(1:12, \(i) list(
      label = MONTHS[i],
      num   = fmt_int(m$num[i]),
      den   = fmt_int(m$den[i]),
      rate  = fmt_num(m$rate[i], d),
      below = isTRUE(!is.na(target) && !is.na(m$rate[i]) && round(m$rate[i], d) < target)
    )),
    quarters = map(1:4, \(i) list(
      label = paste0("Q", i),
      rate  = fmt_num(q$rate[q$q == i], d) %||% "",
      below = isTRUE(!is.na(target) && length(q$rate[q$q == i]) && !is.na(q$rate[q$q == i]) &&
                     round(q$rate[q$q == i], d) < target)
    )),
    analysis       = as.list(analysis),
    analysis_auto  = auto,
    action         = as.list(split_lines(ind$action)),
    data_provider  = ind$data_provider %||% "",
    responsible    = ind$responsible %||% "",
    chart          = chart_path
  )
}

# ---- Typst ------------------------------------------------------------------

find_typst <- function() {
  if (nzchar(Sys.which("typst")))  return(list(bin = "typst",  pre = character(0)))
  if (nzchar(Sys.which("quarto"))) return(list(bin = "quarto", pre = "typst"))  # Quarto bundles Typst
  stop("Typst not found. Install it (Windows: winget install --id Typst.Typst) or install Quarto.")
}

typst_compile <- function(template, output, root, data_rel) {
  t <- find_typst()
  q <- function(x) shQuote(x, type = if (.Platform$OS.type == "windows") "cmd" else "sh")
  args <- c(t$pre, "compile", "--root", q(root), "--input", q(paste0("data=", data_rel)),
            q(template), q(output))
  res <- suppressWarnings(system2(t$bin, args, stdout = TRUE, stderr = TRUE))
  status <- attr(res, "status") %||% 0
  if (status != 0) stop("Typst failed:\n", paste(res, collapse = "\n"))
  invisible(output)
}

# ---- Main entry point -------------------------------------------------------

build_deck <- function(cfg) {
  root <- normalizePath(cfg$project_dir, winslash = "/")
  old <- setwd(root); on.exit(setwd(old))

  build_dir <- file.path("build"); dir.create(file.path(build_dir, "charts"), recursive = TRUE, showWarnings = FALSE)
  dir.create(cfg$out_dir, recursive = TRUE, showWarnings = FALSE)

  meta <- read_indicators(cfg)
  df <- if (cfg$input_type == "raw") read_raw(cfg) else read_aggregate(cfg)

  # Columns treated as criteria for "all"/"pooled" when no explicit list is given
  criteria_cols <- meta |> filter(tolower(coalesce(rule, "column")) == "column") |>
    mutate(col = coalesce(as.character(source_column), indicator)) |>
    pull(col)

  groups <- if (blank(cfg$group_col)) list(All = df) else {
    if (!cfg$group_col %in% names(df)) stop("Group column '", cfg$group_col, "' not found.")
    g <- split(df, df[[cfg$group_col]])
    if (isTRUE(cfg$include_overall_deck)) c(list(All = df), g) else g
  }

  outputs_wanted <- tolower(cfg$output %||% "pptx")
  if ("pptx" %in% outputs_wanted) source(file.path(root, "R", "kpi_pptx.R"))
  if ("pdf" %in% outputs_wanted) suppressPackageStartupMessages({ library(ggplot2); library(jsonlite) })

  outputs <- unlist(imap(groups, function(gdf, gname) {
    tag <- if (gname == "All") "" else paste0("_", slug(gname))
    items <- map(seq_len(nrow(meta)), function(i) {
      ind <- as.list(meta[i, ])
      counts <- if (cfg$input_type == "raw") counts_from_raw(gdf, ind, criteria_cols, cfg)
                else counts_from_aggregate(gdf, ind)
      st <- compute_indicator(counts, cfg)
      chart_title <- ind$chart_title %||% paste("Compliance to", ind$title)
      chart_file <- file.path(build_dir, "charts", paste0(slug(ind$indicator), tag, ".svg"))
      if ("pdf" %in% outputs_wanted) make_chart(st, chart_title, as_target(ind$target), chart_file, cfg)
      list(rec = slide_record(ind, st, paste0("/", chart_file), cfg), st = st,
           chart_title = chart_title, target = as_target(ind$target))
    })

    months_with_data <- sort(unique(gdf$.month))
    period <- if (!blank(cfg$period)) cfg$period else if (length(months_with_data))
      period_label(min(months_with_data), max(months_with_data), cfg$year %||% "")
      else as.character(cfg$year %||% "")

    deck_meta <- list(
      deck_title = cfg$deck_title,
      subtitle   = if (gname == "All") cfg$subtitle %||% "" else gname,
      period     = trimws(period),
      facility   = cfg$facility,
      footer     = cfg$footer,
      logo       = paste0("/", cfg$logo),
      full_logo  = paste0("/", cfg$full_logo %||% "assets/logo.svg"),
      prepared   = cfg$prepared_by %||% "",
      indicator_list = isTRUE(cfg$indicator_list %||% TRUE),
      title_slide = isTRUE(cfg$title_slide)
    )
    stem <- file.path(cfg$out_dir, paste0(slug(cfg$deck_title), tag, if (!is.null(cfg$year)) paste0("_", cfg$year)))
    made <- character(0)

    if ("pdf" %in% outputs_wanted) {
      deck <- list(meta = deck_meta, slides = map(items, "rec"))
      json_file <- file.path(build_dir, paste0("deck", tag, ".json"))
      js <- toJSON(deck, auto_unbox = TRUE, pretty = TRUE, na = "string")
      con <- file(json_file, open = "wb"); writeBin(charToRaw(enc2utf8(as.character(js))), con); close(con)
      typst_compile(cfg$template, paste0(stem, ".pdf"), root, paste0("/", json_file))
      made <- c(made, paste0(stem, ".pdf"))
    }
    if ("pptx" %in% outputs_wanted) {
      build_pptx(items, deck_meta, paste0(stem, ".pptx"), cfg)
      made <- c(made, paste0(stem, ".pptx"))
    }
    message("Created ", paste(made, collapse = ", "))
    made
  }), use.names = FALSE)
  invisible(outputs)
}

# =============================================================================
# Portable entry points
# =============================================================================

kpi_defaults <- function() list(
  project_dir    = KPI_HOME,
  input_type     = "auto",      # "raw", "aggregate" or "auto" (detects numerator/denominator columns)
  data_sheet     = NULL,        # NULL = first sheet that isn't "indicators"
  meta_sheet     = "indicators",
  date_col       = NULL,        # NULL = first date column (or a column called "month")
  year           = as.integer(format(Sys.Date(), "%Y")),
  yes_values     = c("Yes", "Y", "1", "True", "Compliant", "Met"),
  no_values      = c("No", "N", "0", "False", "Non-compliant", "Not met"),
  group_col      = NULL,
  include_overall_deck = TRUE,
  quarter_method = "pooled",
  limits_method  = "sd",        # mean +/- 3 SD, as in the Excel sheet ("p" / "laney" also available)
  cap_limits     = FALSE,       # TRUE = cut UCL/LCL off at 100% / 0%
  axis_pad       = 0.2,         # y-axis padding (20%)
  axis_pad_mode  = "range",     # "range" = 20% of the spread; "value" = highest x 1.2, lowest x 0.8
  rate_digits    = 1,
  auto_analysis  = TRUE,
  subtitle       = "",
  facility       = "Mji Mpya Referral Hospital, Dodoma",     # title slide box: name, location
  period         = NULL,        # title slide period; NULL = worked out from the data (e.g. "Q3 2026")
  indicator_list = TRUE,        # slide 2: list of the indicators in the deck
  footer         = "Mji Mpya Health Network",
  prepared_by    = "",
  title_slide    = TRUE,
  logo           = "assets/mark.svg",          # PDF footer (the PPTX logos live in the master)
  full_logo      = "assets/logo.svg",          # PDF title slide
  chart_font     = "sans",
  chart_size     = c(8.0, 3.35),
  output         = "pptx",      # "pptx", "pdf" or c("pptx", "pdf")
  template       = "templates/kpi-deck.typ",
  pptx_template  = file.path(KPI_HOME, "templates", "kpi_master.pptx")
)

#' Build a KPI deck from one dataset.
#'
#' @param data        Excel file with the data (raw records, or monthly numerator/denominator).
#' @param indicators  Excel file with the indicator sheet. Default: an "indicators" sheet
#'                    inside the data file, else <data name>_indicators.xlsx next to it.
#' @param title       Deck title (title slide + file name).
#' @param out_dir     Where to save. Default: an "output" folder next to the data file.
#' @param ...         Any setting from kpi_defaults(), e.g. date_col = "audit_date",
#'                    group_col = "unit", year = 2026, output = c("pptx", "pdf").
make_kpi_deck <- function(data, indicators = NULL, title = "KPI Report", out_dir = NULL, ...) {
  data <- normalizePath(data, winslash = "/", mustWork = TRUE)
  if (is.null(indicators)) {
    side <- sub("\\.xlsx?$", "_indicators.xlsx", data)
    indicators <- if ("indicators" %in% excel_sheets(data)) data else if (file.exists(side)) side else
      stop("No indicator sheet found. Create one with: make_indicator_sheet(\"", data, "\")")
  }
  indicators <- normalizePath(indicators, winslash = "/", mustWork = TRUE)
  cfg <- modifyList(kpi_defaults(), list(...))
  cfg$deck_title <- title
  cfg$data_file <- data
  cfg$meta_file <- indicators
  cfg$out_dir <- normalizePath(out_dir %||% file.path(dirname(data), "output"), winslash = "/", mustWork = FALSE)

  if (!cfg$meta_sheet %in% excel_sheets(indicators))
    stop("No sheet called '", cfg$meta_sheet, "' in ", basename(indicators),
         ". Create one with make_indicator_sheet(\"", basename(data), "\").")
  if (is.null(cfg$data_sheet)) {
    sh <- setdiff(excel_sheets(data), cfg$meta_sheet)
    cfg$data_sheet <- sh[1]
  }
  if (cfg$input_type == "auto") {
    cols <- tolower(names(read_excel(data, sheet = cfg$data_sheet, n_max = 0)))
    cfg$input_type <- if (all(c("indicator", "numerator", "denominator") %in% cols)) "aggregate" else "raw"
  }
  build_deck(cfg)
}

#' Start an indicator sheet from a data file: one row per Yes/No column, plus an
#' "Overall" row. Saved as <data name>_indicators.xlsx next to the data (the data file
#' itself is never changed). Fill in the text in Excel, then run make_kpi_deck().
make_indicator_sheet <- function(data, file = NULL, sheet = NULL, target = 100,
                                 yes_values = kpi_defaults()$yes_values, no_values = kpi_defaults()$no_values) {
  data <- normalizePath(data, winslash = "/", mustWork = TRUE)
  file <- file %||% sub("\\.xlsx?$", "_indicators.xlsx", data)
  if (file.exists(file)) stop(file, " already exists; delete it or pass another file name.")
  sheet <- sheet %||% setdiff(excel_sheets(data), "indicators")[1]
  df <- read_excel(data, sheet = sheet)
  vals <- tolower(c(yes_values, no_values, "n/a", "na"))
  yn <- names(df)[vapply(df, function(x) {
    v <- tolower(trimws(as.character(x[!is.na(x)])))
    length(v) > 0 && mean(v %in% vals) > 0.9
  }, logical(1))]
  if (!length(yn)) stop("No Yes/No columns found in sheet '", sheet, "'.")

  ind <- data.frame(
    order = seq_len(length(yn) + 1), include = "Yes",
    indicator = c(yn, "Overall"), title = c(yn, "Overall Compliance"),
    measure = c(paste(yn, "compliance"), "Overall compliance"), target = target,
    reporting_schedule = "Monthly", definition = "", numerator = "", denominator = "",
    rule = c(rep("column", length(yn)), "all"), criteria = "", source_column = "",
    data_provider = "", responsible = "", analysis = "", action = "", chart_title = "",
    check.names = FALSE
  )
  writexl::write_xlsx(list(indicators = ind), file)
  message("Found ", length(yn), " Yes/No columns: ", paste(yn, collapse = ", "),
          "\nIndicator sheet written to ", file, "\nFill in the text columns, then run make_kpi_deck().")
  invisible(file)
}
