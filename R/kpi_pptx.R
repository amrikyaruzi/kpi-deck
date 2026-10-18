# =============================================================================
# PowerPoint output: fills templates/kpi_master.pptx by placeholder name.
# Table = flextable (native PowerPoint table); chart = mschart (native, editable chart).
# =============================================================================

suppressPackageStartupMessages({
  library(officer)
  library(flextable)
  library(mschart)
})

PPT_FONT <- "Arial"

ph <- function(label) ph_location_label(ph_label = label)

# Size (inches) of a named placeholder in the master
ph_box <- function(doc, layout, label) {
  lp <- layout_properties(doc, layout = layout)
  lp[lp$ph_label == label, c("offx", "offy", "cx", "cy")][1, ]
}

# Rough fit: shrink font until the text fits the box (PowerPoint won't auto-shrink
# text it didn't type itself until someone edits it)
fit_size <- function(items, width_in, height_in, base = 11, min = 8, indent_in = 0) {
  if (length(items) == 0) return(base)
  for (sz in seq(base, min, by = -0.5)) {
    chars_per_line <- (width_in - indent_in) * 72 / (sz * 0.5)
    lines <- sum(ceiling(pmax(nchar(items), 1) / chars_per_line))
    if (lines * sz * 1.2 / 72 + length(items) * 2 / 72 <= height_in) return(sz)
  }
  min
}

# Runs use a fully specified fp_text (some officer versions drop partial "lite" formats)
txt <- function(x, bold = FALSE, size = 11) {
  ftext(as.character(x), fp_text(bold = bold, font.size = size, font.family = PPT_FONT, color = "black"))
}

label_value <- function(label, value, size = 11) {
  fpar(txt(label, TRUE, size), txt(value, FALSE, size), fp_p = fp_par_lite())
}

# Lists: written as plain paragraphs so they inherit the numbering/bullets defined
# in the master; font size is then set directly on the runs.
ph_with_numbered <- function(doc, items, label, size) {
  doc <- ph_with(doc, items, location = ph(label))
  x <- doc$slide$get_slide(doc$cursor)$get()
  runs <- xml2::xml_find_all(x, sprintf(".//p:sp[p:nvSpPr/p:cNvPr[@name='%s']]//a:rPr", label), xml2::xml_ns(x))
  xml2::xml_set_attr(runs, "sz", as.character(round(size * 100)))
  doc
}

# Width (inches) of text in Arial/Helvetica, used to size the fraction line
text_width_in <- function(x, size = 11) {
  x <- x[nzchar(x)]
  if (!length(x)) return(0)
  grDevices::pdf(NULL); on.exit(grDevices::dev.off())
  graphics::par(family = "Helvetica", ps = size)
  max(graphics::strwidth(x, units = "inches"))
}

# ---- Monthly table ----------------------------------------------------------

make_table <- function(rec, box) {
  m <- rec$months; q <- rec$quarters
  df <- data.frame(Period = c("Numerator", "Denominator", "Rate, %", "Quarterly Rate, %"), check.names = FALSE)
  for (i in 1:12) {
    df[[MONTHS[i]]] <- c(m[[i]]$num, m[[i]]$den, m[[i]]$rate, if (i %% 3 == 1) q[[(i + 2) / 3]]$rate else "")
  }
  w_unit <- box$cx / (1.55 + 12)
  b <- fp_border(color = "black", width = 0.75)

  ft <- flextable(df)
  for (k in 1:4) ft <- merge_at(ft, i = 4, j = 1 + (3 * (k - 1) + 1):(3 * k))
  ft <- ft |>
    font(fontname = PPT_FONT, part = "all") |>
    fontsize(size = 11, part = "all") |>
    color(color = "black", part = "all") |>
    bg(bg = "#A6A6A6", part = "header") |>
    bold(bold = FALSE, part = "header") |>
    bold(i = 4, j = 2:13) |>
    align(align = "center", part = "all") |>
    align(j = 1, align = "left", part = "all") |>
    valign(valign = "center", part = "all") |>
    border_outer(border = b, part = "all") |>
    border_inner(border = b, part = "all") |>
    padding(padding.top = 0, padding.bottom = 0, padding.left = 5, padding.right = 5, part = "all") |>
    width(j = 1, width = 1.55 * w_unit) |>
    width(j = 2:13, width = w_unit) |>
    height_all(height = box$cy / 5, part = "all") |>
    hrule(rule = "exact", part = "all")

  below_m <- which(vapply(m, `[[`, logical(1), "below"))
  if (length(below_m)) ft <- ft |> color(i = 3, j = 1 + below_m, color = "#C00000") |> bold(i = 3, j = 1 + below_m)
  below_q <- which(vapply(q, `[[`, logical(1), "below"))
  if (length(below_q)) ft <- color(ft, i = 4, j = 2 + 3 * (below_q - 1), color = "#C00000")
  ft
}

# ---- Control chart (native PowerPoint line chart) ---------------------------

make_ms_chart <- function(st, title, target, cfg) {
  m <- st$months
  ser <- list(m$rate, rep(target, 12), if (cfg$limits_method == "sd") m$cl else ifelse(is.na(m$den), NA, m$cl),
              m$ucl, m$lcl)
  names(ser) <- names(chart_colours)
  cols  <- chart_colours
  width <- c(1.25, 1, 1, 1, 1)           # points: compliance line 1.25 pt, others 1 pt
  sym   <- c("circle", "none", "none", "none", "none")
  if (any(m$signal)) {
    ser[["Special cause"]] <- ifelse(m$signal, m$rate, NA)
    cols <- c(cols, "Special cause" = "#C00000")
    width <- c(width, 0); sym <- c(sym, "circle")
  }
  names(width) <- names(sym) <- names(ser)

  dat <- data.frame(
    month  = factor(rep(MONTHS, length(ser)), levels = MONTHS),
    series = factor(rep(names(ser), each = 12), levels = names(ser)),
    value  = unlist(ser, use.names = FALSE)
  )
  # Drop series that are entirely empty (e.g. limits with < 2 months of data)
  keep <- tapply(!is.na(dat$value), dat$series, any)
  dat <- droplevels(dat[dat$series %in% names(keep)[keep], ])
  lv <- levels(dat$series)

  ax <- axis_limits(c(dat$value, target), cfg$axis_pad %||% 0.2, cfg$axis_pad_mode %||% "range")
  lims <- ax$lims; step <- ax$step

  grey_txt <- fp_text(color = "#595959", font.size = 9, font.family = PPT_FONT)
  th <- mschart_theme(
    main_title = fp_text(color = "#595959", font.size = 13, font.family = PPT_FONT),
    axis_text_x = grey_txt, axis_text_y = grey_txt, legend_text = grey_txt,
    grid_major_line_y = fp_border(color = "#D9D9D9", width = 0.5),
    grid_major_line_x = fp_border(style = "none"),
    axis_ticks_x = fp_border(style = "none"), axis_ticks_y = fp_border(style = "none"),
    legend_position = "b"
  )
  line_style <- setNames(ifelse(lv == "Special cause", "none", "solid"), lv)

  ch <- ms_linechart(dat, x = "month", y = "value", group = "series") |>
    chart_settings(style = "lineMarker") |>
    chart_data_stroke(values = cols[lv]) |>
    chart_data_fill(values = cols[lv]) |>
    chart_data_line_width(values = width[lv]) |>
    chart_data_line_style(values = line_style) |>
    chart_data_symbol(values = sym[lv]) |>
    chart_data_size(values = setNames(ifelse(lv == "Special cause", 9, 6), lv)) |>
    chart_ax_y(limit_min = lims[1], limit_max = lims[2], major_unit = step, num_fmt = "0") |>
    chart_labels(title = title, xlab = "", ylab = "") |>
    set_theme(th)
  # Straight lines between months (newer mschart versions smooth by default)
  if (exists("chart_data_smooth", envir = asNamespace("mschart")))
    ch <- mschart::chart_data_smooth(ch, values = setNames(rep(0L, length(lv)), lv))
  ch
}

# ---- Deck -------------------------------------------------------------------

build_pptx <- function(items, deck_meta, file, cfg) {
  doc <- read_pptx(cfg$pptx_template)
  master <- layout_summary(doc)$master[1]

  has_ph <- function(layout, label) label %in% layout_properties(doc, layout)$ph_label

  # ---- Slide 1: title (branded title layout) ----
  if (isTRUE(deck_meta$title_slide)) {
    doc <- add_slide(doc, layout = "Brand Title", master = master)
    if (has_ph("Brand Title", "ph_facility")) {
      # "Mji Mpya Referral Hospital, Dodoma" -> MJI MPYA REFERRAL HOSPITAL (36 pt bold) / Dodoma (28 pt)
      parts <- trimws(strsplit(deck_meta$facility, ",")[[1]])
      centre <- fp_par(text.align = "center")
      paras <- list(fpar(txt(toupper(parts[1]), TRUE, 36), fp_p = centre))
      if (length(parts) > 1)
        paras <- c(paras, list(fpar(txt(paste(parts[-1], collapse = ", "), FALSE, 28), fp_p = centre)))
      doc <- ph_with(doc, do.call(block_list, paras), location = ph("ph_facility"))
    }
    doc <- ph_with(doc, deck_meta$deck_title, location = ph("ph_deck_title"))
    if (nzchar(deck_meta$subtitle) && has_ph("Brand Title", "ph_subtitle"))
      doc <- ph_with(doc, deck_meta$subtitle, location = ph("ph_subtitle"))
    doc <- ph_with(doc, deck_meta$period, location = ph("ph_period"))
    if (nzchar(deck_meta$prepared) && has_ph("Brand Title", "ph_prepared"))
      doc <- ph_with(doc, deck_meta$prepared, location = ph("ph_prepared"))
  }

  # ---- Slide 2: list of indicators in this deck ----
  if (isTRUE(cfg$indicator_list %||% TRUE) && "Indicator List" %in% layout_summary(doc)$layout) {
    titles <- vapply(items, function(it) it$rec$title, character(1))
    box <- ph_box(doc, "Indicator List", "ph_indicator_list")
    # 14 pt as in the template; smaller when there are too many indicators to fit
    size <- max(8, min(14, ((box$cy - 0.1) * 72 / length(titles) - 10) / 1.5))
    doc <- add_slide(doc, layout = "Indicator List", master = master)
    doc <- ph_with_numbered(doc, titles, "ph_indicator_list", floor(size * 2) / 2)
  }

  lay <- "KPI Indicator"
  tbl_box <- ph_box(doc, lay, "ph_table")
  an_box  <- ph_box(doc, lay, "ph_analysis")
  ac_box  <- ph_box(doc, lay, "ph_action")
  def_box <- ph_box(doc, lay, "ph_definition")
  ms_box  <- ph_box(doc, lay, "ph_measure")
  num_box <- ph_box(doc, lay, "ph_numerator")

  for (it in items) {
    s <- it$rec
    doc <- add_slide(doc, layout = lay, master = master)
    doc <- ph_with(doc, s$title, location = ph("ph_title"))
    doc <- ph_with(doc, label_value("Measure: ", s$measure, fit_size(paste("Measure:", s$measure), ms_box$cx, ms_box$cy)), location = ph("ph_measure"))
    doc <- ph_with(doc, label_value("Target: ", s$target), location = ph("ph_target"))
    doc <- ph_with(doc, label_value("Reporting schedule: ", s$schedule), location = ph("ph_schedule"))
    doc <- ph_with(doc, label_value("Definition: ", s$definition, fit_size(paste("Definition:", s$definition), def_box$cx, def_box$cy)), location = ph("ph_definition"))
    # Fraction: numerator (+ "x 100%"), a line shape under it, denominator below.
    # The block starts at the left edge of ph_numerator and is only as wide as its
    # longest text, so it sits right after the "Numerator:" / "Denominator:" labels.
    times <- if (isTRUE(l10n_info()[["UTF-8"]])) " \u00d7 100%" else " x 100%"
    num_text <- paste0(s$numerator_text, times)
    w <- min(num_box$cx, text_width_in(c(num_text, s$denominator_text)) + 0.12)
    doc <- ph_with(doc, fpar(txt(num_text), fp_p = fp_par(text.align = "center")), location = ph("ph_numerator"))
    doc <- ph_with(doc, s$denominator_text, location = ph("ph_denominator"))
    doc <- ph_with(doc, "", location = ph_location(
      left = num_box$offx, top = num_box$offy + num_box$cy + 0.02,
      width = w, height = 0, newlabel = "fraction_line",
      geom = "line", ln = sp_line(color = "black", lwd = 0.75)))

    x <- doc$slide$get_slide(doc$cursor)$get()
    ns <- xml2::xml_ns(x)
    find_sp <- function(name) xml2::xml_find_first(x, sprintf(".//p:sp[p:nvSpPr/p:cNvPr[@name='%s']]", name), ns)
    # officer adds the line as an empty placeholder; make it a plain line shape
    xml2::xml_remove(xml2::xml_find_all(find_sp("fraction_line"),
      "./p:nvSpPr/p:nvPr/p:ph | ./p:nvSpPr/p:cNvSpPr/a:spLocks | ./p:txBody", ns))
    # narrow the numerator/denominator boxes to the fraction width (left edge unchanged)
    for (nm in c("ph_numerator", "ph_denominator")) {
      ext <- xml2::xml_find_first(find_sp(nm), "./p:spPr/a:xfrm/a:ext", ns)
      if (!inherits(ext, "xml_missing")) xml2::xml_set_attr(ext, "cx", as.character(round(w * 914400)))
    }

    doc <- ph_with(doc, make_table(s, tbl_box), location = ph("ph_table"))

    an <- unlist(s$analysis)
    # An empty placeholder still goes on the slide, so PowerPoint shows "Click to add ..." there
    doc <- if (length(an)) ph_with_numbered(doc, an, "ph_analysis", fit_size(an, an_box$cx, an_box$cy, base = 10.5, indent_in = 0.25))
           else ph_with(doc, "", location = ph("ph_analysis"))
    if (isTRUE(s$analysis_auto)) {
      doc <- ph_with(doc, "(auto-drafted, review before circulating)", location = ph("ph_analysis_note"))
    }
    ac <- unlist(s$action)
    doc <- if (length(ac)) ph_with_numbered(doc, ac, "ph_action", fit_size(ac, ac_box$cx, ac_box$cy, base = 10.5, indent_in = 0.25))
           else ph_with(doc, "", location = ph("ph_action"))
    pv <- ph_box(doc, lay, "ph_provider")
    doc <- ph_with(doc, label_value("Data Provider: ", s$data_provider,
                                    fit_size(paste("Data Provider:", s$data_provider), pv$cx, pv$cy, base = 10.5)),
                   location = ph("ph_provider"))
    doc <- ph_with(doc, label_value("Responsible: ", s$responsible,
                                    fit_size(paste("Responsible:", s$responsible), pv$cx, pv$cy, base = 10.5)),
                   location = ph("ph_responsible"))

    chart <- make_ms_chart(it$st, it$chart_title, it$target, cfg)
    doc <- ph_with(doc, chart, location = ph("ph_chart"))
    doc <- ph_with(doc, deck_meta$footer, location = ph("ph_footer"))
  }
  print(doc, target = file)
  invisible(file)
}
