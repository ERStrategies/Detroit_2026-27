# =============================================================================
# explore_helpers.R
# Config-FREE profiling for the Tier 0 "Explore" pass.
#
# Purpose: look at TRULY RAW data (original district headers, nothing mapped)
# so you can (a) answer the structural Tier 0 questions and (b) author the
# column_map in district_config.yml. This file depends on NO config and NO
# ERS D_ names on purpose — that's what breaks the chicken-and-egg problem
# where the review artifact needed the map that Tier 0 is supposed to produce.
#
# Nothing here makes a decision or assigns meaning. It only describes what's in
# the file. Meaning (which column is the term? is "Rotation" really the
# expression?) is a human/skill judgment made FROM this output.
# =============================================================================

suppressPackageStartupMessages({
  library(gt)
})

# ---- One row per column: type, distinctness, missingness, sample values -----
explore_profile <- function(df) {
  n <- nrow(df)
  miss <- vapply(df, function(x) {
    m <- is.na(x)
    if (is.character(x)) m <- m | trimws(x) %in% c("", "NA", "MISSING", "NULL")
    sum(m)
  }, integer(1))
  data.frame(
    column       = names(df),
    type         = vapply(df, function(x) class(x)[1], character(1)),
    n_distinct   = vapply(df, function(x) length(unique(x[!is.na(x)])), integer(1)),
    n_missing    = miss,
    pct_missing  = round(100 * miss / n, 1),
    sample_values = vapply(df, function(x) {
      u <- unique(x[!is.na(x)])
      paste(as.character(utils::head(u, 5)), collapse = ", ")
    }, character(1)),
    stringsAsFactors = FALSE, check.names = FALSE, row.names = NULL
  )
}

# ---- Full value set (with counts) for one column ----------------------------
explore_value_counts <- function(df, col, top = NULL) {
  x <- df[[col]]
  t <- as.data.frame(table(value = x, useNA = "ifany"), stringsAsFactors = FALSE)
  names(t) <- c(col, "n")
  t <- t[order(-t$n), , drop = FALSE]
  rownames(t) <- NULL
  if (!is.null(top)) t <- utils::head(t, top)
  t
}

# ---- Auto value-dumps for every low-cardinality column ----------------------
# Captures the categorical fields you map by eye: term, school, rotation, class
# format, grade, flags. High-cardinality columns (expression, course name) are
# skipped here — profile them with explore_value_counts(df, "col", top = 40).
explore_low_card <- function(df, max_distinct = 25) {
  nd <- vapply(df, function(x) length(unique(x[!is.na(x)])), integer(1))
  cols <- names(df)[nd > 1 & nd <= max_distinct]
  setNames(lapply(cols, function(c) explore_value_counts(df, c)), cols)
}

# ---- Shared table styling (same visual family as the review artifact) -------
explore_gt <- function(df, title = NULL) {
  g <- gt(df)
  if (!is.null(title)) g <- tab_header(g, title = md(paste0("**", title, "**")))
  g |>
    tab_options(table.font.size = px(12), data_row.padding = px(4),
                column_labels.font.weight = "bold", table.width = pct(100),
                heading.align = "left", heading.title.font.size = px(13),
                table.border.top.style = "none") |>
    opt_table_font(font = "system-ui")
}

# Render a named list of frames (e.g. explore_low_card output) as stacked tables.
explore_render <- function(named_frames, prefix = "Values \u00B7 ") {
  htmltools::tagList(lapply(names(named_frames), function(nm) {
    htmltools::HTML(as_raw_html(explore_gt(named_frames[[nm]], paste0(prefix, nm))))
  }))
}
