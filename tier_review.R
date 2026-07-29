# =============================================================================
# tier_review.R
# Reusable "tier block" for the Course Schedule processing review artifact.
#
# A tier = a title, a question, and a list of checks. Each check is a DECISION
# (yardstick + recorded call) and may own EVIDENCE (small summary tables the
# analyst eyeballs to make that call), rendered right beneath it.
#
# Design tenets:
#   - Flag, don't drop      : a check reports a STATUS; never deletes, never halts.
#   - Non-blocking          : a missing/failed input (or unloaded data) renders
#                             as a flag / "data not loaded", not a crash.
#   - Auto-flag + override  : a computed flag can carry an analyst note.
#   - Config is the template: district knobs live in the YAML config.
#
# Decision lines render as compact HTML blocks; evidence renders as gt tables.
# render_tier() returns an htmltools tagList — return it from a Quarto chunk.
# =============================================================================

suppressPackageStartupMessages({
  library(gt)
  library(glue)
  library(htmltools)
})

# ---- Status vocabulary ------------------------------------------------------
tr_status_levels <- c("ok", "flag", "missing", "todo", "na")

tr_status_meta <- list(
  ok      = list(label = "OK",        icon = "\u2713"),
  flag    = list(label = "Flag",      icon = "\u26A0"),
  missing = list(label = "Missing",   icon = "\u2717"),
  todo    = list(label = "To decide", icon = "\u270E"),
  na      = list(label = "N/A",       icon = "\u2014")
)

tr_badge_colors <- list(
  ok      = list(bg = "#ecfdf3", fg = "#067647"),
  flag    = list(bg = "#fffaeb", fg = "#b54708"),
  missing = list(bg = "#fef3f2", fg = "#b42318"),
  todo    = list(bg = "#eff8ff", fg = "#175cd3"),
  na      = list(bg = "#f2f4f7", fg = "#667085")
)

# ---- Safe evaluation (this is what makes checks + evidence "non-blocking") --
tr_safe <- function(expr, otherwise = NA) {
  tryCatch(expr, error = function(e) otherwise)
}

# ---- Status + rate helpers --------------------------------------------------
tr_threshold_status <- function(value, threshold, ok_when = c("below", "above")) {
  ok_when <- match.arg(ok_when)
  if (is.null(value) || length(value) != 1 || is.na(value)) return("missing")
  ok <- if (ok_when == "below") value <= threshold else value >= threshold
  if (isTRUE(ok)) "ok" else "flag"
}

tr_presence_status <- function(present, partial = FALSE) {
  if (isTRUE(partial)) return("flag")
  if (isTRUE(present)) "ok" else "missing"
}

# % of a flag column that is "true", robust to Y/N, 1/0, TRUE/FALSE encodings.
tr_pct_true <- function(x) {
  round(100 * mean(x %in% c(1, "1", "Y", "y", "TRUE", TRUE), na.rm = TRUE), 1)
}

# ---- Constructors -----------------------------------------------------------
# evidence: NULL, OR a function (thunk) returning a data.frame, OR a named list
#           of data.frames. Deferred so it evaluates at render time (and can be
#           tr_safe()'d). Each frame renders as its own captioned gt table.
tr_check <- function(id, question,
                     good     = NA_character_,
                     value    = NULL,
                     status   = "todo",
                     records  = NA_character_,
                     note     = NA_character_,
                     evidence = NULL) {
  if (!status %in% tr_status_levels) {
    stop(glue("Unknown status '{status}'. Use one of: ",
              paste(tr_status_levels, collapse = ", ")))
  }
  structure(
    list(id = id, question = question, good = good, value = value,
         status = status, records = records, note = note, evidence = evidence),
    class = "tr_check"
  )
}

tr_tier <- function(number, title, question, checks, enabled = TRUE) {
  structure(
    list(number = number, title = title, question = question,
         checks = checks, enabled = enabled),
    class = "tr_tier"
  )
}

# ---- Formatters -------------------------------------------------------------
tr_fmt_value <- function(v) {
  if (is.null(v) || (length(v) == 1 && is.na(v))) return("\u2014")
  if (is.numeric(v) && length(v) == 1) return(format(round(v, 2), big.mark = ","))
  paste(as.character(v), collapse = ", ")
}

.tr_lbl <- "font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:.04em;"

tr_decision_block <- function(chk, disabled = FALSE) {
  status <- if (disabled) "na" else chk$status
  col <- tr_badge_colors[[status]]; meta <- tr_status_meta[[status]]

  badge <- tags$span(
    style = paste0("display:inline-block;background:", col$bg, ";color:", col$fg,
                   ";font-weight:500;font-size:11px;padding:2px 8px;border-radius:999px;",
                   "white-space:nowrap;flex:none;"),
    paste(meta$icon, meta$label))

  value <- if (!is.null(chk$value) && !(length(chk$value) == 1 && is.na(chk$value)))
    tags$span(style = "color:#475467;font-size:12px;", paste0("  \u2014  ", tr_fmt_value(chk$value)))
  else NULL

  good <- if (!is.na(chk$good) && nzchar(chk$good))
    tags$div(style = "color:#667085;font-size:12px;margin-top:3px;",
             tags$span(style = paste0(.tr_lbl, "color:#98a2b3;"), "good \u00B7 "), chk$good)
  else NULL

  recorded <- tags$div(style = "font-size:13px;margin-top:4px;color:#1d2939;",
    tags$span(style = paste0(.tr_lbl, "color:#667085;"), "recorded \u00B7 "),
    tr_fmt_value(chk$records))

  note <- if (!is.na(chk$note) && nzchar(chk$note))
    tags$div(style = paste0("margin-top:5px;background:#fff7ed;border-left:3px solid #f79009;",
                            "padding:4px 8px;color:#7a4708;font-size:12px;"),
             tags$span(style = "font-weight:600;", "note \u00B7 "), chk$note)
  else NULL

  tags$div(style = "padding:11px 0;border-bottom:1px solid #eef1f4;",
    tags$div(style = "display:flex;align-items:baseline;gap:8px;",
      tags$span(style = "color:#98a2b3;font-size:12px;min-width:28px;flex:none;", chk$id),
      tags$span(style = "font-weight:500;color:#1d2939;flex:1;", chk$question, value),
      badge),
    good, recorded, note)
}

.tr_evidence_gt <- function(df) {
  gt(df) |>
    tab_options(table.font.size = px(12), data_row.padding = px(4),
                column_labels.font.weight = "bold", table.width = pct(100),
                table.border.top.style = "none") |>
    opt_table_font(font = "system-ui")
}

tr_render_evidence <- function(chk) {
  if (is.null(chk$evidence)) return(NULL)
  res <- tr_safe(chk$evidence(), otherwise = NULL)

  # normalize to a named list of data.frames
  if (is.data.frame(res)) res <- setNames(list(res), glue("Evidence \u00B7 {chk$id}"))

  if (is.null(res) || !is.list(res) || length(res) == 0) {
    return(tags$div(style = "margin:2px 0 10px 36px;color:#98a2b3;font-size:12px;font-style:italic;",
                    glue("Evidence \u00B7 {chk$id} \u2014 data not loaded")))
  }

  blocks <- lapply(seq_along(res), function(i) {
    cap <- names(res)[i]; df <- res[[i]]
    if (is.null(df) || !is.data.frame(df) || nrow(df) == 0) {
      return(tags$div(style = "color:#98a2b3;font-size:12px;font-style:italic;margin-bottom:8px;",
                      glue("{cap} \u2014 no rows")))
    }
    tags$div(style = "margin-bottom:12px;",
      tags$div(style = paste0(.tr_lbl, "color:#667085;margin-bottom:4px;"), cap),
      HTML(as_raw_html(.tr_evidence_gt(df))))
  })

  tags$div(style = "margin:8px 0 10px 36px;", blocks)
}

# ---- The one renderer -------------------------------------------------------
render_tier <- function(tier) {
  disabled <- !isTRUE(tier$enabled)

  st <- vapply(tier$checks, function(c) if (disabled) "na" else c$status, character(1))
  counts <- table(factor(st, levels = tr_status_levels))
  labs <- vapply(tr_status_levels, function(s) {
    n <- as.integer(counts[[s]])
    if (n == 0) NA_character_ else glue("{n} {tr_status_meta[[s]]$label}")
  }, character(1))
  summ <- paste(labs[!is.na(labs)], collapse = "  \u00B7  ")

  q <- if (disabled) glue("{tier$question}  \u2014  tier not run (entered downstream)")
       else tier$question

  header <- tags$div(
    style = paste0("background:", if (disabled) "#8794a3" else "#0f2a43",
                   ";color:#fff;padding:12px 16px;"),
    tags$div(style = "display:flex;align-items:baseline;gap:8px;",
      tags$span(style = "font-weight:700;font-size:12px;background:rgba(255,255,255,.15);padding:2px 8px;border-radius:999px;",
                glue("Tier {tier$number}")),
      tags$span(style = "font-weight:600;font-size:16px;", tier$title)),
    tags$div(style = "margin-top:5px;font-size:13px;opacity:.9;font-style:italic;", q),
    tags$div(style = "margin-top:7px;font-size:12px;opacity:.95;",
             tags$span(style = "font-weight:600;", "Status: "), summ))

  body_blocks <- lapply(tier$checks, function(chk) {
    tagList(tr_decision_block(chk, disabled),
            if (!disabled) tr_render_evidence(chk) else NULL)
  })

  tags$section(
    style = "border:1px solid #e3e6ea;border-radius:10px;overflow:hidden;margin:1.4rem 0;font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;",
    header,
    tags$div(style = "padding:2px 16px 10px 16px;background:#fff;", body_blocks))
}
