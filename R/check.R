#' Perform basic integrity checks on MRIO table
#'
#' @description Checks MRIO tables for anomalies. Optionally exports results in
#' an Excel file. Automatically detects MRIO version (62 vs 72 economies).
#'
#' List of checks performed:
#' 1. Non-numeric cells.
#' 1. Blank cells.
#' 1. Negative inputs.
#' 1. Negative consumption items.
#' 1. Negative gross value added (GVA).
#' 1. Negative direct purchases abroad by residents.
#' 1. Positive purchases on the domestic territory by non-residents.
#' 1. Discrepancy between row and column output.
#' 1. Negative sum of value added terms.
#' 1. Negative final sales.
#' 1. Negative total final sales.
#' 1. Negative total exports.
#' 1. GVA = 0 with inputs > 0.
#' 1. GVA = 0 with exports > 0.
#'
#' @param path Path to file or directory. Files must be in Excel and have "MRIO"
#'   (case insensitive) in the name. Function ignores all other files in the
#'   directory.
#' @param precision Number of decimal places to consider in checking zero
#'   values.
#' @param export If `TRUE`, saves results in an Excel file in the same directory
#'   as `path`.
#' @param limit Number of flagged entries to print. Set to `Inf` to print all
#'   entries (not recommended).
#'
#' @importFrom rlang .data
#'
#' @export
check <- function(path, precision = 10, export = TRUE, limit = 10) {

  # Set up
  savedir <- ifelse(file.info(path)$isdir, path, dirname(path))
  threshold <- as.numeric(glue::glue("1e{-precision}"))
  delay <- .1
  exceed_limit_message <- c(" " = " {.warn {nrow(hits)} detected, printing first {limit}:}")
  report_cols <- c("type", "cell", "row_entity", "col_entity", "value")
  cli::cli_div(theme = list(
    span.header = list(color = "cyan"),
    span.check = list(color = "darkgreen"),
    span.warn = list(color = "darkgoldenrod4"),
    span.cell = list(color = "gray65")
  ))
  headerstyle <- openxlsx::createStyle(
    fontColour = "white",
    fgFill = "#007DB7",
    textDecoration = "bold"
  )

  reports = openxlsx::createWorkbook()

  # Preamble
  timestamp <- print_timestamp()
  cli::cli_text("")
  cli::cli_rule()
  cli::cli_text("")
  cli::cli_text("Checking started at {timestamp}.")
  Sys.sleep(delay)
  cli::cli_bullets(c("i" = " `precision = {precision}`: threshold for zero set at {threshold}."))
  Sys.sleep(delay)
  cli::cli_bullets(c("i" = ifelse(
    export,
    ' `export = TRUE`: results will be saved in "{.header {savedir}}".',
    " `export = FALSE`: results will not be saved."
  )))
  Sys.sleep(delay)
  cli::cli_bullets(c("i" = " `limit = {limit}`: first {limit} flagged entries will be printed."))
  Sys.sleep(delay)

  # Get files
  if (file.info(path)$isdir) {
    files <- list.files(path, pattern = "^[^~].*(MRIO|mrio|Mrio).*(xls|xlsx)$")
  } else {
    files <- path
  }

  for (file in files) {

    report <- data.frame(
      type = c(),
      cell = c(),
      row_entity = c(),
      col_entity = c(),
      value = c()
    )

    # Load MRIO and check if OK -----------------------------------------------

    if (file.info(path)$isdir) file <- file.path(path, file)

    cli::cli_text("")
    cli::cli_rule(left = "Checks on {.header {basename(file)}}")
    cli::cli_progress_message("Reading file...")

    mrio <- readxl::read_excel(
        file,
        range = readxl::cell_limits(c(8, 4), c(NA, NA)),
        col_names = FALSE,
        progress = FALSE
      ) |>
      suppressMessages() |>
      suppressWarnings()

    N <- 35; f <- 5
    nrow <- which(mrio[, 1] == "r69")[1]
    G <- (nrow - 8) / N
    mrio <- mrio[1:nrow, 2:(1 + G*N + G*f + 1)] |> as.matrix()

    xrow <- mrio[nrow(mrio), 1:(G*N)]
    xcol <- mrio[1:(G*N), ncol(mrio)]
    diff <- xrow - xcol

    vasum <- colSums(mrio[(G*N + 2):(G*N + 7), 1:(G*N)])

    dpa <- mrio[G*N + 4, (G*N + 1):(G*N + 1 + G*f)]
    pnr <- mrio[G*N + 5, (G*N + 1):(G*N + 1 + G*f)]
    gva <- mrio[G*N + 6, 1:(G*N)]

    Y_big <- mrio[1:(G*N), (G*N + 1):(ncol(mrio) - 1)]
    Y <- Y_big %*% kronecker(diag(G), rep(1, f))
    y <- rowSums(Y)

    Z <- mrio[1:(G * N), 1:(G * N)]
    zuse <- colSums(Z)

    Zf <- Z %*% kronecker(diag(G), rep(1, N))
    Yf <- Y
    for (s in 1:G) {
      Zf[(s-1)*N + 1:N, s] <- 0
      Yf[(s-1)*N + 1:N, s] <- 0
    }
    E <- Zf + Yf
    e <- rowSums(E)

    test1 <- G %% 1 == 0
    test2 <- nrow(mrio) >= 2205
    test3 <- ncol(mrio) >= 2205
    if (any(!c(test1, test2, test3))) {
      cli::cli_bullets(c("x" = " This file appears to differ from the standard MRIO format."))
      next
    }

    cli::cli_progress_done()

    # Cells -------------------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Cells have expected values?}")
    Sys.sleep(delay)

    # [*] Non-numeric cells ----

    hits <- data.frame(row = c(), col = c(), value = c())

    for (col in 1:ncol(mrio)) {
      isblank <- is.na(mrio[, col])
      ischar <- is.na(as.numeric(mrio[, col]) |> suppressWarnings())
      hit <- which(!isblank & ischar)
      if (length(hit) > 0) {
        hits <- rbind(hits, data.frame(row = hit, col = col, value = mrio[hit, col]))
      }}

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check All cells are numeric.}")
      Sys.sleep(delay)

    } else {

      cli::cli_bullets(c("!" = " {.warn Non-numeric cells}"))
      Sys.sleep(delay)

      hits$type <- "Blank cells"
      hits$cell <- paste0(excelcol(hits$col + 4), 7 + hits$row)
      hits$row_entity <- NA
      hits$col_entity <- NA

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{.cell {hits$cell[hit]}}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))

      cli::cli_bullets(c("x" = " Presence of non-numeric cells makes other checks unreliable. Aborting."))
      next
    }

    # [*] Blank cells ----

    hits <- data.frame(row = c(), col = c(), value = c())

    for (col in 1:ncol(mrio)) {
      hit <- which(is.na(mrio[, col]))
      if (length(hit) > 0) {
        hits <- rbind(hits, data.frame(row = hit, col = col, value = mrio[hit, col]))
      }}

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No blank cells.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Blank cells}"))
      Sys.sleep(delay)

      hits$type <- "Blank cells"
      hits$cell <- paste0(excelcol(hits$col + 4), 7 + hits$row)
      hits$row_entity <- NA
      hits$col_entity <- NA

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{.cell {hits$cell[hit]}}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] Negative inputs ----

    hits <- data.frame(row = c(), col = c(), value = c())

    for (col in 1:(G*N)) {
      hit <- which((mrio[1:(G*N), col] < -threshold))
      if (length(hit) > 0) {
        hits <- rbind(hits, data.frame(row = hit, col = col, value = mrio[hit, col]))
      }}

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative inputs.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Negative inputs}"))
      Sys.sleep(delay)

      hits$type <- "Negative inputs"
      hits$cell <- paste0(excelcol(hits$col + 4), 7 + hits$row)
      hits$row_entity <- mrio_entity(
        s = hits$row %/% N + 1,
        i = ifelse(hits$row %% N == 0, N, hits$row %% N),
        g = G
      )
      hits$col_entity <- mrio_entity(
        s = hits$col %/% N + 1,
        i = ifelse(hits$row %% N == 0, N, hits$row %% N),
        g = G
      )

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{.cell {hits$cell[hit]}} {hits$row_entity[hit]} => {hits$col_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] Negative consumption ----

    hits <- data.frame(row = c(), col = c(), value = c())
    col_index <- G*N + subset(1:(G*f), 1:(G*f) %% 5 != 0)
    for (col in col_index) {
      hit <- which((mrio[1:(G*N), col] < -threshold))
      if (length(hit) > 0) {
        hits <- rbind(hits, data.frame(row = hit, col = col - G*N, value = mrio[hit, col]))
      }}

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative consumption items.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Negative consumption items}"))
      Sys.sleep(delay)

      hits$"type" <- "Negative consumption"
      hits$cell <- paste0(excelcol(hits$col + G*N + 4), 7 + hits$row)
      hits$row_entity <- mrio_entity(
        s = hits$row %/% N + 1,
        i = ifelse(hits$row %% N == 0, N, hits$row %% N),
        g = G
      )
      hits$col_entity <- mrio_entity(
        s = hits$col %/% f + 1,
        f = hits$col %% f,
        g = G
      )

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{.cell {hits$cell[hit]}} {hits$row_entity[hit]} => {hits$col_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] Negative GVA ----

    hits <- data.frame(
      col = which(gva < -threshold),
      value = gva[which(gva < -threshold)]
    )

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative GVA.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Negative GVA}"))
      Sys.sleep(delay)

      hits$"type" <- "Negative GVA"
      hits$cell <- paste0(excelcol(hits$col + 4), 7 + G*N + 6)
      hits$row_entity <- NA
      hits$col_entity <- mrio_entity(
        s = hits$col %/% N + 1,
        i = ifelse(hits$col %% N == 0, N, hits$col %% N),
        g = G
      )

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{.cell {hits$cell[hit]}} {hits$col_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] Negative direct purchases abroad ----

    hits <- data.frame(
      col = which(dpa < -threshold),
      value = dpa[which(dpa < -threshold)]
    )

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative direct purchases abroad.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Negative direct purchases abroad}"))
      Sys.sleep(delay)

      hits$"type" <- "Negative direct purchases abroad"
      hits$cell <- paste0(excelcol(hits$col + 4), 7 + G*N + 4)
      hits$row_entity <- NA
      hits$col_entity <- mrio_entity(
        s = hits$col %/% N + 1,
        i = ifelse(hits$col %% N == 0, N, hits$col %% N),
        g = G
      )

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{.cell {hits$cell[hit]}} {hits$col_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] Positive purchases by non-residents ----

    hits <- data.frame(
      col = which(pnr > threshold),
      value = pnr[which(pnr > threshold)]
    )

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No positive purchases by non-residents.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Positive purchases by non-residents}"))
      Sys.sleep(delay)

      hits$"type" <- "Positive purchases by non-residents"
      hits$cell <- paste0(excelcol(hits$col + 4), 7 + G*N + 5)
      hits$row_entity <- NA
      hits$col_entity <- mrio_entity(
        s = hits$col %/% N + 1,
        i = ifelse(hits$col %% N == 0, N, hits$col %% N),
        g = G
      )

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{.cell {hits$cell[hit]}} {hits$col_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # Balanced table ----------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Table is balanced?}")
    Sys.sleep(delay)

    hits <- data.frame(
      row = which(abs(diff) > threshold),
      col = which(abs(diff) > threshold),
      value = abs(diff[which(abs(diff) > threshold)])
    )

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No row/column discrepancies in outputs.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Unbalanced row/column outputs}"))
      Sys.sleep(delay)

      hits$type <- "Unbalanced outputs"
      excel_cell_row <- paste0(excelcol(hits$col + 4), 7 + nrow(mrio))
      excel_cell_col <- paste0(excelcol(4 + ncol(mrio)), 7 + hits$row)
      hits$cell <- paste0(excel_cell_row, ", ", excel_cell_col)
      hits$row_entity <- mrio_entity(
        s = hits$row %/% N + 1,
        i = ifelse(hits$row %% N == 0, N, hits$row %% N),
        g = G
      )
      hits$col_entity <- mrio_entity(
        s = hits$col %/% N + 1,
        i = ifelse(hits$row %% N == 0, N, hits$row %% N),
        g = G
      )

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{.cell {hits$cell[hit]}} {hits$row_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # Aggregates --------------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Aggregates have expected values?}")
    Sys.sleep(delay)

    # [*] Negative sum of GVA terms ----

    hits <- data.frame(
      col = which(vasum < -threshold),
      value = vasum[which(vasum < -threshold)]
    )

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative sum of GVA terms.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Negative sum of GVA terms}"))
      Sys.sleep(delay)

      hits$"type" <- "Negative sum of GVA terms"
      hits$cell <- NA
      hits$row_entity <- NA
      hits$col_entity <- mrio_entity(
        s = hits$col %/% N + 1,
        i = ifelse(hits$col %% N == 0, N, hits$col %% N),
        g = G
      )

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{hits$col_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] Negative final sales ----

    hits <- data.frame(row = c(), col = c(), value = c())
    for (col in 1:G) {
      hit <- which((Y[, col] < -threshold))
      if (length(hit) > 0) hits <- rbind(hits, data.frame(row = hit, col = col, value = Y[hit, col]))
    }

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative final sales.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Negative final sales}"))
      Sys.sleep(delay)

      hits$type <- "Negative final sales"
      hits$cell <- NA
      hits$row_entity <- mrio_entity(
        s = hits$row %/% N + 1,
        i = ifelse(hits$row %% N == 0, N, hits$row %% N),
        g = G
      )
      hits$col_entity <- mrio_entity(s = hits$col, g = G)

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{hits$row_entity[hit]} => {hits$col_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] Negative total final sales ----

    hits <- data.frame(
      row = which(y < -threshold),
      value = y[which(y < -threshold)]
    )

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative total final sales.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Negative total final sales}"))
      Sys.sleep(delay)

      hits$type <- "Negative total final sales"
      hits$cell <- NA
      hits$row_entity <- mrio_entity(
        s = hits$row %/% N + 1,
        i = ifelse(hits$row %% N == 0, N, hits$row %% N),
        g = G
      )
      hits$col_entity <- NA

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{hits$row_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] Negative total exports ----

    hits <- data.frame(row = c(), col = c(), value = c())
    for (col in 1:G) {
      hit <- which((E[, col] < -threshold))
      if (length(hit) > 0) hits <- rbind(hits, data.frame(row = hit, col = col, value = E[hit, col]))
    }

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative total exports.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn Negative total exports}"))
      Sys.sleep(delay)

      hits$type <- "Negative exports"
      hits$cell <- NA
      hits$row_entity <- mrio_entity(
        s = hits$row %/% N + 1,
        i = ifelse(hits$row %% N == 0, N, hits$row %% N),
        g = G
      )
      hits$col_entity <- mrio_entity(s = hits$col, g = G)

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:min(c(limit, nrow(hits)))) {
        cli::cli_bullets(c(" " = "{hits$row_entity[hit]} => {hits$col_entity[hit]}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] GVA = 0, inputs > 0 ----

    hits <- data.frame(
      col = which(abs(gva) < threshold & zuse > threshold),
      value = zuse[which(abs(gva) < threshold & zuse > threshold)]
    )

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No GVA = 0 with inputs > 0.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn GVA = 0 but inputs > 0}"))
      Sys.sleep(delay)

      hits$type <- "GVA = 0 but inputs > 0"
      hits$cell <- NA
      hits$row_entity <- NA
      hits$col_entity <- mrio_entity(s = hits$col, g = G)

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:nrow(hits)) {
        cli::cli_bullets(c(" " = "{hits$col_entity[hit]} total inputs: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    # [*] GVA = 0, exports > 0 ----

    hits <- data.frame(
      col = which(abs(gva) < threshold & e > threshold),
      value = e[which(abs(gva) < threshold & e > threshold)]
    )

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No GVA = 0 with exports > 0.}")
      Sys.sleep(delay)

    } else {
      cli::cli_bullets(c("!" = " {.warn GVA = 0 but exports > 0}"))
      Sys.sleep(delay)

      hits$type <- "GVA = 0 but exports > 0"
      hits$cell <- NA
      hits$row_entity <- NA
      hits$col_entity <- mrio_entity(s = hits$col, g = G)

      if (nrow(hits) > limit) {
        cli::cli_bullets(exceed_limit_message)
        Sys.sleep(delay)
      }
      for (hit in 1:nrow(hits)) {
        cli::cli_bullets(c(" " = "{hits$col_entity[hit]} total exports: {hits$value[hit]}"))
        Sys.sleep(delay)
      }
      report <- rbind(report, subset(hits, select = report_cols))
    }

    if (nrow(report) > 0) {
      sheetname <- sub("\\.\\w+$", "", basename(file)) |> substr(1, 31)
      openxlsx::addWorksheet(reports, sheetname)
      openxlsx::writeData(reports, sheetname, report, headerStyle = headerstyle)
    }
  }

  cli::cli_text("")
  cli::cli_rule("")
  cli::cli_text("")
  Sys.sleep(delay)

  # Save report in Excel file -----------------------------------------------

  if (export & length(openxlsx::sheets(reports)) == 0) {
    cli::cli_bullets(c("x" = ' No report saved..'))
    cli::cli_text("")
    Sys.sleep(delay)
  }
  if (export & length(openxlsx::sheets(reports)) > 0) {
    filename <- paste0("Report ", gsub(":", "", timestamp), ".xlsx")
    openxlsx::saveWorkbook(reports, file.path(savedir, filename), overwrite = TRUE)
    cli::cli_text('Report successfully saved at "{.header {savedir}}".')
    cli::cli_text("")
    Sys.sleep(delay)
  }
}

# Helper functions --------------------------------------------------------

excelcol <- function(n) {

  N <- length(LETTERS); cell <- c()

  for (m in n) {
    j <- m %/% N
    if (j == 0) {
      cell <- c(cell, LETTERS[m])
    } else {
      cell <- c(cell, paste0(Recall(j), LETTERS[m %% N]))
  }}

  return(cell)
}

mrio_entity <- function(s, i = NULL, f = NULL, g) {

  if (g > 63) {
    dict_countries1 <- dict_countries |>
      dplyr::distinct(.data$mrio_ind, .keep_all = TRUE) |>
      dplyr::filter(!is.na(.data$mrio_ind)) |>
      dplyr::rename(ind = "mrio_ind")
  } else {
    dict_countries1 <- dict_countries |>
      dplyr::distinct(.data$mrio_ind62, .keep_all = TRUE) |>
      dplyr::filter(!is.na(.data$mrio_ind62)) |>
      dplyr::rename(ind = "mrio_ind62")
  }

  entity <- c()
  for (j in s) entity <- c(entity, dict_countries1$code[dict_countries1$ind == j])
  if (!is.null(i)) entity <- paste0(entity, "_", paste0("c", i))
  if (!is.null(f)) entity <- paste0(entity, "_", paste0("f", f))

  return(entity)
}

print_timestamp <- function() format(Sys.time(), "%Y-%m-%d %R")
