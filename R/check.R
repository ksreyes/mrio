#' Perform basic integrity checks on MRIO Excel files
#'
#' Given directory `path`, loads all Excel files with "MRIO" (case-insensitive)
#' in the file name and examines them for anomalies.
#'
#' The number of countries in the MRIO is automatically detected. However, note
#' that the MRIO table must follow the standard Excel format.
#'
#' @param path Path to file or directory of files.
#' @param precision Number of decimal places to consider in checking whether a
#'   table is balanced.
#' @param export If `TRUE`, saves results in an Excel file in the same directory
#'   as `path`.
#'
#' @importFrom rlang .data
#'
#' @export
check <- function(path, precision = 8, export = TRUE) {

  savedir <- ifelse(file.info(path)$isdir, path, dirname(path))
  threshold <- as.numeric(glue::glue("1e{-precision}"))
  delay <- .25
  exportcheck <- ifelse(
    export,
    " `export=TRUE`: results will be saved in {savedir}.",
    " `export=FALSE`: results will not be saved."
  )
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

  timestamp <- print_timestamp()
  cli::cli_text("")
  cli::cli_rule()
  cli::cli_text("")
  cli::cli_text("Checking started on {timestamp}.")
  Sys.sleep(delay)
  cli::cli_bullets(c("i" = " `precision = {precision}`: threshold for zero set at {threshold}."))
  Sys.sleep(delay)
  cli::cli_bullets(c("i" = ifelse(
    export,
    ' `export = TRUE`: results will be saved in "{.header {savedir}}".',
    " `export = FALSE`: results will not be saved."
  )))
  Sys.sleep(delay)

  reports = openxlsx::createWorkbook()

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
      suppressMessages()

    # Check if file is MRIO ---------------------------------------------------

    N <- 35; f <- 5
    nrow <- length(which(!is.na(mrio[, 1])))
    G <- (nrow - 8) / N
    mrio <- mrio[1:nrow, 2:(1 + G*N + G*f + 1)] |> as.matrix()

    test1 <- G %% 1 == 0
    test2 <- nrow(mrio) >= 2205
    test3 <- ncol(mrio) >= 2205
    test4 <- is.numeric(mrio)
    if (any(!c(test1, test2, test3, test4))) {
      cli::cli_bullets(c("x" = " This file appears to differ from the standard MRIO format."))
      next
    }

    xrow <- mrio[nrow(mrio), 1:(G*N)]
    xcol <- mrio[1:(G*N), ncol(mrio)]
    dpa <- mrio[G*N + 4, (G*N + 1):(G*N + 1 + G*f)]
    pnr <- mrio[G*N + 5, (G*N + 1):(G*N + 1 + G*f)]
    vabp <- mrio[G*N + 6, 1:(G*N)]

    Z <- mrio[1:(G * N), 1:(G * N)]
    zuse <- colSums(Z)
    va <- colSums(mrio[(G*N + 2):(G*N + 7), 1:(G*N)])
    Y_big <- mrio[1:(G*N), (G*N + 1):(ncol(mrio) - 1)]
    Y <- Y_big %*% kronecker(diag(G), rep(1, f))
    y <- rowSums(Y)

    Zf <- Z %*% kronecker(diag(G), rep(1, N))
    Yf <- Y
    for (s in 1:G) {
      Zf[(s-1)*N + 1:N, s] <- 0
      Yf[(s-1)*N + 1:N, s] <- 0
    }
    E <- Zf + Yf
    e <- rowSums(E)

    cli::cli_progress_done()

    # Balanced table ----------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Table is balanced?}")
    Sys.sleep(delay)

    diff <- xrow - xcol
    hits <- data.frame(
      row = which(abs(diff) > threshold),
      col = which(abs(diff) > threshold),
      value = abs(diff[which(abs(diff) > threshold)])
    )
    excel_cells <- c()
    row_entities <- c()
    col_entities <- c()

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No row/column discrepancies in outputs.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Unbalanced row/column outputs}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {

        excel_cell_row <- paste0(excelcol(hits$col[hit] + 4), 7 + nrow(mrio))
        excel_cell_col <- paste0(excelcol(4 + ncol(mrio)), 7 + hits$row[hit])
        excel_cell <- paste0(excel_cell_row, ", ", excel_cell_col)
        row_entity = mrio_entity(
          s = hits$row[hit] %/% N + 1,
          i = ifelse(hits$row[hit] %% N == 0, N, hits$row[hit] %% N),
          g = G
        )
        col_entity = mrio_entity(
          s = hits$col[hit] %/% N + 1,
          i = ifelse(hits$row[hit] %% N == 0, N, hits$row[hit] %% N),
          g = G
        )
        excel_cells <- c(excel_cells, excel_cell)
        row_entities <- c(row_entities, row_entity)
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{.cell {excel_cell}} {row_entity}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$type <- "Unbalanced outputs"
      hits$cell <- excel_cells
      hits$row_entity <- row_entities
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }

    # Cells -------------------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Cells have expected values?}")
    Sys.sleep(delay)

    # Blank cells

    hits <- data.frame(row = c(), col = c(), value = c())
    excel_cells <- c()

    for (col in 1:ncol(mrio)) {
      hit <- which(is.na(mrio[, col]))
      if (length(hit) > 0) hits <- rbind(hits, data.frame(row = hit, col = col, value = mrio[hit, col]))
    }
    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No blank cells.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Blank cells}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {
        excel_cell <- paste0(excelcol(hits$col[hit] + 4), 7 + hits$row[hit])
        excel_cells <- c(excel_cells, excel_cell)
        cli::cli_bullets(c(" " = "{.cell {excel_cell}}"))
        Sys.sleep(delay)
      }

      hits$type <- "Blank cells"
      hits$cell <- excel_cells
      hits$row_entity <- NA
      hits$col_entity <- NA
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }

    # Negative inputs

    hits <- data.frame(row = c(), col = c(), value = c())
    excel_cells <- c()
    row_entities <- c()
    col_entities <- c()

    for (col in 1:(G*N)) {
      hit <- which((mrio[1:(G*N), col] < -threshold))
      if (length(hit) > 0) hits <- rbind(hits, data.frame(row = hit, col = col, value = mrio[hit, col]))
    }
    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative inputs.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Negative inputs}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {
        excel_cell <- paste0(excelcol(hits$col[hit] + 4), 7 + hits$row[hit])
        row_entity = mrio_entity(
          s = hits$row[hit] %/% N + 1,
          i = ifelse(hits$row[hit] %% N == 0, N, hits$row[hit] %% N),
          g = G
        )
        col_entity = mrio_entity(
          s = hits$col[hit] %/% N + 1,
          i = ifelse(hits$row[hit] %% N == 0, N, hits$row[hit] %% N),
          g = G
        )
        excel_cells <- c(excel_cells, excel_cell)
        row_entities <- c(row_entities, row_entity)
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{.cell {excel_cell}} {row_entity} => {col_entity}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$type <- "Negative inputs"
      hits$cell <- excel_cells
      hits$row_entity <- row_entities
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }

    # Negative consumption

    hits <- data.frame(row = c(), col = c(), value = c())
    excel_cells <- c()
    row_entities <- c()
    col_entities <- c()
    col_index <- G*N + subset(1:(G*f), 1:(G*f) %% 5 != 0)

    for (col in col_index) {
      hit <- which((mrio[1:(G*N), col] < -threshold))
      if (length(hit) > 0) hits <- rbind(hits, data.frame(row = hit, col = col - G*N, value = mrio[hit, col]))
    }
    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative consumption items.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Negative consumption items}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {
        excel_cell <- paste0(excelcol(hits$col[hit] + G*N + 4), 7 + hits$row[hit])
        row_entity = mrio_entity(
          s = hits$row[hit] %/% N + 1,
          i = ifelse(hits$row[hit] %% N == 0, N, hits$row[hit] %% N),
          g = G
        )
        col_entity = mrio_entity(
          s = hits$col[hit] %/% f + 1,
          f = hits$col[hit] %% f,
          g = G
        )
        excel_cells <- c(excel_cells, excel_cell)
        row_entities <- c(row_entities, row_entity)
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{.cell {excel_cell}} {row_entity} => {col_entity}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$"type" <- "Negative consumption"
      hits$cell <- excel_cells
      hits$row_entity <- row_entities
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }

    # Negative value added at basic prices

    hits <- data.frame(
      col = which(vabp < -threshold),
      value = vabp[which(vabp < -threshold)]
    )
    excel_cells <- c()
    col_entities <- c()

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative value added at basic prices.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Negative value added at basic prices}"))
      Sys.sleep(delay)

      for (hit in 1:nrow(hits)) {
        excel_cell <- paste0(excelcol(hits$col[hit] + 4), 7 + G*N + 6)
        col_entity = mrio_entity(
          s = hits$col[hit] %/% N + 1,
          i = ifelse(hits$col[hit] %% N == 0, N, hits$col[hit] %% N),
          g = G
        )
        excel_cells <- c(excel_cells, excel_cell)
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{.cell {excel_cell}} {col_entity}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$"type" <- "Negative value added at basic prices"
      hits$cell <- excel_cells
      hits$row_entity <- NA
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }

    # Negative direct purchases abroad

    hits <- data.frame(
      col = which(dpa < -threshold),
      value = dpa[which(dpa < -threshold)]
    )
    excel_cells <- c()
    col_entities <- c()

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative direct purchases abroad.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Negative direct purchases abroad}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {
        excel_cell <- paste0(excelcol(hits$col[hit] + 4), 7 + G*N + 4)
        col_entity = mrio_entity(
          s = hits$col[hit] %/% N + 1,
          i = ifelse(hits$col[hit] %% N == 0, N, hits$col[hit] %% N),
          g = G
        )
        excel_cells <- c(excel_cells, excel_cell)
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{.cell {excel_cell}} {col_entity}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$"type" <- "Negative direct purchases abroad"
      hits$cell <- excel_cells
      hits$row_entity <- NA
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }

    # Positive purchases by non-residents

    hits <- data.frame(
      col = which(pnr > threshold),
      value = pnr[which(pnr > threshold)]
    )
    excel_cells <- c()
    col_entities <- c()

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No positive purchases by non-residents.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Positive purchases by non-residents}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {
        excel_cell <- paste0(excelcol(hits$col[hit] + 4), 7 + G*N + 5)
        col_entity = mrio_entity(
          s = hits$col[hit] %/% N + 1,
          i = ifelse(hits$col[hit] %% N == 0, N, hits$col[hit] %% N),
          g = G
        )
        excel_cells <- c(excel_cells, excel_cell)
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{.cell {excel_cell}} {col_entity}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$"type" <- "Positive purchases by non-residents"
      hits$cell <- excel_cells
      hits$row_entity <- NA
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }

    # Aggregates --------------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Aggregates have expected values?}")
    Sys.sleep(delay)

    # Value added

    hits <- data.frame(
      col = which(va < -threshold),
      value = va[which(va < -threshold)]
    )
    col_entities <- c()

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative value added.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Negative value added}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {
        col_entity = mrio_entity(
          s = hits$col[hit] %/% N + 1,
          i = ifelse(hits$col[hit] %% N == 0, N, hits$col[hit] %% N),
          g = G
        )
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{col_entity}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$"type" <- "Negative value added"
      hits$cell <- NA
      hits$row_entity <- NA
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }

    # Exports

    hits <- data.frame(row = c(), col = c(), value = c())
    row_entities <- c()
    col_entities <- c()

    for (col in 1:G) {
      hit <- which((E[, col] < -threshold))
      if (length(hit) > 0) hits <- rbind(hits, data.frame(row = hit, col = col, value = E[hit, col]))
    }
    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No negative exports.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Negative exports}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {
        row_entity = mrio_entity(
          s = hits$row[hit] %/% N + 1,
          i = ifelse(hits$row[hit] %% N == 0, N, hits$row[hit] %% N),
          g = G
        )
        col_entity = mrio_entity(
          s = hits$col[hit],
          g = G
        )
        excel_cells <- c(excel_cells, excel_cell)
        row_entities <- c(row_entities, row_entity)
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{row_entity} => {col_entity}: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$type <- "Negative exports"
      hits$cell <- NA
      hits$row_entity <- row_entities
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }
    Sys.sleep(delay)

    # GVA = 0, inputs > 0

    hits <- data.frame(
      col = which(abs(va) < threshold & zuse > threshold),
      value = zuse[which(abs(va) < threshold & zuse > threshold)]
    )
    col_entities <- c()

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No GVA = 0 with inputs > 0.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn GVA = 0 but inputs > 0}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {
        col_entity = mrio_entity(
          s = hits$col[hit],
          g = G
        )
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{col_entity} total inputs: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$type <- "GVA = 0 but inputs > 0"
      hits$cell <- NA
      hits$row_entity <- NA
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }
    Sys.sleep(delay)

    # GVA = 0, exports > 0

    hits <- data.frame(
      col = which(abs(va) < threshold & e > threshold),
      value = e[which(abs(va) < threshold & e > threshold)]
    )
    col_entities <- c()

    if (nrow(hits) == 0) {
      cli::cli_alert_success(" {.check No GVA = 0 with exports > 0.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn GVA = 0 but exports > 0}"))
      Sys.sleep(delay)
      for (hit in 1:nrow(hits)) {
        col_entity = mrio_entity(
          s = hits$col[hit],
          g = G
        )
        col_entities <- c(col_entities, col_entity)

        cli::cli_bullets(c(" " = "{col_entity} total exports: {hits$value[hit]}"))
        Sys.sleep(delay)
      }

      hits$type <- "GVA = 0 but exports > 0"
      hits$cell <- NA
      hits$row_entity <- NA
      hits$col_entity <- col_entities
      report <- rbind(
        report,
        subset(hits, select = c("type", "cell", "row_entity", "col_entity", "value"))
      )
    }
    Sys.sleep(delay)

    if (nrow(report) > 0) {
      openxlsx::addWorksheet(reports, basename(file))
      openxlsx::writeData(reports, basename(file), report, headerStyle = headerstyle)
    }
  }

  cli::cli_text("")
  cli::cli_rule("")
  Sys.sleep(delay)

  # Save report in Excel file
  if (export) {
    filename <- paste0("Report ", gsub(":|\\s", "-", timestamp), ".xlsx")
    openxlsx::saveWorkbook(reports, file.path(savedir, filename), overwrite = TRUE)
    cli::cli_text("")
    cli::cli_text('Report successfully saved at "{.header {savedir}}".')
  }

  cli::cli_text("")
  Sys.sleep(delay)
}

excelcol <- function(n) {
  if (n < 1 | n %% 1 != 0) cli::cli_abort("`n` must be a positive integer.")
  N <- length(LETTERS)
  j <- n %/% N
  if (j == 0) LETTERS[n] else paste0(Recall(j), LETTERS[n %% N])
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

  entity <- dict_countries1$code[dict_countries1$ind == s]
  if (!is.null(i)) entity <- paste0(entity, "_", paste0("c", i))
  if (!is.null(f)) entity <- paste0(entity, "_", paste0("f", f))
  return(entity)
}

print_result <- function(hits, success, fail) {
  if (length(hits) == 0) {
    return(cli::cli_alert_success(" {.check {success}.}"))
  } else {
    hits <- paste(hits, collapse = ", ")
    warnings <- paste0(" {.warn {fail}: }", hits)
    return(cli::cli_bullets(c("!" = warnings)))
  }
}

print_timestamp <- function() {
  Sys.time()
}
