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
#'
#' @importFrom rlang .data
#'
#' @export
check <- function(path, precision = 8) {

  threshold <- as.numeric(glue::glue("1e{-precision}"))
  delay <- .25
  cli::cli_div(theme = list(
    span.header = list(color = "cyan"),
    span.check = list(color = "darkgreen"),
    span.warn = list(color = "darkgoldenrod4")
  ))

  # Get files
  if (file.info(path)$isdir) {
    files <- list.files(path, pattern = "^[^~].*(MRIO|mrio|Mrio).*(xls|xlsx)$")
  } else {
    files <- path
  }

  for (file in files) {

    if (file.info(path)$isdir) file <- file.path(path, file)

    cli::cli_text("")
    cli::cli_rule(left = "Checks on {.header {basename(file)}}")
    cli::cli_progress_message("Reading file...")

    mrio <- readxl::read_excel(
        file,
        range = readxl::cell_limits(c(8, 5), c(NA, NA)),
        col_names = FALSE,
        progress = FALSE
      ) |>
      as.matrix() |>
      suppressMessages()

    if (nrow(mrio) < 2205 | ncol(mrio) < 2205) {
      cli::cli_bullets(c("x" = " This file is not an MRIO table or it does not follow the standard MRIO format."))
      next
    }

    N <- 35
    f <- 5
    G <- (ncol(mrio) - 1) / (N + f)
    mrio <- mrio[1:(G*N + 8), ]

    xrow <- mrio[nrow(mrio), 1:(G*N)]
    xcol <- mrio[1:(G*N), ncol(mrio)]
    dpa <- mrio[G*N + 4, (G*N + 1):(G*N + 1 + G*f)]
    pnr <- mrio[G*N + 5, (G*N + 1):(G*N + 1 + G*f)]
    vabp <- mrio[G*N + 6, 1:(G*N)]

    Z <- mrio[1:(G * N), 1:(G * N)]
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

    cli::cli_progress_done()

    # Balanced table ----------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Table is balanced?}")
    Sys.sleep(delay)

    diff <- xrow - xcol
    hits <- which(abs(diff) > threshold)

    if (length(hits) == 0) {
      cli::cli_alert_success(" {.check No discrepancies greater than {threshold}.}")
    }
    else {
      row_hits <- sapply(4 + hits, excelcol) |> paste0(7 + nrow(mrio))
      col_hits <- paste0(excelcol(4 + ncol(mrio)), 7 + hits)
      warnings <- c()
      for (h in 1:length(hits)) {
        warnings <- c(
          warnings,
          " " = glue::glue(" Difference in output between {row_hits[h]} and {col_hits[h]} is {abs(diff[hits[h]])}")
        )
      }
      cli::cli_bullets(c("!" = " {.warn Table may be unbalanced.}", warnings))
    }
    Sys.sleep(delay)

    # Cells -------------------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Cells have expected values?}")
    Sys.sleep(delay)

    # Blank cells
    hits <- c()
    for (i in 1:ncol(mrio)) {
      hit <- which(is.na(mrio[, i]))
      if (length(hit) > 0) hits <- c(hits, "*" = paste0(excelcol(i + 4), 7 + hit))
    }
    print_result(
      hits,
      success = "No blank cells",
      fail = "Blank cells"
    )
    Sys.sleep(delay)

    # Negative inputs
    hits <- c()
    for (i in 1:(G*N)) {
      hit <- which((mrio[1:(G*N), i] < 0))
      if (length(hit) > 0) hits <- c(hits, "*" = paste0(excelcol(i + 4), 7 + hit))
    }
    print_result(
      hits,
      success = "No negative inputs",
      fail = "Cells with negative inputs"
    )
    Sys.sleep(delay)

    # Negative consumption
    hits <- c()
    colindex <- G*N + subset(1:(G*f), 1:(G*f) %% 5 != 0)
    for (i in colindex) {
      hit <- which((mrio[1:(G*N), i] < 0))
      if (length(hit) > 0) hits <- c(hits, "*" = paste0(excelcol(i + 4), 7 + hit))
    }
    print_result(
      hits,
      success = "No negative consumption",
      fail = "Cells with negative consumption"
    )
    Sys.sleep(delay)

    # Negative value added at basic prices
    hits <- which(vabp < 0)
    if (length(hits) != 0) {
      hits <- sapply(4 + hits, excelcol) |> paste0(7 + G*N + 6)
    }
    print_result(
      hits,
      success = "No negative value added at basic prices",
      fail = "Cells with negative value added at basic prices"
    )
    Sys.sleep(delay)

    # Negative direct purchases abroad
    hits <- which(dpa < 0)
    if (length(hits) != 0) {
      hits <- sapply(4 + hits + G*N, excelcol) |> paste0(7 + G*N + 4)
    }
    print_result(
      hits,
      success = "No negative direct purchases abroad",
      fail = "Cells with negative direct purchases abroad"
    )
    Sys.sleep(delay)

    # Positive purchases by non-residents
    hits <- which(pnr > 0)
    if (length(hits) != 0) {
      hits <- sapply(4 + hits + G*N, excelcol) |> paste0(7 + G*N + 5)
    }
    print_result(
      hits,
      success = "No positive purchases by non-residents",
      fail = "Cells with positive purchases by non-residents"
    )
    Sys.sleep(delay)

    # Aggregates --------------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Aggregates have expected values?}")
    Sys.sleep(delay)

    # Value added
    hits <- which(va < 0)
    if (length(hits) == 0) {
      cli::cli_alert_success(" {.check No negative value added.}")
      Sys.sleep(delay)
    } else {
      cli::cli_bullets(c("!" = " {.warn Negative value added}"))
      Sys.sleep(delay)
      for (h in hits) {
        entity <- mrio_entity(s = h %/% N + 1, i = h %% N, g = G)
        cli::cli_bullets(c(" " = "{entity}: {va[h]}"))
        Sys.sleep(delay)
      }
    }

    # Exports
    running_hits <- 0
    for (col in 1:G) {

      hits <- which((E[, col] < 0))

      if (length(hits) > 0) {
        if (running_hits == 0) cli::cli_bullets(c("!" = " {.warn Negative exports}"))
        running_hits <- running_hits + length(hits)
        Sys.sleep(delay)

        for (h in hits) {
          exporter <- mrio_entity(s = h %/% N + 1, i = h %% N, g = G)
          partner <- mrio_entity(s = col, g = G)
          cli::cli_bullets(c(" " = "{exporter} => {partner}: {E[h, col]}"))
          Sys.sleep(delay)
        }
      }
    }
    if (running_hits == 0) cli::cli_alert_success(" {.check No negative exports.}")
    Sys.sleep(delay)
  }

  cli::cli_text("")
}

excelcol <- function(n) {
  if (n < 1 | n %% 1 != 0) cli::cli_abort("`n` must be a positive integer.")
  N <- length(LETTERS)
  j <- n %/% N
  if (j == 0) LETTERS[n] else paste0(Recall(j), LETTERS[n %% N])
}

mrio_entity <- function(s, i = NULL, g) {

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

  country <- dict_countries1$code[dict_countries1$ind == s]
  sector <- paste0("c", i)
  if (is.null(i)) return(country) else return(paste0(country, "_", sector))
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
