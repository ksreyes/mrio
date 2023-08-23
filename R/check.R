#' Perform basic integrity checks on MRIO Excel file
#'
#' Given a directory `path`, this will load all Excel files with "MRIO"
#' (case-insensitive) in the file name and examine them for the following
#' anomalies:
#'  \enumerate{
#'    \item Blank cells.
#'    \item Negative values for inputs and consumption.
#'    \item
#'  }
#' The number of countries in the MRIO is automatically detected. However, note
#' that the MRIO table must follow the standard Excel format.
#'
#' @param path Directory of files to check.
#' @param precision Largest difference tolerated for equality.
#'
#' @importFrom rlang .data
#'
#' @export
check <- function(path, precision = 8) {

  threshold <- as.numeric(sprintf("1e%d", -precision))
  delay <- .25
  cli::cli_div(theme = list(
    span.header = list(color = "cyan"),
    span.check = list(color = "darkgreen")
  ))

  files <- list.files(path, pattern = "^[^~].*(MRIO|mrio|Mrio).*(xls|xlsx)$")

  for (file in files) {

    cli::cli_text("")
    cli::cli_rule(left = "Checks on {.header {file}}")

    cli::cli_progress_message("Loading file...")

    mrio <- paste0(path, file) |>
      readxl::read_excel(
        range = readxl::cell_limits(c(8, 5), c(NA, NA)),
        col_names = FALSE,
        progress = FALSE
      ) |>
      suppressMessages() |>
      as.matrix()

    cli::cli_progress_done()

    if (nrow(mrio) < 2205 | ncol(mrio) < 2205) {
      cli::cli_bullets(c("x" = " This file is not an MRIO table or it does not follow the standard MRIO format."))
      next
    }

    N <- 35
    f <- 5
    G <- (nrow(mrio) - 8) / N

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

    # Balanced table ----------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Table is balanced?}")

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
          "*" = sprintf(" Difference in output between %s and %s is %.2f", row_hits[h], col_hits[h], abs(diff[hits[h]]))
        )
      }
      cli::cli_bullets(c("!" = " Table may be unbalanced.", warnings))
    }
    Sys.sleep(delay)

    # Cells -------------------------------------------------------------------

    cli::cli_text("")
    cli::cli_text("{.emph Cells have expected values?}")

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

    # Value added
    warnings <- c()
    hits <- which(va < 0)
    if (length(hits) == 0) {
      cli::cli_alert_success(" {.check No negative value added.}")
    } else {
      for (h in hits) {
        entity <- mrio_entity(s = h %/% N + 1, i = h %% N, g = G)
        warnings <- c(
          warnings,
          "*" = paste0(" ", entity, ": ", va[h])
        )
      }
      cli::cli_bullets(c("!" = " Negative value added", warnings))
    }
    Sys.sleep(delay)

    # Exports
    warnings <- c()
    for (col in 1:G) {
      hits <- which((E[, col] < 0))
      if (length(hits) > 0) {
        for (h in hits) {
          exporter <- mrio_entity(s = h %/% N + 1, i = h %% N, g = G)
          partner <- mrio_entity(s = col, g = G)
          warnings <- c(
            warnings,
            "*" = paste0(" ", exporter, " -> ", partner , ": ", E[h, col])
          )
        }
      }
    }
    if (length(warnings) == 0) {
      cli::cli_alert_success(" {.check No negative exports.}")
    } else {
      cli::cli_bullets(c("!" = " Negative exports", warnings))
    }

    rm(dpa, E, f, file, G, hits, mrio, N, pnr, s, va, vabp, xcol, xrow, y, Y, Y_big, Yf, Z, Zf)
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
      dplyr::rename(ind = "mrio_ind")
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
    warnings <- paste0(" {fail}: ", hits)
    return(cli::cli_bullets(c("!" = warnings)))
  }
}
