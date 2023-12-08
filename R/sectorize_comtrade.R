#' Aggregate Comtrade to MRIO sectors
#'
#' @description Aggregates raw product-level data from a UN Comtrade bulk
#'   download to the MRIO sectors and BEC5 end use categories.
#'
#' @param path File or directory of files of Comtrade bulk data.
#'
#' @details Each file must be a Comtrade bulk download of one country for one
#'   year compressed via Gzip. The function searches for file names in `path`
#'   with the pattern "ABC 20XX.gz", where "ABC" is the uppercase three-letter
#'   country code and "20XX" is the selected year. All other files are ignored.
#'
#' @export
sectorize_comtrade <- function(path) {

  headerstyle <- openxlsx::createStyle(fgFill = "#E7E6E6",
                                       textDecoration = "bold")
  cli::cli_div(theme = list(span.header = list(color = "cyan"),
                            span.check = list(color = "darkgreen")))

  savedir <- ifelse(file.info(path)$isdir, path, dirname(path))
  if (file.info(path)$isdir) {
    files <- list.files(path, pattern = "^[A-Z]{3}.*(20)[0-9]{2}.*(gz)$")
  } else {
    files <- path
  }

  br()
  cli::cli_text("Found the following files:")
  for (file in files) cli::cli_bullets(c("*" = " {.header {file}}"))
  br()

  for (file in files) {

    cli::cli_text("Working on {.header {file}}...")

    conn <- DBI::dbConnect(duckdb::duckdb(), dbdir = ":memory:")

    df_cleaned <- DBI::dbGetQuery(conn, glue::glue(
      "SELECT period,
        reporterCode,
        flowCode,
        partnerCode,
        classificationCode,
        cmdCode,
        primaryValue
      FROM read_csv_auto('{file.path(savedir, file)}')
      WHERE partnerCode <> 0
        AND partner2Code = 0
        AND (flowCode = 'X' OR flowCode = 'M')
        AND length(cmdCode) = 6
        AND motCode = 0
        AND customsCode = 'C00'"
      ))

    hscode <- hscodes$edition[hscodes$code == unique(df_cleaned$classificationCode)]

    # Assume that unknown products (cmdCode 999999) are military weapons
    hsbecsitc$BEC5[hsbecsitc$BEC5 == 8] <- "812020"

    df_sectorized <- df_cleaned |>
      dplyr::left_join(
        hsbecsitc |> dplyr::select({{hscode}}, .data$HS02, .data$BEC5),
        by = c("cmdCode" = hscode),
        multiple = "any"
      ) |>
      dplyr::left_join(
        hsbec |> dplyr::select(BEC5 = .data$BEC5Code1, .data$BEC5EndUse),
        multiple = "any"
      ) |>
      dplyr::left_join(hs02cpc, multiple = "any") |>
      dplyr::left_join(cpcisic, multiple = "any") |>
      dplyr::left_join(isicmrio, multiple = "any") |>
      dplyr::summarise(
        primaryValue = sum(.data$primaryValue),
        .by = c(.data$period, .data$reporterCode, .data$flowCode, .data$partnerCode, .data$BEC5EndUse, .data$MRIO)
      ) |>
      suppressMessages()

    filename <- paste0(sub("\\.gz$", "", file), ".xlsx")
    sheetname <- "Sheet1"
    output <- openxlsx::createWorkbook()
    openxlsx::addWorksheet(output, sheetname)
    openxlsx::writeData(output, sheetname, df_sectorized, withFilter = TRUE, headerStyle = headerstyle)
    openxlsx::saveWorkbook(output, file.path(savedir, filename), overwrite = TRUE)

    DBI::dbDisconnect(conn, shutdown = TRUE)

    cli::cli_alert_success("{.check Completed.}")
    br()
  }
}

# Helper functions --------------------------------------------------------

br <- function() cli::cli_text("")
