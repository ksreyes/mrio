#' Download Eurostat macro data
#'
#' @description Accesses the Eurostat API using the `restatapi` package to
#'   extract macroeconomic data. These are used in the control totals updating
#'   phase of the MRIO compilation process.
#'
#' @param path Path to file or directory. Files must be in Excel and file
#'   names must start with the uppercase three-letter country code. Function
#'   ignores all other files in the directory.
#' @param reorder_sheets If `TRUE`, arranges the sheets of the workbook
#'   alphabetically upon saving.
#' @param starting_year Earliest year of data to download.
#'
#' @export
download_eurostat <- function(path, reorder_sheets = TRUE, starting_year = 1995) {

  cli::cli_div(theme = list(span.header = list(color = "cyan")))
  headerstyle <- openxlsx::createStyle(fontColour = "white", fgFill = "#007DB7", textDecoration = "bold")
  latest_date <- format(Sys.time(), "%Y-%m-%d")
  latest_year <- format(Sys.time(), "%Y") |> as.numeric()
  stem <- "1. ES"

  cli::cli_text("")
  cli::cli_progress_message("Downloading Eurostat metadata...")
  tables <- c("nama_10_gdp", "nama_10_a64", "namq_10_gdp", "namq_10_a10")
  meta <- restatapi::get_eurostat_toc() |> dplyr::filter(code %in% tables)
  cli::cli_progress_done()

  if (file.info(path)$isdir) {
    files <- list.files(path, pattern = "^[A-Z]{3}.*(xls|xlsx)$")
  } else {
    files <- path
  }

  for (file in files) {

    filepath <- file.path(path, file)
    wb <- openxlsx::loadWorkbook(filepath)
    sheets <- openxlsx::sheets(wb)

    filecode <- stringr::str_extract(file, "[A-Z]{3}")
    country <- get_iso_a2(filecode)
    if (!(country %in% es_countries$code)) {
      cli::cli_bullets(c("x" = " {filecode} is not a Eurostat country."))
      cli::cli_text("")
      next
    }

    # nama_10_gdp -------------------------------------------------------------

    meta1 <- subset(meta, code == tables[1])
    sheetname1 <- glue::glue("{stem} {tables[1]}")
    if (!(sheetname1 %in% sheets)) {
      openxlsx::addWorksheet(wb, sheetname1, tabColour = "#BDD7EE", zoom = 80)
    }
    openxlsx::deleteData(wb, sheetname1, cols = 1:999, rows = 1:9999, gridExpand = TRUE)

    df1 <- restatapi::get_eurostat_data(
        tables[1],
        filters = c(country, "CP_MNAC", "CLV15_MNAC"),
        date_filter = starting_year:latest_year,
        label = TRUE
      ) |>
      tidyr::pivot_wider(names_from = time, values_from = values) |>
      dplyr::mutate(na_item = forcats::fct_expand(na_item, esdsd_nama_10_gdp$name)) |>
      tidyr::complete(geo, na_item, unit) |>
      dplyr::arrange(unit) |>
      dplyr::right_join(esdsd_nama_10_gdp, by = c("na_item" = "name")) |>
      dplyr::select(geo, code, na_item, unit, matches("[0-9]{4}"))

    units <- paste(unique(df1$unit), collapse = "; ")

    writepair("Source", "Eurostat", 1, sheet = sheetname1, workbook = wb)
    writepair("Title", meta1$title, 2, sheet = sheetname1, workbook = wb)
    writepair("Code", meta1$code, 3, sheet = sheetname1, workbook = wb)
    writepair("Units", units, 4, sheet = sheetname1, workbook = wb)
    writepair("Last updated", meta1$lastUpdate, 5, sheet = sheetname1, workbook = wb)
    writepair("Date extracted", latest_date, 6, sheet = sheetname1, workbook = wb)

    openxlsx::writeData(
      wb, sheetname1, df1, startRow = 8,
      keepNA = TRUE, na.string = ":",
      withFilter = TRUE, headerStyle = headerstyle
    )
    openxlsx::setColWidths(wb, sheetname1, c(1, 3, 4), c(20, 40, 20))
    sheets <- openxlsx::sheets(wb)

    # nama_10_a64 -------------------------------------------------------------

    meta2 <- subset(meta, code == tables[2])
    sheetname2 <- glue::glue("{stem} {tables[2]}")
    if (!(sheetname2 %in% sheets)) {
      openxlsx::addWorksheet(wb, sheetname2, tabColour = "#BDD7EE", zoom = 80)
    }
    openxlsx::deleteData(wb, sheetname2, cols = 1:999, rows = 1:9999, gridExpand = TRUE)

    df2 <- restatapi::get_eurostat_data(
        tables[2],
        filters = c(country, "CP_MNAC", "CLV15_MNAC", "B1G", "P1", "P2"),
        date_filter = starting_year:latest_year,
        label = TRUE
      ) |>
      tidyr::pivot_wider(names_from = time, values_from = values) |>
      dplyr::mutate(nace_r2 = forcats::fct_expand(nace_r2, esdsd_nama_10_a64$name)) |>
      tidyr::complete(geo, na_item, unit, nace_r2) |>
      dplyr::right_join(esdsd_nama_10_a64, by = c("nace_r2" = "name")) |>
      dplyr::select(geo, na_item, unit, code, nace_r2, matches("[0-9]{4}")) |>
      dplyr::group_by(na_item, unit) |>
      dplyr::arrange(code, .by_group = TRUE) |>
      dplyr::ungroup()

    items <- paste(unique(df2$na_item), collapse = "; ")
    units <- paste(unique(df2$unit), collapse = "; ")

    writepair("Source", "Eurostat", 1, sheet = sheetname2, workbook = wb)
    writepair("Title", meta2$title, 2, sheet = sheetname2, workbook = wb)
    writepair("Code", meta2$code, 3, sheet = sheetname2, workbook = wb)
    writepair("Items", items, 4, sheet = sheetname2, workbook = wb)
    writepair("Units", units, 5, sheet = sheetname2, workbook = wb)
    writepair("Last updated", meta2$lastUpdate, 6, sheet = sheetname2, workbook = wb)
    writepair("Date extracted", latest_date, 7, sheet = sheetname2, workbook = wb)

    openxlsx::writeData(
      wb, sheetname2, df2, startRow = 9,
      keepNA = TRUE, na.string = ":",
      withFilter = TRUE, headerStyle = headerstyle
    )
    openxlsx::setColWidths(wb, sheetname2, c(1, 2, 3, 5), c(20, 20, 20, 50))
    sheets <- openxlsx::sheets(wb)

    # namq_10_gdp -------------------------------------------------------------

    meta3 <- subset(meta, code == tables[3])
    sheetname3 <- glue::glue("{stem} {tables[3]}")
    if (!(sheetname3 %in% sheets)) {
      openxlsx::addWorksheet(wb, sheetname3, tabColour = "#BDD7EE", zoom = 80)
    }
    openxlsx::deleteData(wb, sheetname3, cols = 1:999, rows = 1:9999, gridExpand = TRUE)

    df3 <- restatapi::get_eurostat_data(
        tables[3],
        filters = c(country, "CP_MNAC", "CLV15_MNAC", "NSA"),
        date_filter = starting_year:latest_year,
        label = TRUE
      ) |>
      tidyr::pivot_wider(names_from = time, values_from = values) |>
      dplyr::mutate(na_item = forcats::fct_expand(na_item, esdsd_namq_10_gdp$name)) |>
      tidyr::complete(geo, na_item, unit, s_adj) |>
      dplyr::arrange(unit) |>
      dplyr::right_join(esdsd_nama_10_gdp, by = c("na_item" = "name")) |>
      dplyr::select(geo, code, na_item, unit, s_adj, matches("[0-9]{4}-Q[0-9]{1}"))

    units <- paste(unique(df3$unit), collapse = "; ")

    writepair("Source", "Eurostat", 1, sheet = sheetname3, workbook = wb)
    writepair("Title", meta3$title, 2, sheet = sheetname3, workbook = wb)
    writepair("Code", meta3$code, 3, sheet = sheetname3, workbook = wb)
    writepair("Units", units, 4, sheet = sheetname3, workbook = wb)
    writepair("Adjustment", unique(df3$s_adj), 5, sheet = sheetname3, workbook = wb)
    writepair("Last updated", meta3$lastUpdate, 6, sheet = sheetname3, workbook = wb)
    writepair("Date extracted", latest_date, 7, sheet = sheetname3, workbook = wb)

    df3 <- df3 |> dplyr::select(-s_adj)
    openxlsx::writeData(
      wb, sheetname3, df3, startRow = 9,
      keepNA = TRUE, na.string = ":",
      withFilter = TRUE, headerStyle = headerstyle
    )
    openxlsx::setColWidths(wb, sheetname3, c(1, 3, 4), c(20, 40, 20))
    sheets <- openxlsx::sheets(wb)

    # namq_10_a10 -------------------------------------------------------------

    meta4 <- subset(meta, code == tables[4])
    sheetname4 <- glue::glue("{stem} {tables[4]}")
    if (!(sheetname4 %in% sheets)) {
      openxlsx::addWorksheet(wb, sheetname4, tabColour = "#BDD7EE", zoom = 80)
    }
    openxlsx::deleteData(wb, sheetname4, cols = 1:999, rows = 1:9999, gridExpand = TRUE)

    df4 <- restatapi::get_eurostat_data(
        tables[4],
        filters = c(country, "CP_MNAC", "CLV15_MNAC", "NSA", "B1G"),
        date_filter = starting_year:latest_year,
        label = TRUE
      ) |>
      tidyr::pivot_wider(names_from = time, values_from = values) |>
      dplyr::mutate(nace_r2 = forcats::fct_expand(nace_r2, esdsd_namq_10_a10$name)) |>
      tidyr::complete(geo, na_item, unit, s_adj, nace_r2) |>
      dplyr::arrange(unit) |>
      dplyr::right_join(esdsd_namq_10_a10, by = c("nace_r2" = "name")) |>
      dplyr::select(geo, na_item, unit, s_adj, code, nace_r2, matches("[0-9]{4}-Q[0-9]{1}"))

    items <- paste(unique(df4$na_item), collapse = "; ")
    units <- paste(unique(df4$unit), collapse = "; ")

    writepair("Source", "Eurostat", 1, sheet = sheetname4, workbook = wb)
    writepair("Title", meta4$title, 2, sheet = sheetname4, workbook = wb)
    writepair("Code", meta4$code, 3, sheet = sheetname4, workbook = wb)
    writepair("Item", unique(df4$na_item), 4, sheet = sheetname4, workbook = wb)
    writepair("Units", units, 5, sheet = sheetname4, workbook = wb)
    writepair("Adjustment", unique(df4$s_adj), 6, sheet = sheetname4, workbook = wb)
    writepair("Last updated", meta4$lastUpdate, 7, sheet = sheetname4, workbook = wb)
    writepair("Date extracted", latest_date, 8, sheet = sheetname4, workbook = wb)

    df4 <- df4 |> dplyr::select(-na_item, -s_adj)
    openxlsx::writeData(
      wb, sheetname4, df4, startRow = 10,
      keepNA = TRUE, na.string = ":",
      withFilter = TRUE, headerStyle = headerstyle
    )
    openxlsx::setColWidths(wb, sheetname4, c(1, 2, 4), c(20, 20, 50))
    sheets <- openxlsx::sheets(wb)

    if (reorder_sheets) openxlsx::worksheetOrder(wb) <- order(sheets)
    openxlsx::saveWorkbook(wb, filepath, overwrite = TRUE)

    cli::cli_text('"{.header {file}}" successfully updated.')
    cli::cli_text("")
  }
}

# Helper functions --------------------------------------------------------

writepair <- function(x, y, row, sheet, coloffset = 0, workbook) {
  openxlsx::writeData(
    workbook, sheet, x,
    startCol = coloffset + 1,
    startRow = row,
    colNames = FALSE
  )
  openxlsx::writeData(
    workbook, sheet, y,
    startCol = coloffset + 2,
    startRow = row,
    colNames = FALSE
  )
}