#' Download Eurostat macro data
#'
#' @description Downloads macroeconomic data from Eurostat and saves them in a
#'   control totals workbook.
#'
#' @param workbook Control totals workbook or directory of workbooks. File names
#'   must start with the uppercase three-letter country code. All other files in
#'   the directory are ignored.
#' @param reorder_sheets If `TRUE`, arranges the sheets of the workbook
#'   alphabetically upon saving.
#' @param starting_year Earliest year of data to download.
#'
#' @details Uses the `restatapi` package to access the Eurostat API. Downloaded
#'   data are saved in a series of sheets with name pattern "1. ES <dataset
#'   code>". Sheets are created if they do not exist.
#'
#' @export
download_eurostat <- function(workbook,
                              reorder_sheets = TRUE,
                              starting_year = 1995) {

  cli::cli_div(theme = list(span.header = list(color = "cyan")))
  headerstyle <- openxlsx::createStyle(fontColour = "white",
                                       fgFill = "#007DB7",
                                       textDecoration = "bold")
  latest_date <- format(Sys.time(), "%Y-%m-%d")
  latest_year <- format(Sys.time(), "%Y") |> as.numeric()

  br()
  cli::cli_text("Downloading Eurostat metadata...")
  br()

  tables <- c("nama_10_gdp", "nama_10_a64", "namq_10_gdp", "namq_10_a10")
  metadata <- restatapi::get_eurostat_toc()
  metadata <- metadata[metadata$code %in% tables, ]

  if (file.info(workbook)$isdir) {
    files <- list.files(workbook, pattern = "^[A-Z]{3}.*(xls|xlsx)$")
  } else {
    files <- workbook
  }

  for (file in files) {

    wb_path <- file.path(workbook, file)
    wb <- openxlsx::loadWorkbook(wb_path)

    # filecode <- stringr::str_extract(file, "[A-Z]{3}")
    filecode <- regmatches(file, gregexpr("[A-Z]{3}", file))
    country <- get_iso_a2(filecode)
    if (!(country %in% es_countries$code)) {
      cli::cli_bullets(c("x" = " {filecode} is not a Eurostat country."))
      br()
      next
    }

    # nama_10_gdp -------------------------------------------------------------

    info1 <- get_table_info("nama_10_gdp", metadata, wb)

    df1 <- restatapi::get_eurostat_data(info1$code,
                                        filters = c(country, "CP_MNAC", "CLV15_MNAC"),
                                        date_filter = starting_year:latest_year,
                                        label = TRUE) |>
      tidyr::pivot_wider(names_from = .data$time, values_from = .data$values) |>
      dplyr::mutate(na_item = forcats::fct_expand(.data$na_item, esdsd_nama_10_gdp$name)) |>
      tidyr::complete(.data$geo, .data$na_item, .data$unit) |>
      dplyr::arrange(.data$unit) |>
      dplyr::right_join(esdsd_nama_10_gdp, by = c("na_item" = "name")) |>
      dplyr::select(.data$geo, .data$code, .data$na_item, .data$unit, tidyselect::matches("[0-9]{4}"))

    units <- paste(unique(df1$unit), collapse = "; ")

    writepair("Source",         "Eurostat",       1, info1)
    writepair("Title",          info1$title,      2, info1)
    writepair("Code",           info1$code,       3, info1)
    writepair("Units",          units,            4, info1)
    writepair("Last updated",   info1$lastUpdate, 5, info1)
    writepair("Date extracted", latest_date,      6, info1)

    openxlsx::writeData(wb, info1$sheet, df1,
                        startRow = 8, keepNA = TRUE, na.string = ":",
                        withFilter = TRUE, headerStyle = headerstyle)
    openxlsx::setColWidths(wb, info1$sheet, c(1, 3, 4), c(20, 40, 20))

    # nama_10_a64 -------------------------------------------------------------

    info2 <- get_table_info("nama_10_a64", metadata, wb)

    df2 <- restatapi::get_eurostat_data(info2$code,
                                        filters = c(country, "CP_MNAC", "CLV15_MNAC", "B1G", "P1", "P2"),
                                        date_filter = starting_year:latest_year,
                                        label = TRUE) |>
      tidyr::pivot_wider(names_from = .data$time, values_from = .data$values) |>
      dplyr::mutate(nace_r2 = forcats::fct_expand(.data$nace_r2, esdsd_nama_10_a64$name)) |>
      tidyr::complete(.data$geo, .data$na_item, .data$unit, .data$nace_r2) |>
      dplyr::right_join(esdsd_nama_10_a64, by = c("nace_r2" = "name")) |>
      dplyr::select(.data$geo, .data$na_item, .data$unit, .data$code, .data$nace_r2, tidyselect::matches("[0-9]{4}")) |>
      dplyr::group_by(.data$na_item, .data$unit) |>
      dplyr::arrange(.data$code, .by_group = TRUE) |>
      dplyr::ungroup()

    items <- paste(unique(df2$na_item), collapse = "; ")
    units <- paste(unique(df2$unit), collapse = "; ")

    writepair("Source",         "Eurostat",       1, info2)
    writepair("Title",          info2$title,      2, info2)
    writepair("Code",           info2$code,       3, info2)
    writepair("Items",          items,            4, info2)
    writepair("Units",          units,            5, info2)
    writepair("Last updated",   info2$lastUpdate, 6, info2)
    writepair("Date extracted", latest_date,      7, info2)

    openxlsx::writeData(wb, info2$sheet, df2,
                        startRow = 9, keepNA = TRUE, na.string = ":",
                        withFilter = TRUE, headerStyle = headerstyle)
    openxlsx::setColWidths(wb, info2$sheet, c(1, 2, 3, 5), c(20, 20, 20, 50))

    # namq_10_gdp -------------------------------------------------------------

    info3 <- get_table_info("namq_10_gdp", metadata, wb)

    df3 <- restatapi::get_eurostat_data(info3$code,
                                        filters = c(country, "CP_MNAC", "CLV15_MNAC", "NSA"),
                                        date_filter = starting_year:latest_year,
                                        label = TRUE) |>
      tidyr::pivot_wider(names_from = .data$time, values_from = .data$values) |>
      dplyr::mutate(na_item = forcats::fct_expand(.data$na_item, esdsd_namq_10_gdp$name)) |>
      tidyr::complete(.data$geo, .data$na_item, .data$unit, .data$s_adj) |>
      dplyr::arrange(.data$unit) |>
      dplyr::right_join(esdsd_nama_10_gdp, by = c("na_item" = "name")) |>
      dplyr::select(.data$geo, .data$code, .data$na_item, .data$unit, .data$s_adj, tidyselect::matches("[0-9]{4}-Q[0-9]{1}"))

    units <- paste(unique(df3$unit), collapse = "; ")

    writepair("Source",         "Eurostat",        1, info3)
    writepair("Title",          info3$title,       2, info3)
    writepair("Code",           info3$code,        3, info3)
    writepair("Units",          units,             4, info3)
    writepair("Adjustment",     unique(df3$s_adj), 5, info3)
    writepair("Last updated",   info3$lastUpdate,  6, info3)
    writepair("Date extracted", latest_date,       7, info3)

    df3 <- df3 |> dplyr::select(-.data$s_adj)
    openxlsx::writeData(wb, info3$sheet, df3,
                        startRow = 9, keepNA = TRUE, na.string = ":",
                        withFilter = TRUE, headerStyle = headerstyle)
    openxlsx::setColWidths(wb, info3$sheet, c(1, 3, 4), c(20, 40, 20))

    # namq_10_a10 -------------------------------------------------------------

    info4 <- get_table_info("namq_10_a10", metadata, wb)

    df4 <- restatapi::get_eurostat_data(info4$code,
                                        filters = c(country, "CP_MNAC", "CLV15_MNAC", "NSA", "B1G"),
                                        date_filter = starting_year:latest_year,
                                        label = TRUE) |>
      tidyr::pivot_wider(names_from = .data$time, values_from = .data$values) |>
      dplyr::mutate(nace_r2 = forcats::fct_expand(.data$nace_r2, esdsd_namq_10_a10$name)) |>
      tidyr::complete(.data$geo, .data$na_item, .data$unit, .data$s_adj, .data$nace_r2) |>
      dplyr::arrange(.data$unit) |>
      dplyr::right_join(esdsd_namq_10_a10, by = c("nace_r2" = "name")) |>
      dplyr::select(.data$geo, .data$na_item, .data$unit, .data$s_adj, .data$code, .data$nace_r2, tidyselect::matches("[0-9]{4}-Q[0-9]{1}"))

    items <- paste(unique(df4$na_item), collapse = "; ")
    units <- paste(unique(df4$unit), collapse = "; ")

    writepair("Source",         "Eurostat",          1, info4)
    writepair("Title",          info4$title,         2, info4)
    writepair("Code",           info4$code,          3, info4)
    writepair("Item",           unique(df4$na_item), 4, info4)
    writepair("Units",          units,               5, info4)
    writepair("Adjustment",     unique(df4$s_adj),   6, info4)
    writepair("Last updated",   info4$lastUpdate,    7, info4)
    writepair("Date extracted", latest_date,         8, info4)

    df4 <- dplyr::select(df4, -.data$na_item, -.data$s_adj)
    openxlsx::writeData(wb, info4$sheet, df4,
                        startRow = 10, keepNA = TRUE, na.string = ":",
                        withFilter = TRUE, headerStyle = headerstyle)
    openxlsx::setColWidths(wb, info4$sheet, c(1, 2, 4), c(20, 20, 50))

    sheets <- openxlsx::sheets(wb)
    if (reorder_sheets) openxlsx::worksheetOrder(wb) <- order(openxlsx::sheets(wb))
    openxlsx::saveWorkbook(wb, wb_path, overwrite = TRUE)

    cli::cli_text('"{.header {file}}" successfully updated.')
    br()
  }
}

# Helper functions --------------------------------------------------------

get_table_info <- function(table_code, meta, workbook) {

  meta_subset <- meta[meta$code == table_code, ]
  sheetname <- paste0("1. ES ", table_code)
  sheets <- openxlsx::sheets(workbook)
  if (!(sheetname %in% sheets)) {
    openxlsx::addWorksheet(workbook, sheetname, tabColour = "#BDD7EE", zoom = 80)
  }
  openxlsx::deleteData(workbook, sheetname, cols = 1:999, rows = 1:9999, gridExpand = TRUE)

  return(list(title = meta_subset$title,
              code = meta_subset$code,
              lastUpdate = meta_subset$lastUpdate,
              sheet = sheetname,
              workbook = workbook))
}

writepair <- function(x, y, row, info) {

  openxlsx::writeData(info$workbook, info$sheet, x,
                      startRow = row, colNames = FALSE)
  openxlsx::writeData(info$workbook, info$sheet, y,
                      startCol = 2, startRow = row, colNames = FALSE)
}

br <- function() cli::cli_text("")
