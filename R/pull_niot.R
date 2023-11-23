#' Pull national IOT from MRIO
#'
#' @description Pulls a national input-output table for a given year from the
#'   Multi-regional Input-Output database and saves it in a control totals
#'   workbook.
#'
#' @param workbook Control totals workbook or directory of workbooks. File names
#'   must start with the uppercase three-letter country code. All other files in
#'   the directory are ignored.
#' @param source Directory of MRIO tables. Must be stored locally.
#' @param year Years of NIOTs to pull. If `NULL`, infers years from sheets with
#'   name pattern "5. MRIO 20XX".
#'
#' @details Workbook must have at least one sheet with name pattern "5. MRIO
#'   20XX" with the appropriate formatting. Function overwrites any existing
#'   data.
#'
#' @export
pull_niot <- function(workbook, source, year = NULL) {

  N <- 35; f <- 5

  if (file.info(workbook)$isdir) {
    workbooks_all <- list.files(workbook, pattern = "^[A-Z]{3}.*(xls|xlsx)$")
  } else {
    workbooks_all <- workbook
  }

  # Get workbooks with appropriate sheets and relevant years

  cli::cli_text("")
  cli::cli_text("Reading files...")
  cli::cli_text("")

  years <- c(); workbooks <- c()

  for (wb in workbooks_all) {

    if (file.info(workbook)$isdir) wb_path <- file.path(workbook, wb)
    wb_file <- openxlsx::loadWorkbook(wb_path) |> suppressWarnings()
    wb_file_sheets <- openxlsx::sheets(wb_file)

    relevant_sheets <- regmatches(
        wb_file_sheets,
        regexpr("(5. MRIO 20)[0-9]{2}", wb_file_sheets)
      )

    if (length(relevant_sheets) == 0) next
    workbooks <- c(workbooks, wb)

    wb_file_years <- regmatches(relevant_sheets, regexpr("[0-9]{4}", relevant_sheets)) |>
      as.numeric() |> unique()

    years <- c(years, wb_file_years)
  }

  if (length(workbooks) == 0) {
    cli::cli_abort("No valid workbooks found.")
    cli::cli_text("")
  }

  cli::cli_text("The following valid workbooks were found:")
  for (wb in workbooks) {
    cli::cli_bullets(c("*" = "{wb}"))
  }
  cli::cli_text("")

  countries <- regmatches(workbooks, regexpr("[A-Z]{3}", workbooks)) |> unique()
  if (is.null(year)) year <- years

  # Load relevant MRIOs

  mrios <- list.files(source, pattern = "^[^~].*(MRIO|mrio|Mrio).*(xls|xlsx)$")
  mrio_years <- regmatches(mrios, regexpr("[0-9]{4}", mrios))
  mrios <- mrios[match(year, mrio_years)]

  for (i in 1:length(mrios)) {

    cli::cli_text("Working on {year[i]}...")
    cli::cli_text("")

    mrio <- readxl::read_excel(
        file.path(source, mrios[i]),
        range = readxl::cell_limits(c(6, 2), c(NA, NA)),
        progress = FALSE
      ) |>
      suppressMessages() |>
      suppressWarnings()

    for (j in 1:length(workbooks)) {

      niot <- extract_niot(mrio, countries[j])

      if (file.info(workbook)$isdir) wb_path_j <- file.path(workbook, workbooks[j])
      wb_j <- openxlsx::loadWorkbook(wb_path_j)
      sheet <- paste0("5. MRIO ", year[i])

      openxlsx::writeData(
        wb_j, sheet, niot[1:N, 4:(4 + N + f)],
        startCol = 5, startRow = 7, colNames = FALSE
      )

      openxlsx::writeData(
        wb_j, sheet, niot[(N + 3):(N + 4), 4:(4 + N + f)],
        startCol = 5, startRow = 44, colNames = FALSE
      )

      openxlsx::writeData(
        wb_j, sheet, niot[(N + 6):(N + 11), 4:(4 + N + f)],
        startCol = 5, startRow = 47, colNames = FALSE
      )

      openxlsx::saveWorkbook(wb_j, wb_path_j, overwrite = TRUE)
    }}

  cli::cli_text("All tasks done!")
  cli::cli_text("")
}

# Helper functions --------------------------------------------------------

extract_niot <- function(mrio_source, country_code) {

  country_code <- get_country_code(country_code)

  # Keep relevant columns
  niot1 <- mrio_source |>
    dplyr::select(.data$`...1`, .data$`...2`, .data$`...3`, tidyselect::starts_with(country_code))
  colnames(niot1) <- c("desc", "s", "i", unlist(niot1[1, ])[-(1:3)])
  niot1 <- niot1[-1, ]
  niot1 <- dplyr::mutate(niot1, dplyr::across(tidyselect::matches("[cF][0-9]+"), as.numeric))

  # Get domestic use matrix
  domestic <- subset(niot1, .data$s == country_code)
  domestic <- dplyr::bind_rows(
    domestic,
      cbind(
        desc = "TOTAL DOMESTIC AT BASIC PRICE",
        s = "ToT",
        i = "D",
        dplyr::summarise(domestic, dplyr::across(dplyr::matches("[cF][0-9]+"), sum))
      ))

  # Imports
  imports <- niot1 |>
    dplyr::left_join(sector_classes) |>
    subset(.data$s != country_code & !is.na(.data$s) & .data$s != "ToT") |>
    dplyr::summarise(dplyr::across(tidyselect::matches("[cF][0-9]+"), sum), .by = class) |>
    dplyr::mutate(desc = c("GOODS", "SERVICES"), s = "M", i = c("G", "S")) |>
    suppressMessages()
  imports <- dplyr::bind_rows(
      cbind(
        desc = "IMPORTS OF GOODS AND SERVICES, FOB", s = "ToT", i = "M",
        dplyr::summarise(imports, dplyr::across(tidyselect::matches("[cF][0-9]+"), sum))
      ),
      imports
    ) |>
    dplyr::select(.data$desc, .data$s, .data$i, tidyselect::matches("[cF][0-9]+"))

  # Value added and totals
  gva <- niot1 |>
    subset((is.na(.data$s) | .data$s == "ToT") & !is.na(.data$i)) |>
    dplyr::mutate(
      s = "ToT",
      desc = dplyr::case_when(
        .data$desc == "Intermediate input total" ~ "TOTAL AT BASIC PRICE",
        .data$desc == "TOTAL" ~ "TOTAL OUTPUT",
        .default = .data$desc
      ))

  # Exports
  exports <- mrio_source |>
    dplyr::select(.data$`...1`, .data$`...2`, .data$`...3`, !tidyselect::starts_with(country_code))
  colnames(exports)[1:3] <- c("desc", "s", "i")
  exports <- exports[-1, -ncol(exports)]
  exports <- exports |>
    subset(.data$s == country_code) |>
    dplyr::mutate(dplyr::across(tidyselect::matches("[A-Za-z]{3}...[0-9]+"), as.numeric)) |>
    dplyr::rowwise(.data$desc, .data$s, .data$i) |>
    dplyr::summarise(X = sum(dplyr::c_across(tidyselect::matches("[A-Za-z]{3}...[0-9]+")))) |>
    dplyr::ungroup() |>
    suppressMessages()
  exports <- exports |>
    dplyr::add_row(desc = "TOTAL DOMESTIC AT BASIC PRICE", s = "ToT", i = "D", X = sum(exports$X)) |>
    dplyr::bind_rows(
      cbind(imports[, 1:3], X = 0),
      cbind(gva[, 1:3], X = c(sum(exports$X), rep(0, 6), sum(exports$X)))
    )

  # Consolidate
  niot <- domestic |>
    dplyr::bind_rows(imports, gva) |>
    dplyr::left_join(exports) |>
    dplyr::rowwise(.data$desc, .data$s, .data$i) |>
    dplyr::mutate(TOTAL = sum(dplyr::c_across(tidyselect::where(is.numeric)))) |>
    dplyr::ungroup() |>
    suppressMessages()

  return(niot)
}
