#' Add MRIO labels
#'
#' Convert between various labels for countries and sectors in the
#' Multiregional Input-Output (MRIO) tables.
#'
#' @param .data A data frame.
#' @param .col Reference column.
#' @param from Field type of reference column.
#' @param to Target field type.
#' @param replace If TRUE, target replaces reference column. Default = FALSE
#'   adds target as a new column next to the reference column.
#'
#' @return An object of the same type as `.data`.
#' @details See `?dict_countries` and `?dict_sectors` for valid field types to
#'   convert between.
#'
#' @examples
#' df <- data.frame(
#'   t = c(2022, 2022, 2022),
#'   s = c(11, 25, 46),
#'   value = c(100, 90, 120)
#' )
#' df |> add_mrio_labels(s, from = "mrio_ind", to = "mrio_name")
#'
#' # A named vector renames the new column.
#' df |> add_mrio_labels(s, from = "mrio_ind", to = c("country" = "mrio_name"))
#'
#' @export
add_mrio_labels <- function(.data, .col, from, to, replace = FALSE) {

  df_string <- deparse(substitute(.data))
  col_string <- deparse(substitute(.col))

  # Input errors

  if(!is.data.frame(.data)) {
    cli::cli_abort(c(
            "{.var .data} is not a {.cls data.frame} or similar.",
      "x" = "You've supplied a {.cls {class(.data)}}."
    ))
  }

  errors <- c()

  if (!(col_string %in% colnames(.data))) {
    errors <- c(errors,
      "!" = "Column `{col_string}` is not in `{df_string}`."
    )
  }

  if (!(from %in% c(colnames(dict_countries), colnames(dict_sectors)))) {
    errors <- c(errors,
      "!" = "`{from}` is not in any dictionary. See `?dict_countries` and `?dict_sectors` for valid fields."
    )
  }

  for (field in to) {

    if (from %in% colnames(dict_countries) & !(field %in% colnames(dict_countries))) {
      errors <- c(errors,
        "!" = "`{field}` is not in the same dictionary as `{from}`. See `?dict_countries` for valid fields."
      )
    }

    if (from %in% colnames(dict_sectors) & !(field %in% colnames(dict_sectors))) {
      errors <- c(errors,
        "!" = "`{field}` is not in the same dictionary as `{from}`. See `?dict_sectors` for valid fields."
      )
    }
  }

  if(length(errors) > 0) cli::cli_abort(errors)

  # Determine new columns

  old_columns <- colnames(.data)
  col_position <- which(old_columns == col_string)
  new_columns <- append(old_columns, to, after = col_position)

  if (replace) {
    new_columns <- new_columns[new_columns != col_string]
  }

  df <- df |>
    dplyr::rename("{ from }" := {{ .col }}) |>
    dplyr::left_join(dict_countries) |>
    dplyr::rename("{{ .col }}" := { from }) |>
    dplyr::select(dplyr::all_of(new_columns)) |>
    suppressMessages()

  if (all(is.na(df[names(to)]))) {
    cli::cli_warn("Multiple NAs in created column/s suggest dictionary mismatch.")
  }

  return(df)
}
