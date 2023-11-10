#' Get MRIO indices, codes, and names
#'
#' Functions for easily converting one MRIO key to another.
#'
#' @param key An MRIO country or sector index, code, or name. Accepts vectors
#'   and lists.
#' @param unmatched What value to set for unmatched keys? Default is `NA`.
#'
#' @return An object of the same size as `key`.
#'
#' @examples
#' get_country_name(41)
#'
#' # Multiple inputs with different key types (including the target key)
#' # may be entered
#' get_country_ind(c("PRC", "Japan", "IND", 62))
#'
#' # Inputs are not case sensitive
#' get_country_code("philippines")
#'
#' @name get-functions

#' @export
#' @rdname get-functions
get_country_name <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_country(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_countries$mrio_name[match])
    }}
  return(result)
}

#' @export
#' @rdname get-functions
get_country_ind <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_country(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_countries$mrio_ind[match])
    }}
  return(result)
}

#' @export
#' @rdname get-functions
get_country_code <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_country(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_countries$mrio_code[match])
    }}
  return(result)
}

#' @export
#' @rdname get-functions
get_iso_num <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_country(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_countries$iso_num[match])
    }}
  return(result)
}

#' @export
#' @rdname get-functions
get_iso_a2 <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_country(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_countries$iso_a2[match])
    }}
  return(result)
}

#' @export
#' @rdname get-functions
get_iso_a3 <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_country(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_countries$iso_a3[match])
    }}
  return(result)
}

#' @export
#' @rdname get-functions
get_sector_name <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_sector(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_sectors$c_name[match])
    }}
  return(result)
}

#' @export
#' @rdname get-functions
get_sector_shortname <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_sector(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_sectors$c_name_short[match])
    }}
  return(result)
}

#' @export
#' @rdname get-functions
get_sector_ind <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_sector(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_sectors$c_ind[match])
    }}
  return(result)
}

#' @export
#' @rdname get-functions
get_sector_code <- function(key, unmatched = NA) {
  result <- c()
  for (k in key) {
    match <- match_sector(k)
    if (length(match) == 0) {
      result <- c(result, unmatched)
    } else {
      result <- c(result, dict_sectors$c_abv[match])
    }}
  return(result)
}

# Helper functions --------------------------------------------------------

match_country <- function(j) {

  inds <- dict_countries$mrio_ind
  names <- dict_countries$mrio_name |> tolower()
  codes <- dict_countries$mrio_code |> tolower()
  iso_num <- dict_countries$iso_num
  iso_a2 <- dict_countries$iso_a2 |> tolower()
  iso_a3 <- dict_countries$iso_a3 |> tolower()

  if (is.na(as.numeric(j)) |> suppressWarnings()) {
    j <- tolower(j)
    match <- c(which(names == j), which(codes == j), which(iso_a2 == j), which(iso_a3 == j))
  } else {
    match <- c(which(inds == as.numeric(j)), which(iso_num == as.numeric(j)))
  }

  return(match[1])
}

match_sector <- function(j) {

  inds <- dict_sectors$c_ind
  names <- dict_sectors$c_name |> tolower()
  shortnames <- dict_sectors$c_name_short |> tolower()
  codes <- dict_sectors$c_abv |> tolower()

  if (is.na(as.numeric(j)) |> suppressWarnings()) {
    j <- tolower(j)
    match <- c(which(names == j), which(shortnames == j), which(codes == j))
  } else {
    match <- which(inds == as.numeric(j))
  }

  return(match[1])
}
