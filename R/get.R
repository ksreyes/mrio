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

  inds <- dict_countries$mrio_ind
  names <- dict_countries$mrio_name |> tolower()
  codes <- dict_countries$mrio_code |> tolower()

  result <- c()

  for (k in key) {
    if (is.na(as.numeric(k)) |> suppressWarnings()) {
      k <- tolower(k)
      match <- c(which(names == k), which(codes == k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_countries$mrio_name[match])
      }
    } else {
      match <- which(inds == as.numeric(k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_countries$mrio_name[match])
      }
    }
  }
  return(result)
}

#' @export
#' @rdname get-functions
get_country_ind <- function(key, unmatched = NA) {

  inds <- dict_countries$mrio_ind
  names <- dict_countries$mrio_name |> tolower()
  codes <- dict_countries$mrio_code |> tolower()

  result <- c()

  for (k in key) {
    if (is.na(as.numeric(k)) |> suppressWarnings()) {
      k <- tolower(k)
      match <- c(which(names == k), which(codes == k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, inds[match])
      }
    } else {
      match <- which(inds == as.numeric(k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, inds[match])
      }
    }
  }
  return(result)
}

#' @export
#' @rdname get-functions
get_country_code <- function(key, unmatched = NA) {

  inds <- dict_countries$mrio_ind
  names <- dict_countries$mrio_name |> tolower()
  codes <- dict_countries$mrio_code |> tolower()

  result <- c()

  for (k in key) {
    if (is.na(as.numeric(k)) |> suppressWarnings()) {
      k <- tolower(k)
      match <- c(which(names == k), which(codes == k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_countries$mrio_code[match])
      }
    } else {
      match <- which(inds == as.numeric(k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_countries$mrio_code[match])
      }
    }
  }
  return(result)
}

#' @export
#' @rdname get-functions
get_sector_name <- function(key, unmatched = NA) {

  inds <- dict_sectors$c_ind
  names <- dict_sectors$c_name |> tolower()
  shortnames <- dict_sectors$c_name_short |> tolower()
  codes <- dict_sectors$c_abv |> tolower()

  result <- c()

  for (k in key) {
    if (is.na(as.numeric(k)) |> suppressWarnings()) {
      k <- tolower(k)
      match <- c(which(names == k), which(shortnames == k), which(codes == k))
      match <- c(which(names == k), which(codes == k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_sectors$c_name[match])
      }
    } else {
      match <- which(inds == as.numeric(k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_sectors$c_name[match])
      }
    }
  }
  return(result)
}

#' @export
#' @rdname get-functions
get_sector_shortname <- function(key, unmatched = NA) {

  inds <- dict_sectors$c_ind
  names <- dict_sectors$c_name |> tolower()
  shortnames <- dict_sectors$c_name_short |> tolower()
  codes <- dict_sectors$c_abv |> tolower()

  result <- c()

  for (k in key) {
    if (is.na(as.numeric(k)) |> suppressWarnings()) {
      k <- tolower(k)
      match <- c(which(names == k), which(shortnames == k), which(codes == k))
      match <- c(which(names == k), which(codes == k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_sectors$c_name_short[match])
      }
    } else {
      match <- which(inds == as.numeric(k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_sectors$c_name_short[match])
      }
    }
  }
  return(result)
}

#' @export
#' @rdname get-functions
get_sector_ind <- function(key, unmatched = NA) {

  inds <- dict_sectors$c_ind
  names <- dict_sectors$c_name |> tolower()
  shortnames <- dict_sectors$c_name_short |> tolower()
  codes <- dict_sectors$c_abv |> tolower()

  result <- c()

  for (k in key) {
    if (is.na(as.numeric(k)) |> suppressWarnings()) {
      k <- tolower(k)
      match <- c(which(names == k), which(shortnames == k), which(codes == k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, inds[match])
      }
    } else {
      match <- which(inds == as.numeric(k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, inds[match])
      }
    }
  }
  return(result)
}

#' @export
#' @rdname get-functions
get_sector_code <- function(key, unmatched = NA) {

  inds <- dict_sectors$c_ind
  names <- dict_sectors$c_name |> tolower()
  shortnames <- dict_sectors$c_name_short |> tolower()
  codes <- dict_sectors$c_abv |> tolower()

  result <- c()

  for (k in key) {
    if (is.na(as.numeric(k)) |> suppressWarnings()) {
      k <- tolower(k)
      match <- c(which(names == k), which(shortnames == k), which(codes == k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_sectors$c_abv[match])
      }
    } else {
      match <- which(inds == as.numeric(k))
      if (length(match) == 0) {
        result <- c(result, unmatched)
      } else {
        result <- c(result, dict_sectors$c_abv[match])
      }
    }
  }
  return(result)
}
