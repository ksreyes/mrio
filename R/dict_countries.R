#' Countries dictionary
#'
#' Correspondence table of country names and codes. Note that the use of the
#' term "country" here is a convenience and intends no judgment on the political
#' status of the entities listed.
#'
#' @source
#'  * ISO Online Browsing Platform. https://www.iso.org/obp/ui.
#'  * United Nations M49 Standard. https://unstats.un.org/unsd/methodology/m49.
#'  * World Bank Indicators. https://data.worldbank.org.
#'  * Asian Development Bank, *Handbook of Style and Usage*, July 2022. https://www.adb.org/documents/handbook-style-and-usage.
#'  * Asian Development Bank MRIO Team.
#' @format Data frame with columns
#' \describe{
#' \item{iso_num, iso_a2, iso_a3}{Numeric, two-letter, and three-letter country
#'   codes according to the International Standard for country codes (ISO 3166).}
#' \item{un_name}{Country name according to the United Nations.}
#' \item{adb_code, adb_name}{Country three-letter codes and names according
#'   to the Asian Development Bank.}
#' \item{mrio_ind, mrio_ind62, mrio_code, mrio_name}{Country indexing, codes,
#'   and names according to the ADB Multiregional Input-Output Tables.
#'   `mrio_ind` is the index for the 72-country version while `mrio_ind62` is
#'   for the 62-country version.}
#' \item{wb_name, wb_region}{Country names and regional classifications
#'   according to the World Bank.}
#' \item{code, name}{Consolidated country codes and names based on the ADB and
#'   MRIO conventions and, if missing, the ISO and UN conventions.}
#' \item{region}{Consolidated regional classification that is essentially the
#'   World Bank classification with the inclusion of Taipei,China.}
#' \item{continent}{Country continent; manually assigned.}
#' }
"dict_countries"
