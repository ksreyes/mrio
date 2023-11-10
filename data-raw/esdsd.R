
tables <- c("nama_10_gdp", "nama_10_a64", "namq_10_gdp", "namq_10_a10")

esdsd_nama_10_gdp <- restatapi::get_eurostat_dsd(tables[1]) |>
  dplyr::filter(concept == "na_item") |>
  dplyr::select(code, name)

esdsd_nama_10_a64 <- restatapi::get_eurostat_dsd(tables[2]) |>
  dplyr::filter(concept == "nace_r2") |>
  dplyr::select(code, name)

esdsd_namq_10_gdp <- restatapi::get_eurostat_dsd(tables[3]) |>
  dplyr::filter(concept == "na_item") |>
  dplyr::select(code, name)

esdsd_namq_10_a10 <- restatapi::get_eurostat_dsd(tables[4]) |>
  dplyr::filter(concept == "nace_r2") |>
  dplyr::select(code, name)

es_countries <- restatapi::get_eurostat_dsd(tables[1]) |>
  dplyr::filter(concept == "geo") |>
  dplyr::select(code, name) |>
  dplyr::slice(8:45)

usethis::use_data(
    esdsd_nama_10_gdp, esdsd_nama_10_a64, esdsd_namq_10_gdp, esdsd_namq_10_a10,
    es_countries,
    internal = TRUE, overwrite = TRUE
  )
