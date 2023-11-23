
# Which MRIO sectors are goods and which are services

sector_classes <- dict_sectors |>
  dplyr::mutate(
    c_ind = paste0("c", c_ind),
    class = c(rep("good", 17), rep("service", 18))
  ) |>
  dplyr::select(i = c_ind, class)

usethis::use_data(sector_classes, internal = TRUE, overwrite = TRUE)
