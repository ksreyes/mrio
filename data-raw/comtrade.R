
# HS/BEC/SITC mappings

hsbecsitc <- readxl::read_excel(
    "../Automation Project/2021 MRIO-72 Bilaterals_XXX.xlsx",
    sheet = "HSBECSITC"
  )

# BEC categories

hsbec <- readxl::read_excel(
    "../Automation Project/2021 MRIO-72 Bilaterals_XXX.xlsx",
    sheet = "HS17BEC5"
  )

# HS02/CPC mapping

hs02cpc <- readxl::read_excel(
    "../Automation Project/2021 MRIO-72 Bilaterals_XXX.xlsx",
    sheet = "CPC"
  ) |>
  dplyr::select(HS02 = HS2002, CPC = CPCcode...7)

# CPC/ISIC mapping

cpcisic <- readxl::read_excel(
    "../Automation Project/2021 MRIO-72 Bilaterals_XXX.xlsx",
    sheet = "ISIC"
  ) |>
  dplyr::select(CPC = CPCcode, ISIC = ISICcode)

# ISIC/MRIO mapping

isicmrio <- readxl::read_excel(
    "../Automation Project/2021 MRIO-72 Bilaterals_XXX.xlsx",
    sheet = "MRIO",
    skip = 2
  ) |>
  dplyr::select(ISIC = Code...1, MRIO = No.)

# HS classification codes

hscodes <- tibble::tribble(
  ~code,  ~edition,
  "H0",   "HS92",
  "H1",   "HS96",
  "H2",   "HS02",
  "H3",   "HS07",
  "H4",   "HS12",
  "H5",   "HS17",
  "H6",   "HS22"
)

sysdata_filenames <- load("R/sysdata.rda")
save(
    list = c(sysdata_filenames, "hsbecsitc", "hsbec", "hs02cpc", "cpcisic", "isicmrio", "hscodes"),
    file = "R/sysdata.rda",
    compress = "xz"
  )
