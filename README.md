# mrio

<!-- badges: start -->
[![R-CMD-check](https://github.com/ksreyes/mrio/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/ksreyes/mrio/actions/workflows/R-CMD-check.yaml)
<!-- badges: end -->

## Overview

Tools for working with the Asian Development Bank (ADB) [Multiregional Input-Output (MRIO) Database](https://kidb.adb.org/mrio), a time series of intercountry input-output tables available for up to 72 countries and disaggregated into 35 sectors.

### check

`check(path, precision = 8, export = TRUE, limit = 10)`

Performs basic integrity checks on an MRIO Excel file or a directory of such files. 

### get_*

* `get_country_name(key, unmatched = NA)`
* `get_country_ind(key, unmatched = NA)`
* `get_country_code(key, unmatched = NA)`
* `get_sector_name(key, unmatched = NA)`
* `get_sector_shortname(key, unmatched = NA)`
* `get_sector_ind(key, unmatched = NA)`
* `get_sector_code(key, unmatched = NA)`

Functions for converting an MRIO key to another.

### dict_countries

Correspondence table of country names and codes, including MRIO country keys where applicable.

### dict_sectors

Correspondence table of sector keys in the MRIO.

## Installation

```r
devtools::install_github("ksreyes/mrio")
```
