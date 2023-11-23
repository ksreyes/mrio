# mrio

<!-- badges: start -->
[![R-CMD-check](https://github.com/ksreyes/mrio/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/ksreyes/mrio/actions/workflows/R-CMD-check.yaml)
<!-- badges: end -->

## Overview

Tools for working with the Asian Development Bank (ADB) [Multiregional Input-Output (MRIO) Database](https://kidb.adb.org/mrio), a time series of intercountry input-output tables available for up to 72 countries and disaggregated into 35 sectors.

### check

`check(workbook, reorder_sheets = TRUE, starting_year = 1995)`

Performs basic integrity checks on an MRIO Excel file or a directory of such files. 

### sectorize_comtrade

`sectorize_comtrade(path)`

Aggregates raw product-level data from a UN Comtrade bulk download to the MRIO sectors.

### download_eurostat

`download_eurostat(path, reorder_sheets = TRUE, starting_year = 1995)`

Downloads macroeconomic data from Eurostat and saves them in a control totals workbook.

### pull_niot

Pulls a national input-output table for a given year from the Multi-regional Input-Output database and saves it in a control totals workbook.

### get_*

* `get_country_name(key, unmatched = NA)`
* `get_country_ind(key, unmatched = NA)`
* `get_country_code(key, unmatched = NA)`
* `get_iso_num(key, unmatched = NA)`
* `get_iso_a2(key, unmatched = NA)`
* `get_iso_a3(key, unmatched = NA)`
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
