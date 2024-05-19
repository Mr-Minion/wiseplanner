##############################
### Holiday datapull sheet ###
##############################

# 1. Description------
# This script pulls, sorts and modifies holiday calendar data over multiple years in an English and German version.
# The output are automatically generated as an English and German Excel file.
# Import into PowerBI is possible as well.

## 1.1 Clear space, set options-------------------------------------------------
rm(list = ls())
#options(java.parameters = "- Xmx1024m")
Sys.setlocale("LC_CTYPE", "english")
Sys.setlocale("LC_TIME", "English")   #Adjusting date to english: Dez instead of Dec!

## 1.2 Load/install dependencies------------------------------------------------
packages <- c("bookdown", "data.table", "dplyr", "flextable", "furrr", "ggplot2", "ggrepel", "janitor",
              "lubridate", "magrittr", "naniar", "openxlsx", "plotly", "readxl", "rvest", "tidyr", "tidyverse")

# #install packages (only needed when script used for the first time !takes a while!)
# for (i in 1: length(packages)) {
# install.packages(packages[i])}

## For installing tabulizer, github-access is required, which is access limited.
## The following code allows remote access to github to install the package
# install.packages("remotes")
# remotes::install_github("ropensci/tabulizer")

## Load packages(!always needed!)
to.install <- setdiff(packages, rownames(installed.packages()))
if (length(to.install) > 0) {
  install.packages(to.install)}

lapply(packages, library, character.only = TRUE)

## 1.3 Define years-------------------------------------------------------------------------
current_year <- as.integer(format(Sys.Date(), "%Y"))
years <- c((current_year - 4):(current_year + 4))


# 2. Holiday preparation for English calendar Visp------------------------------------------
## 2.1 Access calendars within the homepage Feiertagskalender.ch----------------------------
combined_years_en <- data.frame()

for (year in years) {
  url <- paste0("https://www.feiertagskalender.ch/index.php?geo=2840&jahr=", year, "&hl=en")
  content <- read_html(url)
  year_holidays <- content %>% 
    html_table(fill = TRUE)
  combined_years_en <- bind_rows(year_holidays, combined_years_en)
}

colnames(combined_years_en) <- colnames(combined_years_en) %>%
  str_replace_all(" ", "_") %>% 
  tolower()

## 2.2 Cleaning the df of unnecessary holidays---------------------------------------
pattern_en <- "^(January|February|March|April|May|June|July|August|September|October|November|December)"
cleaned_holidays_en <- combined_years_en %>%
  filter(!(grepl(pattern_en, class) | class =="5" | grepl("%", class) | grepl("Sa", day) | grepl("Su", day)))

cleaned_holidays_en$date <- as.Date(cleaned_holidays_en$date, format = "%d.%m.%Y")

## 2.3 Add a Bridge day after Ascension Day--------------------------------
ascension_bridgeday_row_en <- cleaned_holidays_en %>%
  filter (public_holiday == "Ascension Day") %>%
  mutate(date = date + 1, public_holiday = "Ascension Day Bridge Day")

siteVisp_holidays_en <- bind_rows(cleaned_holidays_en, ascension_bridgeday_row_en)


# 3. Holiday preparation for English calendar Visp------------------------------------------
## 3.1 Access calendars within the homepage Feiertagskalender.ch----------------------------
combined_years_de <- data.frame()

for (year in years) {
  url <- paste0("https://www.feiertagskalender.ch/index.php?geo=2840&jahr=", year, "&hl=de")
  content <- read_html(url)
  year_holidays <- content %>% 
    html_table(fill = TRUE)
  combined_years_de <- bind_rows(year_holidays, combined_years_de)
}

colnames(combined_years_de) <- colnames(combined_years_de) %>%
  str_replace_all(" ", "_") %>% 
  tolower()

## 3.2 Cleaning the df of unnecessary holidays---------------------------------------
pattern_de <- "^(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)"
cleaned_holidays_de <- combined_years_de %>%
  filter(!(grepl(pattern_de, klasse) | klasse =="5" | grepl("%", klasse) | grepl("Sa", tag) | grepl("So", tag)))

cleaned_holidays_de$datum <- as.Date(cleaned_holidays_de$datum, format = "%d.%m.%Y")

## 3.3 Add a Bridge day after Ascension Day--------------------------------
ascension_bridgeday_row_de <- cleaned_holidays_de %>%
  filter (feiertag == "Auffahrt") %>%
  mutate(datum = datum + 1, feiertag = "Auffahrt Brückentag")

siteVisp_holidays_de <- bind_rows(cleaned_holidays_de, ascension_bridgeday_row_de)
