library(tidyverse)
library(readxl)
sheets <- paste("Sheet", 1:22, sep = "")
files <- dir(pattern = ".xlsx")

df <- map2(.x = sheets, .y = files, ~read_xlsx(path = .y, sheet = .x, col_names = c("Dutch_Orig", "English_gloss", "High_Language_HL", "Contemporary_Equiv_HL", "Low_Language_LL", "Contemporary_Equiv_LL", "Remarks")))



