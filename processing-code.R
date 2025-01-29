library(tidyverse)
library(readxl)
library(googledrive)
library(googlesheets4)

sheets <- paste("Sheet", 1:22, sep = "")
files <- dir(path = "data-source/", pattern = ".xlsx", full.names = TRUE)

df <- map2(.x = sheets, 
           .y = files, 
           ~read_xlsx(path = .y, 
                      sheet = .x, 
                      col_names = c("Dutch_Orig", "English_gloss", 
                                    "High_Language_HL", "Contemporary_Equiv_HL", 
                                    "Low_Language_LL", "Contemporary_Equiv_LL", 
                                    "Remarks"))) |> 
  map2(.y = 1:22,
       ~mutate(.x, pagenum = .y))

walland <- df |> 
  list_rbind() |> 
  mutate(across(where(is.character), ~replace_na(., "")))

walland1 <- walland |> 
  # replace \r\n
  mutate(Dutch_Orig = str_replace_all(Dutch_Orig, "(?<!\\-)\\r\\n", " "),
         Dutch_Orig = str_replace_all(Dutch_Orig, "(?<=\\-)\\r\\n", ""),
         English_gloss = str_replace_all(English_gloss, "([[:punct:]])\\r\\n", "\\1"),
         English_gloss = str_replace_all(English_gloss, "(?<![[:punct:]])\\r\\n", " "),
         High_Language_HL = str_replace_all(High_Language_HL, "(?<=[[:punct:]])\\r\\n", ""),
         High_Language_HL = str_replace_all(High_Language_HL, "(?<![[:punct:]])\\r\\n", " "),
         Contemporary_Equiv_HL = str_replace_all(Contemporary_Equiv_HL, "(\\~)\\r\\n", "\\1"),
         Contemporary_Equiv_HL = str_replace_all(Contemporary_Equiv_HL, "(\\!)\\r\\n", "\\1; "),
         Low_Language_LL = str_replace_all(Low_Language_LL, "(?<=\\-)\\r\\n", "")) |> 
  mutate(Dutch_Orig = str_replace_all(Dutch_Orig, "\\-\\s+", "-"),
         English_gloss = str_replace_all(English_gloss, "\\-\\s+", "-"),
         High_Language_HL = str_replace_all(High_Language_HL, "\\-\\s+", "-"),
         Low_Language_LL = str_replace_all(Low_Language_LL, "\\-\\s+", "-"))

# googledrive::drive_create(name = "Walland-1864-annotated",
#                           path = "https://drive.google.com/drive/u/0/folders/1GEcnGtmDGHdJO2AH8wWqXTd_1aaNcE8W",
#                           type = "spreadsheet")
# Created Drive file:
#   • Walland-1864-annotated <id: 1VCUJNGysXkmZ31eOTruHjfFYrp0_o5qCvOe3Ule67A4>
#   With MIME type:
#   • application/vnd.google-apps.spreadsheet

googlesheets4::sheet_write(walland1, ss = "1VCUJNGysXkmZ31eOTruHjfFYrp0_o5qCvOe3Ule67A4", sheet = "main")
write_csv(walland1, "data-output/walland-1864-annotated.csv")
write_tsv(walland1, "data-output/walland-1864-annotated.tsv")
write_rds(walland1, "data-output/walland-1864-annotated.rds")