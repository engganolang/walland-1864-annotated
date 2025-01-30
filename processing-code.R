# G. Rajeg, 2025 (University of Oxford/Udayana University)

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
         Low_Language_LL = str_replace_all(Low_Language_LL, "\\-\\s+", "-")) |> 
  mutate(ID = 1:nrow(walland1)) |> 
  select(ID, everything())

# Split the remark by carriage return and new line markers
walland2 <- walland1 |> 
  mutate(Remarks_Split = str_split(Remarks, "\\r\\n")) |> 
  unnest_longer(Remarks_Split) |> 
  mutate(Remarks_Category = "Rm_Others",
         Remarks_Category = if_else(str_detect(Remarks_Split, "^HL\\:"),
                                    "Rm_HL",
                                    Remarks_Category),
         Remarks_Category = if_else(str_detect(Remarks_Split, "^LL\\:"),
                                    "Rm_LL",
                                    Remarks_Category),
         Remarks_Category = if_else(str_detect(Remarks_Split, "^OE\\s"),
                                    "Rm_Sound_Changes",
                                    Remarks_Category),
         Remarks_Category = if_else(str_detect(Remarks_Split, "^ME\\:"),
                                    "Rm_Contemporary_Enggano_Translation",
                                    Remarks_Category),
         Remarks_Category = if_else(str_detect(Remarks_Split, "^K.+hler"),
                                    "Rm_Kahler",
                                    Remarks_Category),
         Remarks_Category = factor(Remarks_Category,
                                   levels = c("Rm_HL",
                                              "Rm_LL",
                                              "Rm_Sound_Changes",
                                              "Rm_Contemporary_Enggano_Translation",
                                              "Rm_Kahler",
                                              "Rm_Others")))

## multiple remarks of the same Remarks_Category are joined into one line
### then pivoted into wider table with Remarks_Category values as column names
walland3 <- walland2 |> 
  group_by(Dutch_Orig, English_gloss, High_Language_HL, Contemporary_Equiv_HL, 
           Low_Language_LL, Contemporary_Equiv_LL, Remarks, pagenum, 
           Remarks_Category) |> 
  mutate(Remarks_Joined = str_c(Remarks_Split, collapse = "__")) |> 
  ungroup() |> 
  select(-Remarks_Split,
         -Remarks) |> 
  distinct() |> 
  pivot_wider(names_from = "Remarks_Category", 
              values_from = "Remarks_Joined", 
              values_fill = "") |> 
  mutate(across(matches("^Rm_"), ~str_replace_all(., "__", " ; "))) |> 
  relocate(Rm_LL, .after = Rm_HL) |> 
  relocate(Rm_Sound_Changes, .after = Rm_LL) |> 
  relocate(Rm_Others, .after = Rm_Kahler)

## multiple remarks of the same Remarks_Category have their own rows (long-table format)
walland4 <- walland2 |> 
  relocate(pagenum, .after = Remarks_Category) |> 
  relocate(Remarks, .after = Remarks_Category) |> 
  rename(Remarks_Orig = Remarks) |> 
  arrange(ID, Remarks_Category)

# googledrive::drive_create(name = "Walland-1864-annotated",
#                           path = "https://drive.google.com/drive/u/0/folders/1GEcnGtmDGHdJO2AH8wWqXTd_1aaNcE8W",
#                           type = "spreadsheet")
# Created Drive file:
#   • Walland-1864-annotated <id: 1VCUJNGysXkmZ31eOTruHjfFYrp0_o5qCvOe3Ule67A4>
#   With MIME type:
#   • application/vnd.google-apps.spreadsheet

googlesheets4::sheet_write(walland3, ss = "1VCUJNGysXkmZ31eOTruHjfFYrp0_o5qCvOe3Ule67A4", sheet = "rm_wide")
googlesheets4::sheet_write(walland4, ss = "1VCUJNGysXkmZ31eOTruHjfFYrp0_o5qCvOe3Ule67A4", sheet = "rm_long")
write_csv(walland3, "data-output/walland-1864-annotated.csv")
write_tsv(walland3, "data-output/walland-1864-annotated.tsv")
write_rds(walland3, "data-output/walland-1864-annotated.rds")

write_csv(walland4, "data-output/walland-1864-annotated-remark-long.csv")
write_tsv(walland4, "data-output/walland-1864-annotated-remark-long.tsv")
write_rds(walland4, "data-output/walland-1864-annotated-remark-long.rds")