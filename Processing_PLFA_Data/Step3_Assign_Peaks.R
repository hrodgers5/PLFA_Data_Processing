require(tidyverse)
require(striprtf)
require(openxlsx)
library(readxl)
require(writexl)


#NAME THE PEAKS

#first, import an excel file with the named peaks, then merge it with your cleaned PLFA data. The file included is according to (JOERGENSEN, 2021), but only includes the peaks present in my samples. Therefore, you should make your own sheet.

peak_names <- read_excel("PLFA Data Processing 2023/peak_matching.xlsx")

#import your combined, clean dataset and merge with the peak names

PLFAs_cleaned <- (read_excel("PLFA Data Processing 2023/Cleaned_Files/PLFA_all_clean.xlsx")) %>% 
  pivot_longer(cols = 3:81, names_to = "Peak", values_to = "nmol")

PLFAs_named <- merge(PLFAs_cleaned, peak_names, by = "Peak")

#Next, pivot wider and sum the totals
PLFAs <- PLFAs_named %>% 
  select(!Peak) %>% 
  pivot_wider(names_from = "Name", values_from = "nmol", values_fn = ~sum(.x, na.rm = TRUE))

#do some math: calculate totals

PLFAs$Gram_Pos <- PLFAs$Firmicutes + PLFAs$Actinobacteria

PLFAs$Fungi <- PLFAs$AMF + PLFAs$Zygomycota + PLFAs$Asco_and_Basidio #+ unspecified_Fungi

PLFAs$Bacteria <- PLFAs$Gram_Neg + PLFAs$Gram_Pos

PLFAs$Total_Micro <- PLFAs$Bacteria + PLFAs$Fungi + PLFAs$Unspecified_microbe

#more math: calculate ratios and rel abundances
PLFAs$F_to_B <- PLFAs$Fungi / PLFAs$Bacteria

#FINAL SAVE!!
write_xlsx(PLFAs, "PLFAs_FINAL_2023.xlsx")
