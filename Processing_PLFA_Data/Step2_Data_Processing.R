require(tidyverse)
require(textreadr)
require(openxlsx)
library(readxl)
require(writexl)

#### DATA PROCESSING ####

# I am using this equation, taken from the NEON PLFA protocol, to process the data:

# PLFA (nmol/g) = (area of peak/areaISTD * nmol of ISTD) * (std area in blank/std area in sample) / sample weight

## EXPLANATION OF THE EQUATION:

#PART 1: (area of peak/areaISTD * nmol of ISTD) 
#converts peak area to nmol for all PLFAs, based on the nmol of standard that we added

#PART 2: (std area in blank/std area in sample)
#standardizes all values based on the blank. We assume that the blank had 100% extraction efficiency, and that the samples had lower extraction efficiency due to some inhibition by soil particles.

#PART 3: lastly, subtract any peaks in the blank from all samples, and divide by sample weight to get our results in nmol/g of soil


#### DATA CLEANING ####
#select only the data we'll need, pull out our internal standard and blank values, and subtract the blank peaks out from all samples

my_PLFAs <- read_excel("parsed_files/2023_PLFA_batch_F.xlsx")

#select only sampleID, Response, and Peak Name columns
my_PLFAs <- my_PLFAs %>% select("SampleID", "Response", "Peak Name")

#remove calibration peaks
my_PLFAs <- subset(my_PLFAs, SampleID != 'Calibration Mix MIDI in 2mL Hexane')

#Pivot wider
my_PLFAs_2 <- my_PLFAs %>% 
  filter(`Peak Name` != '') %>% 
  pivot_wider(names_from = "Peak Name", values_from = "Response", 
              values_fn=function(x) sum(as.numeric(x)))

my_PLFAs_2 <- as.data.frame(my_PLFAs_2)

#set rownames as Sample ID. Convert data to numeric. Set NAs to 0
names <- my_PLFAs_2$SampleID
my_PLFAs_2 <- select (my_PLFAs_2, -c(SampleID, "SOLVENT PEAK"))

my_PLFAs_2 <- as.data.frame(lapply(my_PLFAs_2, as.numeric))

rownames(my_PLFAs_2) <- names
my_PLFAs_2[is.na(my_PLFAs_2)] = 0
view(my_PLFAs_2)

#fix sample ID names if you need to
#rownames(my_PLFAs_2) <- c("1-1", "1-2", "1-3", "1-4", "1-5", "1-6", "1-7", "1-8","1-9", "1-10-B")

##pull out the internal standard and blank values that we'll use later. CHECK THEM

#internal standard- is it nearly the same between samples?
(internal_standard_values <- my_PLFAs_2$X19.0)

#SPECIFY YOUR BLANK and check out the values as they print. Did your blank have much contamination?
  #if you have two blanks in this batch, you'll need to average them like so
(blank_internal_standard <- 
    (my_PLFAs_2['9-16B', "X19.0"] + my_PLFAs_2['10-16B', "X19.0"])/2)

(blank_all_values <- 
    (t(my_PLFAs_2["9-16B",]) + t(my_PLFAs_2["10-16B",]))/2)

#subtract blank from all other rows, then set negatives to 0. CHECK: Does blank have all 0s now?
my_PLFAs_3 <- sweep(x = my_PLFAs_2, MARGIN = 2, STATS = blank_all_values, FUN = "-")

my_PLFAs_3[my_PLFAs_3 < 0] <- 0 

#PART 1: (area of peak/areaISTD * nmol of ISTD) 
#Convert from peak area to nmol.

my_PLFAs_3 <- (my_PLFAs_3 / internal_standard_values) * 6.1

#PART 2: * (std area in blank/std area in sample)
    
#look at your correction factors to see if your extraction efficiency was ok (close to 1)
(correction_factor <- (blank_internal_standard/internal_standard_values))

#apply correction factor
my_PLFAs_4 <- my_PLFAs_3 * correction_factor

#add sample ID back
my_PLFAs_4$ID <- rownames(my_PLFAs_4)

#PART 3: Divide by g of soil
  #if you used exactly 1g soil, you can skip this step

#input an excel sheet with "ID" and "Weigh" columns
weights <- as.data.frame(read_excel(""))

#merge by rownames
my_PLFAs_4 <- merge(weights, my_PLFAs_4, by = "ID")
view(my_PLFAs_4)

#divide by sample weight. make sure to only select numeric columns w biomarkers.
my_PLFAs_4[,4:66] <- (my_PLFAs_4[,4:66]/ my_PLFAs_4$Weight)

#SAVING
#Save the cleaned file. 
write_xlsx(my_PLFAs_4, "cleaned_files/PLFA_batch_F_clean.xlsx")

#Clear the environment, then clean your next file
rm(list=ls()) 

#AFTER YOU'VE CLEANED ALL YOUR FILES:
#read in all cleaned files, then merge
PLFAs_batchA <- read_excel("cleaned_files/PLFA_batch_A_clean.xlsx")
PLFAs_batchB <- read_excel("cleaned_files/PLFA_batch_B_clean.xlsx")
PLFAs_batchC <- read_excel("cleaned_files/PLFA_batch_C_clean.xlsx")
PLFAs_batchD <- read_excel("cleaned_files/PLFA_batch_D_clean.xlsx")
PLFAs_batchE <- read_excel("cleaned_files/PLFA_batch_E_clean.xlsx")
PLFAs_batchF <- read_excel("cleaned_files/PLFA_batch_F_clean.xlsx")

PLFA_all_clean <- bind_rows(PLFAs_batchA, PLFAs_batchB, PLFAs_batchC, PLFAs_batchD, PLFAs_batchE, PLFAs_batchF)

#order columns
new_order = sort(colnames(PLFA_all_clean))
PLFA_all_clean <- PLFA_all_clean[, new_order]

write_xlsx(PLFA_all_clean, "cleaned_files/PLFA_all_clean.xlsx")
