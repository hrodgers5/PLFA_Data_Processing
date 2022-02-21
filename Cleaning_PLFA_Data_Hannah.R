require(tidyverse)
require(textreadr)
require(openxlsx)
library(readxl)
require(writexl)

#Authors: This R script was written by Shannon Albeke (the parsing RTF file part) and Hannah Rodgers (the data processing part).

#Aim: To take an RTF file with PLFA results from the GC, convert to excel, clean, and process the data.

#Last updated 1/31/2022

#### PARSING THE RTF FILE ####
  setwd("C:/Users/hanna/Desktop/GRAD PROJECTS/OREI/OREI 2021 PLFA Data/")

# Read in the rtf file, each row is a character within a vector
rtf<- read_rtf("OREI batch5 2.9.22.rtf") %>% 
  # split elements by newline character, if present
  str_split(pattern='\n') %>% 
  unlist()

# Find the headers to each section (headers start with 'Volume')
vol<- which(str_detect(rtf, pattern = "Volume"))

# Loop through each of the groups in vol, try and extract the tabular data
outRTF<- data.frame()
outPrim<- data.frame()
for(i in 1:length(vol)){
  if(i != length(vol)){
    # start with the first line of a sample,
    # grab all lines in file until the start of the next sample
    tmp<- rtf[vol[i]:(vol[(i + 1)] - 1)]
  } else
  {
    tmp<- rtf[vol[i]:length(rtf)]
  }
  
  
  # Find the indices that aren't tabular
  v<- tmp[-which(str_detect(tmp, pattern = "\\|"))]
  
  # Smash into a single character string
  v<- paste0(v, collapse = " ")
  
  # Now for each key:value pair, try and locate the values
  # File
  File<- str_squish(str_sub(v, str_locate(v, "File:")[2] + 1, str_locate(v, "Samp Ctr:")[1] - 1))
  # Sample Center
  SampCtr<- str_squish(str_sub(v, str_locate(v, "Samp Ctr:")[2] + 1, str_locate(v, "ID Number:")[1] - 1))
  # Id number
  ID<- str_squish(str_sub(v, str_locate(v, "ID Number:")[2] + 1, str_locate(v, "Type:")[1] - 1))
  # Type
  Typ<- str_squish(str_sub(v, str_locate(v, "Type:")[2] + 1, str_locate(v, "Bottle:")[1] - 1))
  # Bottle
  Bottle<- str_squish(str_sub(v, str_locate(v, "Bottle:")[2] + 1, str_locate(v, "Method:")[1] - 1))
  # Method
  Meth<- str_squish(str_sub(v, str_locate(v, "Method:")[2] + 1, str_locate(v, "Created:")[1] - 1))
  # Created
  Created<- str_squish(str_sub(v, str_locate(v, "Created:")[2] + 1, str_locate(v, "Sample ID:")[1] - 1))
  
  # Need to use logic to deal with times when the ECL doesn't exist
  if(str_detect(v, "ECL Deviation")){
    # Sample ID
    SampID<- str_squish(str_sub(v, 
                                str_locate(v, "Sample ID:")[2] + 1, 
                                str_locate(v, "ECL Deviation:")[1] - 1))
    # ECL Deviation
    ECL<- str_squish(str_sub(v, 
                             str_locate(v, "ECL Deviation:")[2] + 1, 
                             str_locate(v, "Reference ECL Shift:")[1] - 1))
    # ECL Ref
    RefECL<- str_squish(str_sub(v, 
                                str_locate(v, "Reference ECL Shift:")[2] + 1, 
                                str_locate(v, "Number Reference Peaks:")[1] - 1))
    # Ref Peaks
    RefPeaks<- str_squish(str_sub(v, 
                                  str_locate(v, "Number Reference Peaks:")[2] + 1, 
                                  str_locate(v, "Total Response:")[1] - 1))
    
  } else
  {
    # Sample ID
    SampID<- str_squish(str_sub(v, 
                                str_locate(v, "Sample ID:")[2] + 1, 
                                str_locate(v, "Total Response:")[1] - 1)) 
    
  }
  
  # Total Response
  TotResp<- str_squish(str_sub(v, str_locate(v, "Total Response:")[2] + 1, str_locate(v, "Total Named:")[1] - 1))
  # Total named
  TotNamed<- str_squish(str_sub(v, str_locate(v, "Total Named:")[2] + 1, str_locate(v, "Percent Named:")[1] - 1))
  # Percent Named
  PctNamed<- str_squish(str_sub(v, str_locate(v, "Percent Named:")[2] + 1, str_locate(v, "Total Amount:")[1] - 1))
  # Total Amount
  TotAmnt<- str_squish(str_sub(v, str_locate(v, "Total Amount:")[2] + 1, str_locate(v, "Profile Comment:")[1] - 1))
  # Profile Comment
  ProfComment<- str_squish(str_sub(v, str_locate(v, "Profile Comment:")[2] + 1, nchar(v)))
  
  # Build the table and append
  outPrim<- rbind(outPrim, data.frame(File = File,
                                      SampleCenter = SampCtr,
                                      IDNumber = ID,
                                      Type = Typ,
                                      Bottle = Bottle,
                                      Method = Meth,
                                      Created = Created,
                                      SampleID = SampID,
                                      ECLDeviation = ifelse(exists("ECL"), ECL, NA),
                                      ReferenceECLShift = ifelse(exists("RefECL"), RefECL, NA),
                                      NumberReferencePeaks = ifelse(exists("RefPeaks"), RefPeaks, NA),
                                      TotalResponse = TotResp,
                                      TotalNamed = TotNamed,
                                      PercentNamed = PctNamed,
                                      TotalAmount = TotAmnt,
                                      ProfileComment = ProfComment))
  
  # now we have to iterate through the remaining rows for the group
  # Create a blank df
  tmpDF<- data.frame()
  # get the tabular information
  myTab<- tmp[str_detect(tmp, pattern = "\\|")]
  # splitting by \n created issues when creating new table rows:
  # must remove "hanging chads"
  myTab <- myTab[which(myTab != '*| ')]
    
  # set the column names for the entire table
  if(i == 1){
    nm<- str_squish(unlist(str_split(myTab[1], pattern = "\\|")))
  }
  # Iterate through each row, skip the first
  for(j in 2:length(myTab)){
    df<- data.frame(t(str_squish(unlist(str_split(myTab[j], pattern = "\\|")))))
    names(df)<- nm
    tmpDF<- rbind(tmpDF, df)
  }# Close j loop
    
  # now update the meta data for the rows
  tmpDF$SampleCenter <- SampCtr
  tmpDF$IDNumber <- ID
  tmpDF$Created <- Created
  tmpDF$SampleID <- SampID
  tmpDF<- tmpDF %>% 
    select(SampleCenter, IDNumber, SampleID, Created, RT:Comment2)
  
  # append to primary output
  outRTF<- rbind(outRTF, tmpDF)
  
}
# (hannah's part): SAVE THOSE FILES ####

#check "outRTF." It should have all your sample names in "SampleID"

#Select only the rows with your samples and blanks, then save.
my_PLFAs_batch5 <- outRTF[202:2329,]
write_xlsx(my_PLFAs_batch5, "parsedRTFs/PLFA_batch5.xlsx")

#STOP HERE. Process all your rtf files and save them as different names.

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
#here, we select only the data we'll need, pull out our internal standard and blank values, subtract the blank peaks out from everything, and put it into a nice datatable

my_PLFAs <- read_excel("parsedRTFs/PLFA_batch5.xlsx")

#select only sampleID, Response, and Peak Name columns
my_PLFAs <- my_PLFAs %>% select("SampleID", "Response", "Peak Name")
my_PLFAs <- subset(my_PLFAs, SampleID != 'Calibration Mix MIDI in 2mL Hexane')

#Pivot wider
my_PLFAs_2 <- my_PLFAs %>% 
  filter(`Peak Name` != '') %>% 
  pivot_wider(names_from = "Peak Name", values_from = "Response")

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
(blank_internal_standard <- my_PLFAs_2['5-20 (blank)', "X19.0"])
(blank_all_values <- t(my_PLFAs_2["1-10-B",]))

#subtract blank from all other rows, then set negatives to 0. CHECK: Does blank have all 0s now?
my_PLFAs_3 <- sweep(x = my_PLFAs_2, MARGIN = 2, STATS = blank_all_values, FUN = "-")

my_PLFAs_3[my_PLFAs_3 < 0] <- 0 

#PART 1: (area of peak/areaISTD * nmol of ISTD) 
  #Convert from peak area to nmol.

my_PLFAs_3 <- (my_PLFAs_3 / internal_standard_values) * 6.1

#PART 2: * (std area in blank/std area in sample)
  #after this, every sample 19:0 should be the same as blank 19:0

correction_factor <- (blank_internal_standard/internal_standard_values)
my_PLFAs_4 <- my_PLFAs_3 * correction_factor

#check out this correction factor to see if your extraction efficiency was ok (close to 1)
(my_PLFAs_4$correction_factor <- correction_factor)

#PART 3: Divide by g of soil

#input a dataframe (from excel) with sample weights, where Sample ID matches the IDs in my_PLFAs_4
weights <- as.data.frame(read_excel("PLFA IDs and Weights OREI 2021.xlsx", sheet = "Sheet2"))

#merge by rownames
my_PLFAs_4$ID <- rownames(my_PLFAs_4)
my_PLFAs_5 <- merge(weights, my_PLFAs_4, by = "ID")
view(my_PLFAs_5)

#divide by sample weight. make sure to only select numeric columns w biomarkers.
my_PLFAs_5[,4:66] <- (my_PLFAs_5[,4:66]/ my_PLFAs_5$Weight)

#here you can save the cleaned file, then merge all files. 
write_xlsx(my_PLFAs_5, "parsedRTFs/PLFA_batch5_clean.xlsx")

#Clear the environment before you start the next file
rm(list=ls()) 

#read in all cleaned files, then merge
PLFAs_batch1 <- read_excel("parsedRTFs/PLFA_batch1_clean.xlsx")
PLFAs_batch2 <- read_excel("parsedRTFs/PLFA_batch2_clean.xlsx")
PLFAs_batch3 <- read_excel("parsedRTFs/PLFA_batch3_clean.xlsx")
PLFAs_batch4 <- read_excel("parsedRTFs/PLFA_batch4_clean.xlsx")
PLFAs_batch5 <- read_excel("parsedRTFs/PLFA_batch5_clean.xlsx")

PLFAs_cleaned <- bind_rows(PLFAs_batch1, PLFAs_batch2, PLFAs_batch3, PLFAs_batch4, PLFAs_batch5)

write_xlsx(PLFAs_cleaned, "parsedRTFs/PLFA_cleaned_all.xlsx")

#### FINAL DATA ANALYSES ####
PLFAs_cleaned <- read_excel("parsedRTFs/PLFA_cleaned_all.xlsx")

#rename the peaks!
PLFAs <- PLFAs_cleaned[,1:2]

attach(PLFAs_cleaned)

PLFAs$bacteria <- `X14.0` + `X15.0.anteiso` + `X17.1.iso.w9c` + `X17.0.iso` + 
  `X17.0.anteiso` + `X17.1.w8c`

PLFAs$acinomycetes <- `X16.0.10.methyl` + `X17.0.10.methyl` + `X18.0.10.methyl`

PLFAs$gram_neg <- `X16.1.w7c` + `X17.0.cyclo` + `X18.1.w5c` + `X18.1.w7c`

PLFAs$gram_pos <- `X15.0.iso` + `X16.0.iso`

PLFAs$AMF <- `X16.1.w5c`

PLFAs$sapro_fungi <- `X18.2.w6c` + `X18.1.w9c`

PLFAs$other_mic <- `X15.0` + `X16.0` + `X17.0` + `X18.0`
                           
detach (PLFAs_cleaned)

#calculate sums
attach(PLFAs)

PLFAs$total_MB <- rowSums(PLFAs[,3:9])

PLFAs$total_fungi <- sapro_fungi + AMF

PLFAs$total_bacteria <- bacteria + gram_neg + gram_pos + acinomycetes

PLFAs$F_to_B <- PLFAs$total_fungi / PLFAs$total_bacteria

detach(PLFAs)

#FINAL SAVE!!
write_xlsx(PLFAs, "PLFAs_FINAL_OREI_2021.xlsx")

#### OTHER THINGS- NOT USED ####

PLFAs_cleaned_2 <- rename(PLFAs_cleaned, 
                          "bacteria" = X14.0, 
                          "gram +" = X15.0.iso,
                          "bacteria" = X15.0.anteiso,
                          "total" = X15.0,
                          "gram +" = X16.0.iso,
                          "gram -" = X16.1.w7c,
                          "AMF" = X16.1.w5c,
                          "total" = X16.0,
                          "actinomycetes" = X16.0.10.methyl,  
                          "bacteria" = X17.1.iso.w9c,
                          "bacteria" = X17.0.iso, 
                          "bacteria" = X17.0.anteiso, 
                          "bacteria" = X17.1.w8c, 
                          "gram -" = X17.0.cyclo, 
                          "total" = X17.0, 
                          "actinobacteria" = X17.0.10.methyl, 
                          "fungi" = X18.2.w6c, 
                          "fungi" = X18.1.w9c, 
                          "gram -" = X18.1.w7c,
                          "gram -" = X18.1.w5c,  
                          "total" = X18.0, 
                          "actinomycetes" = X18.0.10.methyl)

#code to make a peak names vector
PEAK_NAMES <- structure(list(X14.0 = "bacteria", X15.0.iso = "gram +", X15.0.anteiso =  "bacteria", X15.0 = "total", X16.0.iso = "gram +", X16.1.w7c = "gram -", 
                             X16.1.w5c = "AMF", X16.0 = "total", X16.0.10.methyl =  "actinomycetes", X17.1.iso.w9c = "bacteria", X17.0.iso = "bacteria", X17.0.anteiso = "bacteria", X17.1.w8c = "bacteria", X17.0.cyclo = "gram -", X17.0 =  "total", X17.0.10.methyl = "actinobacteria", X18.2.w6c = "fungi", X18.1.w9c = "fungi", X18.1.w7c = "gram -", X18.1.w5c = "gram -",  X18.0 = "total", X18.0.10.methyl = "actinomycetes"), row.names = "", class = "data.frame")
