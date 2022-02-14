require(tidyverse)
require(textreadr)
require(openxlsx)
library(readxl)
require(writexl)

#Authors: This R script was written by Shannon Albeke (the parsing RTF file part) and Hannah Rodgers (the data processing part).

#Aim: To take an RTF file with PLFA results from the GC, convert to excel, clean, and process the data.

#Last updated 1/31/2022

#### PARSING THE RTF FILE ####
setwd("C:/Users/hanna/Desktop/Work in Progress/PLFA/OREI 2021 PLFA Data/")

# Read in the rtf file, each row is a character within a vector
rtf<- read_rtf("OREI batch1 1.05.22.rtf") %>% 
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
# (hannah's part): MERGING YOUR RTF FILES ####

#check "outRTF." It should have all your sample names in "SampleID"

#Select only the rows with your samples and blanks, then save. 
my_PLFAs_batch1 <- outRTF[1643:2186,]
write_xlsx(my_PLFAs_batch1, "parsedRTFs/PLFA_batch1.xlsx")
#STOP HERE. Process all your rtf files and save them as different names.

#Next, merge all your PLFA rtf files into one big file, and save.
my_PLFAs <- rbind(my_PLFAs_batch1, my_PLFAs_batch2, my_PLFAs_batch3, my_PLFAs_batch4, use.names = TRUE)
write_xlsx(my_PLFAs, "parsedRTFs/PLFA_all.xlsx")

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

my_PLFAs <- read_excel("parsedRTFs/PLFA_all.xlsx")

#select only sampleID, Response, and Peak Name columns
my_PLFAs <- my_PLFAs %>% select("SampleID", "Response", "Peak Name")

#Pivot wider
my_PLFAs_2 <- my_PLFAs %>% 
  filter(`Peak Name` != '') %>% 
  pivot_wider(names_from = "Peak Name", values_from = "Response")

my_PLFAs_2 <- as.data.frame(my_PLFAs_2)

#convert Sample ID to "rownames" and convert all other data to numeric
names <- my_PLFAs_2$SampleID
my_PLFAs_2 <- select (my_PLFAs_2, -c(SampleID, "SOLVENT PEAK"))

my_PLFAs_2 <- as.data.frame(lapply(my_PLFAs_2,as.numeric))
rownames(my_PLFAs_2) <- names

#pull out the internal standard and blank values that we'll use later
internal_standard_values <- my_PLFAs_2$X19.0
blank_internal_standard <- my_PLFAs_2["3-20-BLANK", "X19.0"]
blank_all_values <- as.numeric(my_PLFAs_2["3-20-BLANK",])

#check out blank_all_values. Did you have much contamination in the blank?
blank_all_values

#subtract blank from all other rows. Then check: Does my_PLFAs_2 have all 0s in the blank column now?
my_PLFAs_2 <- sweep(x = my_PLFAs_2, MARGIN = 2, STATS = blank_all_values, FUN = "-")

#PART 1: (area of peak/areaISTD * nmol of ISTD) 
#Convert from peak area to nmol.

my_PLFAs_3 <- (my_PLFAs_2 / internal_standard_values) * 6.1

#PART 2: * (std area in blank/std area in sample)
  #after this, every sample 19:0 should be the same as blank 19:0

correction_factor <- (blank_internal_standard/internal_standard_values)
my_PLFAs_4 <- my_PLFAs_3 * correction_factor

#check out this correction factor to see if your extraction efficiency was ok (close to 1)
(my_PLFAs_4$correction_factor <- correction_factor)

#PART 3: Divide by g of soil
#add back in correction factor, ID, and internal standard values
my_PLFAs_4$internal_standard_peakarea <- internal_standard_values
my_PLFAs_4$correction_factor <- correction_factor
my_PLFAs_4$ID <- rownames(my_PLFAs_4)

#input a dataframe (from excel) with sample weights, where row names match the row names in my_PLFAs_4 and 
weights <- as.data.frame(read_excel("PLFA IDs and Weights OREI 2021.xlsx", sheet = "Sheet2"))
my_PLFAs_5 <- merge(my_PLFAs_4, weights, by = "ID")

#divide by sample weight. make sure to only select numeric columns.
my_PLFAs_6 <- (my_PLFAs_5[,2:70]/ my_PLFAs_5$Weight)

#FINAL CLEANUP:
   #set all NAs and negatives to 0
my_PLFAs_6[is.na(my_PLFAs_6)] = 0
my_PLFAs_6[my_PLFAs_6 < 0] <- 0 

# Save the output to Excel
write_xlsx(list("PLFA_nmol_per_g" = my_PLFAs_6), "PLFA_cleaned_data_DATE.xlsx")