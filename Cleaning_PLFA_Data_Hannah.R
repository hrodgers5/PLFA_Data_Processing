require(tidyverse)
require(textreadr)
require(openxlsx)

#Authors: This R script was written by Shannon Albeke (the parsing RTF file part) and Hannah Rodgers (the data processing part).

#Aim: To take an RTF file with PLFA results from the GC, convert to excel, clean, and process the data.

#Last updated 1/31/2022

#### PARSING THE RTF FILE ####
setwd("C:/Users/hanna/Desktop/Work in Progress/PLFA")

# Read in the rtf file, each row is a character within a vector
rtf<- read_rtf("PLFA_OREI_1.5.2022.rtf")

# Find the headers to each section
vol<- which(str_detect(rtf, pattern = "Volume"))

# Loop through each of the groups in vol, try and extract the tabular data
outRTF<- data.frame()
outPrim<- data.frame()
for(i in 1:length(vol)){
  if(i != length(vol)){
    # Get the subset of data
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
    SampID<- str_squish(str_sub(v, str_locate(v, "Sample ID:")[2] + 1, str_locate(v, "ECL Deviation:")[1] - 1))
    # ECL Deviation
    ECL<- str_squish(str_sub(v, str_locate(v, "ECL Deviation:")[2] + 1, str_locate(v, "Reference ECL Shift:")[1] - 1))
    # ECL Ref
    RefECL<- str_squish(str_sub(v, str_locate(v, "Reference ECL Shift:")[2] + 1, str_locate(v, "Number Reference Peaks:")[1] - 1))
    # Ref Peaks
    RefPeaks<- str_squish(str_sub(v, str_locate(v, "Number Reference Peaks:")[2] + 1, str_locate(v, "Total Response:")[1] - 1))
    
  } else
  {
    # Sample ID
    SampID<- str_squish(str_sub(v, str_locate(v, "Sample ID:")[2] + 1, str_locate(v, "Total Response:")[1] - 1)) 
    
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
    
  # set the column names for the entire table
  if(i == 1){
    nm<- str_squish(unlist(str_split(myTab[1], pattern = "\\|")))
  }
  # Iterate through each row, skip the first
  for(j in 2:length(myTab)){
    df<- data.frame(t(str_squish(unlist(str_split(myTab[j], pattern = "\\|")))))
    names(df)<- nm
    tmpDF<- rbind(tmpDF, df)
  } # Close j loop
    
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


#### DATA PROCESSING ####

# I am using this equation, taken from the NEON PLFA protocol, to process the data:

# PLFA (nmol/g) = (area of peak/areaISTD * nmol of ISTD) * (std area in blank/std area in sample) / sample weight

## EXPLANATION OF THE EQUATION:
#PART 1: area of peak/areaISTD * nmol of ISTD : 
#this converts peak area to nmol for all PLFAs, based on the nmol of standard that we added

#PART 2: std area in blank/std area in sample : 
#This standardizes all values based on the blank. We assume that the blank had 100% extraction efficiency, and that the samples had lower extraction efficiency due to some inhibition by soil particles.

#PART 3: lastly, we subtract any peaks in the blank from everything, and divide by sample weight to get our results in nmol/g of soil


#this part of the code will select only the data you need, standardize your peak values by comparing the sample ISTD to the blank ISTD, subtract the blank peaks from all sample peaks, then save to excel

###DATA CLEANING 
#select only sampleID, Response, and Peak Name columns
my_PLFAs <- outRTF %>% select("SampleID", "Response", "Peak Name")

#select only the rows with your samples and blanks
my_PLFAs <- my_PLFAs[1643:2186,]

#Pivot wider
my_PLFAs_2 <- my_PLFAs %>% 
  filter(`Peak Name` != '') %>% 
  pivot_wider(names_from = "Peak Name", values_from = "Response")

my_PLFAs_2 <- as.data.frame(my_PLFAs_2)

#convert Sample ID to "rownames"
names <- my_PLFAs_2$SampleID
my_PLFAs_2 <- select (my_PLFAs_2, -SampleID)

my_PLFAs_2 <- as.data.frame(lapply(my_PLFAs_2,as.numeric))
rownames(my_PLFAs_2) <- names

#PART 1: Convert from peak area to nmol. We added 6.1 nmol of standard to each sample.
my_PLFAs_3 <- my_PLFAs_2 * my_PLFAs_2$x19.0 * 6.1


area of peak/areaISTD * nmol of ISTD

#PART 2: Standardize by blank by multiplying every PLFA by "blank ISTD"/"sample ISTD"
  #after this, every sample 19:0 should be the same as blank 19:0

#note to self: I want to save the value sample/blank in the final excel table

my_PLFAs_4 <- my_PLFAs_3 %>% 
  mutate(across(everything(), ~ .x * my_PLFAs_3["1-10_Blank_Hannah", "X19.0"]/X19.0)) 

#PART 3: Subtract blank peaks from samples and divide by g of soil
    #set all NAs to 0
outRTF4[is.na(outRTF4)] = 0

    #specify where your blank is
blank <- as.numeric(outRTF4["1-10_Blank_Hannah",])

    #substract blank from all other rows
outRTF5 <- sweep(x = outRTF4, MARGIN = 2, STATS = blank, FUN = "-")
  
    #add back in ISTD values- don't want to subtract those! also change all negatives to 0
outRTF5$X19.0 <- outRTF4$X19.0
outRTF5[outRTF5 < 0] <- 0 

# Save the output to Excel
wb <- createWorkbook()
addWorksheet(wb, "OverviewData")
addWorksheet(wb, "TabularData")

writeDataTable(wb, "OverviewData", x = outPrim, tableStyle = "none", withFilter = FALSE)
writeDataTable(wb, "TabularData", x = outRTF5, rowNames = TRUE, tableStyle = "none", withFilter = FALSE)

saveWorkbook(wb, file = "Processed_1.4.2022.xlsx", overwrite = TRUE)

#test change
