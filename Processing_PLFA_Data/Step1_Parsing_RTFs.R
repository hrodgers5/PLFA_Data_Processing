require(tidyverse)
require(striprtf)
require(openxlsx)
library(readxl)
require(writexl)


#Authors: This R script was written by Shannon Albeke (the parsing RTF file part) and Hannah Rodgers (the data processing part).

#Aim: To take an RTF file with PLFA results from the GC, convert to excel, clean, and process the data.
#processed in R 4.2.2
#Last updated 1/31/2022

#### PARSING THE RTF FILE ####
setwd("C:/Users/hanna/Desktop/Work in Progress/PLFA Data 2023/2023_RTFs/")

# Read in the rtf file, each row is a character within a vector
  rtf<- striprtf::read_rtf("2023_batch1_2_3.rtf") %>% 
  # split elements by newline character, if present
  str_split(pattern='\n') %>% 
  unlist()

# Find the headers to each section (headers start with 'Volume')
vol <- which(str_detect(rtf, pattern = "Volume"))

# Loop through each of the groups in vol, try and extract the tabular data
outRTF<- data.frame()
outPrim<- data.frame()
for(i in 1:length(vol)){
  if(i != length(vol)){ 
    # start with the first line of a sample,
    # grab all lines in file until the start of the next sample
    tmp <- rtf[vol[i]:(vol[(i + 1)] - 1)]
  } else
  {
    tmp <- rtf[vol[i]:length(rtf)]
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
  tmpDF<- tmpDF  %>% 
  dplyr::select(SampleCenter, IDNumber, SampleID, Created, Response:Comment2)
  
  # append to primary output
  outRTF<- rbind(outRTF, tmpDF)
  
}
# (hannah's part): SAVE THOSE FILES ####

#check "outRTF." It should have all your sample names in "SampleID" # 

#Select only the rows with your samples and blanks, then save. Save each "batch" of samples you ran together in a separate file.
my_PLFAs_batchA <- outRTF[236:2188,]
write_xlsx(my_PLFAs_batchA, "2023_PLFA_batchA.xlsx")

# For now I renamed the columns in excel. Will write script later to rename in R so that it flows in one workflow... although not sure if all RTF's will be parsed the same way by this script. 
#STOP HERE. Process all your rtf files and save them as different names.
