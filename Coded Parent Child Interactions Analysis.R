# Title of analysis: LAPE Lab Analysis
# Date of analysis: Started 5/12/2022
# Purpose: Calculate reliability statistics and summary statistics for child-adult interactions


# Setting working directory
library(readxl)
library(stringr)
library(dplyr)
library(irr)
library(tidyr)
library(vcd)
library(writexl)
library(purrr)
setwd(dirname(rstudioapi::getSourceEditorContext()$path))

# Function to Read in Coder 1 Spreadsheet---------
## Input: File Name
## Output: Dataframe of Combined Child, Adult and Transcription tabs
readCoder1File <- function(X) {
   ## General Information Tab------
   # Read in the spreadsheet
   suppressMessages(relDataAdult <- read_xlsx(X, sheet = "General Information"))
   
   # Get the name of the coder
   coder1 = str_remove(relDataAdult[[14,1]], "Coder 1: |Coder 1 name: ")
   
   ## Adult Behaviors Tab--------
   # Read in the spreadsheet
   suppressMessages(relDataAdult <- read_xlsx(X, sheet = "Adult Behaviors (1)"))
   
   # Set the names of the columns
   namesAdult = c("TurnNum", "TurnDesc", "TimeStamp", 
                  
                  "A_AdultMLU_VerbUtt_C1", "A_AdultMLU_NumMorph_C1", "A_AdultMLU_RespAdTurn_C1",  
                  "A_AdultMLU_VerbUttResp_C1", "A_AdultMLU_MorphVerbRespUtt_C1", 
                  
                  "A_AdultCP_ForceCh_C1", "A_AdultCP_YesNo_C1", "A_AdultCP_TestQ_C1", 
                  "A_AdultCP_RealQ_C1", "A_AdultCP_Comment_C1", "A_AdultCP_Direct_C1", "A_AdultCP_Prompt_C1", 
                  
                  "A_CaregiveRE_VerbUttResp_C1", "A_CaregiveRE_VerbMap_C1", 
                  "A_CaregiveRE_Rep_C1", "A_CaregiveRE_VerbExp_C1", "A_CaregiveRE_VerbRecast_C1",
                  
                  "A_EnvArr_EA_C1","A_EnvArr_ReqStrat_C1", "A_EnvArr_ComStrat_C1",
                  "A_EnvArr_TimeDel_C1",
                  
                  "A_EATimeDelCom_C1", "A_Engagement_C1")
   
   # Set the names of the columns to the names specified above
   names(relDataAdult) <- namesAdult
   
   # Filter out rows that just have variable names
   relDataAdult <- relDataAdult %>% 
      filter(!is.na(TurnNum) & !grepl("Turn #|Reliability", TurnNum))
   
   # Function that checks if a variable can be converted to a number
   is_all_numeric <- function(x) {
      !any(is.na(suppressWarnings(as.numeric(na.omit(x))))) & is.character(x)
   }
   
   # Converting character/string variables to numeric
   relDataAdult_C1 <- relDataAdult %>%
      mutate_if(is_all_numeric, as.numeric) 
   
   
   
   ## Child Behaviors Tab--------
   # Read in the spreadsheet
   suppressMessages(relDataChild <- read_xlsx(X, sheet = "Child Behaviors (1)"))
   
   # Set the names of the columns
   namesChild = c("TurnNum", "TurnDesc", "TimeStamp", 
                  "C_TurnInt_Int_C1", "C_TurnInt_Unint_C1", 
                  
                  "C_ModeCom_Spoke_C1", "C_ModeCom_Sign_C1", "C_ModeCom_NonSym_C1", "C_ModeCom_NumCB_C1", 
                  
                  "C_SymUtt_NumMorph_C1", "C_SymUtt_SWInv_C1", "C_SymUtt_SignInv_C1", 
                  "C_SymUtt_InvSignSpoke_C1", 
                  
                  "C_Gest_ContDi_C1", "C_Gest_DistDi_C1", "C_Gest_ObjIcon_C1", "C_Gest_Convent_C1", 
                  "C_Gest_GestInv_C1", 
                  
                  "C_Vocal_Jarg_C1", "C_Vocal_ConsVow_C1", "C_Vocal_VowOnly_C1", 
                  
                  "C_ConsInv_C1", 
                  
                  "C_IntentInit_ChildInit_C1", "C_IntentInit_Intent_C1", 
                  "C_IntentInit_IntentInit_C1", 
                  
                  "C_ComPurp_Req_C1", "C_ComPurp_Com_C1")
   
   # Set the names of the columns to the names specified above
   names(relDataChild) <- namesChild
   
   # Filter out rows that just have variable names
   relDataChild <- relDataChild %>% 
      filter(!is.na(TurnNum) & !grepl("Turn #|Reliability", TurnNum))
   
   # Converting character/string variables to numeric
   relDataChild_C1 <- relDataChild %>%
      mutate_if(is_all_numeric, as.numeric) 
   
   
   
   ## Transcription Tab--------
   # Read in the spreadsheet
   suppressMessages(
      relDataTranscription <- read_xlsx(X, sheet = "Transcription")
   )
   
   # Set the names of the columns
   namesTranscription = c("TurnNum", "TurnDesc", "TimeStamp", 
                          "T_AdultInit_C1", "T_AdultResp_C1", "T_ChildInit_C1", 
                          "T_ChildResp_C1",  "T_OtherPerson_C1")
   
   # Set the names of the columns to the names specified above
   names(relDataTranscription) <- namesTranscription
   
   # Filter out rows that just have variable names
   relDataTranscription <- relDataTranscription %>% 
      filter(!is.na(TurnNum) & !grepl("Turn #|Reliability", TurnNum))
   
   # Converting character/string variables to numeric
   relDataTranscription_C1 <- relDataTranscription %>%
      mutate_if(is_all_numeric, as.numeric) 
   
   ## Merging Adult, Child and Transcription Tabs-----
   AdultChildDataset_C1 <- merge(relDataAdult_C1, relDataChild_C1, by=c("TurnNum", "TurnDesc", "TimeStamp"))
   AdultChildTranscriptDataset_C1 <- merge(AdultChildDataset_C1, relDataTranscription_C1, 
                                           by=c("TurnNum", "TurnDesc", "TimeStamp"))
   # Add the coder name to the dataframe as a column
   AdultChildTranscriptDataset_C1 <- AdultChildTranscriptDataset_C1 %>% 
      mutate(Coder1 = coder1)
   return(AdultChildTranscriptDataset_C1)
}

# Function to Read in Coder 2 Spreadsheet---------
## Input: File Name
## Output: Dataframe of Combined Child, Adult and Transcription tabs
readCoder2File <- function(X) {
   ## General Information Tab------
   # Read in the spreadsheet
   suppressMessages(relDataAdult <- read_xlsx(X, sheet = "General Information"))
   
   # Get the name of the coder
   coder2 = str_remove(relDataAdult[[14,1]], "Coder 2: |Coder 2 name: ")
   
   ## Adult Behaviors Tab--------
   # Read in the spreadsheet
   suppressMessages(relDataAdult <- read_xlsx(X, sheet = "Adult Behaviors (1)"))
   # Set the names of the columns
   namesAdult = c("TurnNum", "TurnDesc", "TimeStamp", 
                  
                  "A_AdultMLU_VerbUtt_C2", "A_AdultMLU_NumMorph_C2", "A_AdultMLU_RespAdTurn_C2",  
                  "A_AdultMLU_VerbUttResp_C2", "A_AdultMLU_MorphVerbRespUtt_C2", 
                  
                  "A_AdultCP_ForceCh_C2", "A_AdultCP_YesNo_C2", "A_AdultCP_TestQ_C2", 
                  "A_AdultCP_RealQ_C2", "A_AdultCP_Comment_C2", "A_AdultCP_Direct_C2", "A_AdultCP_Prompt_C2", 
                  
                  "A_CaregiveRE_VerbUttResp_C2", "A_CaregiveRE_VerbMap_C2", 
                  "A_CaregiveRE_Rep_C2", "A_CaregiveRE_VerbExp_C2", "A_CaregiveRE_VerbRecast_C2",
                  
                  "A_EnvArr_EA_C2","A_EnvArr_ReqStrat_C2", "A_EnvArr_ComStrat_C2",
                  "A_EnvArr_TimeDel_C2",
                  
                  "A_EATimeDelCom_C2", "A_Engagement_C2")
   
   # Set the names of the columns to the names specified above
   names(relDataAdult) <- namesAdult
   
   # Filter out rows that just have variable names
   relDataAdult <- relDataAdult %>% 
      filter(!is.na(TurnNum) & !grepl("Turn #|Reliability", TurnNum))
   
   # Function that checks if a variable can be converted to a number
   is_all_numeric <- function(x) {
      !any(is.na(suppressWarnings(as.numeric(na.omit(x))))) & is.character(x)
   }
   
   # Converting character/string variables to numeric
   relDataAdult_C2 <- relDataAdult %>%
      mutate_if(is_all_numeric, as.numeric) 
   
   
   ## Child Behaviors Tab--------
   # Read in the spreadsheet
   suppressMessages(relDataChild <- read_xlsx(X, sheet = "Child Behaviors (1)"))
   
   # Set the names of the columns
   namesChild = c("TurnNum", "TurnDesc", "TimeStamp", 
                  "C_TurnInt_Int_C2", "C_TurnInt_Unint_C2", 
                  
                  "C_ModeCom_Spoke_C2", "C_ModeCom_Sign_C2", "C_ModeCom_NonSym_C2", "C_ModeCom_NumCB_C2", 
                  
                  "C_SymUtt_NumMorph_C2", "C_SymUtt_SWInv_C2", "C_SymUtt_SignInv_C2", 
                  "C_SymUtt_InvSignSpoke_C2", 
                  
                  "C_Gest_ContDi_C2", "C_Gest_DistDi_C2", "C_Gest_ObjIcon_C2", "C_Gest_Convent_C2", 
                  "C_Gest_GestInv_C2", 
                  
                  "C_Vocal_Jarg_C2", "C_Vocal_ConsVow_C2", "C_Vocal_VowOnly_C2", 
                  
                  "C_ConsInv_C2", 
                  
                  "C_IntentInit_ChildInit_C2", "C_IntentInit_Intent_C2", 
                  "C_IntentInit_IntentInit_C2", 
                  
                  "C_ComPurp_Req_C2", "C_ComPurp_Com_C2")
   
   # Set the names of the columns to the names specified above
   names(relDataChild) <- namesChild
   
   # Filter out rows that just have variable names
   relDataChild <- relDataChild %>% 
      filter(!is.na(TurnNum) & !grepl("Turn #|Reliability", TurnNum))
   
   # Converting character/string variables to numeric
   relDataChild_C2 <- relDataChild %>%
      mutate_if(is_all_numeric, as.numeric) 
   
   
   ## Transcription Tab--------
   # Read in the spreadsheet
   suppressMessages(relDataTranscription <- read_xlsx(X, sheet = "Transcription"))
   
   # Set the names of the columns
   namesTranscription = c("TurnNum", "TurnDesc", "TimeStamp", 
                          "T_AdultInit_C2", "T_AdultResp_C2", "T_ChildInit_C2", 
                          "T_ChildResp_C2",  "T_OtherPerson_C2")
   
   # Set the names of the columns to the names specified above
   names(relDataTranscription) <- namesTranscription
   
   # Filter out rows that just have variable names
   relDataTranscription <- relDataTranscription %>% 
      filter(!is.na(TurnNum) & !grepl("Turn #|Reliability", TurnNum))
   
   # Converting character/string variables to numeric
   relDataTranscription_C2 <- relDataTranscription %>%
      mutate_if(is_all_numeric, as.numeric) 
   
   ## Merging Adult, Child and Transcription Tabs-----
   AdultChildDataset_C2 <- merge(relDataAdult_C2, relDataChild_C2, by=c("TurnNum", "TurnDesc", "TimeStamp"))
   
   AdultChildTranscriptDataset_C2 <- merge(AdultChildDataset_C2, relDataTranscription_C2, 
                                           by=c("TurnNum", "TurnDesc", "TimeStamp"))
   # Add the coder name to the dataframe as a column
   AdultChildTranscriptDataset_C2 <- AdultChildTranscriptDataset_C2 %>% 
      mutate(Coder2 = coder2)
   
   return(AdultChildTranscriptDataset_C2)
}

# Function to read in and merge all subjects' coder 1 and 2 files-------
# Input: This function takes a list of subject numbers (in increasing order)
# Output: This function returns a new list of dataframes for coders 1 and 2
mergeCoderFiles <- function(X) {
   # Create empty dataframe to put all of the spreadsheets into
   FullCombinedDataset = data.frame()
   
   # Loop through subject numbers
   for (i in 1:length(X)) {
      num = X[i]
      
      # Read in and format coder files using functions defined above
      coder1Dataset = readCoder1File(paste(num, "Coding 1 Excel File.xlsx"))
      coder2Dataset = readCoder2File(paste(num, "Coding 2 Excel File.xlsx"))
      
      # Merge Coders 1 and 2 spreadsheets and format
      AdultChildTranscript_C1and2 <- merge(coder1Dataset, coder2Dataset, 
                                           by=c("TurnNum", "TurnDesc", "TimeStamp"))
      
      # Add in a subject number column
      AdultChildTranscript_C1and2 <- AdultChildTranscript_C1and2 %>% 
         mutate(subNum = num)
      
      # Merge all of the subjects' datasets
      if (!(exists("FullCombinedDataset")&&is.data.frame(get("FullCombinedDataset")))) {
         FullCombinedDataset = data.frame(colnames = names(AdultChildTranscript_C1and2))
      } else {
         FullCombinedDataset = rbind(FullCombinedDataset, AdultChildTranscript_C1and2)
      }
      print(paste("Finished reading in spreadsheets for subject number", num))
   }
   
   # Replace NAs in the coded variables with 0s
   FullCombinedDataset <- FullCombinedDataset %>% 
      mutate(across(where(is.numeric), ~replace(., is.na(.), 0)))
   return(FullCombinedDataset)
}

# Function to export a dataframe with overall reliability statistics---------
# Input: Dataframe, index of the first variable to calculate reliability for (index for var1coder1),
#        index for the other coder variable (i.e. index for var1coder2)
# Output: Dataframe with the variable names, the statistic calculated for each and its value
exportSumRelFun <- function(dataset, index1, index2) {
   # Filter out rows that are empty
   dataset <- dataset %>% 
      filter(!is.na(TurnDesc)) 
   
   # Set the indexes for the paired coder vars
   indexVarC1 = index1
   indexVarC2 = index2
   
   # Create an empty dataframe to put the reliability values into
   outputDf = data.frame(matrix(nrow = 0, ncol = 5))
   colnames(outputDf) = c("Var1", "Var2", "Statistic", "Value", "NumTurns")

   # Loop through all of the pairs of vars
   for (i in 1:52) {
      # Get the names of the vars
      nameVarC1 = colnames(dataset)[indexVarC1] 
      nameVarC2 = colnames(dataset)[indexVarC2] 
      
      # Filter the dataset for just adult or child turns depending on the variable
      if (startsWith(nameVarC1, "A_")) {
         datasetFiltered <- dataset %>% 
            filter(startsWith(TurnDesc, "A: ")) 
      } else if (startsWith(nameVarC1, "C_")) {
         datasetFiltered <- dataset %>% 
            filter(startsWith(TurnDesc, "C: ")) 
      } else {
         datasetFiltered <- dataset
      }
         
      # If the variable is numeric then print the unique values
      if (is.numeric(datasetFiltered[,nameVarC1]) & is.numeric(datasetFiltered[,nameVarC2])) {
         
         # If the variable is categorical (only has 2 unique values),
         # but only has 1 unique value then Fleiss' kappa is 1 (perfect agreement between coders)
         if (length(unique(datasetFiltered[,nameVarC1])) == 1 & length(unique(datasetFiltered[,nameVarC2])) == 1) {
            # Add Fleiss' kappa=1 as a row to the output dataframe
            newRow <- data.frame("Var1"=nameVarC1, "Var2"=nameVarC2, "Statistic"="Fleiss' Kappa", "Value"=1, "NumTurns"=nrow(datasetFiltered))
            outputDf = rbind(outputDf, newRow)
         } # Otherwise if the variable is categorical (only has 2 unique values) then calculate Fleiss' Kappa
         else if (length(unique(datasetFiltered[,nameVarC1])) < 3 & length(unique(datasetFiltered[,nameVarC2])) < 3) {
            # Calculate Fleiss' kappa and add it as a row to the output dataframe
            kap = kappam.fleiss(datasetFiltered %>% select(nameVarC1, nameVarC2))
            newRow <- data.frame("Var1"=nameVarC1, "Var2"=nameVarC2, "Statistic"="Fleiss' Kappa", "Value"=kap$value, "NumTurns"=kap$subjects)
            outputDf = rbind(outputDf, newRow)
         }  # Otherwise if the variable is continuous (>2 unique values) then calculate ICC
         else if ((length(unique(datasetFiltered[,nameVarC1])) >= 3 | length(unique(datasetFiltered[,nameVarC2])) >= 3)) {
            # Calculate ICC and add it as a row to the output dataframe
            iccOut = icc(datasetFiltered %>% select(nameVarC1, nameVarC2), model = "twoway", type = "agreement", unit = "single")
            newRow <- data.frame("Var1"=nameVarC1, "Var2"=nameVarC2, "Statistic"="ICC", "Value"=iccOut$value, "NumTurns"=iccOut$subjects)
            outputDf = rbind(outputDf, newRow)
         }
      }
      # Increment the index variables
      indexVarC1 = indexVarC1 + 1
      indexVarC2 = indexVarC2 + 1
   }
   return(outputDf)
}

# Function to export a dataframe with individual coder reliability statistics---------
# Input: Dataframe, index of the first variable to calculate reliability for (index for var1coder1),
#        index for the other coder variable (i.e. index for var1coder2)
# Output: Dataframe with the variable names, the coder names, subject numbers,
#         the statistic calculated for each and its value, and the number of turns used to calculate that statistic
exportCoderRelFun <- function(dataset, index1, index2) {
   # Filter out rows that are empty
   dataset <- dataset %>% 
      filter(!is.na(TurnDesc)) 
   print(dataset)
   # Set indexes for the paired coder vars
   indexVarC1 = index1
   indexVarC2 = index2
   
   # Create a new dataframe to put the reliability values into
   outputDf = data.frame(matrix(nrow = 0, ncol = 9))
   colnames(outputDf) = c("subNum", "Var1", "Var2", "Coder1", "Coder2", "Statistic", "Value", "NumTurns", "FullAgree")
   
   # Loop through all of the pairs of vars
   for (i in 1:52) {
      # Get the names of the vars
      nameVarC1 = colnames(dataset)[indexVarC1] 
      nameVarC2 = colnames(dataset)[indexVarC2] 
      
      # Filter the dataset for just adult or child turns depending on the variable
      if (startsWith(nameVarC1, "A_")) {
         datasetFiltered <- dataset %>% 
            filter(startsWith(TurnDesc, "A: ")) 
      } else if (startsWith(nameVarC1, "C_")) {
         datasetFiltered <- dataset %>% 
            filter(startsWith(TurnDesc, "C: ")) 
      } else {
         datasetFiltered <- dataset
      }
      print(datasetFiltered[,nameVarC1])
      print(datasetFiltered[,nameVarC2])
      # Check if the vars are numeric (necessary to calculate reliability)
      if (is.numeric(datasetFiltered[,nameVarC1]) & is.numeric(datasetFiltered[,nameVarC2])) {
         # Get all of the subNum, coder1, and coder2 combos
         relNames <- datasetFiltered %>% 
            group_keys(subNum, Coder1, Coder2) 
         
         # Compute if the coders had agreement (all values matched) - will use this variable later
         fullAgree <- datasetFiltered %>% 
            select(subNum, Coder1, Coder2, nameVarC1, nameVarC2) %>%
            mutate(dupCheck = ifelse(datasetFiltered[,nameVarC1]==datasetFiltered[,nameVarC2], 1, 0)) %>% 
            group_split(subNum) %>%
            map(~ifelse(sum(.[6])==nrow(.[6]), TRUE, FALSE)) %>% 
            unlist()
         
         relNames = cbind(relNames, "FullAgree"=fullAgree)
         

         # If the variable is categorical (has 0-2 unique values) then calculate Cohen's Kappa
         if (length(unique(datasetFiltered[,nameVarC1])) < 3 & length(unique(datasetFiltered[,nameVarC2])) < 3) {
            # Calculate Cohen's kappa and add it as a row to the output dataframe
            kap <- datasetFiltered %>% 
               select(subNum, Coder1, Coder2, nameVarC1, nameVarC2) %>%
               group_split(subNum) %>%
               map(~kappa2(.[-c(1,2,3)])[["value"]]) %>% 
               unlist()
            
            sub <- datasetFiltered %>% 
               select(subNum, Coder1, Coder2, nameVarC1, nameVarC2) %>%
               group_split(subNum) %>%
               map(~kappa2(.[-c(1,2,3)])[["subjects"]]) %>% 
               unlist()
            
            kapFull = cbind(relNames, "Value"=kap, "NumTurns"=sub)
            
            newRow <- data.frame("subNum"=kapFull$subNum, "Var1"=nameVarC1, "Var2"=nameVarC2, 
                                 "Coder1"=kapFull$Coder1, "Coder2"=kapFull$Coder2, "Statistic"="Cohen's Kappa",
                                 "Value"=kapFull$Value, "NumTurns"=kapFull$NumTurns, "FullAgree"=kapFull$FullAgree)
         }  # Otherwise if the variable is continuous (>2 unique values) then calculate ICC
         else if ((length(unique(datasetFiltered[,nameVarC1])) >= 3 | length(unique(datasetFiltered[,nameVarC2])) >= 3)) {
            # Calculate ICC and add it as a row to the output dataframe
            iccOut <- datasetFiltered %>% 
               select(subNum, Coder1, Coder2, nameVarC1, nameVarC2) %>%
               group_split(subNum) %>%
               map(~icc(.[-c(1,2,3)], model = "twoway", type = "agreement", unit = "single")[["value"]]) %>% 
               unlist()
            
            sub <- datasetFiltered %>% 
               select(subNum, Coder1, Coder2, nameVarC1, nameVarC2) %>%
               group_split(subNum) %>%
               map(~icc(.[-c(1,2,3)], model = "twoway", type = "agreement", unit = "single")[["subjects"]]) %>% 
               unlist()

            iccFull = cbind(relNames, "Value"=iccOut, "NumTurns"=sub)
            
            newRow <- data.frame("subNum"=iccFull$subNum, "Var1"=nameVarC1, "Var2"=nameVarC2,
                                 "Coder1"=iccFull$Coder1, "Coder2"=iccFull$Coder2, "Statistic"="ICC",
                                 "Value"=iccFull$Value, "NumTurns"=iccFull$NumTurns, "FullAgree"=kapFull$FullAgree)
         }
         # Add the new row to the output dataframe
         outputDf = rbind(outputDf, newRow)
      }
      # Increment the index variables
      indexVarC1 = indexVarC1 + 1
      indexVarC2 = indexVarC2 + 1
   }
   # Note:  If the coders had 100% agreement then Cohen's kappa outputs NaN 
   # Replacing those NaNs with "1" here
   outputDf$Value <- ifelse(outputDf$FullAgree==T & is.na(outputDf$Value), 1, outputDf$Value)
   outputDf <- outputDf %>% select(-FullAgree) 
   return(outputDf)
}

# Using the functions on a small dataset------
subNums = c(170, 329, 371)
FullDataset <- mergeCoderFiles(subNums)
individRelDf <- exportCoderRelFun(FullDataset, 4, 57)
# write_xlsx(individRelDf, "Individual Reliability Statistics.xlsx")
overallReliabilityDf <- exportSumRelFun(FullDataset, 4, 57)
# write_xlsx(overallReliabilityDf, "Overall Reliability Statistics.xlsx")
