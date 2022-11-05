rm(list=ls());
#Load libraries
library(readxl)
library(tidyverse)
library(lubridate)

setwd("C:/Users/mmamm/Desktop/Data Manipulation")
#Read all sheets from the excel file
sheet_names <- excel_sheets('E1094_08122020b.xlsx')

#Read data from the first sheet "Details". 
filter <- dplyr::filter
data <- read_excel('E1094_08122020b.xlsx', sheet = 'TwinDetails')

table(data$Type_Desc) # Summary of data by description e.g parent, sibling or twin.
nrow(data) #Number of participants before data processing.
#Omit siblings, parents and triplets.
data<-data[!(data$Type_Desc == "Parent"|data$Type_Desc== "Sibling"|data$Type_Desc== "Triplet") ,]
summary(as.factor(data$ACTUAL_ZYGOSITY) ) # or table(data$ACTUAL_ZYGOSITY) # Summary of data by zygosity.
data <- data %>% filter(!(is.na(ACTUAL_ZYGOSITY))  )#Remove individuals with NA's in zygosity
data<-data[(data$ACTUAL_ZYGOSITY == "DZ"),] ##Omit MZ and UZ.
#data<-data[(data$SEX == "F"),] #
data <- data %>% filter(!(is.na(YEAR_BIRTH))  )#Remove individuals with NA's in YEAR_BIRTH
data <- data %>% filter(!(is.na(SEX))  )#Remove individuals with NA's in YEAR_BIRTH

data <- data %>%
  mutate(twinID = sub('.$', '', PublicID),
         twinID = match(twinID, unique(twinID)),
         #unique ID for each person.
         newID = row_number())

head(data)

#Read all the remaining sheets. 
list_data <- sapply(sheet_names[-c(1)], function(.x) {
  tmp <- read_excel('E1094_08122020b.xlsx', .x)
  tmp %>%  
    filter(if_any(3:last_col(), .fns = ~.x != 'NULL')) %>%
    distinct(PublicID, .keep_all = TRUE)
}, simplify = FALSE) %>%
  imap(function(x, y) {
    #Get the date column from the data frame. 
    col <- str_subset(names(x), regex('date', ignore_case = TRUE))
    if(length(col)) {
      #Extract the year from the date column
      x <- x %>%  
        mutate(!!col := year(.data[[col]])) %>%
        rename_with(~paste(., y, sep = '_'), -PublicID)
    } 
    x
  })
list_data[[length(list_data) + 1]] <- data
list_data <- rev(list_data) 

#Join all the sheets together. 
all_data <- list_data %>% reduce(left_join, by = 'PublicID')
#df1 <- do.call(data.frame, all_data); head(df1) # To view data as a dataframe

all_data <- all_data %>%
  #Get the age from each sheet
  mutate(across(contains('Date'), ~. - YEAR_BIRTH, .names = '{sub("date", "_age", .col, ignore.case = TRUE)}'), .before = 1, 
         #For Q2 sheet separately
         Response_age_Q20 = ResponseYear - YEAR_BIRTH) %>% 
  #rename Q2 to Q20 so that columns can be easily selected. 
  #contains('Q2') also selects Q21, Q22 and others.
  rename(Q20_179 = Q2_179) %>%
  #drop the rows that have no response
  filter(if_any(contains('age'), Negate(is.na))) %>%
  #drop the date columns
  select(-contains('Date'), -ResponseYear#, -Q11A_522_Q11A 
  ) %>%
  #arrange the column in required order. 
  select(PublicID:newID, contains('PCodes'), contains('Q20'), contains('Q17B'), 
         contains('Q17D'), contains('Q22'), contains('Q23'), contains('Q11A'), 
         contains('Q29'), contains('Q36')) %>%
  mutate(across(-(PublicID:newID), as.numeric), 
         across(-(PublicID:newID), ~replace(., . > 100, NA))) %>%
  data.frame()
#write_csv(all_data, 'output3.csv')

all_data$first_response <- NA
all_data$last_response <- NA
question_sheets <- sheet_names[-1]
#Change Q2 to Q20
question_sheets[question_sheets == "Q2"] <- "Q20"

for (i in seq_len(nrow(all_data))) {
  x <- all_data[i, ]
  #get all the age columns
  all_ages <- x %>% select(contains('age')) %>% unlist()
  #Sort them in ascending order, dropping NA values
  all_ages <- sort(all_ages)
  #Keep only values greater than 50
  all_ages <- all_ages[all_ages >= 50]
  #Get the group names
  all_ages_names <- sub('.*_', '', names(all_ages))
  #get min age 
  for(j in all_ages_names) {
    #get all the values column for j and drop NULL values
    val <- x %>%
      select(contains(j)) %>%
      select(-contains('age')) %>%
      unlist() %>% na.omit() %>% as.numeric()
    #If there is a valid value
    if(length(val)) {
      #Get the corresponding age
      all_data$first_response[i] <- x %>% 
        select(contains(paste0("_age_", j))) %>% 
        unlist(use.names = FALSE)
      #Break the loop as we have already found the minimum value so no need to look further.
      break
    }
  }
  
#get max age
#Do the same as minimum values but start from last value in descending order. 
  for(j in rev(all_ages_names)) {
    val <- x %>%
      select(contains(j)) %>%
      select(-contains('age')) %>%
      unlist() %>% na.omit() %>% as.numeric()
    if(length(val)) {
      all_data$last_response[i] <- x %>% 
        select(contains(paste0("_age_", j))) %>% 
        unlist(use.names = FALSE)
      break
    }
  }
}
sum(all_data$first_response == all_data$last_response, na.rm = TRUE) # Join the study but no follow-up
sum(all_data$first_response < all_data$last_response, na.rm = TRUE) # Join the study and are followed up

#Valid first and last response values, drop NA values. 
all_data <- all_data %>% filter(!(is.na(first_response) | is.na(last_response)))
all_data<- all_data[(all_data$first_response < all_data$last_response),]
table(all_data$SEX) 
all_data<- all_data[(all_data$SEX=="F"),] #Omit males

fracture_data <- all_data %>%
  #Select fractures questions
   select(matches('Q11A|Q17D|Pcodes|Q17B|Q29|Q36'), -Q11A_527_Q11A)
fractures_sheets <- c("Q11A", "Q17D", "Pcodes", "Q17B", "Q29", "Q36")

#get all the fractures age
all_data$fracture <- ''
for(i in seq_len(nrow(fracture_data))) {
  x <-  fracture_data[i, ]
  ans <- ''
  
  for(j in fractures_sheets) {
    val <- x %>% select(contains(j)) %>% select(-contains('age')) %>% unlist %>% na.omit() %>% as.numeric()
    #age <- x %>% select(contains(paste0('age_', j))) %>% unlist(use.names = FALSE)
    ans <- c(ans, val)
    
  }
  all_data$fracture[i] <- toString(sort(as.numeric(unique(ans[ans != ""]))))
}

all_data <- all_data %>% 
  mutate(HD_after_1st_response = map2_chr(strsplit(fracture, ', '), first_response, ~{
    .x <- as.numeric(.x)
    toString(.x[.x >= .y])
  }),
  
  first_fracture = map_chr(strsplit(HD_after_1st_response, ', '), 
                           ~{
                             .x <- as.numeric(.x)
                             if(length(.x)) 
                               as.character(min(.x, na.rm = TRUE)) 
                             else NA_character_
                           }),
  first_fracture = ifelse(is.na(first_fracture), last_response, first_fracture), 
  Status = as.integer(first_fracture != last_response)) %>%
  filter(first_response != first_fracture)

sum(all_data$first_response < all_data$first_fracture, na.rm = TRUE) 

names(all_data)

# Omit participants who said they have a fracture but did not give the age at which the fracture happened.
cond <- with(all_data, rowSums(all_data %>% select(matches('Q20|Q22|Q23|Q11A_527'), -contains('age')) == 1, na.rm = TRUE) > 0  &
               all_data$fracture =="")
complement_cond <- !cond | is.na(cond)
alldata3<-subset(all_data, complement_cond) 


df1 <- do.call(data.frame, alldata3); head(df1)
write.csv(df1, 'result_F.csv', row.names = FALSE)

#5 new columns
# - first_response          First valid response and age greater than 50.
# - last_response           Last valid response and age greater than 50.
# - fracture           All the ages when fractures was observed
# - first_fracture     first fracture after inclusion to study
# - Status                  Status 1/0.


