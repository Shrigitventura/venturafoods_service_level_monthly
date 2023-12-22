library(tidyverse)
library(readr)
library(lubridate)
library(readxl)
library(bizdays)

#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# Function Creation 

create.calendar(name='NoWeekends', weekdays=c('saturday', 'sunday'))
current_date <- Sys.Date()
previous_month <- current_date %m-% months(1)
start_date <- floor_date(previous_month, "month")
end_date <- ceiling_date(previous_month, "month") - days(1)


lookup_table <- data.frame(
    request_type = c("Distribution: Product Availability - Backorders", 
                         "Distribution: Within-campus transfer requests to fulfill future orders", 
                         "Other Issues", 
                         "Planning: BT Not Sche- Late", 
                         "Planning: Move Up Production Date Within Leadtime", 
                         "Planning: Request Production Dates", 
                         "QA: COA Not Provided to Customer", 
                         "QA: Missing COA", 
                         "Transportation: Changing from LTL to full truckload", 
                         "Transportation: Scheduling backorders once appointment time confirmed"),
    department = c("Distribution", 
                             "Distribution", 
                             "Other", 
                             "Planning", 
                             "Planning", 
                             "Planning", 
                             "QA", 
                             "QA", 
                             "Transportation", 
                             "Transportation")
)


#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# Work with master files (original)


# master_data_closed <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/master file_closed.xlsx",
#                           sheet = "Raw Data")
# 
# master_data_closed %>%
#   janitor::clean_names() %>%
#   data.frame() %>%
#   dplyr::mutate(created = format(created, "%m/%d/%Y"),
#                   due_by = format(due_by, "%m/%d/%Y"),
#                   closed = format(closed, "%m/%d/%Y")) -> master_data_closed_rds
# 
# saveRDS(master_data_closed_rds, "master_data_closed.rds")

readRDS("master_data_closed.rds") -> master_data_closed_rds




# master_data_completed <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/master file_completed.xlsx",
#                           sheet = "data")
# 
# master_data_completed %>%
#     janitor::clean_names() %>%
#     mutate(across(c(created, task_response_entered, closed), as.Date)) %>%
#     data.frame() -> master_data_completed_rds

# saveRDS(master_data_completed_rds, "master_data_completed.rds")

readRDS("master_data_completed.rds") -> master_data_completed_rds






#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# Original File Read


# Closed 
# https://venturafoods.sharepoint.com/sites/operations/SUPPLYCHAIN/SCS/Lists/Service%20Escalation%20Tracker/Closed.aspx?xsdata=MDV8MDJ8fGE4YWJlYTRkZTJiOTQ0YWQ1MWIyMDhkYmZhNzdlNGNkfGU5ZGRmYzNlZWE2ODQyNGZiNTU1ZDlhMmI0MTIxM2FkfDB8MHw2MzgzNzkxNjUwNDk3NzA0MTh8VW5rbm93bnxWR1ZoYlhOVFpXTjFjbWwwZVZObGNuWnBZMlY4ZXlKV0lqb2lNQzR3TGpBd01EQWlMQ0pRSWpvaVYybHVNeklpTENKQlRpSTZJazkwYUdWeUlpd2lWMVFpT2pFeGZRPT18MXxMMk5vWVhSekx6RTVPbTFsWlhScGJtZGZXVEpWTVUxcVRYZGFiVVYwV1cxTk5GcHBNREJaVkVVd1RGUnNiVTVVUlhST1JFcHRXVlJCZDA1dFdtcGFSRUUwUUhSb2NtVmhaQzUyTWk5dFpYTnpZV2RsY3k4eE56QXlNekU1TnpBME1ERTN8ZTQ3NGViY2VmYjk0NGMyMTUxYjIwOGRiZmE3N2U0Y2R8NzU0OTZhZThlNjljNDQ5M2E3Nzg2MWViNDFhNzkxYzg%3D&sdata=Zld0a2E3eWR5NyswbFk4aWpwbDVtcmFhUU5VM0FTU3h5NDFHOC96V0ZDQT0%3D&ovuser=e9ddfc3e-ea68-424f-b555-d9a2b41213ad%2CSLee%40venturafoods.com&OR=Teams-HL&CT=1703184672345&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiIyNy8yMzExMDIyNDcwNSIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D

# Completed
# https://venturafoods.sharepoint.com/sites/operations/SUPPLYCHAIN/SCS/Lists/SERVICEESCALATIONTASKS/Complete.aspx?xsdata=MDV8MDJ8fGQ4NTI4OGEyMWEwMDRmODA0NzJlMDhkYmZhNzg1N2RifGU5ZGRmYzNlZWE2ODQyNGZiNTU1ZDlhMmI0MTIxM2FkfDB8MHw2MzgzNzkxNjY5ODAwNjUwNzB8VW5rbm93bnxWR1ZoYlhOVFpXTjFjbWwwZVZObGNuWnBZMlY4ZXlKV0lqb2lNQzR3TGpBd01EQWlMQ0pRSWpvaVYybHVNeklpTENKQlRpSTZJazkwYUdWeUlpd2lWMVFpT2pFeGZRPT18MXxMMk5vWVhSekx6RTVPbTFsWlhScGJtZGZXVEpWTVUxcVRYZGFiVVYwV1cxTk5GcHBNREJaVkVVd1RGUnNiVTVVUlhST1JFcHRXVlJCZDA1dFdtcGFSRUUwUUhSb2NtVmhaQzUyTWk5dFpYTnpZV2RsY3k4eE56QXlNekU1T0RrM01EWTB8ZGJlYjQ3NGRkYjkyNGRjYTQ3MmUwOGRiZmE3ODU3ZGJ8YTI3YTg1MzY0NzkyNGQ2NWE4YWEyMjczNmFiMjJlMjc%3D&sdata=blEweHBPcDdWTVNmcVpmUVN3YnEraXdubHpxKzJJbGdXdllkMGFEVHBlTT0%3D&ovuser=e9ddfc3e-ea68-424f-b555-d9a2b41213ad%2CSLee%40venturafoods.com&OR=Teams-HL&CT=1703184695160&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiIyNy8yMzExMDIyNDcwNSIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D

# OTIF
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/7E766C3EDA4DC0E70BD08586AF56C8CB/K1717--K77



closed <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/12.21.2023_for_Nov/closed.xlsx")
completed <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/12.21.2023_for_Nov/completed.xlsx")
otif <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/12.21.2023_for_Nov/otif.xlsx")





#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# ETL: closed file


# clean up closed
closed %>% 
    janitor::clean_names() %>% 
    dplyr::mutate(days_to_close = round(days_to_close, 0)) %>% 
    dplyr::mutate(created = as.Date(format(created, "%m/%d/%Y"), "%m/%d/%Y"),
                                due_by = as.Date(format(due_by, "%m/%d/%Y"), "%m/%d/%Y"),
                                closed = as.Date(format(closed, "%m/%d/%Y"), "%m/%d/%Y")) -> closed_cleaned

completed %>% 
    janitor::clean_names() %>% 
    dplyr::mutate(created = as.Date(format(created, "%m/%d/%Y"), "%m/%d/%Y"),
                                closed = as.Date(format(closed, "%m/%d/%Y"), "%m/%d/%Y")) -> completed_cleaned

otif %>% 
  janitor::clean_names() %>% 
  dplyr::rename(customer_number = customer_ship_to_ship_to,
                customer_ship_to_name = customer_ship_to_name_1,
                selling_region = selling_region_name) -> otif
  
# Crete columns to match 
closed_cleaned %>% 
  dplyr::filter(closed >= start_date & closed <= end_date) %>% 
  dplyr::select(-request_type, -created_day) %>% 
  dplyr::mutate(month = lubridate::month(closed, label = TRUE)) %>% 
  dplyr::left_join(otif %>% select(customer_number,
                                   customer_ship_to_name,
                                   customer_sold_to_name,
                                   selling_region)) %>% 
  dplyr::mutate(close_time_period = ifelse(is.na(created) | is.na(closed), 
                                            0, 
                                            bizdays::bizdays(created, closed, 'NoWeekends') - 1)) %>% 
  dplyr::mutate(column2 = ifelse(close_time_period > 2, ">2", "<= 2"),
                column1 = ifelse(close_time_period < 2, "<=1", 
                                 ifelse(close_time_period < 3, "<=2",
                                        ifelse(close_time_period < 4, "<=3", ">4")))) %>% 
  dplyr::relocate(month, .after = created) %>%
  dplyr::relocate(customer_number, .after = closed) %>% 
  dplyr::relocate(customer_ship_to_name, .after = customer_number) %>% 
  dplyr::relocate(customer_sold_to_name, .after = customer_ship_to_name) %>%
  dplyr::relocate(selling_region, .after = customer_sold_to_name) %>% 
  dplyr::relocate(close_time_period, .after = created_month) %>% 
  dplyr::relocate(column2, .after = close_time_period) %>%
  dplyr::relocate(column1, .after = column2) -> closed_cleaned


rbind(master_data_closed_rds, closed_cleaned) -> master_data_closed_rds

master_data_closed_rds %>%
  mutate(row_id = apply(., 1, paste, collapse = "")) %>%
  distinct(row_id, .keep_all = TRUE) %>%
  select(-row_id) -> master_data_closed_rds
  
saveRDS(master_data_closed_rds, "master_data_closed.rds")





#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# ETL: completed file




completed_cleaned %>%
  dplyr::filter(closed >= start_date & closed <= end_date) %>% 
  dplyr::select(-task_status, -request_status, -created_by, -created_month, -item_type, -path, -close_day, -created_day) %>% 
  dplyr::left_join(lookup_table) %>% 
  dplyr::mutate(month = lubridate::month(closed, label = TRUE)) %>% 
  dplyr::mutate(task_response_time_days = ifelse(is.na(created) | is.na(task_response_entered), 
                                           0, 
                                           bizdays::bizdays(created, task_response_entered, 'NoWeekends') - 1)) %>% 
  dplyr::mutate(bracket = ifelse(task_response_time_days < 2, "<=1", 
                                 ifelse(task_response_time_days < 3, "<=2",
                                        ifelse(task_response_time_days < 4, "<=3", ">4")))) %>% 
                                   
  dplyr::relocate(number_days_to_close, task_id, task_owner, created, task_response_entered, month, closed,
                  escalation_number, request_type, department, task_response_time_days, bracket) -> completed_cleaned

rbind(master_data_completed_rds, completed_cleaned) -> master_data_completed_rds



master_data_completed_rds %>%
  mutate(row_id = apply(., 1, paste, collapse = "")) %>%
  distinct(row_id, .keep_all = TRUE) %>%
  select(-row_id) -> master_data_completed_rds

saveRDS(master_data_completed_rds, "master_data_completed.rds")




#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# KPI: Closed












#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# KPI: Completed




