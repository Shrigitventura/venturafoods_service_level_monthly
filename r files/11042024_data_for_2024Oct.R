library(tidyverse)
library(readr)
library(lubridate)
library(readxl)
library(bizdays)
library(RColorBrewer)
library(scales)
library(rlang)

#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# Function Creation 

create.calendar(name='CustomWeekends', weekdays=c('saturday', 'sunday'), holidays=character(0))


current_date <- Sys.Date()
previous_month <- current_date %m-% months(1)
start_date <- floor_date(previous_month, "month")
end_date <- ceiling_date(previous_month, "month") - days(1)

target_date <- current_date %m-% months(1)
target_month <- lubridate::month(target_date)
target_year <- lubridate::year(target_date)




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
# https://venturafoods.sharepoint.com/sites/operations/SUPPLYCHAIN/SCS/Lists/Service%20Escalation%20Tracker/stan.aspx

# Completed
# https://venturafoods.sharepoint.com/sites/operations/SUPPLYCHAIN/SCS/Lists/SERVICEESCALATIONTASKS/All.aspx#InplviewHashbe3cd0d3-d91e-49b7-b0a3-5ab4f7704192=FilterField1%3DTask%255Fx0020%255FStatus-FilterValue1%3DComplete

# OTIF
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/7E766C3EDA4DC0E70BD08586AF56C8CB/K1717--K77



closed <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.10/closed.xlsx")
completed <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.10/completed.xlsx")
otif <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.10/otif.xlsx")





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
  dplyr::mutate(year_month = format(closed, "%Y/%m")) %>% 
  dplyr::left_join(otif %>% select(customer_number,
                                   customer_ship_to_name,
                                   customer_sold_to_name,
                                   selling_region)) %>% 
  dplyr::mutate(selling_region = ifelse(is.na(selling_region), "Multiple", selling_region)) %>%
  dplyr::mutate(close_time_period = bizdays(created, closed, 'CustomWeekends')) %>% 
  dplyr::relocate(year_month, .after = created) %>%
  dplyr::relocate(customer_number, .after = closed) %>% 
  dplyr::relocate(customer_ship_to_name, .after = customer_number) %>% 
  dplyr::relocate(customer_sold_to_name, .after = customer_ship_to_name) %>%
  dplyr::relocate(selling_region, .after = customer_sold_to_name) %>% 
  dplyr::relocate(close_time_period, .after = created_month) %>% 
  dplyr::mutate(closure_in_2_days = ifelse(close_time_period <= 2, "Met", "Not Met")) %>% 
  dplyr::relocate(closure_in_2_days, .before = issue) %>% 
  dplyr::select(-id, -customer_name) %>% 
  dplyr::relocate(days_to_close, ticket_owner, ticket, request_status, created, year_month, due_by, closed, customer_number,
                  customer_ship_to_name, customer_sold_to_name, selling_region, location, created_month, close_time_period,
                  closure_in_2_days, issue)-> closed_cleaned

master_data_closed_rds %>% 
  dplyr::filter(!(closed >= start_date & closed <= end_date)) -> master_data_closed_rds


rbind(master_data_closed_rds, closed_cleaned) -> master_data_closed_rds

master_data_closed_rds %>%
  mutate(row_id = apply(., 1, paste, collapse = "")) %>%
  distinct(row_id, .keep_all = TRUE) %>%
  select(-row_id) -> master_data_closed_rds


master_data_closed_rds %>% 
  dplyr::mutate(created = as.Date(created),
                due_by = as.Date(due_by),
                closed = as.Date(closed)) -> master_data_closed_rds






saveRDS(master_data_closed_rds, "master_data_closed.rds")



writexl::write_xlsx(master_data_closed_rds, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.10/master_data_closed.xlsx")

#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# ETL: completed file




completed_cleaned %>%
  dplyr::filter(closed >= start_date & closed <= end_date) %>% 
  dplyr::select(-task_status, -request_status, -created_by, -created_month, -item_type, -path) %>% 
  dplyr::mutate(department = stringr::str_extract(request_type, "^[^:]*")) %>% 
  dplyr::mutate(year_month = format(created, "%Y/%m")) %>% 
  mutate(task_response_time_days = bizdays(created, task_response_entered, 'CustomWeekends')) %>% 
  dplyr::mutate(bracket = ifelse(task_response_time_days < 2, "<=1", 
                                 ifelse(task_response_time_days < 3, "<=2", 
                                        ifelse(task_response_time_days < 4, "<=3",
                                               ifelse(task_response_time_days < 5,  "<=4", ">=5"))))) %>% 
  dplyr::mutate(number_days_to_close = "",
                number_days_to_close = as.double(number_days_to_close)) %>% 
  dplyr::select(number_days_to_close, task_id, task_owner, created, task_response_entered, year_month, closed,
                escalation_number, request_type, department, task_response_time_days, bracket) %>% 
  mutate(closure_in_2_days = ifelse(task_response_time_days <= 1, "Met", "Not met")) -> completed_cleaned



if (has_name(completed_cleaned, "location")) {
  completed_cleaned <- completed_cleaned %>% dplyr::select(-location)
}

if (has_name(completed_cleaned, "sku")) {
  completed_cleaned <- completed_cleaned %>% dplyr::select(-sku)
}


master_data_completed_rds %>% 
  dplyr::filter(!(closed >= start_date & closed <= end_date)) -> master_data_completed_rds




rbind(master_data_completed_rds, completed_cleaned) -> master_data_completed_rds


###
master_data_completed_rds %>% str()
completed_cleaned %>% str()

###



master_data_completed_rds %>%
  mutate(row_id = apply(., 1, paste, collapse = "")) %>%
  distinct(row_id, .keep_all = TRUE) %>%
  select(-row_id) -> master_data_completed_rds



master_data_completed_rds %>% 
  dplyr::mutate(created = as.Date(created),
                closed = as.Date(closed),
                task_response_entered = as.Date(task_response_entered)) -> master_data_completed_rds



saveRDS(master_data_completed_rds, "master_data_completed.rds")

writexl::write_xlsx(master_data_completed_rds, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.10/master_data_completed.xlsx")




## Copy the master file into the new folder
file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.09/MASTER_CLOSED.xlsx", 
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.10/MASTER_CLOSED.xlsx")

file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.09/MASTER_COMPLETED.xlsx", 
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.10/MASTER_COMPLETED.xlsx")

file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.09/Service Escalation Performance - September.pptx", 
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.10/Service Escalation Performance - October.pptx")



