library(tidyverse)
library(readr)
library(lubridate)
library(readxl)
library(bizdays)
library(RColorBrewer)
library(scales)

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
# https://venturafoods.sharepoint.com/sites/operations/SUPPLYCHAIN/SCS/Lists/Service%20Escalation%20Tracker/Closed.aspx?xsdata=MDV8MDJ8fGE4YWJlYTRkZTJiOTQ0YWQ1MWIyMDhkYmZhNzdlNGNkfGU5ZGRmYzNlZWE2ODQyNGZiNTU1ZDlhMmI0MTIxM2FkfDB8MHw2MzgzNzkxNjUwNDk3NzA0MTh8VW5rbm93bnxWR1ZoYlhOVFpXTjFjbWwwZVZObGNuWnBZMlY4ZXlKV0lqb2lNQzR3TGpBd01EQWlMQ0pRSWpvaVYybHVNeklpTENKQlRpSTZJazkwYUdWeUlpd2lWMVFpT2pFeGZRPT18MXxMMk5vWVhSekx6RTVPbTFsWlhScGJtZGZXVEpWTVUxcVRYZGFiVVYwV1cxTk5GcHBNREJaVkVVd1RGUnNiVTVVUlhST1JFcHRXVlJCZDA1dFdtcGFSRUUwUUhSb2NtVmhaQzUyTWk5dFpYTnpZV2RsY3k4eE56QXlNekU1TnpBME1ERTN8ZTQ3NGViY2VmYjk0NGMyMTUxYjIwOGRiZmE3N2U0Y2R8NzU0OTZhZThlNjljNDQ5M2E3Nzg2MWViNDFhNzkxYzg%3D&sdata=Zld0a2E3eWR5NyswbFk4aWpwbDVtcmFhUU5VM0FTU3h5NDFHOC96V0ZDQT0%3D&ovuser=e9ddfc3e-ea68-424f-b555-d9a2b41213ad%2CSLee%40venturafoods.com&OR=Teams-HL&CT=1703184672345&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiIyNy8yMzExMDIyNDcwNSIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D

# Completed
# https://venturafoods.sharepoint.com/sites/operations/SUPPLYCHAIN/SCS/Lists/SERVICEESCALATIONTASKS/Complete.aspx?xsdata=MDV8MDJ8fGQ4NTI4OGEyMWEwMDRmODA0NzJlMDhkYmZhNzg1N2RifGU5ZGRmYzNlZWE2ODQyNGZiNTU1ZDlhMmI0MTIxM2FkfDB8MHw2MzgzNzkxNjY5ODAwNjUwNzB8VW5rbm93bnxWR1ZoYlhOVFpXTjFjbWwwZVZObGNuWnBZMlY4ZXlKV0lqb2lNQzR3TGpBd01EQWlMQ0pRSWpvaVYybHVNeklpTENKQlRpSTZJazkwYUdWeUlpd2lWMVFpT2pFeGZRPT18MXxMMk5vWVhSekx6RTVPbTFsWlhScGJtZGZXVEpWTVUxcVRYZGFiVVYwV1cxTk5GcHBNREJaVkVVd1RGUnNiVTVVUlhST1JFcHRXVlJCZDA1dFdtcGFSRUUwUUhSb2NtVmhaQzUyTWk5dFpYTnpZV2RsY3k4eE56QXlNekU1T0RrM01EWTB8ZGJlYjQ3NGRkYjkyNGRjYTQ3MmUwOGRiZmE3ODU3ZGJ8YTI3YTg1MzY0NzkyNGQ2NWE4YWEyMjczNmFiMjJlMjc%3D&sdata=blEweHBPcDdWTVNmcVpmUVN3YnEraXdubHpxKzJJbGdXdllkMGFEVHBlTT0%3D&ovuser=e9ddfc3e-ea68-424f-b555-d9a2b41213ad%2CSLee%40venturafoods.com&OR=Teams-HL&CT=1703184695160&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiIyNy8yMzExMDIyNDcwNSIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D

# OTIF
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/7E766C3EDA4DC0E70BD08586AF56C8CB/K1717--K77



closed <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.01/closed.xlsx")
completed <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.01/completed.xlsx")
otif <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.01/otif.xlsx")





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
  dplyr::relocate(closure_in_2_days, .before = issue) -> closed_cleaned

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



writexl::write_xlsx(master_data_closed_rds, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.01/master_data_closed.xlsx")

#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
# ETL: completed file




completed_cleaned %>%
  dplyr::filter(closed >= start_date & closed <= end_date) %>% 
  dplyr::select(-task_status, -request_status, -created_by, -created_month, -item_type, -path, -close_day, -created_day) %>% 
  dplyr::mutate(department = stringr::str_extract(request_type, "^[^:]*")) %>% 
  dplyr::mutate(year_month = format(created, "%Y/%m")) %>% 
  mutate(task_response_time_days = bizdays(created, task_response_entered, 'CustomWeekends')) %>% 
  dplyr::mutate(bracket = ifelse(task_response_time_days < 2, "<=1", 
                                 ifelse(task_response_time_days < 3, "<=2", 
                                        ifelse(task_response_time_days < 4, "<=3",
                                               ifelse(task_response_time_days < 5,  "<=4", ">=5"))))) %>% 
  dplyr::relocate(number_days_to_close, task_id, task_owner, created, task_response_entered, year_month, closed,
                  escalation_number, request_type, department, task_response_time_days, bracket) %>% 
  mutate(closure_in_2_days = ifelse(task_response_time_days <= 1, "Met", "Not met")) -> completed_cleaned


master_data_completed_rds %>% 
  dplyr::filter(!(closed >= start_date & closed <= end_date)) -> master_data_completed_rds



rbind(master_data_completed_rds, completed_cleaned) -> master_data_completed_rds





master_data_completed_rds %>%
  mutate(row_id = apply(., 1, paste, collapse = "")) %>%
  distinct(row_id, .keep_all = TRUE) %>%
  select(-row_id) -> master_data_completed_rds



master_data_completed_rds %>% 
  dplyr::mutate(created = as.Date(created),
                closed = as.Date(closed),
                task_response_entered = as.Date(task_response_entered)) -> master_data_completed_rds



saveRDS(master_data_completed_rds, "master_data_completed.rds")

writexl::write_xlsx(master_data_completed_rds, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.01/master_data_completed.xlsx")




## Copy the master file into the new folder
file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2023.12/MASTER_CLOSED.xlsx", 
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.01/MASTER_CLOSED.xlsx")

file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2023.12/MASTER_COMPLETED.xlsx", 
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.01/MASTER_COMPLETED.xlsx")

file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2023.12/MASTER_PPTX.pptx", 
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Service Level Escalation/monthly/2024.01/MASTER_PPTX.pptx")

















# #####################################################################################################################################
# #####################################################################################################################################
# #####################################################################################################################################
# # KPI: Closed
# master_data_closed_rds %>%
#   dplyr::mutate(year_month = paste(year(closed), month(closed, label = TRUE), sep = "/")) %>%
#   dplyr::mutate(year_month = as.Date(paste(year_month, "01", sep = "-"), format = "%Y/%b-%d")) %>%
#   dplyr::group_by(year_month) %>%
#   dplyr::summarise(count = n()) %>% 
#   ggplot2::ggplot(aes(x = year_month, y = count)) +
#   ggplot2::geom_bar(stat = "identity", fill = "#4B8BBE") +
#   ggplot2::geom_text(aes(label = count), vjust = -0.3) +
#   ggplot2::scale_x_date(date_labels = "%Y/%b", date_breaks = "1 month") +
#   ggplot2::labs(title = "# of Tickets Closed", x = NULL, y = NULL) +
#   ggplot2::theme_classic() +
#   ggplot2::theme(plot.title = element_text(hjust = 0.5, size = 20)) -> closed_kpi_1
# 
# # master_data_closed_rds %>%
# #   dplyr::filter(selling_region != "Canada Co Pack") %>%
# #   dplyr::group_by(selling_region, column1) %>%
# #   dplyr::summarise(count = n(), .groups = "drop") %>%
# #   dplyr::group_by(selling_region) %>%
# #   dplyr::mutate(percentage = count / sum(count)) %>%
# #   tidyr::pivot_wider(names_from = column1, values_from = percentage)
# #   
# #   
# #   dplyr::mutate(percentage = count / sum(count)) %>%
# #   ggplot2::ggplot(aes(x = selling_region, y = percentage, fill = selling_region)) +
# #   ggplot2::geom_bar(stat = "identity", position = "stack") +
# #   ggplot2::geom_text(aes(label = scales::percent(percentage, accuracy = 1)), position = position_stack(vjust = 0.5), color = "black") +
# #   ggplot2::scale_fill_brewer(palette = "Set3") +
# #   ggplot2::scale_y_continuous(labels = scales::percent) +
# #   ggplot2::labs(title = NULL, x = NULL, y = NULL) +
# #   ggplot2::theme_classic() +
# #   ggplot2::theme(plot.title = element_text(hjust = 0.5)) -> closed_kpi_2
# 
# 
# # Create a new variable for the created month
# master_data_closed_rds %>%
#   dplyr::mutate(created_month = format(created, "%Y-%m")) %>%
#   dplyr::filter(selling_region != "Canada Co Pack") %>%
#   dplyr::group_by(created_month, selling_region) %>%
#   dplyr::summarise(count = n()) %>%
#   ggplot2::ggplot(aes(x = created_month, y = count, fill = selling_region)) +
#   ggplot2::geom_bar(stat = "identity", position = "stack") +
#   ggplot2::geom_text(aes(label = count), position = position_stack(vjust = 0.5), color = "black") +
#   ggplot2::scale_fill_brewer(palette = "Set3") +
#   ggplot2::labs(title = NULL, x = NULL, y = NULL) +
#   ggplot2::theme_classic() +
#   ggplot2::theme(plot.title = element_text(hjust = 0.5)) -> closed_kpi_3
# 
# master_data_closed_rds %>%
#   mutate(year = lubridate::year(closed)) %>%
#   mutate(month = sprintf("%02d", lubridate::month(closed, label = FALSE))) %>%
#   mutate(year_month = paste0(year, "/", month)) %>%
#   mutate(year_month = as.Date(paste0(year_month, "/01"), format = "%Y/%m/%d")) %>%
#   dplyr::group_by(year_month, closure_in_2_days) %>%
#   dplyr::summarise(count = n(), .groups = "drop") %>%
#   tidyr::pivot_wider(names_from = closure_in_2_days, values_from = count, values_fill = 0) %>%
#   dplyr::mutate(percentage_met = met / (met + notmet)) %>%
#   dplyr::select(year_month, percentage_met) %>%
#   mutate(year_month = format(year_month, "%Y/%m")) %>%
#   mutate(year_month = factor(year_month, levels = sort(unique(year_month)))) %>%
#   ggplot(aes(x = year_month, y = percentage_met, group = 1)) +
#   geom_line() +
#   geom_hline(yintercept = 0.9, linetype = "dashed", color = "red") +
#   scale_y_continuous(labels = scales::percent_format(), limits = c(0, 1)) +
#   labs(x = "Year/Month", y = " ", title = "% of Tickets within 2 Days") +
#   theme_minimal() +
#   theme(plot.title = element_text(hjust = 0.5)) -> closed_kpi_4
# 
# 
# #####################################################################################################################################
# #####################################################################################################################################
# #####################################################################################################################################
# # KPI: Completed
# 
# ### Bracket column needs to be re-examined
# master_data_completed_rds %>%
#   dplyr::filter(lubridate::month(closed) == target_month, lubridate::year(closed) == target_year) %>%
#   dplyr::group_by(department, bracket) %>%
#   dplyr::summarise(count = n()) %>%
#   dplyr::mutate(percentage = count / sum(count)) %>%
#   ggplot2::ggplot(aes(x = department, y = percentage, fill = fct_rev(bracket))) +
#   ggplot2::geom_bar(stat = "identity", position = "fill") +
#   ggplot2::geom_text(aes(label = paste0(round(percentage * 100), "%")), position = position_fill(vjust = 0.5)) +
#   ggplot2::scale_y_continuous(labels = scales::percent) +
#   ggplot2::scale_fill_brewer(palette = "Set3", name = "Bracket") +
#   ggplot2::labs(title = "Response compliance % by Department", x = NULL, y = NULL) +
#   ggplot2::theme_classic() +
#   ggplot2::theme(plot.title = element_text(hjust = 0.5, size = 20),
#                  axis.text.x = element_text(size = 14, face = "bold"),
#                  axis.text.y = element_text(size = 14, face = "bold")) +
#   ggplot2::coord_cartesian(xlim = c(NA, NA), ylim = c(NA, NA), expand = TRUE) -> completed_kpi_1
# 
# 
# 
# 
# master_data_completed_rds %>%
#   dplyr::filter(lubridate::month(closed) == target_month, lubridate::year(closed) == target_year) %>%
#   dplyr::group_by(request_type) %>%
#   dplyr::summarise(count = n()) %>%
#   dplyr::arrange(desc(count)) %>%
#   ggplot2::ggplot(aes(x = reorder(request_type, -count), y = count)) +
#   ggplot2::geom_bar(stat = "identity", fill = "lightblue4") +
#   ggplot2::geom_text(aes(label = count), vjust = -0.3, color = "white") +
#   ggplot2::theme_classic() +
#   ggplot2::theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
#   ggplot2::labs(title = "Count of Types of Request", x = NULL, y = NULL) +
#   ggplot2::geom_text(aes(label = count), vjust = -0.3, color = "black") +
#   ggplot2::theme(axis.text.x = element_text(angle = 0, hjust = 1)) +
#   ggplot2::scale_x_discrete(labels = function(x) str_wrap(x, width = 10)) +
#   theme(plot.title = element_text(hjust = 0.5))-> completed_kpi_2
# 
# 
# 
# master_data_completed_rds %>%
#   mutate(year = lubridate::year(closed)) %>%
#   mutate(month = sprintf("%02d", lubridate::month(closed, label = FALSE))) %>%
#   mutate(year_month = paste0(year, "/", month)) %>%
#   mutate(year_month = as.Date(paste0(year_month, "/01"), format = "%Y/%m/%d")) %>%
#   dplyr::group_by(year_month, closure_in_2_days) %>%
#   dplyr::summarise(count = n(), .groups = "drop") %>%
#   tidyr::pivot_wider(names_from = closure_in_2_days, values_from = count, values_fill = 0) %>%
#   dplyr::mutate(percentage_met = met / (met + notmet)) %>%
#   dplyr::select(year_month, percentage_met) %>%
#   mutate(year_month = format(year_month, "%Y/%m")) %>%
#   mutate(year_month = factor(year_month, levels = sort(unique(year_month)))) %>%
#   ggplot(aes(x = year_month, y = percentage_met, group = 1)) +
#   geom_line() +
#   geom_hline(yintercept = 0.9, linetype = "dashed", color = "red") +
#   scale_y_continuous(labels = scales::percent_format(), limits = c(0, 1)) +
#   labs(x = NULL, y = NULL, title = "Response Time Performance") +
#   theme_minimal() +
#   theme(plot.title = element_text(hjust = 0.5),
#         axis.text.x = element_text(face = "bold"),
#         axis.text.y = element_text(face = "bold")) -> completed_kpi_3
# 
# ############################################################################################################
# completed_kpi_3
# completed_kpi_1
# completed_kpi_2
# 
# closed_kpi_4
# closed_kpi_1
# # closed_kpi_2
# closed_kpi_3 
# 
# 
