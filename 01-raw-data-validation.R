# Converted from 01-raw-data-validation(1).Rmd
# Each original executable R Markdown chunk is a Positron # %% code cell.

# R Markdown metadata ----
# title: "1_Raw_Data_Validation"
# author: "Zach Friedman + Liz"
# date: "`r Sys.Date()`"
# output:
#   html_document:
#     toc: yes
#     df_print: paged
#   word_document:
#     keep_md: yes
#     toc: yes
# file summary: This file loads in course schedule data and validates the data
# reviewer: Jenny Katz
# editor_options:
#   markdown:
#     wrap: 72


# %% setup
# Original R Markdown options: include=FALSE
knitr::opts_chunk$set(echo = TRUE)


# Key Links ----

# - [Data
#   Dictionary:](https://erstrategies1.sharepoint.com/:x:/s/orgfiles/EW4fU7_js69Hl9DzoYTuB7gBVUamadRtarFtyz6-5jampA?e=c5s2Ev)

# - [Style
#   Guide:](https://app.tettra.co/teams/ersknowledge/pages/coding-style-guide)

# - [Data Request and
#   Validation](https://erstrategies1.sharepoint.com/:x:/s/orgfiles/EcSQElbD3mlMsc_HzVYj_KUBJfCavz6HVlHWD7cN3PAAKw?e=CuNtD8)

# Overview of Script ----

# Part I: Setup ----

# This part of the script allows you to:

# - Load the functions that R will be using to work with your data (Excel
#   automatically loads all its functions - if, countif, index, match,
#   etc. - when you open it, but R requires manual loading of functions if
#   you want full access to what it offers)

# - Specify the SharePoint folder where your raw data lives so R can pull
#   it in

# - Actually load in the raw data you'll be working with throughout the
#   rest of the script

# 1.1 Load Packages ----

# **What:** This code block loads in all relevant function packages
# (a.k.a. R functions you'll be needing in this script). Note that if you
# haven't installed the packages into your computer yet, R will print an
# error and you'll need to install them first by running
# install.packages(tidyverse) and install.packages(formattable), etc.

# **Check:** no error message coming out of code block

# **Project Team Action:** do not change (run as is)

# %% setup & load data
# Original R Markdown options: include=FALSE
# Remove everything from any previous sessions
rm(list = ls())

# This section sets up our file to correctly output the data for the school we are interested in; I am going to work on just testing for one school in order to make sure that the code is working correctly.
knitr::opts_chunk$set(
  echo = FALSE,
  fig.width = 7, 
  fig.height = 5,
  message = FALSE,  # Hide messages
  warning = FALSE
)

# Load necessary libraries
library(tidyverse)
library(ggplot2)

#load data
library(openxlsx) # Provides functions for reading, writing, and manipulating Excel files: read.xlsx, write.xlsx
library(here) # A simpler way to manage file paths: here() generates paths relative to the project root
library(Microsoft365R) # Use Microsoft365R to connect to SharePoint and load files
library(readxl) # Provides functions for reading Excel files: read_excel()

#manipulate data
library(tidyverse) # Includes ggplot2, dplyr, tidyr, readr, etc.
library(data.table) # fread(): load large CSVs quickly
library(dplyr) # Functions for data manipulation: filter(), select(), mutate(), etc. (part of tidyverse)
library(safejoin) # Provides safer join operations for 1:many matching - safe_left_join()
library(lubridate) # Provides functions for working with date-times: ymd(), hms(), today()
library(scales) # Provides functions for percents

#format and display data
library(flextable) # Creates and formats tables in a publication-ready format: flextable()
library(grid)
library(gtable)
library(gridExtra)
library(glue)

# Load and register fonts (run only once per session)
library(extrafont)
loadfonts(device = "win")  # For Windows; use loadfonts() for Mac

library(erstools) # Load erstools

### 0.2 Connect to internal and external sharepoint sites

#set sharepoint site URL for access to our org-wide and external files
orgfiles_site_url <- "https://erstrategies1.sharepoint.com/sites/orgfiles"

#access sharepoint site URL you may be prompted to sign in here
orgfiles_site <- get_sharepoint_site(site_url = orgfiles_site_url)

### 0.3 Connect to Client Work and Internal drives within sharepoint sites

#see all folders in client work drive
client_work_drive <- orgfiles_site$get_drive("Client Work")

#see all folders in internal work drive
internal_drive <- orgfiles_site$get_drive("Internal")


# 1.2 Declare Variables ----

# **What:** For R to know where to access your raw data, it needs to know
# where the raw data is located in SharePoint and the file names. The code
# block below allows you to first name your username (in the structure
# "C:/Users/username", which is typically the beginning of a folder path)
# and then the path to reach your folder with the raw data.

# **Check:** The code block should print the folder path to your raw data.

# **Project Team Action:** definitely change (see action guidance in code
# block)

# b. File Path ----

# Take green comments and move directly above code chunk

# %% b. File Path
# Create folder path
raw_data_folder_path <- "District Partners/Detroit Public Schools/26-27 HS Redesign Implementation/1. Data & Analysis - Secure/Fall 2026 Course Schedule Analysis"


# 1.3 Load Data ----

# **What:** Now that R knows how to access the folder path, it's time to
# load in your raw data

# **Check:** the code block should print the names of all data files you
# need.

# **Project Team Action:** None

# %% 1.3 Load Data
# Load course schedule data
cs_raw_file <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/0. Raw Data/course_sections_251103.csv")

expression_coded <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/1. Coding/expression_coded.xlsx")

location_coded <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/1. Coding/location_coded.xlsx")

#cs_validated_file <- ers_read_sharepoint(
#  folder_path = raw_data_folder_path,
#  file_name_with_extension =
#    "/1.1 Validated Data/course_sections_24_25_with_CTC.csv")

# Load student demographic data
#stu_demographics <- ers_read_sharepoint(
#  folder_path = raw_data_folder_path,
#  file_name_with_extension =
#    "/Raw Data/student_enrollment_and_demographic.xlsx")



# Part II: Pre-Validation Setup ----

# 2.1 Create Checks Table ----

# - **What:** Creates a function that adds a data validation check to a
#   main table and that groups the checks table by all types of checks and
#   outputs just the most recent row for each check.

# - **Check:** No output should be printed but you still must run the
#   code. Running the code more than once is OK and will not cause any
#   issues.

# - **Project Team Action:** do not change (run as is)

# %% 2.1 Create Checks Table
# Original R Markdown options: warning = F, message = F

# create checks table in the current GitHub repo
checks_file_csv <- "checks_table.csv"

# delete existing file so every run starts fresh
if (file.exists(checks_file_csv)) {
  file.remove(checks_file_csv)
}

# create checks table function
init_checks_table <- function(file_path = checks_file_csv) {
  
  if (file.exists(file_path)) {
    checks_table <- read_csv(
      file_path,
      show_col_types = FALSE,
      col_types = cols(
        check_id = col_character(),
        check_name = col_character(),
        value = col_character(),
        notes = col_character(),
        status = col_character(),
        updated_at = col_datetime()
      )
    )
  } else {
    checks_table <- tibble(
      check_id = character(),
      check_name = character(),
      value = character(),
      notes = character(),
      status = character(),
      updated_at = as.POSIXct(character())
    )
    
    write_csv(checks_table, file_path)
  }
  
  checks_table
}

update_check <- function(check_id_value,
                         check_name,
                         value,
                         notes = NA_character_,
                         status = NA_character_,
                         file_path = checks_file_csv) {
  
  checks_table <- init_checks_table(file_path)
  
  new_row <- tibble(
    check_id = as.character(check_id_value),
    check_name = as.character(check_name),
    value = as.character(value),
    notes = as.character(notes),
    status = as.character(status),
    updated_at = now()
  )
  
  checks_table <- checks_table %>%
    filter(check_id != as.character(check_id_value)) %>%
    bind_rows(new_row)
  
  write_csv(checks_table, file_path)
  
  invisible(checks_table)
}

get_checks_table <- function(file_path = checks_file_csv) {
  init_checks_table(file_path)
}


# 2.2 Understand File Structure ----

# a. Number of total rows ----

# - **What:** Understand the structure of your file by loading the number
#   of rows

# - **Check:** The number of rows should be roughly 8 times the number of
#   students in the district

# - **Project Team Action:** Update status if necessary

# %% a. Number of total rows
# Count number of rows
update_check(
  check_id = 1,
  check_name = "Number of rows in data frame",
  value = nrow(cs_raw_file),
      notes = "Depends on number of students and number of course records you'd expect",
  status = "Done"
)

checks_table <- get_checks_table()

get_checks_table()

# Print list of column names
names(cs_raw_file)


# b. Number of total students ----

# - **What:** Transform your data if necessary. You can also leave a note
#   about district follow-up required if the team did not receive the full
#   data set.

# - **Check:** The number of rows should be roughly 8 times the number of
#   students in the district

# - **Project Team Action:** maybe change (see guidance in code block)

# %% b. Number of total students
# Filter to just the most recent year
cs_raw_file <- cs_raw_file |>
  filter(Year_End == 2026) |>
  mutate(
    D_stu_id = MaskedStudentID,
    .keep = "unused")

update_check(
  check_id = 2,
  check_name = "Number of students", 
  value = n_distinct(cs_raw_file$D_stu_id),
  notes = "Should match what you'd expect from public data",
  status = "Done"
  )

checks_table <- get_checks_table()

get_checks_table()

# 2.3 Column Mapping ----

# a. Review current columns ----

# - **What:** View the data as a table to understand the values in each
#   column.

# - **Check:** A new tab is created with a table view of your data

# - **Project Team Action:** do not change (run as is)

# b. Map project columns to ERS Standard Names ----

# - **What:** Assign your project's column names to ERS standard column
#   names. ERS Standard names are on the left side of the equal sign, your
#   district's columns are on the right. See resolution of
#   critical/important fields that are missing in step 2.2b. Having
#   standard column names is critical for this code running without error.

# - **Check:** In your data, all the critical fields are present and most
#   of the important fields are present. Align with your project lead on
#   which fields are critical based on your project's context. For
#   exmaple:

#   - If you received active course enrollments from snapshot in time vs.
#     all historical course enrollments from that year

#   - If you received Period and Rotation as an "Expression" - e.g.,
#     "2-8(A)", "7(B)", "1-4(A), 6-8(B)", etc. This is how it's typically
#     exported from the Student Information System (SIS) PowerSchool. If a
#     record has Expression "2-8(A)", then that student has that course
#     during Periods 2 through 8 on A days. You'll deal with "exploding"
#     this record in the 04-db-explosions script, so don't worry about it
#     for now!

# - **Project Team Action:** definitely change (see guidance in code
#   block)

# %% b. Map project columns to ERS Standard Names
# If the ERS Standard column is present in your data then type your district's column on the right hand side (it is ok if the names are identical)
# If the ERS Standard column is not present in your data then type "MISSING"

# CRITICAL for ALL course schedule data
cs_raw_file <- cs_raw_file |>
  select(
    -`Grade Scale ID`,
    -`Grade Scale Name`,
    -`AP Flag`,
    -`Accelerated Course Flag`,
    -`Virtual Flag`,
    -`Virtual Delivery Type`
  ) |>
  mutate(
    # Core fields
    # D_stu_id = `DPSCD Masked Student Id`,
    D_employee_id = MaskedTeacherID,
    D_year_id = Year_End,
    D_ccid = CCID,
    #D_users_dcid = USERS_DCID,

    D_course_id = `Course Number`,
    D_course_name = `Course Name`,
    #D_course_subject = "MISSING",
    D_sced_subject_code = `SCED Code`,

    D_term = "MISSING",
    D_expression = `Period Expression`,
    D_period = "MISSING",
    D_rotation = "MISSING",
    D_location_id = EEM_CODE,
    D_location_name = "MISSING",

    D_course_section = "MISSING",
    #D_course_section_id = "MISSING",
    D_stu_grade = "MISSING",

    # Enrollment dates
    D_stu_enter_date = `Date Enrolled`,
    D_stu_exit_date = `Date Left`,
    #D_course_start_date = "MISSING",
    #D_course_end_date = "MISSING",

    # Important fields not provided
    #D_course_credit_recovery = "MISSING",
    #D_course_ell_flag = "MISSING",
    #D_course_swd_flag = "MISSING",
    #D_course_rigor = "MISSING",
    #D_course_format = "MISSING",
    #D_home_school = "MISSING",

    #D_stu_swd_flag = "MISSING",
    #D_stu_ell_flag = "MISSING",
    #D_stu_poverty_flag = "MISSING",
    #D_stu_race = "MISSING",

    # Nice-to-have fields not provided
    #D_course_pathway = "MISSING",
    #D_cte_flag = "MISSING",
    #D_class_size_max = "MISSING",
    #D_credits_earned = "MISSING",
    #D_credits_possible = "MISSING",
    #D_period_start_time = "MISSING",
    #D_period_end_time = "MISSING",
    #D_room = "MISSING",

    .keep = "unused"
  )


# %% Check expression --> period and rotation conversion
# Check expression --> period and rotation conversion
expression_check <- cs_raw_file |>
  group_by(
    D_expression,
    D_period,
    D_rotation
  ) |>
    summarise(count = n())



# c. Review Missing Columns ----

# - **What:** List all of the columns that are missing

# - **Check:** Ensure that none of these fields are necessary for your
#   analysis. If they are necessary, then the missing columns must be
#   resolved in step 2.2b

# - **Project Team Action:** do not change (run as is)

# %% c. Review Missing Columns
# Create function called "check_missing_columns" that prints out all columns that have any values of "MISSING"

update_check(
     check_id = 3,
     check_name = "Missing Columns in Data Frame", 
     value = (cs_raw_file %>%
      select(where(is.character)) %>%
      select_if(~ any(. == "MISSING")) %>%
      colnames() %>%
      paste0(collapse = ", ")),
     notes = "There should be no critical fields missing",
     status = "In Progress" # Update status if it has changed. Do not change any other values
)

checks_table <- get_checks_table()

get_checks_table()


# d. Add or edit critical columns ----

# - **What:** Resolve missing columns that are necessary for your analysis and add them so the code will run smoothly.

# - **Check:** Columns that you resolved do not show up in the code chunk below that checks for missing columns

# - **Project Team Action:** change depending on your files needs. The below example is from Springfield that needed to update Period to "MISSING" and pull in student demographic data to update the student-level flags. You do not need to do either of these things, but this is an example of how you might resolve missing critical fields. If you have other critical fields that are missing, you should resolve them here and then re-run the check for missing columns to ensure they are resolved.

# **Project Example:** Updating column D_period

# %% Project Example: Updating column D_period
# Convert to string
#cs_raw_file <- cs_raw_file |>
#  mutate(D_course_section = as.character(D_course_section))

# Remove blank perriod and cycle day ID columns
#cs_raw_file <- cs_raw_file %>%
#  select(-CYCLE_DAY_ID, -Period)



# %% d. Add or edit critical columns
# View student demographic data
#View(stu_demographics)

# Print list of column names
#names(stu_demographics)

# Convert decimals into integers
#stu_demographics <- stu_demographics |>
#  mutate(across(where(is.numeric), ~ if (all(. %% 1 == 0, na.rm = TRUE)) as.integer(.) else .))

# Update column names in student demographic file
#stu_demographics <- stu_demographics %>% 
#  rename(D_stu_id = Stu_ID,
#         D_stu_grade = `Current Grade Level ID`,
#         D_location_id = `SIS School ID`,
#         D_stu_swd_flag = `Stu_SWD flag`,
#         D_stu_ell_flag = `Stu_ELL flag`,
#         D_stu_poverty_flag = `Stu_Poverty Flag`)

# Check that student demographics file has only one row per stu_id
#stu_demographics %>% 
#  group_by(D_stu_id) %>% 
#  summarise(count = n()) %>% 
#  group_by(count) %>% 
#  summarise(count_sum = sum(count),
#            stu_count = n_distinct(D_stu_id))


# %% Join demographics file
# Join demographics file
# Prepare demographics data for merge
#stu_demographics <- stu_demographics |>
#  mutate(across(
#    c(D_stu_swd_flag, D_stu_ell_flag, D_stu_poverty_flag),
#    ~ recode(.x, "Y" = 1L, "N" = 0L) |> as.integer()
#  )) |> 
#  select(D_stu_id, D_stu_grade, D_stu_swd_flag, D_stu_ell_flag, D_stu_poverty_flag)

# Remove placeholder columns, then merge
#cs_raw_file <- cs_raw_file %>%
#  select(-D_stu_grade,
#         -D_stu_swd_flag,
#         -D_stu_ell_flag,
#         -D_stu_poverty_flag) %>%
#  left_join(stu_demographics,
#            by = "D_stu_id")



# %%
# Detroit: filter for rows with D_location_name (other rows are for elementary and middle schools)
cs_raw_file <- cs_raw_file |>
  filter(!is.na(D_location_name))


# %%
# Merge in flag from location_coded: if C_eem_code is in location_coded, create new column C_location_type = "Neighborhood School", else "Non-Neighborhood School"
cs_raw_file <- cs_raw_file |>
  left_join(
    location_coded,
    by = c(
      "D_location_name",
      "D_location_id"
  )
)

# Filter out data from non-neighborhood schools
cs_raw_file <- cs_raw_file |>
  filter(!is.na(C_location_type))

# %%
# Update D_term based on course name
cs_raw_file <- cs_raw_file |>
  mutate(
    D_term = case_when(
      str_detect(D_course_name, "-\\s*A\\s*\\*?\\s*$") ~ "S1",
      str_detect(D_course_name, "-\\s*B\\s*\\*?\\s*$") ~ "S2",
      TRUE ~ "FY"
    )
  )

# Check course names for records with term FY
check_fy_courses <- cs_raw_file |>
  filter(D_term == "FY") |>
  group_by(D_course_name) |>
  summarise(count = n())

# %%
# Join period and rotation data by expression
cs_raw_file <- cs_raw_file |>
  select(
    -D_period,
    -D_rotation
  ) |>
    left_join(
      expression_coded,
      by = "D_expression"
  )

# check distribution of term, period, rotation
check_tpr <- cs_raw_file |>
  group_by(
    D_term,
    D_expression,
    D_period,
    D_rotation
  ) |>
    summarise(count = n())

# e. Re-run Validation after resolution ----

# - **What:** check the missing columns once again to ensure the issues
#   were resolved

# - **Check:** No important columns should be missing

# - **Project Team Action:** Update the status as necessary

# %% e. Re-run Validation after resolution
# Rerun check
update_check(
     check_id = 3,
     check_name = "Missing Columns in Data Frame", 
     value = (cs_raw_file %>%
      select(where(is.character)) %>%
      select_if(~ any(. == "MISSING")) %>%
      colnames() %>%
      paste0(collapse = ", ")),
     notes = "There should be no critical fields missing",
     status = "In Progress" # Update status if it has changed. Do not change any other values
)

checks_table <- get_checks_table()

get_checks_table()


# Part III: Validation checks ----

# 3.1 What year(s) of data do we have? ----

# - **What:** Understand if we have the correct year of data.
#   **IMPORTANT!** This is only necessary if you if receive all historical
#   course enrollments (i.e., if your data is NOT a snapshot in time)

# - **Check:** Ensure you know how many years of data are present and if
#   it the years are correct.

# - **Project Team Action:** Change status as necessary

# %% 3.1 What year(s) of data do we have?
# Update the Date columns from character to date 
cs_raw_file <- cs_raw_file %>% mutate(D_stu_exit_date = mdy(D_stu_exit_date))
cs_raw_file <- cs_raw_file %>% mutate(D_stu_enter_date = mdy(D_stu_enter_date))

# Now get the max and min dates
max_date <- max(cs_raw_file$D_stu_enter_date, na.rm = TRUE)
min_date <- min(cs_raw_file$D_stu_enter_date, na.rm = TRUE)


# Then add them to the data validation check table
update_check(
     check_id = 3,
     check_name = "Year of Data", 
     value = paste0(min_date, " through ", max_date),
     notes = "Our year should align with expectations",
     status = "Done" # Update if it has changed. Do not change any other values
)
checks_table <- get_checks_table()

get_checks_table()



# 3.2 Missing & Unique values ----

# a. Update Blanks or NULL to NA ----

# - **What**: This will update all blanks ("") or nulls with NA, a special
#   value in R that denotes missing data. This will make it easier to deal
#   with missing data throughout the code.
# - **Check**: see below
# - **Project Team Action**: None

# %% a. Update Blanks or NULL to NA
# Update blanks or NULL to NA
cs_raw_file <- cs_raw_file %>%
  mutate(across(everything(), ~if_else(. %in% c("", NULL), NA, .)))


# b. Review data with blanks or NAs ----

# - **What:** Understand if we have any blanks or NAs in key fields

# - **Check:** -- Are there any critical columns with NA or blank values?
#   -- Do the Unique values for things like Teacher ID, Student ID, and
#   D_location match expectations?

# - **Project Team Action:** review output and note any key columns with a
#   large % of missing values. Are they critical? Do you need to ask
#   district about missing data? Discuss any issues with project team.

# %% b. Review data with blanks or NAs
# Create new data frame with column information
col_info <-  data.frame(
  Column_Name = names(cs_raw_file),
  Data_Type = sapply(cs_raw_file, class),
  Count_Rows = sapply(cs_raw_file, function(x) length(na.omit(x))),
  NA_Rows = sapply(cs_raw_file, function(x) sum(is.na(x))),
  Unique_Rows = sapply(cs_raw_file, function(x) length(unique(na.omit(x)))),
  stringsAsFactors = FALSE,
  row.names= NULL
)

# Calc % missing
col_info <- col_info %>%
            mutate(Percent_NA = NA_Rows / Count_Rows * 100)

# Print new data frame
col_info

# Add primary keys to data checks table (employee, teacher, school)
# Teachers
update_check(
     check_id = 4,
     check_name = "Unique count of teacher IDs", 
     value = cs_raw_file %>% pull(D_employee_id) %>% unique %>% length(),
     notes = "Compare to public data",
     status = "Done" 
) %>% head
get_checks_table() #print a record of all the most recent checks

# Students
update_check(
     check_id = 5,
     check_name = "Unique count of student IDs", 
     value = cs_raw_file %>% pull(D_stu_id) %>% unique %>% length(),
     notes = "Compare to public data",
     status = "Done" 
) %>% head
get_checks_table() #print a record of all the most recent checks

# Schools (LK: changed from ID to name)
update_check(
     check_id = 6,
     check_name = "Unique count of school names", 
     value = cs_raw_file %>% pull(D_location_name) %>% unique %>% length(),
     notes = "Compare to public data",
     status = "Done" 
) %>% head
get_checks_table() #print a record of all the most recent checks

# Ratio of rows per student at each school (LK: changed from ID to name)
location_ratios <- cs_raw_file |>
  group_by(D_location_name) |>
  summarize(
    n_rows = n(),
    n_students = n_distinct(D_stu_id),
    ratio = n_rows / n_students,
    .groups = "drop"
  ) |>
  mutate(
    ratio_str = paste0(D_location_name, ": ", round(ratio, 2))
  )

value_str <- paste(location_ratios$ratio_str, collapse = ", ")

update_check(
  check_id = 7,
  check_name = "Avg. rows per student by school",
  value = value_str,
  notes = "Format: location_ID: ratio",
  status = "Done"
) %>% head
get_checks_table() #print a record of all the most recent checks

checks_table <- get_checks_table()
get_checks_table()
View(col_info)



# %% Check rows that have any NAs to see if these are records we should keep or remove
# Check rows that have any NAs to see if these are records we should keep or remove
# Identify columns with NA values
cols_with_na <- col_info %>%
  filter(NA_Rows > 0) %>%
  pull(Column_Name)

# Create NA summary by course
na_summary_by_course <- cs_raw_file %>%
  select(D_course_name, all_of(cols_with_na)) %>%
  group_by(D_course_name) %>%
  summarise(
    
    # Total rows for the course
    total_rows = n(),
    
    # Rows where ANY tracked column is NA
    rows_with_any_na = sum(if_any(all_of(cols_with_na), is.na)),
    
    # NA count for each column
    across(
      all_of(cols_with_na),
      ~ sum(is.na(.)),
      .names = "NA_{.col}"
    ),
    
    .groups = "drop"
  ) %>%
  filter(rows_with_any_na > 0) %>%
  arrange(desc(rows_with_any_na), D_course_name)


# c. Remove records for students without demographic data ----

# - **What:** Remove rows that are missing key fields

# - **Check:** -- Are there any critical columns with NA or blank values?
#   -- Do the Unique values for things like Teacher ID, Student ID, and
#   D_location match expectations?

# - **Project Team Action:** review output and note any key columns with a
#   large % of missing values. Are they critical? Do you need to ask
#   district about missing data? Discuss any issues with project team.

# %% c. Remove records for students without demographic data
# Count unique students with NA in any demographic column
n_students_na_demog <- cs_raw_file %>%
  filter(if_any(all_of(names(stu_demographics)), is.na)) %>%
  pull(D_stu_id) %>%
  unique() %>%
  length()

# Add to checks table
update_check(
  check_id = 8,
  check_name = "Students with NA in demographics",
  value = n_students_na_demog,
  notes = "Count of unique students with NA in any demographic column",
  status = "In Progress"
) %>% head
get_checks_table() #print a record of all the most recent checks

checks_table <- get_checks_table()



# %% Filter out records with missing student demographic data
# Filter out records with missing student demographic data
cs_raw_file <- cs_raw_file %>%
  filter(if_all(all_of(names(stu_demographics)), ~ !is.na(.)))

# Update checks table
update_check(
  check_id = 8,
  check_name = "Students with NA in demographics",
  value = n_students_na_demog,
  notes = "Count of unique students with NA in any demographic column",
  status = "Done"
) %>% head
get_checks_table() #print a record of all the most recent checks

# Students
update_check(
     check_id = 9,
     check_name = "Final unique count of student IDs", 
     value = cs_raw_file %>% pull(D_stu_id) %>% unique %>% length(),
     notes = "Compare to public data",
     status = "Done" 
) %>% head
get_checks_table() #print a record of all the most recent checks

checks_table <- get_checks_table()



# 3.3 Are there any duplicates? ----

# a. Review Duplicates Table ----

# - **What:** Understand if we have any fully duplicated rows

# - **Check:** Review the cs_raw_file_dups table to understand the
#   duplicates and why they are appearing before deleting

# %% a. Review Duplicates Table
#View duplicates
dups_cs_raw_file <- cs_raw_file %>%
  group_by_all() %>%
 mutate(n = n()) %>%
    filter(n>1)

#add nrows to checks table
update_check(
     check_id = 10,
     check_name = "Check count of full duplicates", 
     value = dups_cs_raw_file %>% nrow(),
     notes = "There should be no full duplicates",
     status = "Initial check" #Update status if it has changed. Do not change any other values
) %>% head
get_checks_table() #print a record of all the most recent checks

checks_table <- get_checks_table()
get_checks_table()


# b. Remove Duplicates ----

# - **What:** The below code chunk will remove the duplicate rows, so only
#   run once you feel confident they are true duplicates

# - **Check:** That the number of duplicates is now 0

# %% b. Remove Duplicates
# Removes duplicates
cs_raw_file <- cs_raw_file[!duplicated(cs_raw_file), ]

# Update checks table
update_check(
     check_id = 10,
     check_name = "Check count of full duplicates", 
     value = cs_raw_file[duplicated(cs_raw_file,by = NULL),] %>% nrow,
     notes = "There should be no full duplicates",
     status = "Done" # Update status if it has changed. Do not change any other values
) %>% head
get_checks_table() # Print a record of all the most recent checks

# ------------
  # 3.4 Count of Periods at each school ----

# - **What:** Count of records by period at each school

# - **Check:** Understand if we have similar # of records per period, by
#   school for "core periods" - after school, HR, or other non-core
#   periods might have fewer/more records

# - **Project Action:** Review the tables to understand the different
#   periods at each school and if there is an even distribution of periods
#   in the school

# %% 3.4 Count of Periods at each school
# View count of records by D_term/school
count_periods_by_school <- cs_raw_file %>%
  filter(complete.cases(`D_period`)) %>%
  group_by(D_location_name, `D_period`) %>% 
  count(sort = TRUE) %>%
  arrange(D_location_name, D_period) %>%
  group_by(D_location_name) %>% 
  mutate(perc_periods_at_school = round(100*n/sum(n),1)) %>%
  select(D_location_name,`D_period`,perc_periods_at_school)%>%
  pivot_wider(names_from = `D_period`, 
              values_from = perc_periods_at_school, 
              names_prefix = "Period_") %>%
  rename_all(~ ifelse(. == "", "Period_unknown", .))

# Springfield finding: schools do NOT have similar number of records per period
checks_table <- update_check(
  check_id = 11,
  check_name = "Check count of records by period by school. What % are 1-8?", 
  value = cs_raw_file %>%
    group_by(D_period) %>%
    count(sort = TRUE) %>%
    ungroup() %>%
    mutate(perc_periods_at_school = round(100 * n / sum(n), 1)) %>%
    filter(D_period %in% c("1", "2", "3", "4", "5", "6", "7", "8")) %>%
    pull(perc_periods_at_school) %>%
    sum(),
  notes = "There should be a similar # of records by period per school IF schools have a similar schedule (e.g., all have periods 1-8)",
  status = "Done"
)

get_checks_table() # Print a record of all the most recent checks



# 3.5 Count of Term at each school ----

# - **What:** Count of records by term at each school

# - **check:** Review the tables to understand the different terms at each
#   school and if there is an even distribution of semester 1 and semester
#   2 classes

# - **Project Action:** No updates

# %% 3.5 Count of Term at each school
# View count of records by D_term/school
count_term_by_school <- cs_raw_file %>%
  filter(complete.cases(`D_term`)) %>%
  group_by(D_location_name, `D_term`) %>% 
  count(sort = TRUE) %>%
  arrange(D_location_name) %>%
  group_by(D_location_name) %>% 
  mutate(perc_D_term_at_school = round(100*n/sum(n),1)) %>%
  select(D_location_name,`D_term`,perc_D_term_at_school)%>%
  pivot_wider(names_from = `D_term`, 
              values_from = perc_D_term_at_school, 
              names_prefix = "D_term_") %>%
  rename_all(~ ifelse(. == "", "D_term_unknown", .))


cs_raw_file %>% 
  group_by(D_term, D_location_name) %>% 
  summarize(student_count = n_distinct(D_stu_id)) %>% 
  arrange(D_location_name) %>% 
  group_by(D_location_name) %>% 
  mutate(percent_students = round(100*student_count/sum(student_count),0))


# 3.6 Distinct Count of Course IDs by Teacher ID ----

# a. Return the count ----

# - **What:** Count of teachers by distinct course count

# - **check:** Review the counts teachers by course count. We would expect
#   majority of teachers to have between 3-10 classes.

# - **Project Action:** No updates

# %% a. Return the count
# Original R Markdown options: echo=TRUE
# Creates a table that is the employee ID by number of courses they are teaching
course_count_list <- cs_raw_file %>% 
  group_by(D_employee_id) %>% 
  summarize(course_count = n_distinct(D_course_name),
            student_load = n_distinct(D_stu_id)) %>% 
  mutate(course_count_bucket = case_when(course_count == 0 ~ "0",
          course_count <= 2 ~ "01-02",
          course_count <= 4 ~ "03-04",
          course_count <= 6 ~ "05-06",
          course_count <= 8 ~ "07-08",
          course_count <= 10 ~ "09-10",
          course_count <= 12 ~ "11-12",
          course_count <= 16 ~ "13-16",
          course_count <= 20 ~ "17-20",
          course_count > 20 ~ "21+"
          ))

#returns the number of teachers by course count bucket
total_teachers <- sum(course_count_list$course_count)

course_count_by_bucket <- course_count_list %>%
  group_by(course_count_bucket) %>%
  summarize(n = sum(course_count)) %>%
  mutate(
    total_teachers = total_teachers,
    pct_of_total = round(100 * (n / total_teachers), 2)
  )



# b. Investigate teachers with large number of course IDs ----

# - **What:** Returns a list of teachers with high number of course IDs
#   and their student load

# - **Check:** We usually should not see more than 10 per term. Are there
#   any placeholder teacher ids that you want to exclude?

# - **Project Action:** If you want to include more columns in your
#   investigation, add them at the end of the list in the select() clause

# %% b. Investigate teachers with large number of course IDs
#bring the course count flags back
teacher_course_count_large <- cs_raw_file %>% 
  left_join(course_count_list, by = "D_employee_id") %>%
  filter(course_count > 10) %>%
  group_by(D_employee_id, 
           D_course_name,
           course_count, 
           student_load)%>% #add additional columns here
  summarize(course_size = n_distinct(D_stu_id)) %>% 
  arrange(course_count,
          D_employee_id,
          D_course_name,
          student_load,
          course_size)

print(n_distinct(teacher_course_count_large$D_employee_id))


# c. Investigate teachers with small number of course IDs ----

# - **What:** Returns a list of teachers with high number of course IDS
#   and their student load

# - **check:** We usually should not see more then 10 per D_term. Are
#   there any placeholder teacher ids that you want to exclude?

# - **Project Action:** If you want to include more columns in your
#   investigation, add them at the end of the list in the select() clause

# %% c. Investigate teachers with small number of course IDs
teacher_course_count_small <- cs_raw_file %>% 
  left_join(course_count_list, by = "D_employee_id") %>% 
  filter(course_count < 3) %>%
  group_by(D_employee_id,
           D_course_name,
           course_count,
           student_load)%>% #add additional columns here
  summarize(course_size = n_distinct(D_stu_id)) %>% 
  arrange(D_employee_id, D_course_name, course_count, student_load, course_size)

print(n_distinct(teacher_course_count_small$D_employee_id))


# 3.7 Distinct count of students per teacher ID ----

# a. Return the Count ----

# - **What:** Returns a count of students per teacher id and then groups
#   into buckets

# - **Check:** For elementary, we typically see a load between 15-30
#   students; for secondary, can be 30 up to 180-200

# - **Project Action:** None needed

# %% a. Return the Count
# Calculates the student count by unique courses 
student_count_list <- cs_raw_file %>% 
  group_by(D_employee_id) %>% 
  summarize(student_count = n_distinct(D_stu_id)) %>% 
  mutate(student_count_bucket = case_when(student_count == 0 ~ "0",
          student_count <= 5 ~ "1: 01-05",
          student_count <= 15 ~ "2: 06-15",
          student_count <= 30 ~ "3: 16-30",
          student_count <= 50 ~ "4: 31-50",
          student_count <= 100 ~ "5: 51-100",
          student_count <= 150 ~ "6: 101-150",
          student_count <= 200 ~ "7: 151-200",
          student_count <= 250 ~ "8: 201-250",
          student_count > 250 ~ "9: 251+"
          ))

student_count_list %>% group_by(student_count_bucket) %>% count(sort = TRUE)

#returns the number of teachers by course count bucket
total_teachers <- nrow(student_count_list)

student_count_by_bucket <- student_count_list %>%
  group_by(student_count_bucket) %>%
  summarize(n = n()) %>%
  mutate(
    total_teachers = total_teachers,
    pct_of_total = round(100 * (n / total_teachers), 2)
  )


# b. Investigate teachers with large number of student IDs ----

# - **What:** Returns a list of teachers with lots of students

# - **check:** Do they all make sense because of advisory, study hall,
#   electives, etc.?

# - **Project Action:** None needed

# %% b. Investigate teachers with large number of student IDs
teacher_load_large <- cs_raw_file %>% 
  group_by(D_employee_id, D_course_id, D_course_name) %>% 
  left_join(student_count_list, by = "D_employee_id") %>%
  filter(student_count > 200) %>%
  arrange(-student_count,D_employee_id)
  
print(n_distinct(teacher_load_large$D_employee_id))


# 3.8. Distinct Count of Course IDs per student ID ----

# a. Return the Counts ----

# %% a. Return the Counts
student_course_count_list <- cs_raw_file %>% 
  group_by(D_stu_id) %>% 
  summarize(course_count = n_distinct(D_course_id)) %>% 
  mutate(course_count_bucket = case_when(course_count == 0 ~ "0",
          course_count <= 2 ~ "01-2",
          course_count <= 4 ~ "03-4",
          course_count <= 6 ~ "05-6",
          course_count <= 8 ~ "07-8",
          course_count <= 10 ~ "09-10",
          course_count <= 12 ~ "11-12",
          course_count <= 16 ~ "13-16",
          course_count <= 20 ~ "17-20",
          course_count <= 24  ~ "21-24",
          course_count >  24  ~ "25+",
          ))

#returns the number of students by course count bucket
total_students <- nrow(student_course_count_list)

student_count_by_bucket <- student_course_count_list %>%
  group_by(course_count_bucket) %>%
  summarize(n = n()) %>%
  mutate(
    total_students = total_students,
    pct_of_total = round(100 * (n / total_students), 2)
  )


# b. Investigate students with small number of courses ----

# - **What:** Returns a list of students with a small number of courses

# - **check:** Do they all make sense because of advisory, study hall,
#   electives, etc.?

# - **Project Action:** If you want to add more columns for your
#   investigation add to the group_by clause

# %% b. Investigate students with small number of courses
student_course_count_small <- cs_raw_file %>% 
  group_by(D_stu_id, D_course_section, D_course_name) %>%
  count() %>%
  left_join(student_course_count_list, by = "D_stu_id") %>%
  filter(course_count < 3) %>%
  select(-n)

student_course_count_small %>% 
  group_by(D_course_name) %>% 
  summarise(avg_course_count = round(mean(course_count), digits = 0),
            stu_count = n_distinct(D_stu_id))


# 3.9 Distinct Count of Students per Subject per Grade ----

# - **What:** Returns the % of students per grade per school taking a
#   subject

# - **check:** You are looking for wonky things - such as a low percentage
#   (less then 75%) of 5th graders taking a math class. We want to ensure
#   we have complete data

# - **Project Action:** None

#   ```{r}
#   View(student_subject_grade_count <- cs_raw_file %>% 
#     select(D_location_name,
#            D_stu_grade,
#            D_stu_id,
#            D_course_subject)%>% 
#     group_by(D_location_name,
#              D_stu_grade) %>% 
#     mutate(student_grade_count = n_distinct(D_stu_id)) %>% 
#     group_by(D_location_name,
#              D_stu_grade,
#              D_course_subject,
#              student_grade_count)%>% 
#     summarise(student_subject_count = n_distinct(D_stu_id)) %>% 
#     mutate(pct_of_total= 100*(student_subject_count / student_grade_count)) %>% 
#     arrange(D_location_name)
#     )

#   ```

# Part IV. Export Tables ----

# 5.1 Delete or Rename Extra Columns ----

# - **What:**

#   - Before exporting your data to use for all the upcoming scripts, take
#     out any columns that you won't need and rename any variables that
#     weren't given a standard name earlier.

#   - Note: renaming the non-standard columns kept in your data is
#     optional but can help you more easily identify these columns further
#     down the line - label the variables in the same way as the standard
#     variables with a D (district) tag and then the variable name all in
#     lower case and with underscores between words (e.g.
#     D_extra_variable)

# - **Check:** look at your data in the environment after running the code
#   block below and be sure that all your variables have the follow the
#   standard naming convention (D tag, lower case, underscores)

# - **Project Team Action:**

#   ```{r}
#   #cs_raw_file <- cs_raw_file %>% 
#     # Remove columns you don't need
#   #  select(-c(EXAMPLE_FIELD)) %>% 
#     # Rename columns to potentially use down the line
#   #  rename(D_example = EXAMPLE_FIELD)
#   ```

# 5.2 Export ----

# - **What:** The below exports all your check tables to an excel file in
#   the folder path you set above

# - **Check:** n/a

# - **Project Action:** Optional - change the file name

# %% 5.2 Export
# Create a new Excel workbook
wb <- createWorkbook()

# Add sheets to the workbook
addWorksheet(wb, "checks_table")
addWorksheet(wb, "col_info")
addWorksheet(wb, "cnt_periods_by_scl")
addWorksheet(wb, "cnt_term_by_scl")
addWorksheet(wb, "tchr_course_cnt")
addWorksheet(wb, "tchr_load")
addWorksheet(wb, "stu_course_cnt")
#addWorksheet(wb, "stud_subject_grade_cnt")

# Write each data frame to a different sheet
writeData(wb, sheet = "checks_table", checks_table)
writeData(wb, sheet = "col_info", col_info)
writeData(wb, sheet = "cnt_periods_by_scl", count_periods_by_school)
writeData(wb, sheet = "cnt_term_by_scl", count_term_by_school)
writeData(wb, sheet = "tchr_course_cnt", rbind(teacher_course_count_large,
                                               teacher_course_count_small))
writeData(wb, sheet = "tchr_load", teacher_load_large)
writeData(wb, sheet = "stu_course_cnt", student_course_count_small)
#writeData(wb, sheet = "stud_subject_grade_cnt", student_course_count)



# %% Save the checks workbook locally first
# Save the checks workbook locally first
temp_file <- tempfile(fileext = ".xlsx")
saveWorkbook(wb, temp_file, overwrite = TRUE)

file.exists(temp_file)

# Upload to SharePoint using existing drive object
internal_drive$upload_file(
  src = temp_file,
  dest = paste0(raw_data_folder_path, "/2. Processed Data/checks_workbook.xlsx")
)


# %% Save course schedule data for next script
# Save course schedule data for next script
ers_write_sharepoint(
  data = cs_raw_file,
  folder_path = raw_data_folder_path,
  file_name_with_extension = 
    "/2. Processed Data/01_course_data_post_validation.csv")



# Great job!!!
