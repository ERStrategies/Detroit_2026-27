# 07-db-unification
# Converted from R Markdown to a Positron-ready R script.
# Original narrative and instructions are retained as comments.


# Part I: Setup ----

# This part of the script allows you to:

# - load the functions that R will be using to work with your data (Excel automatically loads all its functions - if, countif, index, match, etc. - when you open it, but R requires manual loading of functions if you want full access to what it offers)

# - declare the folder path that will tell R where to find your data

# - load in the data you'll be working with throughout the rest of the script

# 1.1 Load Packages ----

# What: This code block loads in all relevant function packages (a.k.a. R functions you'll be needing in this script). Note that if you haven't installed the packages into your computer yet, R will print an error and you'll need to install them first by running install.packages(tidyverse) and install.packages(data.frame).

# Check: no error message coming out of code block

# Project Team Action: do not change (run as is)

# %% 1.1 Load Packages
# Remove everything from any previous sessions
rm(list = ls())

# This section sets up our file to correctly output the data for the school we are interested in; I am going to work on just testing for one school in order to make sure that the code is working correctly.

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


# 1.2 Load data ----

# What: Now that R knows how to access the folder path, it's time to input the name of the data files you want to load. Note that your files should be in CSV format and that your file names should always end with .csv

# Check:

# - The code block should print the name of the data file you're looking to load

# - Confirm that your data file did load and that the columns and row numbers look right (do this by accessing the data in the environment window in the top right of RStudio)

# Project Team Action: maybe change (see guidance in code block)

# %% 1.2 Load data
# Create folder path
raw_data_folder_path <- "District Partners/Detroit Public Schools/26-27 HS Redesign Implementation/1. Data & Analysis - Secure/Fall 2026 Course Schedule Analysis"

# Load double-bookings processed data
course_data_resolved <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/2. Processed Data/06_cs_data_post_db_resolutions.csv")

# Load unexploded data
course_data_unexploded <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/2. Processed Data/03_course_data_post_coding_exclusions.csv")

# Part II: Creating Unified File ----

# 2.1 Merge ----

# What:

# - The purpose of this code block is to append the resolved exploded course schedule data with the rows excluded the data in earlier scripts due to not having double bookings. In other words, the code block below allows you to keep your exploded rows exploded and brings in the unexploded version of the rows that were taken out before running the data through the resolutions script.

# - The way this code block achieves that purpose is by first identifying which rows from the exploded data have a match with the full unexploded data (based on the row id created before the explosions) and then appending the rows WITHOUT a match unto the exploded data set (since those are the rows that aren't represented in the exploded data set).

# - Note that this code block might need you to do some alignment of data type in order to allow seamless binding of rows - a placeholder line has been left for this purpose if needed.

# Check:

# - The first check shows you the rows in the unexploded data that had a matching row id in the exploded data, and you should see EXACTLY the same number of rows with "Match" as the number of rows in the course_data_resolved_before_explosion data frame (second check)
# - Third check shows the number of rows in the joined data frame - this number should EXACTLY match the fourth check, which shows the sum of the number of rows with NAs in the first check shown (which is the unexploded rows that don't match the exploded rows) and the number of rows in the exploded data set

# Project Team Action: Maybe change (see guidance in code block)

# %% 2.1 Merge
# STEP 1: Creating match variable with EXPLODED data

course_data_unexploded <- course_data_unexploded %>%
  mutate(H_matcher_unexploded = "Match")

# Project team action: changing data types - use if any issues pop up when trying the joins below
#course_data_exploded <- course_data_exploded %>% 
#  mutate(D_period = as.numeric(D_period),
#         C_stu_grade_ers = as.integer(C_stu_grade_ers))

course_data_resolved <- course_data_resolved %>% 
  mutate(D_stu_enter_date = as.character(D_stu_enter_date),
         D_stu_exit_date  = as.character(D_stu_exit_date))

course_data_unexploded <- course_data_unexploded %>%
  mutate(D_stu_enter_date = as.character(D_stu_enter_date),
         D_stu_exit_date  = as.character(D_stu_exit_date))

# STEP 2: Create a list of original row IDs that made it into the resolved file
# In other words: "Give me the list of original pre-explosion row IDs that are represented in the resolved/exploded data."
# Methods note: variables that got changed during flaggings and resolutions need to be taken out as well so they don't interfere with merge
course_data_resolved_before_explosion <- course_data_resolved %>% 
  select(H_row_id_before_explosions) %>% 
  unique() %>% 
  mutate(H_matcher_resolved = "Match")

# STEP 3: Mark which exploded rows are already represented in resolved data
    # "Match" = this original row already appears in the resolved/exploded data
    # NA = this original row did NOT go through the resolved/exploded data and you need to append back in
course_data_unexploded <- left_join(course_data_unexploded,
                  course_data_resolved_before_explosion,  
                  by = "H_row_id_before_explosions")

# STEP 4: Check how many exploded rows matched resolved rows
# You expect the "Match" count to equal the number of rows in the helper table from Step 2.
course_data_unexploded %>% 
  group_by(H_matcher_resolved) %>% 
  count()

# STEP 5: Check number of original rows represented in resolved file - this should match the "Match" count in the previous check
nrow(course_data_resolved_before_explosion)

# STEP 6: Create the unified file
course_data_unified <- course_data_unexploded %>% 
  filter(is.na(H_matcher_resolved)) %>% 
  rbind(course_data_resolved,
        fill = TRUE)

# STEP 7: Check final row count
nrow(course_data_unified)

# STEP 8: Manually calculate expected row count
# This should exactly match the row count from the previous check since every row in the resolved data should either be represented in the exploded data (and thus have a match) or not be represented (and thus not have a match and get appended in full)
nrow(course_data_unexploded[is.na(course_data_unexploded$H_matcher_resolved)]) +
nrow(course_data_resolved)


# 2.2 Unified Teacher Load Exclusion ----

# What: Teachers that got flagged as needing exclusion during teacher resolutions already have their double booked rows excluded. However, given that these teachers are taken out from analysis altogether, the rest of their rows (a.k.a. the ones that didn't get flagged as double booked in the first place) now need to get marked "Exclude" in the teacher load exclusion column. To do this, we flag as "Exclude" any row from a teacher whose max value in the C_teachers_to_exclude column was 1. We then combine this "Exclude" flag with the original teacher_load_metric column to arrive at a unified column.

# Check:

# - First check shows the number of values that the C_teachers_to_exclude variable takes for each teacher - the result should print 1 since we want each teacher to either have all their rows included or excluded for this variable

# - Second check shows the number of teachers that do and don't get excluded altogether - be sure that these numbers seem reasonable

# - Third and fourth check show the number of rows with "Exclude" in the teacher_load_exclude variable - the number of excludes should increase from the third to the fourth check since the fully excluded teachers have now seen some of their non-exploded rows marked as exclude too

# Project Team Action: Do not change (run as is)

# %% 2.2 Unified Teacher Load Exclusion
# Step 1: Turn missing teacher-exclusion flags into 0
course_data_unified <- course_data_unified %>% 
  mutate(H_teachers_to_exclude = ifelse(
    is.na(
      H_teachers_to_exclude),
      0,
      H_teachers_to_exclude))

# Step 2: Spread the teacher-exclusion flag to all rows for that teacher
course_data_unified <- course_data_unified %>% 
  group_by(D_employee_id) %>% 
  mutate(H_teachers_to_exclude = max(H_teachers_to_exclude))

# Step 3: Check that each teacher has only one exclusion value
    # Output 1: should be 1 since each teacher should either be fully included or fully excluded
course_data_unified %>% 
  group_by(D_employee_id) %>% 
  summarize(n = n_distinct(H_teachers_to_exclude)) %>% 
  pull(n) %>% 
  max()

# Step 4: Count teachers by exclusion flag
    # Output 2: number of teachers that get fully excluded vs. not - be sure these numbers look reasonable
course_data_unified %>% 
  group_by(D_employee_id, H_teachers_to_exclude) %>% 
  count() %>% 
  group_by(H_teachers_to_exclude) %>% 
  count()

# Step 5: Check the original shared exclusion field
    # Output 3: number of rows with "Exclude" in the original class size and teacher load exclusion variable
course_data_unified %>% 
  group_by(C_class_size_and_teacher_load_exclude) %>% 
  count()

# Step 6: Create separate final exclusion flags for teacher load and class size (since some rows are only excluded for class size, some only for teacher load, and some for both - we want to be able to differentiate these cases in analysis)
course_data_unified <- course_data_unified %>% 
  mutate(C_teacher_load_exclude = ifelse(C_class_size_and_teacher_load_exclude == "Exclude" |
                                         H_teachers_to_exclude == 1,
                                         "Exclude",
                                         "Include")) %>% 
  rename(C_class_size_exclude = C_class_size_and_teacher_load_exclude)

# Step 7: Check final teacher-load exclusion flag
    # Output 4: number of rows with "Exclude" in the final teacher load exclusion variable - this number should be HIGHER than the number of excludes in the original shared exclusion variable since now we're also excluding all rows from teachers that got flagged for exclusion
course_data_unified %>% 
  group_by(C_teacher_load_exclude) %>% 
  count()


# %% 2.2 Unified Teacher Load Exclusion (2)
# Check the exclusions for teacher load
excluded_teachers <- course_data_unified %>% 
  filter(C_teacher_load_exclude == "Exclude") %>% 
  group_by(C_course_name) |> 
  summarise(class_count = n_distinct(C_class_id),
            teacher_count = n_distinct(D_employee_id),
            student_count = n_distinct(D_stu_id))


# 2.3 Unified Version of Other Key Variables ----

# What: This code block creates the unified version of several key variables by combining values in the resolved data and the unexploded data. The main idea is: For rows that went through explosion/resolution, use the updated exploded/resolved values. For rows that did not, keep the original values.

# Check: Look at the outputs and make sure the values and distributions of the united variables is what you would have expected

# Project Team Action: Do not change (run as is)

# %% 2.3 Unified Version of Other Key Variables
# Update key variables by taking the resolved/exploded value when it exists and the original unexploded value when it doesn't
course_data_unified <- course_data_unified %>% 
  mutate(D_term = ifelse(is.na(C_term_exploded),
                         D_term,
                         C_term_exploded),
         D_rotation = ifelse(is.na(C_rotation_exploded),
                             D_rotation,
                             C_rotation_exploded),
         D_period = ifelse(is.na(C_period_exploded),
                           D_period,
                           C_period_exploded),
         M_class_weight = ifelse(is.na(M_class_weight_post_student),
                                 M_class_weight,
                                 M_class_weight_post_student),
         C_class_id = ifelse(is.na(C_class_id_post_teachers),
                             C_class_id,
                             C_class_id_post_teachers))

# Checks
course_data_unified %>% 
  group_by(D_term) %>% 
  count()

course_data_unified %>% 
  group_by(D_rotation) %>% 
  count()

course_data_unified %>% 
  group_by(D_period) %>% 
  count()

course_data_unified %>% 
  group_by(M_class_weight) %>% 
  count()

course_data_unified %>% 
  group_by(C_class_id) %>% 
  count()


# 2.4 Class weight checks ----

# What: This code block creates a table with the total class weight by student, period and class id. This allows you to check that the class weights got updated as expected in the unified file - you can check that the total class weight by student, period and class id is 1 for non-double booked rows and greater than 1 for double booked rows, and that the class weights that got updated in the resolved/exploded data are reflected in the unified data.

# Check: Look at the distribution of class weights by term and make sure they look right given what you know about the data and the resolutions that got made.

# 2.4a Check weights by student ----
# %% 2.4a Check weights by student
# Create total class weight by student, period and class id
cw_by_stu_classid <- course_data_unified %>% 
  filter(C_teacher_load_exclude != "Exclude") %>%
  group_by(D_location_name,
           D_stu_id,
           D_term,
           D_rotation,
           D_period,
           C_class_id) %>% 
  summarise(class_weight_by_classid = sum(M_class_weight))

# Check
cw_by_stu_classid %>% 
  group_by(
    D_term,
    class_weight_by_classid) %>% 
  summarise(n())

# 2.4b Check weights by teacher ----
# %% 2.4b Check weights by teacher
# Create one class-weight row per teacher, TPR, and class id
cw_by_teacher_classid <- course_data_unified %>% 
  filter(
    C_teacher_load_exclude != "Exclude",
    C_course_time_exclude != "Exclude") %>%
  group_by(D_location_name,
           D_employee_id,
           D_term,
           D_rotation,
           D_period,
           C_class_id) %>% 
  summarise(
    class_weight_by_classid = max(M_class_weight, na.rm = TRUE),
    .groups = "drop"
  )

# Check distribution of class weights by teacher/class/TPR
cw_by_teacher_classid %>% 
  group_by(class_weight_by_classid) %>% 
  summarise(n = n(), .groups = "drop")

# Optional: sum class weights by teacher/TPR to check if any teachers still exceed expected load
cw_by_teacher_tpr <- cw_by_teacher_classid %>% 
  group_by(D_location_name,
           D_employee_id,
           D_term,
           D_rotation,
           D_period) %>% 
  summarise(
    total_class_weight_by_teacher_tpr = sum(class_weight_by_classid, na.rm = TRUE),
    .groups = "drop"
  )

# Check distribution of total teacher weight by TPR
cw_by_teacher_tpr %>% 
  group_by(total_class_weight_by_teacher_tpr) %>% 
  summarise(n = n(), .groups = "drop")


# Part III: Exports ----

# 3.1 Trim and export data ----

# What: Limiting course_data_unified only to variables of interest and then exporting to SharePoint

# Check: look at data in environment before exporting it to be sure that you feel good about the set of variables you're passing on to the next script - and once you export it check your SharePoint folder to ensure it got loaded correctly

# Project Team Action: likely change (if you have any additional variables that you want to keep at hand, add them to this code block)

# %% 3.1 Trim and export data
course_data_unified <- course_data_unified %>% 
  mutate(H_exploded_row = ifelse(is.na(C_term_exploded),
                                 "Not Exploded",
                                 "Exploded")) %>%
  select(-c(H_row_id_before_explosions,
            H_row_id_after_explosions,
            H_new_record_id,
            H_matcher_unexploded,
            H_matcher_resolved,
            M_class_weight_exploded,
            M_class_weight_post_student,
            H_teachers_to_exclude,
            C_class_id_post_teachers))


# Export to SharePoint

# %% 3.1 Trim and export data (2)
# Save course schedule data for next script
ers_write_sharepoint(
  data = course_data_unified,
  folder_path = raw_data_folder_path,
  file_name_with_extension = 
    "/2. Processed Data/07_cs_data_post_unification.csv")
