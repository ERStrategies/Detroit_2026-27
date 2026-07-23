# 02-data-standardization
# Converted from R Markdown to a Positron-ready R script.
# Original narrative and instructions are retained as comments.


# ==============================================================================
# Key Links ----
# ==============================================================================

# -   [Data
#     Dictionary:](https://erstrategies1.sharepoint.com/:x:/s/orgfiles/EW4fU7_js69Hl9DzoYTuB7gBVUamadRtarFtyz6-5jampA?e=c5s2Ev)

# -   [Style
#     Guide:](https://app.tettra.co/teams/ersknowledge/pages/coding-style-guide)

# -   [Data Request and
#     Validation](https://erstrategies1.sharepoint.com/:x:/s/orgfiles/EcSQElbD3mlMsc_HzVYj_KUBJfCavz6HVlHWD7cN3PAAKw?e=CuNtD8)

# ==============================================================================
# Part I: Setup ----
# ==============================================================================

# This part of the script allows you to:

# -   load the functions that R will be using to work with your data
#     (Excel automatically loads all its functions - if, countif, index,
#     match, etc. - when you open it, but R requires manual loading of
#     functions if you want full access to what it offers)

# -   specify the Sharepoint folder where your raw data lives so R can
#     pull it in

# -   actually load in the raw data you'll be working with throughout the
#     rest of the script

# 1.1 Load Packages ----

# **What:** This code block loads in all relevant function packages
# (a.k.a. R functions you'll be needing in this script). Note that if you
# haven't installed the packages into your computer yet, R will print an
# error and you'll need to install them first by running
# install.packages(tidyverse) and install.packages(formattable), etc.

# **Check:** no error message coming out of code block

# **Project Team Action:** do not change (run as is)

# %% 1.1 Load Packages and Connect to SharePoint
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


# 1.2 Load Data ----

# -   **What:** Now that R knows how to access the folder path, it's time
#     to load in your raw data

# -   **Check:** The code block should print the names of all data files
#     you need.

# -   **Project Team Action:** Review the CS file columns and data to
#     ensure it loaded correctly

# **a. Course Schedule File**

# %% 1.2 Load Data
# Create folder path
raw_data_folder_path <- "District Partners/Detroit Public Schools/26-27 HS Redesign Implementation/1. Data & Analysis - Secure/Fall 2026 Course Schedule Analysis"

# Load processed data
cs_validated_file <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/2. Processed Data/01_course_data_post_validation.csv")


# ==============================================================================
# Part II: Processing ----
# ==============================================================================

# 2.1 Re-code Student Grade ----

# **What:** This set of code will return a list of grade values and allows
# you to recode your district's values into standard ERS values so that
# subsequent code that uses D_stu_grade will work across project teams.

# **Check:** Run the first code chunk and review the grade values so you
# know how to update the following code, then check the output in Code
# Chunk 2 to ensure that the values have been recoded.

# **Project Team Action:**

# 1.  *Code Chunk 1:* Review the grade values so you know which values to
#     include in the recode chunk (e.g., "Kindergarten", "KG", etc.)
# 2.  *Code Chunk 2:* Update the second code chunk with your projects
#     specific grade values as needed and review the list to ensure
#     everything updated properly

# *Code Chunk 1:*

# %% 2.1 Re-code Student Grade
# Converts D_stu_grade to character type
cs_validated_file <- cs_validated_file %>%
  mutate(D_stu_grade = as.character(D_stu_grade))

# Review D_stu_grade values so you know how to update below code
View(cs_validated_file %>%
  distinct(D_stu_grade) %>%
  arrange(D_stu_grade))


# *Code Chunk 2:*

# %% 2.1 Re-code Student Grade [2]
# Recode student grade values
cs_validated_file <- cs_validated_file %>%
  mutate(C_stu_grade_ers = case_when(
    D_stu_grade %in% c("Pre K", "Pre-Kindergarten") ~ "PreK",
    D_stu_grade %in% c("Kindergarten", "K") ~ "KG",
    D_stu_grade %in% paste0("0", 1:9) ~ D_stu_grade,
    D_stu_grade %in% paste0(1:9) ~ paste0("0", D_stu_grade),
    is.na(D_stu_grade) ~ "Uncoded",
    TRUE ~ D_stu_grade
  ))

# View the list of grades again to ensure there is an update
View(cs_validated_file %>%
  distinct(D_stu_grade,C_stu_grade_ers) %>%
  arrange(D_stu_grade))


# 2.2 Create School Exclude Flag ----

# **What:** There are two code chunks here:

# -   The first one creates a table that allows you to explore
#     grade/school configurations. Or how many students are enrolled at
#     each school per grade.

# -   The second code chunk creates a "exclude/include" flag based on the
#     list of schools below. The schools in that list will get a "Exclude"
#     flag. Schools that usually considered for exclusion are Alternative,
#     special ed or schools with specialized programming where the courses
#     and information may not be comparable or reliable. This should be
#     discussed with your project team ahead of time.

# -   **Check:** Review the list of schools by their flag

# -   **Project Team Action:**

#     1.  First, explore the different schools/grade configurations - this
#         might provide some insight into alt/special ed schools and which
#         ones to add to the exclude list

#     2.  Then update the second code chunk **schools_to_drop** section
#         one with the names of the schools in your district that you wish
#         to exclude - and do note that by the end of the code block that
#         data will be excluded, so only run that code block once you feel
#         confident about the schools you are excluding

# *Explore Schools Code Chunk*

# %% 2.2 Create School Exclude Flag
# Review count of students by school across grades
school_enroll_by_grade <- cs_validated_file %>%
  drop_na(D_stu_grade) %>%
  group_by(D_location_name, C_stu_grade_ers) %>%
  summarise(Unique_Count = n_distinct(D_stu_id)) %>%
  pivot_wider(names_from = C_stu_grade_ers,
              values_from = Unique_Count,
              values_fill = 0) %>%
  mutate(total_enrollment = rowSums(across(where(is.integer)), na.rm = TRUE),
  across(where(is.integer),
         ~ round(. / total_enrollment, digits=2),
         .names = "{.col}_pct"))

View(school_enroll_by_grade)


# *Create Exclude School Flag Column and Exclude Schools*

# %% 2.2 Create School Exclude Flag [2]
# Specify the schools you want to drop
schools_to_exclude <- c("SPED - Ivy School",
                     "Summer Programs-PPS",
                     "CTC Northeast",
                     "CTC Southeast",
                     "DART Programs",
                     "Depaul Center",
                     "NAYA Many Nations Academy",
                     "PPS Pioneer Programs",
                     "SPED - ESY",
                     	"Serendipity",
                     "Youth Builders",
                     "Youth Progress Association")

# Create a school exclude flag based on the above lists
school_enroll_by_grade <- school_enroll_by_grade  %>%
  mutate(C_school_exclude_flag = ifelse(D_location_name %in%
                                      schools_to_exclude,
                                      "Exclude",
                                      "Include"))

# Course schedule file
cs_validated_file <- cs_validated_file %>%
  mutate(C_school_exclude_flag = ifelse(D_location_name %in%
                                        schools_to_exclude,
                                      "Exclude",
                                      "Include"))

# Check - return list of schools by flag
cs_validated_file %>%
  distinct(D_location_name,
           C_school_exclude_flag) %>%
  arrange(C_school_exclude_flag)

# Removing excluded schools from data
cs_validated_file <- cs_validated_file %>%
  filter(C_school_exclude_flag == "Include")


# 2.3 Assign School Level ----

# a. Create school level ----

# -   **NOTE!** You do NOT need to run this block if you know that you
#     only have one school level in your data.

# -   **What:** This creates a school level flag (ES/SS) based on the
#     enrollments in each grade. This flag will be used to split the file
#     into ES/SS (two).

# -   **Check:** Review the list of schools and their flag to ensure the
#     code worked correctly - especially those schools that are coded as
#     "other"

# -   **Project Team Action:**

#     -   Update the grade values to ensure it's capturing all grades in
#         your districts

#     -   review the output for "other" as there is always edge cases.

#     -   update the code if necessary to account for those edge cases and
#         run again

# %% a. Create school level
#team - you might need to update the below values (`01_pct`,`02_pct`) with your values
#  school_enroll_by_grade <- school_enroll_by_grade %>%
#    mutate(ES_total= sum(`01_pct`,
#                         `02_pct`,
#                         `03_pct`,
#                         `04_pct`,
#                         `05_pct`)) %>%
#    mutate (MS_total= sum(`06_pct`,
#                          `07_pct`,
#                          `08_pct`)) %>%
#    mutate(HS_total=sum(`09_pct`,
#                        `10_pct`,
#                        `11_pct`,
#                        `12_pct`)) %>%
#    mutate(C_school_level_detail = case_when(
#      sum(ES_total) > 0 & MS_total == 0 & HS_total == 0 ~ "ES",
#      sum(ES_total) > 0 & MS_total >  0 & HS_total == 0 ~ "K8",
#      sum(ES_total) == 0 & MS_total > 0 & HS_total == 0 ~ "MS",
#      sum(ES_total) == 0 & MS_total > 0 & HS_total > 0 ~ "SS",
#      sum(ES_total) == 0 & MS_total == 0 & HS_total > 0 ~ "HS",
#      sum(ES_total) >0 & MS_total > 0 & HS_total > 0 ~ "K12",
#      TRUE ~ "Other")) %>%
#   mutate(C_school_level = case_when (
#     C_school_level_detail %in% c("ES","K8") ~ "ES",
#     C_school_level_detail %in% c("MS","HS","SS") ~ "SS",
#     TRUE ~ "Other"))

#view updates
#school_enroll_by_grade %>%
#  filter(C_school_exclude_flag == "include") %>%
#   select(-matches ("_pct"), -matches("exclude")) %>%
#  view()


# b. Merge School Level back to Course Schedule ----

# -   **What:** This brings over the school level flags to the course
#     schedule file. **NOTE!** You do NOT need to run this block if you
#     know that you only have one school level in your data. And if you
#     do, you should only run this once you are sure about your school
#     levels!

# -   **Check:** The number of rows in the CS file have remained the same.
#     (Take note of the \# of rows in your environment)

# -   **Project Team Action:** None

# %% b. Merge School Level back to Course Schedule
# Merge school level columns back to CS file
#cs_validated_file <- cs_validated_file %>%
#  left_join(select(school_enroll_by_grade, D_location_name, C_school_level, C_school_level_detail),
#            by = "D_location_name")


# 2.4 Create Student Snapshot Dates ----

# **What:**

# -   **NOTE!** You do **NOT** need to run this code block if received
#     course enrollment data from the district that is already a snapshot
#     in time.

# -   When looking at which students are enrolled in a course, best
#     practice is to limit your look to a specific date based on which
#     term the class fell on. The rationale for this is that students
#     might have joined and left the course throughout the duration of the
#     term, which makes enrollment in a single day a more meaningful
#     representation of a student's typical experience of the classroom
#     during the overall length of the term.

# -   To determine which date to use as snapshot for each term, look at
#     the distribution of start and end dates of students for that term
#     and choose a day around which enrollment seems relatively stable.
#     This will typically mean avoiding the first and last week of the
#     term (since that's when students most often join late or drop
#     early) - and you might want to experiment with a couple different
#     dates to see which one preserves the largest number of your student
#     rows.

# **Check:**

# -   Examine the output of the first two tables printed and identify a
#     date in each term for which the majority of students have entered
#     already but haven't left yet
# -   See the third output and assess if you've lost a percentage of
#     students larger than what you'd want - as a rule of thumb anything
#     above 20% excluded for quarter or semester terms is high, whereas
#     for full year you'll likely have a good amount of exclusions given
#     that lots of students might have left from one semester to another
#     (in addition to the natural starts and exits that happen within each
#     semester and quarter)

# **Project Team Action:** definitely update (see code block for
# instructions)


# %%
# Filter for just 2024 data
cs_validated_file <- cs_validated_file |>
  filter(D_year_id == 2024)

# %% 2.4 Create Student Snapshot Dates
# #check
cs_validated_file %>%
   group_by(D_term,
            D_stu_enter_date) %>%
   count() %>%
   arrange(desc(D_stu_enter_date))

# #check
cs_validated_file %>%
   group_by(D_term,
            D_stu_exit_date) %>%
   count() %>%
   arrange(desc(D_stu_exit_date))


# %%

# update term value based on course name
cs_validated_file <- cs_validated_file |>
  mutate(D_term = case_when(
    str_detect(D_course_name, "-\\s*A\\s*$") ~ "S1",
    str_detect(D_course_name, "-\\s*B\\s*$") ~ "S2",
    TRUE ~ "FY"
    )
  )

# creating snapshot variable
cs_validated_file <- cs_validated_file %>%
   mutate(C_stu_snapshot = case_when(
          D_term == "S1" &
          D_stu_enter_date <= as.Date("2023-10-25") &
          D_stu_exit_date >= as.Date("2023-10-25") ~ "Include",
          D_term == "S1" &
          D_stu_enter_date <= as.Date("2024-01-17") &
          D_stu_exit_date >= as.Date("2024-01-17") ~ "Include",
          D_term == "S2" &
          D_stu_enter_date <= as.Date("2024-03-28") &
          D_stu_exit_date >= as.Date("2024-03-28") ~ "Include",
          D_term == "S2" &
          D_stu_enter_date <= as.Date("2024-05-30") &
          D_stu_exit_date >= as.Date("2024-05-30") ~ "Include",
          TRUE ~ "Exclude"))

# check
cs_validated_file %>%
   group_by(D_term, C_stu_snapshot) %>%
   summarize(count_rows = length(D_term)) %>%
   group_by(D_term) %>%
   mutate(total_rows = sum(count_rows),
          percentage = round(count_rows / total_rows, 2))

# Check all records for one random student
student_check <- cs_validated_file |>
  filter(D_stu_id == sample(unique(D_stu_id), 1)) |>
  select(D_stu_id, D_employee_id, D_course_id, D_course_name, D_stu_enter_date, D_stu_exit_date, C_stu_snapshot) |>
  arrange(C_stu_snapshot, D_course_name)

# 2.5 Split into ES/SS ----

# -   **NOTE!** You do **NOT** need to run this block if you know that you
#     only have one school level in your data.

# -   **What:** This will create an ES/SS file. In general we don't look
#     at student time for students in elementary school since they are not
#     departmentalized. Therefore, it does not make sense to churn all
#     students through double bookings, calculating weights, etc so we
#     split the file into SS/ES and will perform a different set of
#     processing steps on ES. General Guidance:

#     -   Airing on the side of exclusivity, right now the rules are:
#         ES=ES, and MS,SS,HS=SS

#     -   Schools that contain both an ES grade and a SS grade (K8 or K12)
#         will remain in both files.

#     -   However, project teams can decide on a case by case basis how
#         they would like to handle these types of schools and update the
#         rules that makes sense for their project.

# -   **Check:** Review the list of schools that *remain* in the file.

# -   **Project Team Action:** Update the values in the filter part of the
#     function if necessary

# %% 2.5 Split into ES/SS
# Original R Markdown chunk option: eval=FALSE
# #create ES file
# cs_elementary_file <- cs_validated_file %>%
#   filter(C_school_level_detail %in% c("ES","K8","K12"))
#
# #create SS file
# cs_secondary_file <- cs_validated_file %>%
#   filter(C_school_level_detail %in% c("MS","SS","HS","K8","K12"))
#
# #review ES Schools List
# cs_elementary_file %>%
#   distinct(D_location_name, C_school_level_detail)
#
# #review SS Schools List
# cs_secondary_file %>%
#   distinct(D_location_name, C_school_level_detail)


# ==============================================================================
# Part III: Export ----
# ==============================================================================

# **What:** use this code block to export your data into SharePoint to
# load unto the next script

# **Check:** look at your SharePoint folder and make sure that the data
# loaded properly

# **Project Team Action:** maybe change (you can tweak the file name to
# your liking)

# %% Part III: Export
# Save course schedule data for next script
ers_write_sharepoint(
  data = cs_validated_file,
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/Processed Data/02_course_data_post_standardization.csv")
