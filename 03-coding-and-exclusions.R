# 03-coding-and-exclusions
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

# %% 1.1 Load Packages
# Remove everything from any previous sessions
rm(list = ls())

# This section sets up our file to correctly output the data for the school we are interested in; I am going to work on just testing for one school in order to make sure that the code is working correctly.

# Load necessary libraries
library(tidyverse)
library(ggplot2)
library(janitor)

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
cs_data <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/2. Processed Data/02_course_data_post_standardization.csv")


# ==============================================================================
# Part II: Create Coding and Weight Tables ----
# ==============================================================================

# 2.1 Coding Tables ----

# **What:** This set of code will create an aggregated file to enable
# coding (e.g, one row per course name, C_subject, etc. and any other
# useful columns from district to help code C_subject, C_rigor, etc.)

# **Check:** Review the table and modify as needed

# **Project Team Action:** Review the columns included in the group_by
# statement and add any others that will help enable coding or remove the
# ones that will not.

# %% 2.1 Coding Tables
# Step 1: create coding table with course coding data provided by the district (i.e., "D_" columns)
cs_coding_data <- cs_data %>%
  group_by(D_course_name) %>%
#           D_course_subject,
#           D_course_subject_desc,
#           D_course_ell_flag,
#           D_course_swd_flag,
#           D_homeroom) %>%
  count()

# Step 2: Add columns you need to code (i.e., "C_" columns) and fill with "UNCODED" for now
cs_coding_data <- cs_coding_data |>
  mutate(C_course_time = "UNCODED",
         C_course_subject_area = "UNCODED",
         C_course_subject = "UNCODED",
         C_course_credit_type = "UNCODED",
         C_course_name = "UNCODED",
         C_course_rigor = "UNCODED",
         C_course_rigor_detail = "UNCODED",
         C_course_intervention = "UNCODED",
         C_course_time_of_day = "UNCODED",
         C_course_format = "UNCODED",
         C_course_pathway = "UNCODED")

# 2.2 Weight Tables ----

# **What:** This will create 3 tables, D_term, rotation and period so that
# you can code your weights in excel

# **Check:** Review the tables and modify as needed

# **Project Team Action:** Review the columns included in the group_by
# statement and add any others that will help enable weight coding or
# remove the ones that will not.

# a. Term Weight ----

# %% a. Term Weight
# Create term table
term_weight <- cs_data %>%
  group_by(D_term) %>%
  count()

# Add columns to code
term_weight <- term_weight %>%
  mutate(C_term_weight = 0) |>
  select(-n)

# b. Rotation Weight ----

# %% b. Rotation Weight
# Create rotation table
rotation_weight <- cs_data %>%
  group_by(D_rotation) %>%
  count()

# Add column to code
rotation_weight <- rotation_weight %>%
  mutate(C_rotation_weight = 0) |>
  select(-n)

# c. Period Weight ----

# %% c. Period Weight
# Create period weight table
period_weight <- cs_data %>%
  group_by(D_period) |>
  count()

# Add column to code
period_weight <- period_weight %>%
  mutate(C_period_weight = 0) |>
  select(-n)

# %% e. Check Periods
# Check to see which courses are offered during each period
period_check <- cs_data %>%
  group_by(D_location_name,
           D_period,
           D_course_name) %>%
  summarise(stu_count = n_distinct(D_stu_id))

# Check to see how many unique courses are offered during each period
period_check %>%
  group_by(D_period) %>%
  summarise(course_count = n_distinct(D_course_name))

# Check which courses are offered during periods with very few unique courses offered. These are likely period that should have a different weight (e.g., Homeroom or Advisory) or courses that are offered after school.
period_check_courses <- period_check %>%
  group_by(D_location_name, D_period) %>%
  mutate(course_count = n_distinct(D_course_name)) %>%
  filter(course_count < 15)

# Check expressions
expression_check <- cs_data %>%
  group_by(D_location_name,
           D_expression,
           D_period,
           D_rotation,
           D_course_name) %>%
  summarise(stu_count = n_distinct(D_stu_id))

expression_check %>%
  group_by(D_expression) %>%
  summarise(course_count = n_distinct(D_course_name))

expression_check_courses <- expression_check %>%
  group_by(D_expression) |>
  summarise(course_count = n_distinct(D_course_name))

# 2.2 Export Tables ----

# **What:** This will export your tables to the folder path you set in
# step 1.2

# **Check:** Make sure the file is in the folder you specified! Note that
# it sometimes takes a few minutes to show up on SharePoint.

# **Project Team Action:** None

# %% 2.2 Export Tables
# Export to excel
# UNHASH THIS TO EXPORT YOUR CODING AND WEIGHT TABLES TO SHAREPOINT!

# ers_write_sharepoint(
#   data = cs_coding_data,
#   folder_path = raw_data_folder_path,
#   file_name_with_extension =
#     "/1. Coding/course_coding_data.csv")

# ers_write_sharepoint(
#   data = period_weight,
#   folder_path = raw_data_folder_path,
#   file_name_with_extension =
#     "/1. Coding/period_weight.csv")

# ers_write_sharepoint(
#   data = rotation_weight,
#   folder_path = raw_data_folder_path,
#   file_name_with_extension =
#     "/1. Coding/rotation_weight.csv")

# ers_write_sharepoint(
#   data = term_weight,
#   folder_path = raw_data_folder_path,
#   file_name_with_extension =
#     "/1. Coding/term_weight.csv")

# \*\***PAUSE HERE**. GO TO YOUR EXCEL FILES AND START CODING!! RESUME
# ONCE YOU WANT TO MERGE CODING BACK ONTO MAIN CS FILE

# ------------------------------------------------------------------------------

# ==============================================================================
# Part III: Merge Coding and Weights ----
# ==============================================================================

# 3.1 Import Coding and Weights ----

# -   **What:** This imports your coded files. Once you update your coding files in Excel, make sure to **save them with the names and .xlsx file types listed below** and then run this code block to pull the coded fields into R.

# -   **Check:** Scan the environment to ensure your coded files have
#     returned

# -   **Project Team Action:** None

# %% 3.1 Import Coding and Weights
# Load back in course coding file
cs_course_coding <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/1. Coding/course_coding_2024.xlsx")

term_weight_coded <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/1. Coding/term_weight_coded.xlsx")

# %% 3.1 Import Coding and Weights [2]
# If your data already has PERIOD and ROTATION values separated
period_weight_coded <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/1. Coding/period_weight_coded.xlsx")

rotation_weight_coded <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/1. Coding/rotation_weight_coded.xlsx")

# 3.2 Merge Tables ----

# -   **What:** This merges your coded and weight tables and brings in
#     your new fields

# -   **Check:** Scan the environment to ensure your coded files have
#     returned

# -   **Project Team Action:** None

# a. Coded Table ----

# %% a. Coded Table
# Check before joining - should be ZERO rows
View(anti_join(
  cs_course_coding,
  cs_data,
  by = "D_course_name")
)

# %% a. Coded Table [2]
# Check records where students are missing demographic info
# stu_dem_missing <- cs_data %>%
#   filter(if_any(c(
#     D_stu_grade,
#     D_stu_swd_flag,
#     D_stu_ell_flag,
#     D_stu_poverty_flag),
#     is.na))

# # Check
# # Springfield - 10 students total are missing demographic data
# View(stu_dem_missing %>%
#   group_by(D_location_name, D_stu_id) %>%
#   summarise(record_count = n()))

# # Check number of rows
# print(paste("Number of rows before:", nrow(cs_data)))

# # Filter out these records
# cs_data <- cs_data %>%
#   filter(
#     if_all(
#       c(D_stu_grade, D_stu_swd_flag, D_stu_ell_flag, D_stu_poverty_flag),
#       ~ !is.na(.)))

# # Check number of rows
# print(paste("Number of rows after:", nrow(cs_data)))

# %% a. Coded Table [3]
# Bring in course coding fields to CS data
cs_course_coding <- cs_course_coding %>%
  mutate(C_course_pathway = coalesce(C_course_pathway, "NA"))

join_fields <- "D_course_id"

cs_data <- cs_data |>
  select(-any_of(setdiff(names(cs_course_coding), join_fields))) |>
  left_join(cs_course_coding, by = join_fields)

# View records where course coding is missing
course_coding_missing <- cs_data %>%
  filter(if_any(all_of(names(cs_course_coding)), is.na))
View(course_coding_missing)

# %%
# Filter out any records missing D_course_name
cs_data <- cs_data |>
  filter(!is.na(D_course_name))

# Recheck missing data
course_coding_missing <- cs_data %>%
  filter(if_any(all_of(names(cs_course_coding)), is.na))
View(course_coding_missing)


# b. Weight Tables ----

# %% b. Weight Tables
# Bring in Term Weight
cs_data <-left_join(cs_data,
                    term_weight_coded,
                    by = "D_term")

# %% b. Weight Tables [2]
# OPTION 1
# Bring in Rotation Weights
cs_data <-left_join(cs_data,
                    rotation_weight_coded,
                    by = "D_rotation")

# Bring in Period Weights
cs_data <-left_join(cs_data,
                    period_weight_coded,
                    by = "D_period")

# %% b. Weight Tables [3]
# Bring in Expression Weights
# cs_data <-left_join(cs_data,
#                     expression_weight_coded,
#                     by= "D_expression")

# 3.3 Calculate the Class Weight ----

# -   **What: calculates** the class weight now that you have the coded
#     weights

# -   **Check:** ensure the weights seems right

# -   **Project Team Action:** None

# %% 3.3 Calculate the Class Weight
# Calculate class weight
cs_data <- cs_data %>%
  mutate(M_class_weight = C_term_weight *
                          C_rotation_weight *
                          C_period_weight)

# %% 3.3 Calculate the Class Weight [2]
# Calculate expression weight
# cs_data <- cs_data %>%
#   mutate(M_expression_weight = C_term_weight * C_expression_weight)

# %% 3.3 Calculate the Class Weight [3]
# Check summary table of distribution of class weight and expression weight
cs_data |>
  group_by(C_stu_snapshot, M_class_weight) |>
  summarise(count = n())

# cs_data |>
#   group_by(M_expression_weight) |>
#   summarise(count = n())

# Check included with weight 0 - Detroit: all virtual Edgenuity afterschool
check_weight <- cs_data |>
  filter(
    C_stu_snapshot == "Include",
    M_class_weight == 0
  )

# ==============================================================================
# Part IV: Calculate Exclusions ----
# ==============================================================================

# -   **What:** This will calculate our standard exclusions based on a set
#     of standard rules (e.g., if C_time = "After School" then
#     C_time_exclude = TRUE)

# -   **Check:** To ensure the exclusions are accurate

# -   **Project Team Action:** Update based on your specific project
#     context if necessary

# %% Part IV: Calculate Exclusions
# Exclusions for time analyses
cs_data <- cs_data %>%
  mutate(
    C_course_time_exclude = case_when(
      C_time_metric == FALSE ~ "Exclude",
      C_course_subject_area == "Untracked" ~ "Exclude",
      C_course_time_of_day != "During school" ~ "Exclude",
      C_course_subject == "Internship"~ "Exclude",
      C_course_format == "Virtual" ~ "Exclude",
      TRUE ~ "Include"))

# Exclusions for class size and teacher load exclusions
cs_data <- cs_data %>%
  mutate(
    C_class_size_and_teacher_load_exclude = case_when(
      C_class_size_metric == FALSE ~ "Exclude",
      C_teacher_load_metric == FALSE ~ "Exclude",
      C_course_subject_area == "Un-prepped" ~ "Exclude",
      C_course_subject %in% c(
        "Maintenance",
        "Internship",
        "Teaching Assistant",
        "Credit Recovery") ~ "Exclude",
      C_course_format == "Virtual" ~ "Exclude",
      TRUE ~ "Include"))

# =====================================================================
#  Course-coding QA check
#  - Distributions of every coded field (distinct-course grain)
#  - Cross-field validation of the coding rules (locks) we built
#  Output: console tables (dplyr + janitor)
#  Assumes `cs_data` is already loaded and coded.
# =====================================================================

# %%
# ---- Class-weight column (section-level; exists after weighting) ---
weight_col <- "M_class_weight"
 
# Make the TIME exclusion weight-aware: a zero-weight class meets no real
# instructional time, so it should be excluded from the time metric.
# (This adds `class_weight == 0` to your original rule; class size / teacher
# load are left as-is, since that's a separate metric.)
if (!is.na(weight_col)) {
  cs_data <- cs_data %>%
    mutate(C_course_time_exclude = case_when(
      .data[[weight_col]] == 0                 ~ "Exclude",   # NEW: zero class weight
      C_course_subject_area == "Untracked"     ~ "Exclude",
      C_course_time_of_day != "During school"  ~ "Exclude",
      C_course_subject == "Internship"         ~ "Exclude",
      C_course_format == "Virtual"             ~ "Exclude",
      TRUE                                     ~ "Include"))
}
 
# ---- 0. Distinct-course grain --------------------------------------
# One row per course; dedupe on D_course_id (falls back to C_course_name).
key <- if ("D_course_id" %in% names(cs_data)) "D_course_id" else "C_course_name"
courses <- cs_data %>% distinct(.data[[key]], .keep_all = TRUE)
 
# Normalize the three metric flags to real logicals (they may read in as
# "TRUE"/"True"/"False" strings from Excel).
truthy <- function(x) as.character(x) %in% c("TRUE", "True", "true", "1")
metric_cols <- c("C_teacher_load_metric", "C_class_size_metric", "C_time_metric")
for (m in metric_cols) if (m %in% names(courses)) courses[[paste0(m, "_lgl")]] <- truthy(courses[[m]])
 
cat("========================================================\n")
cat("QA CHECK  |  distinct courses:", nrow(courses), "\n")
cat("========================================================\n")
 
# ---- 1. FREQUENCY DISTRIBUTIONS ------------------------------------
dist_fields <- c(
  "C_course_time", "C_course_subject_area", "C_course_subject",
  "C_course_credit_type", "C_course_rigor", "C_course_rigor_detail",
  "C_course_intervention", "C_course_time_of_day", "C_course_format",
  "C_teacher_load_metric", "C_class_size_metric", "C_time_metric",
  "D_course_ell_flag", "D_course_swd_flag",
  "C_course_time_exclude", "C_class_size_and_teacher_load_exclude"
)
 
cat("\n\n########## 1. DISTRIBUTIONS ##########\n")
for (f in dist_fields) {
  cat("\n----- ", f, " -----\n", sep = "")
  if (f %in% names(courses)) {
    print(courses %>%
            tabyl(!!sym(f)) %>%
            adorn_pct_formatting(digits = 1))
  } else {
    cat("  [column not present]\n")
  }
}
 
# Top pathways (Vocational/Career only)
if ("C_course_pathway" %in% names(courses)) {
  cat("\n----- C_course_pathway (top 25, non-blank) -----\n")
  print(courses %>%
          filter(!is.na(C_course_pathway), C_course_pathway != "") %>%
          count(C_course_pathway, sort = TRUE) %>%
          head(25))
}
 
# ---- 2. CROSS-FIELD VALIDATION -------------------------------------
# Each check filters the VIOLATING rows; 0 rows = PASS.
cat("\n\n########## 2. RULE VALIDATION (0 rows = PASS) ##########\n")
 
show_cols <- intersect(
  c("C_course_name", "C_course_subject_area", "C_course_subject",
    "C_course_credit_type", "C_course_rigor", "C_course_rigor_detail",
    "C_course_intervention", "C_course_format",
    "C_teacher_load_metric", "C_class_size_metric", "C_time_metric",
    "C_course_pathway", "D_course_swd_flag", "D_course_ell_flag"),
  names(courses))
 
check <- function(label, bad, cols = show_cols, n_show = 8) {
  n <- nrow(bad)
  cat(sprintf("\n[%s] %s  (%d)\n",
              if (n == 0) "PASS" else "REVIEW", label, n))
  if (n > 0) print(head(bad %>% select(any_of(cols)), n_show))
  invisible(n)
}
 
# 2.1 Intervention must be Core
check("Intervention = 1 but Subject Area != Core",
      courses %>% filter(C_course_intervention == 1, C_course_subject_area != "Core"))
 
# 2.2 Intervention -> credit_type Support & Enrichment
check("Intervention = 1 but credit_type != Support & Enrichment",
      courses %>% filter(C_course_intervention == 1,
                         C_course_credit_type != "Support & Enrichment"))
 
# 2.3 Credit Recovery format -> rigor & rigor_detail Below Standard
check("Format = Credit Recovery but rigor != Below Standard",
      courses %>% filter(C_course_format == "Credit Recovery",
                         C_course_rigor != "Below Standard" |
                         C_course_rigor_detail != "Below Standard"))
 
# 2.4 rigor <-> rigor_detail agreement
above <- c("Honors", "AP", "IB", "Dual Enrollment", "College Prep")
check("rigor_detail is Above-level but rigor != Above Standard",
      courses %>% filter(C_course_rigor_detail %in% above,
                         C_course_rigor != "Above Standard"))
check("rigor_detail = Below Standard but rigor != Below Standard",
      courses %>% filter(C_course_rigor_detail == "Below Standard",
                         C_course_rigor != "Below Standard"))
check("rigor_detail = Standard but rigor != Standard",
      courses %>% filter(C_course_rigor_detail == "Standard",
                         C_course_rigor != "Standard"))
 
# 2.5 Non-Standard format -> all three metrics FALSE
if (all(paste0(metric_cols, "_lgl") %in% names(courses))) {
  check("Format != Standard but a metric flag is TRUE",
        courses %>% filter(C_course_format != "Standard",
                           C_teacher_load_metric_lgl | C_class_size_metric_lgl | C_time_metric_lgl))
 
  # 2.6 Un-prepped -> teacher_load FALSE
  check("Subject Area = Un-prepped but teacher_load = TRUE",
        courses %>% filter(C_course_subject_area == "Un-prepped",
                           C_teacher_load_metric_lgl))
 
  # 2.7 Credit Recovery / Unknown subject -> metrics FALSE
  check("Subject in {Credit Recovery, Unknown} but a metric flag is TRUE",
        courses %>% filter(C_course_subject %in% c("Credit Recovery", "Unknown"),
                           C_teacher_load_metric_lgl | C_class_size_metric_lgl | C_time_metric_lgl))
 
  # 2.8 Untracked / Excluded time -> metrics FALSE
  check("Time in {Untracked, Excluded} but a metric flag is TRUE",
        courses %>% filter(C_course_time %in% c("Untracked", "Excluded"),
                           C_teacher_load_metric_lgl | C_class_size_metric_lgl | C_time_metric_lgl))
}
 
# 2.9 Pathway only for Vocational/Career
check("Pathway populated but Subject != Vocational/Career",
      courses %>% filter(!is.na(C_course_pathway), C_course_pathway != "", C_course_pathway != "NA",
                         C_course_subject != "Vocational/Career"))
 
# 2.10 SWD flag should sit on special-ed (Essentials/Adapted) courses
check("SWD flag = 1 but name has no Essentials/Adapted marker (informational)",
      courses %>% filter(D_course_swd_flag == 1,
                         !grepl("Essentials|Adapted|SPED|Resource", C_course_name, ignore.case = TRUE)))
 
# 2.11 Time hierarchy: Untracked/Excluded time should carry matching area
check("Time = Untracked but Subject Area != Untracked",
      courses %>% filter(C_course_time == "Untracked",
                         C_course_subject_area != "Untracked"))
 
# ---- 2b. CLASS-WEIGHT CHECKS (section grain, on full cs_data) ------
# Class weight is section-level, so these run on every row (NOT deduped).
cat("\n\n########## 2b. CLASS WEIGHT (section grain) ##########\n")
if (is.na(weight_col)) {
  cat("\n[no class-weight column detected -- set `weight_col` near the top to enable]\n")
} else {
  cat("\nWeight column:", weight_col, " | sections:", nrow(cs_data), "\n")
 
  # Full value distribution (0, 0.2, 0.5, 1, 2, ...)
  cat("\n----- ", weight_col, " value distribution -----\n", sep = "")
  print(cs_data %>% tabyl(!!sym(weight_col)) %>% adorn_pct_formatting(digits = 1))
 
  # Zero vs non-zero
  cat("\n----- zero vs non-zero -----\n")
  print(cs_data %>%
          mutate(.wt0 = ifelse(.data[[weight_col]] == 0, "weight = 0", "weight > 0")) %>%
          tabyl(.wt0) %>% adorn_pct_formatting(digits = 1))
 
  # VALIDATION: every zero-weight section must be excluded from the time metric
  bad_wt <- cs_data %>% filter(.data[[weight_col]] == 0,
                               C_course_time_exclude != "Exclude")
  cat(sprintf("\n[%s] class_weight == 0 but NOT excluded from time  (%d)\n",
              if (nrow(bad_wt) == 0) "PASS" else "REVIEW", nrow(bad_wt)))
  if (nrow(bad_wt) > 0)
    print(head(bad_wt %>% select(any_of(c("C_course_name", weight_col,
          "C_course_time_exclude", "C_course_subject_area", "C_course_time_of_day"))), 10))
 
  # Resulting time-exclude split (section grain, weight-aware)
  cat("\n----- C_course_time_exclude (section grain, weight-aware) -----\n")
  print(cs_data %>% tabyl(C_course_time_exclude) %>% adorn_pct_formatting(digits = 1))
 
  # Cross-tab: which reason is driving exclusion among zero-weight sections
  cat("\n----- zero-weight sections by subject area (why they're 0) -----\n")
  print(cs_data %>% filter(.data[[weight_col]] == 0) %>%
          count(C_course_subject_area, C_course_time_of_day, sort = TRUE) %>% head(15))
}
 
# ---- 2c. METRIC EXCLUSIONS REVIEW (distinct courses) ---------------
# Every course with ANY of the three metric flags FALSE, shown with all
# three flags + the likely reason, so you can confirm the right things drop.
cat("\n\n########## 2c. METRIC EXCLUSIONS (courses with any flag FALSE) ##########\n")
mlgl <- paste0(metric_cols, "_lgl")
if (all(mlgl %in% names(courses))) {
  excl <- courses %>%
    filter(!(C_teacher_load_metric_lgl & C_class_size_metric_lgl & C_time_metric_lgl)) %>%
    mutate(
      why = case_when(
        C_course_time %in% c("Untracked", "Excluded") ~ "Untracked/Excluded time",
        C_course_subject == "Credit Recovery"         ~ "Credit Recovery subject",
        C_course_format  == "Credit Recovery"         ~ "Credit Recovery format",
        C_course_format  == "Higher Ed"               ~ "Higher Ed / Dual Enrollment",
        C_course_format  == "Virtual"                 ~ "Virtual format",
        C_course_format  != "Standard"                ~ "Non-standard format",
        C_course_subject == "Unknown"                 ~ "Unknown subject",
        C_course_subject_area == "Un-prepped"         ~ "Un-prepped (teacher load only)",
        TRUE                                          ~ "other / manual review"),
      TL = C_teacher_load_metric_lgl,
      CS = C_class_size_metric_lgl,
      Time = C_time_metric_lgl)
 
  cat("\nCourses with >=1 metric FALSE:", nrow(excl), "of", nrow(courses), "distinct courses\n")
 
  # by reason
  cat("\n----- exclusions by reason -----\n")
  print(excl %>% count(why, sort = TRUE))
 
  # which flags are off (reveals partial cases like Un-prepped -> TL only)
  cat("\n----- flag pattern (TL / CS / Time) -----\n")
  print(excl %>%
          mutate(pattern = paste(ifelse(TL, "T", "F"), ifelse(CS, "T", "F"),
                                 ifelse(Time, "T", "F"), sep = "/")) %>%
          tabyl(pattern) %>% adorn_pct_formatting(digits = 1))
 
  # combined table (unique course x flags x exclusion fields x reason)
  cat("\n----- combined table (unique course / flags / exclusion fields / reason) -----\n")
  excl %>%
    mutate(Time_excl = if ("C_course_time_exclude" %in% names(.)) C_course_time_exclude else NA_character_,
           TLCS_excl = if ("C_class_size_and_teacher_load_exclude" %in% names(.)) C_class_size_and_teacher_load_exclude else NA_character_) %>%
    distinct(C_course_name, C_course_subject_area, C_course_subject,
             C_course_format, TL, CS, Time, Time_excl, TLCS_excl, why) %>%
    arrange(why, C_course_subject_area, C_course_name) %>%
    as_tibble() %>% print(n = Inf)
} else {
  cat("\n[metric flag columns not found]\n")
}
 
# ---- 2d. EXCLUSION FIELDS vs METRIC FLAGS (section grain) ----------
# Confirms the two exclusion fields agree with the coded metric flags.
# Runs on cs_data (section grain) since the exclusions are section-level.
# A REVIEW here = a course the coding flags to drop is NOT being excluded
# by the exclusion field (e.g., the "Virtual" vs "Credit Recovery" /
# "Higher Ed" format-value gap).
cat("\n\n########## 2d. EXCLUSION FIELDS vs METRIC FLAGS (section grain) ##########\n")
for (m in metric_cols) if (m %in% names(cs_data)) cs_data[[paste0(m, "_lgl")]] <- truthy(cs_data[[m]])
 
# --- Time metric vs C_course_time_exclude ---
if (all(c("C_course_time_exclude", "C_time_metric_lgl") %in% names(cs_data))) {
  cat("\n----- C_time_metric  x  C_course_time_exclude -----\n")
  print(cs_data %>% tabyl(C_time_metric_lgl, C_course_time_exclude))
  miss_time <- cs_data %>% filter(!C_time_metric_lgl, C_course_time_exclude != "Exclude")
  cat(sprintf("\n[%s] C_time_metric FALSE but C_course_time_exclude != Exclude  (%d)\n",
              if (nrow(miss_time) == 0) "PASS" else "REVIEW", nrow(miss_time)))
  if (nrow(miss_time) > 0)
    print(miss_time %>% count(C_course_format, C_course_subject, sort = TRUE) %>% head(15))
}
 
# --- Class size / teacher load metrics vs C_class_size_and_teacher_load_exclude ---
if (all(c("C_class_size_and_teacher_load_exclude", "C_class_size_metric_lgl",
          "C_teacher_load_metric_lgl") %in% names(cs_data))) {
  cat("\n----- (class size OR teacher load FALSE)  x  C_class_size_and_teacher_load_exclude -----\n")
  tmp <- cs_data %>% mutate(cs_tl_false = (!C_class_size_metric_lgl) | (!C_teacher_load_metric_lgl))
  print(tmp %>% tabyl(cs_tl_false, C_class_size_and_teacher_load_exclude))
  miss_cstl <- tmp %>% filter(cs_tl_false, C_class_size_and_teacher_load_exclude != "Exclude")
  cat(sprintf("\n[%s] class size / teacher load FALSE but exclude field != Exclude  (%d)\n",
              if (nrow(miss_cstl) == 0) "PASS" else "REVIEW", nrow(miss_cstl)))
  if (nrow(miss_cstl) > 0)
    print(miss_cstl %>% count(C_course_format, C_course_subject, sort = TRUE) %>% head(15))
}
 
# ---- 3. ELL / SWD cross-tabs (informational) -----------------------
cat("\n\n########## 3. FLAG CROSS-TABS ##########\n")
if (all(c("D_course_ell_flag", "C_course_subject") %in% names(courses))) {
  cat("\n----- ELL flag x Subject -----\n")
  print(courses %>% filter(D_course_ell_flag == 1) %>% count(C_course_subject, sort = TRUE))
}
if (all(c("D_course_swd_flag", "C_course_subject_area") %in% names(courses))) {
  cat("\n----- SWD flag x Subject Area -----\n")
  print(courses %>% filter(D_course_swd_flag == 1) %>% count(C_course_subject_area, sort = TRUE))
}
 
cat("\n========================================================\n")
cat("Done. Any check labeled REVIEW lists the offending courses.\n")
cat("========================================================\n")
 
# ==============================================================================
# Part V: Create class IDs ----
# ==============================================================================

# **What:** See Tettra for full guidance on why and how to create a class
# id, but in brief:

# -   You create it to to have an identifier that marks which rows belong
#     to the same class

# -   The way to create it is to concatenate values from various different
#     columns such that the class id is unique enough to not include more
#     than one class within it BUT still generic enough such that you're
#     not splitting up a class into multiple IDs unnecessarily

# **Check:**

# -   Do a group_by of the class id and include in the group_by variables
#     that you think might be needed in the class id but aren't yet
#     included in it (do a single group_by for each variable you with to
#     test) - if a single class id takes multiple values for your grouped
#     variable and this suggests that the class should be split, then your
#     class id should likely include that variable in the first place.

# -   Note that this process is iterative - and when in doubt, lean on the
#     side of making the class id more unique since courses with teachers
#     and students enrolled at the same time can be resolved during the
#     double bookings process (whereas if you make the class id less
#     unique the double booking doesn't get marked and you've artificially
#     combined the classes together even in some cases where this might
#     not be desired).

# **Project Team Action:** definitely change (see instructions in code
# block)

# %% Part V: Create class IDs
# Create your class id
cs_data <- cs_data %>%
  mutate(C_class_id = paste(D_location_id,
                            D_term,
                            D_period,
                            D_rotation,
                            D_course_id,
                            # D_course_section_id,
                            sep = "_"))

# ==============================================================================
# Part VI: Create row IDs ----
# ==============================================================================

# **What:** Create a row id before explosions. You will use this to identify records that have been flagged and resolved through double-bookings processing (scripts 05 and 06) and to ensure that you don't lose any records in the process.

# **Check:** The max number should match the number of rows in the R environment. If you lose records in the double-booking process, you can use this to identify which ones and add them back in. You can also use this to check that you are not double counting any records (e.g., if you have 1000 rows in your original data, after explosions you should have more than 1000 rows but all of the row ids should be less than or equal to 1000).

# **Project Team Action:** None

# %% Part VI: Create row IDs
# Create record id
cs_data <- cs_data %>%
  ungroup() %>%
  mutate(H_row_id_before_explosions = row_number())

# Check
max(cs_data$H_row_id_before_explosions)

# ==============================================================================
# Part VII: IF YOU DON'T NEED TO EXPLODE DATA ----
# ==============================================================================
cs_data <- cs_data %>% 
 mutate(C_term_weight_exploded = C_term_weight,
        C_term_exploded = D_term)

cs_data <- cs_data %>% 
  mutate(C_rotation_weight_exploded = C_rotation_weight,
         C_rotation_exploded = D_rotation)

cs_data <- cs_data %>% 
  mutate(C_period_weight_exploded = C_period_weight,
         C_period_exploded = D_period)


# Check distribution of values in exploded columns
exploded_check <- cs_data |>
  filter(C_class_size_and_teacher_load_exclude == "Include") |>
  group_by(
    D_term,
    C_term_exploded,
    C_term_weight_exploded,
    D_period,
    C_period_exploded,
    C_period_weight_exploded,
    D_rotation,
    C_rotation_exploded,
    C_rotation_weight_exploded,
    M_class_weight) |> 
  summarise(count = n()) |>
  ungroup()

# Create record id
cs_data <- cs_data %>%
  ungroup() %>%
  mutate(H_row_id_after_explosions = row_number())

# Check
max(cs_data$H_row_id_after_explosions)
# ==============================================================================
# Part VIII: Export ----
# ==============================================================================

# **What:** Use this code block to export your data into SharePoint to
# load unto the next script

# **Check:** Look at your SharePoint folder and make sure that the data
# loaded properly

# **Project Team Action:** Maybe change (you can tweak the file name to
# your liking)

# %% Part VIII: Export
# Save course schedule data for next script
ers_write_sharepoint(
  data = cs_data,
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/2. Processed Data/03_course_data_post_coding_exclusions.csv")

# GREAT JOB!!!
