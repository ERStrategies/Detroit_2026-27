# 08-rollups
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

# - the code block should print the name of the data file you're looking to load

# - confirm that your data file did load and that the columns and row numbers look right (do this by accessing the data in the environment window in the top right of RStudio)

# Project Team Action: maybe change (see guidance in code block)

# %% 1.2 Load data
# Create folder path
raw_data_folder_path <- "District Partners/Detroit Public Schools/26-27 HS Redesign Implementation/1. Data & Analysis - Secure/Fall 2026 Course Schedule Analysis"

# Load unified data
course_data_unified <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/2. Processed Data/07_cs_data_post_unification.csv")

# Load course grade data
course_grades <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/0. Raw Data/Stored Grades YE26 Q1-Q3 Update.xlsx")

# Load newer course section data
cs_raw_data <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/0. Raw Data/course_sections_251103.csv")

# Load course coding
course_coding <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/1. Coding/course_coding_2024.xlsx")

# Load teacher licensure data
# teacher_licensure <- ers_read_sharepoint(
#   folder_path = raw_data_folder_path,
#   file_name_with_extension =
#     "/Raw Data/teacher_data_2025_2026.xlsx")

# May also need to load other data (e.g., HR data)


# 1.3 Pull in time per period ----

# What:

# Check:

# Project Team Action:

# %% 1.3 Pull in time per period
# Create table with minutes per period by school based on bell schedule
# period_minutes <- tibble(
#   D_location_id = c(
#     rep(103, 8),  # Putnam
#     rep(104, 5)   # Sci-Tech
#   ),
#   D_period = c(
#     1:8,          # Putnam periods
#     1:4, 8        # Sci-Tech periods
#   ),
#   D_minutes_attended = c(
#     rep(43, 8),               # Putnam minutes
#     98, 88, 90, 87, 55        # Sci-Tech minutes
#   )
# )


# %% 1.3 Pull in time per period (2)
# Merge in time data
# course_data_unified <- course_data_unified %>%
#   left_join(
#     period_minutes,
#     by = c("D_location_id", "D_period"))


# 1.4 Check for NAs ----

# What:

# Check:

# Project Team Action:

# %% 1.4 Check for NAs
# Create new data frame with column information
col_info <-  data.frame(
  Column_Name = names(course_data_unified),
  Data_Type = sapply(course_data_unified, class),
  Count_Rows = sapply(course_data_unified, function(x) length(na.omit(x))),
  NA_Rows = sapply(course_data_unified, function(x) sum(is.na(x))),
  Unique_Rows = sapply(course_data_unified, function(x) length(unique(na.omit(x)))),
  stringsAsFactors = FALSE,
  row.names= NULL
)

# Calc % missing
col_info <- col_info %>%
            mutate(Percent_NA = NA_Rows / Count_Rows * 100)

# Print new data frame
col_info

# # Analyze data from students missing demographic info - should be zero
# stu_dem_na <- course_data_unified %>%
#   filter(if_any(c(
#     D_stu_grade,
#     D_stu_swd_flag,
#     D_stu_ell_flag,
#     D_stu_poverty_flag),
#     is.na))

# stu_dem_na %>%
#   group_by(D_location_name, D_stu_id) %>%
#   summarise(class_count = n_distinct(C_class_id))

# # Check rows where D_minutes_attended is NA
# course_data_unified %>%
#   filter(is.na(D_minutes_attended)) %>%
#   group_by(D_location_name,
#            C_course_name,
#            D_period,
#            C_course_time_exclude) %>%
#   summarise(class_count = n_distinct(C_class_id),
#             stu_count = n_distinct(D_stu_id))


# Part II: Class Roll-up ----

# 2.1 Class-level flags ----

# What: this code block creates class roll-ups (and note that code block will need fleshing out once the "MISSING" variables get values based on calculations from course data)

# Check:

# - First check shows number of values for teacher_load_exclude, time_of_day, employee id, and course name for each class id - you would expect to see all values as one for the first three variables and values above 1 for course name only for consolidated courses (look at the check_key_variables dataframe in the environment to confirm that the data does follow these two must-haves)

# - Second check shows the consolidated course names for consolidated classes (confirm that the names look roughly as you would have expected) and third check looks at the count of consolidated class names by class id and then shows the max value of the count across all classes (should have max 1 since all classes - even consolidated classes - should now have a single course name)

# Project Team Action:

# %% 2.1 Class-level flags
# Create function to collapse course names for consolidated classes
collapse_course_names <- function(course_names) {
  course_names <- course_names |>
    stringr::str_squish()

  course_names <- unique(course_names[!is.na(course_names) & course_names != ""])

  # Sort naturally, so 9 comes before 11
  course_names <- stringr::str_sort(course_names, numeric = TRUE)

  if (length(course_names) == 0) return(NA_character_)
  if (length(course_names) == 1) return(course_names)

  # Drop shorter names that are fully contained in longer names
  # Example: "Phys Ed" + "Phys Ed 10" -> "Phys Ed 10"
  keep <- !sapply(seq_along(course_names), function(i) {
    this_name <- course_names[i]
    other_names <- course_names[-i]

    any(stringr::str_detect(other_names, stringr::fixed(this_name)))
  })

  course_names <- course_names[keep]

  # Sort again after dropping contained names
  course_names <- stringr::str_sort(course_names, numeric = TRUE)

  if (length(course_names) == 1) return(course_names)

  # Find shared starting words
  words <- stringr::str_split(course_names, "\\s+")
  min_words <- min(lengths(words))

  prefix_len <- 0

  for (i in seq_len(min_words)) {
    ith_words <- sapply(words, function(x) x[i])

    if (length(unique(ith_words)) == 1) {
      prefix_len <- i
    } else {
      break
    }
  }

  # If no shared beginning, paste full names
  if (prefix_len == 0) {
    return(paste(course_names, collapse = " / "))
  }

  common_prefix <- paste(words[[1]][seq_len(prefix_len)], collapse = " ")

  suffixes <- sapply(words, function(x) {
    suffix <- x[-seq_len(prefix_len)]
    paste(suffix, collapse = " ")
  })

  suffixes <- suffixes[suffixes != ""]

  # Sort suffixes too
  suffixes <- stringr::str_sort(suffixes, numeric = TRUE)

  stringr::str_squish(
    paste0(common_prefix, " ", paste(suffixes, collapse = " / "))
  )
}


# %% 2.1 Class-level flags (2)
# Join
course_data_unified <- course_data_unified %>%
  select(-any_of("C_class_course_name")) %>% 
  group_by(C_class_id) %>%
  mutate(C_class_course_name = collapse_course_names(C_course_name)) %>%
  ungroup()

# Check
class_name_check <- course_data_unified %>%
  group_by(C_class_course_name, C_course_name) %>%
  summarize(class_count = n_distinct(C_class_id))


# %% 2.1 Class-level flags (3)
# Check
check_key_variables <- course_data_unified %>% 
  group_by(C_class_id) %>% 
  summarize(
    count_class_size_exclude = n_distinct(C_class_size_exclude),
    count_teacher_load_exclude = n_distinct(C_teacher_load_exclude),
    count_time_of_day = n_distinct(C_course_time_of_day),
    count_teacher_id = n_distinct(D_employee_id),
    count_course_name = n_distinct(C_course_name))

# Creating concatenation of subject area for consolidated courses
course_data_unified <- course_data_unified %>% 
  group_by(C_class_id, C_course_subject) %>% 
  count() %>% 
  group_by(C_class_id) %>% 
  summarize(C_class_subject = paste0(
    C_course_subject, collapse = " / ")) %>% 
  right_join(course_data_unified,
             by = "C_class_id")

# Creating concatenation of grad required flags for consolidated courses
class_credit_type <- course_data_unified %>%
  filter(!is.na(C_course_credit_type)) %>%
  group_by(C_class_id, C_course_credit_type) %>%
  summarise(n_students = n(), .groups = "drop") %>%
  mutate(priority = case_when(
    C_course_credit_type == "Graduation Required" ~ 1,
    C_course_credit_type == "Support & Enrichment" ~ 2,
    C_course_credit_type == "Elective" ~ 3,
  #  C_course_credit_type == "UNSURE" ~ 4,
    TRUE ~ 5
  )) %>%
  group_by(C_class_id) %>%
  arrange(desc(n_students), priority, .by_group = TRUE) %>%
  slice(1) %>%
  select(
    C_class_id,
    C_class_credit_type = C_course_credit_type
  )

course_data_unified <- course_data_unified %>%
  left_join(class_credit_type, by = "C_class_id")

# Check
course_data_unified %>% 
filter(substr(C_class_id, 1, 12) == "Consolidated") %>% 
group_by(C_class_id,
         C_class_course_name) %>% 
count()

# Check
course_data_unified %>% 
  group_by(C_class_id,
           C_class_course_name) %>% 
  count() %>% 
  group_by(C_class_id) %>% 
  count() %>% 
  pull(n) %>% 
  max()

# Check
course_data_unified %>% 
  group_by(D_stu_swd_flag) %>% 
  count()

# Check
course_data_unified %>% 
  group_by(D_stu_ell_flag) %>% 
  count()

# Check
course_data_unified %>% 
  group_by(D_stu_id,
           D_stu_swd_flag,
           D_stu_ell_flag) %>% 
  count() %>% 
  group_by(D_stu_id) %>% 
  count() %>% 
  pull(n) %>% 
  max()


# %% 2.1 Class-level flags (4)
# Calculating EL and SWD class characteristics
el_swd_class_helper <- course_data_unified %>% 
  group_by(C_class_id, 
           D_stu_id,
           D_stu_ell_flag, 
           D_stu_swd_flag) %>% 
  count() %>% 
  group_by(C_class_id) %>% 
  summarize(row_count = n(),
            M_num_ell = sum(D_stu_ell_flag),
            M_num_swd = sum(D_stu_swd_flag),
            M_pct_ell = sum(D_stu_ell_flag) / n(),
            M_pct_swd = sum(D_stu_swd_flag) / n(),
            C_ell_bucket = case_when(M_pct_ell <= 0.2 ~ "0-20%",
                                  M_pct_ell < 0.5 ~ "20-50%",
                                  M_pct_ell <= 1 ~ ">50%",
                                  TRUE ~ "Error"),
            C_swd_bucket = case_when(M_pct_swd <= 0.2 ~ "0-20%",
                                  M_pct_swd < 0.5 ~ "20-50%",
                                  M_pct_swd <= 1 ~ ">50%",
                                  TRUE ~ "Error"))

# Bringing values back to course_data_unified
course_data_unified <- el_swd_class_helper %>% 
  select(-c(row_count)) %>% 
  right_join(course_data_unified,
             by = "C_class_id")

# Check
course_data_unified %>% 
  group_by(C_class_credit_type) %>% 
  summarise(class_count = n_distinct(C_class_id))


# %% 2.1 Class-level flags (5)
# Clean consolidated class names so they can be used in teacher rollups

# Count the number of students in each class for each value of D_course_name
class_name_check <- course_data_unified %>% 
  group_by(
    C_class_id,
    C_class_subject,
    C_class_course_name,
    D_course_name) %>% 
  summarise(
    n_teachers = n_distinct(D_employee_id),
    n_students = n_distinct(D_stu_id),
    n_stu_swd = n_distinct(D_stu_id[D_stu_swd_flag == 1]),
    n_stu_ell = n_distinct(D_stu_id[D_stu_ell_flag == 1]),
    .groups = "drop"
  ) |> 
    group_by(C_class_id, C_class_course_name) |>
    mutate(
      n_course_names = n_distinct(D_course_name),
      percent_swd = sum(n_stu_swd) / sum(n_students),
      percent_ell = sum(n_stu_ell) / sum(n_students)
    ) |>
    ungroup() |> 
    arrange(desc(n_course_names)) |> 
    select(
      C_class_id,
      C_class_subject,
      C_class_course_name,
      n_course_names,
      percent_swd,
      percent_ell,
      n_teachers,
      D_course_name,
      n_students,
      n_stu_swd,
      n_stu_ell
    )

# count the number of classes with more than one course name
class_name_check |> 
  filter(n_course_names > 1) |> 
  summarise(count = n_distinct(C_class_id))


# 2.2 Primary flags at class level ----

# %% 2.2 Primary flags at class level
# Create function to identify the primary grade / school
primary_class <- function(field, output_name) {
  primary_df <- course_data_unified %>% 
    group_by(C_class_id, {{ field }}) %>% 
    summarize(
      H_sum_class_weight = sum(M_class_weight, na.rm = TRUE),
      .groups = "drop"
    ) %>% 
    group_by(C_class_id) %>% 
    mutate(max_class_weight_sum = max(H_sum_class_weight, na.rm = TRUE)) %>% 
    ungroup() %>% 
    filter(H_sum_class_weight == max_class_weight_sum) %>% 
    group_by(C_class_id) %>% 
    slice(1) %>% 
    ungroup() %>% 
    select(C_class_id, {{ field }})
  
  names(primary_df)[2] <- output_name
  
  course_data_unified %>% 
    left_join(primary_df, by = "C_class_id")
}

# Run the function
course_data_unified <- primary_class(D_location_id, "C_class_location_id")
course_data_unified <- primary_class(D_location_name, "C_class_location_name")
course_data_unified <- primary_class(D_stu_grade, "C_class_primary_grade")
course_data_unified <- primary_class(C_course_subject, "C_class_primary_subject")
course_data_unified <- primary_class(C_course_subject_area, "C_class_primary_subject_area")
course_data_unified <- primary_class(C_course_rigor, "C_class_primary_rigor")
course_data_unified <- primary_class(C_course_rigor_detail, "C_class_primary_rigor_detail")


# 2.2 Update student count adjuster ----

# %% 2.2 Update student count adjuster
student_count_adjuster <- course_data_unified %>%
  filter(!is.na(student_adjuster_post_student)) %>%
  filter(student_adjuster_post_student != 1) %>%
  group_by(D_location_name,
           D_stu_id,
           C_class_id,
           C_course_name,
           D_term,
           D_period,
           D_rotation,
           M_class_weight,
           student_adjuster_post_student) %>%
  summarise(count = n())

# Check distribution of student count adjuster
course_data_unified %>%
  group_by(student_adjuster_post_student) %>%
  summarise(count = n())

# Update student count adjuster from NA to 1
course_data_unified <- course_data_unified %>%
  mutate(student_adjuster_post_student =
           if_else(is.na(student_adjuster_post_student),
                   1,
                   student_adjuster_post_student))

# Rerun check
course_data_unified %>%
  group_by(student_adjuster_post_student) %>%
  summarise(count = n())

# Calculate M_student_count_adjusted
# This is what you'll use to calculate class size
course_data_unified <- course_data_unified %>%
  mutate(M_student_count_adjusted = 1 / student_adjuster_post_student)

# Rerun check using new field
course_data_unified %>%
  group_by(M_student_count_adjusted) %>%
  summarise(count = n())


# 2.3 Calculate class size ----

# %% 2.3 Calculate class size
# Calculate class size
course_data_unified <- course_data_unified %>%
  group_by(
    C_class_id,
    D_term,
    D_period,
    D_rotation) %>%
  mutate(M_num_stu = n_distinct(D_stu_id),
         M_class_size = sum(M_student_count_adjusted),
         M_class_size_diff = M_num_stu - M_class_size)

# Create class size buckets
course_data_unified <- course_data_unified %>%
  mutate(
    M_class_size = as.numeric(M_class_size),
    C_class_size_bucket = case_when(
      is.na(M_class_size) ~ NA_character_,
      M_class_size < 0 ~ NA_character_,
      M_class_size <= 6 ~ "0 to 6",
      M_class_size <= 12 ~ "7 to 12",
      M_class_size <= 18 ~ "13 to 18",
      M_class_size <= 24 ~ "19 to 24",
      M_class_size <= 30 ~ "25 to 30",
      M_class_size <= 36 ~ "31 to 36",
      M_class_size > 36 ~ "37+"
    ),
    C_class_size_bucket = factor(
      C_class_size_bucket,
      levels = c("0 to 6", "7 to 12", "13 to 18", "19 to 24", "25 to 30", "31 to 36", "37+"),
      ordered = TRUE))

# Check distribution
course_data_unified %>%
  group_by(C_class_size_bucket) %>%
  summarise(count = n_distinct(C_class_id))

# Check distribution
course_data_unified %>%
  group_by(M_class_size_diff) %>%
  summarise(count = n_distinct(C_class_id)) |> 
  arrange(desc(M_class_size_diff))

# Check distribution of class weight
course_data_unified %>%
  group_by(M_class_weight) %>%
  summarise(count = n())

# Create class size * class weight and expression weight variables
course_data_unified <- course_data_unified |> 
  group_by(C_class_id) |> 
  mutate(
    M_max_class_weight = max(M_class_weight, na.rm = TRUE),
    M_class_weight_times_class_size = M_max_class_weight * max(M_class_size)) |> 
  ungroup()

# Check distribution of M_expression_weight variable
course_data_unified |> 
  group_by(M_expression_weight) |> 
  summarise(count = n_distinct(C_class_id)) |> 
  arrange(desc(M_expression_weight))

# Final check
class_size_check <- course_data_unified |> 
  filter(M_class_size_diff != 0 & C_class_size_exclude != "Exclude") |> 
  group_by(
    C_class_id,
    D_term,
    D_period,
    D_rotation,
    C_course_name,
    C_class_size_exclude,
    M_num_stu,
    M_class_size,
    M_class_size_diff,
    M_class_weight_times_class_size) |> 
  summarise(count = n())

# Filter for records with M_expression_weight is NA
expression_weight_check <- course_data_unified %>%
  filter(is.na(M_expression_weight)) %>%
  group_by(
    C_class_id,
    D_employee_id,
    D_term,
    D_expression,
    D_period,
    D_rotation,
    C_course_name,
    C_expression_weight,
    C_term_weight,
    C_teacher_load_exclude,
    C_class_size_exclude,
    C_course_time_exclude) %>%
  summarise(count = n())


# 2.4 Create roll-up table ----

# What:

# Check:

# Project Team Action:

# %% 2.4 Create roll-up table
# Filtering to just variables needed and then creating some additional variables that also need to be included
class_rollup <- course_data_unified %>% 
  group_by(C_class_location_id,
           C_class_location_name,
           C_class_id,
           D_term,
           D_period,
           D_rotation,
           C_teacher_load_exclude,
           C_class_size_exclude,
           D_employee_id,
           C_class_primary_subject,
           C_class_primary_subject_area,
           C_class_course_name,
           C_class_credit_type,
           C_class_primary_grade,
           C_class_primary_rigor,
           C_class_primary_rigor_detail,
           M_num_ell,
           M_num_swd,
           M_pct_ell,
           M_pct_swd,
           C_ell_bucket,
           C_swd_bucket,
           M_num_stu,
           M_class_size,
           C_class_size_bucket) |> 
  summarize(M_class_weight = max(M_class_weight),
            D_minutes_attended = max(D_minutes_attended),
            M_class_weight_times_class_size = max(M_class_weight_times_class_size),
            M_expression_weight = max(M_expression_weight))

# Check for duplicates - should be zero
class_rollup %>%
  group_by(C_class_id) %>%
  summarise(
    across(everything(), ~ n_distinct(.x, na.rm = TRUE)),
    .groups = "drop"
  ) %>%
  filter(if_any(-C_class_id, ~ .x > 1))


# Part III: Teacher Roll-ups ----

# 3.1 EL/SWD flags ----

# What:

# Check:

# Project Team Action:

# %% 3.1 EL/SWD flags
# Creating EL and SWD characteristics (I manually checked the values of the helper to make sure they make sense)
el_swd_teacher_helper <- course_data_unified %>% 
  mutate(H_ell_class = ifelse(M_pct_ell == ">50%",
                              1,
                              0),
         H_swd_class = ifelse(M_pct_swd == ">50%",
                              1,
                              0)) %>% 
  group_by(D_employee_id,
           C_class_id,
           M_pct_ell,
           M_pct_swd) %>% 
  count() %>% 
  group_by(D_employee_id) %>% 
  summarize(C_percent_ell_classes = mean(M_pct_ell),
            C_percent_swd_classes = mean(M_pct_swd),
            C_ell_teacher_flag = ifelse(C_percent_ell_classes > 0.5,
                                        "ELL teacher",
                                        "Not ELL teacher"),
            C_swd_teacher_flag = ifelse(C_percent_swd_classes > 0.5,
                                        "SWD teacher",
                                        "Not SWD teacher"),
            C_ell_swd_teacher_flag = case_when(
              C_ell_teacher_flag == "ELL teacher" & C_swd_teacher_flag == "SWD teacher" ~ "ELL and SWD teacher",
              C_ell_teacher_flag == "ELL teacher" & C_swd_teacher_flag == "Not SWD teacher" ~ "ELL teacher",
              C_ell_teacher_flag == "Not ELL teacher" & C_swd_teacher_flag == "SWD teacher" ~ "SWD teacher",
              C_ell_teacher_flag == "Not ELL teacher" & C_swd_teacher_flag == "Not SWD teacher" ~ "Gen Ed teacher",
              TRUE ~ NA_character_))

# Joining back EL/SWD characteristics
course_data_unified <- course_data_unified %>% 
  left_join(el_swd_teacher_helper,
            by = "D_employee_id")


# 3.2 HR flags ----

# What: Includes teacher evaluation ratings, years experience / novice teacher flags, and teacher race

# Check:

# Project Team Action: If you want to run this block, change from "{r, eval=FALSE}" to "{r}"

# %% 3.2 HR flags
# # Joining evaluation score
# course_data_unified <- teacher_evaluation %>%
#   select(EEID,
#          `Overall Eval Rating for ODE`) %>% 
#   mutate(D_employee_id = as.character(EEID),
#          D_employee_rating_18_19 = `Overall Eval Rating for ODE`) %>% 
#   select(D_employee_id,
#          D_employee_rating_18_19) %>% 
#   right_join(course_data_unified,
#              by = "D_employee_id")
# 
# # Check
# course_data_unified %>% 
#   group_by(D_employee_rating_18_19) %>% 
#   count()
# 
# # Teacher years experience and novice teacher flag
# course_data_unified <- teacher_yrs_exp %>% 
#   mutate(D_employee_id = as.character(EEID),
#          D_teacher_years_experience = YrsInORCnt,
#          C_novice_teacher_indicator = ifelse(D_teacher_years_experience <= 2,
#                                              "Novice",
#                                              "Not Novice")) %>%
#   select(D_employee_id,
#          D_teacher_years_experience,
#          C_novice_teacher_indicator) %>%
#   right_join(course_data_unified,
#              by = "D_employee_id")
# 
# # Check
# course_data_unified %>% 
#   group_by(C_novice_teacher_indicator) %>% 
#   count()
# 
# # Teacher race
# course_data_unified <- teacher_race %>% 
#   mutate(D_employee_id = as.character(ID),
#          D_employee_race = `PPS EEO CD`) %>% 
#   select(D_employee_id,
#          D_employee_race) %>% 
#   right_join(course_data_unified,
#              by = "D_employee_id")
# 
# # Check
# course_data_unified %>% 
#   group_by(D_employee_race) %>% 
#   count()


# %% 3.2 HR flags (2)
# Create primary license level by teacher
# teacher_licensure_summary <- teacher_licensure %>% 
#   group_by(D_employee_license_type) %>% 
#   summarise(count = n_distinct(D_employee_id))

# # Assign rank by license type
# teacher_licensure_summary <- teacher_licensure_summary %>%
#   mutate(
#     D_employee_license_type_rank = case_when(
#       D_employee_license_type == "Professional" ~ 1,
#       D_employee_license_type == "Initial - Extension" ~ 2,
#       D_employee_license_type == "Initial" ~ 3,
#       D_employee_license_type == "Provisional" ~ 4,
#       D_employee_license_type == "Preliminary-Extension" ~ 5,
#       D_employee_license_type == "Preliminary" ~ 6,
#       D_employee_license_type == "Temporary" ~ 7,
#       D_employee_license_type == "Emergency-Extension II" ~ 8,
#       TRUE ~ NA_real_)) %>% 
#   arrange(D_employee_license_type_rank) %>% 
#   select(-count)

# # Merge rank into original data set
# teacher_licensure_rank <- teacher_licensure %>% 
#   left_join(teacher_licensure_summary,
#             by = "D_employee_license_type")

# # Identify top rank per teacher
# teacher_licensure_rank <- teacher_licensure_rank %>% 
#   group_by(D_employee_id) %>% 
#   mutate(top_rank = min(D_employee_license_type_rank)) %>% 
#   group_by(D_employee_id,
#            D_employee_license_type_rank) %>% 
#   mutate(match_flag = case_when(
#     D_employee_license_type_rank == top_rank ~ "Top rank"
#   ))

# # Create list of license subjects per teacher
# teacher_licensure_rank <- teacher_licensure_rank %>%
#   group_by(D_employee_id) %>%
#   arrange(D_employee_license_type_rank, .by_group = TRUE) %>%
#   mutate(
#     D_employee_license_subject_concat = paste(
#       unique(D_employee_license_subject[!is.na(D_employee_license_subject) & D_employee_license_subject != ""]),
#       collapse = ", "))

# # Create a list with one row per teacher
# teacher_licensure_merge <- teacher_licensure_rank %>% 
#   group_by(D_employee_id,
#            D_employee_prof_status,
#            top_rank,
#            D_employee_license_subject_concat) %>% 
#   summarise(count = n())

# # Check teacher count
# n_distinct(teacher_licensure$D_employee_id)
# n_distinct(teacher_licensure_merge$D_employee_id)

# # Merge top licensure type back in
# teacher_licensure_merge <- teacher_licensure_merge %>% 
#   left_join(teacher_licensure_summary,
#             by = c("top_rank" = "D_employee_license_type_rank")) %>% 
#   select(-count) %>% 
#   rename(D_employee_license_type_rank = top_rank)


# %% 3.2 HR flags (3)
# Merge teacher licesure data into course data unified
# course_data_unified <- course_data_unified %>% 
#   left_join(teacher_licensure_merge,
#             by = "D_employee_id")

# # Simplify licensure type to use as proxy for teacher experience
# course_data_unified <- course_data_unified |> 
#   mutate(
#     D_employee_license_type = case_when(
#       D_employee_license_type == "Professional" ~ "1 - Professional",
#       D_employee_license_type == "Initial - Extension" ~ "2 - Initial",
#       D_employee_license_type == "Initial" ~ "2 - Initial",
#       D_employee_license_type == "Provisional" ~ "3 - Provisional",
#       D_employee_license_type == "Preliminary-Extension" ~ "3 - Provisional",
#       D_employee_license_type == "Preliminary" ~ "3 - Provisional",
#       D_employee_license_type == "Temporary" ~ "3 - Provisional",
#       D_employee_license_type == "Emergency-Extension II" ~ "4 - Emergency",
#       is.na(D_employee_license_type) ~ "5 - Not licensed",
#       TRUE ~ NA_character_))


# %% 3.2 HR flags (4)
# Create novice teacher indicator based on license type rank (Professional-Preliminary = valid for 5+ years)
# course_data_unified <- course_data_unified |> 
#   mutate(C_novice_teacher_indicator = case_when(
#     D_employee_license_type_rank >= 7 ~ "Novice",
#     D_employee_license_type_rank < 7 ~ "Experienced",
#     is.na(D_employee_license_type_rank) ~ "Not licensed",
#     TRUE ~ NA_character_))

# # Check distribution of novice teacher indicator
# course_data_unified |> 
#   group_by(C_novice_teacher_indicator) |> 
#   summarise(count = n_distinct(D_employee_id))


# 3.3 Primary flags at teacher level ----

# What: Creates flags based on each teachers' primary school, subject area, and grade level

# Check:

# Project Team Action:

# %% 3.3 Primary flags at teacher level
# Primary subject / subject area / grade / school
primary <- function(field, exclude_advisory = FALSE) {
  source_df <- if (exclude_advisory) {
  course_data_unified |> 
    dplyr::filter(is.na(C_course_subject) | C_course_subject != "Advisory")
} else {
  course_data_unified
}

  source_df |>
    group_by(D_employee_id, C_class_id, {{ field }}) |>
    summarise(H_max_class_weight = max(M_class_weight, na.rm = TRUE), .groups = "drop") |>
    group_by(D_employee_id, {{ field }}) |>
    summarise(
      row_count = n(),
      H_sum_class_weights_subject = sum(H_max_class_weight, na.rm = TRUE),
      .groups = "drop"
    ) |>
    group_by(D_employee_id) |>
    mutate(max_class_weight_sum_subject = max(H_sum_class_weights_subject, na.rm = TRUE)) |>
    ungroup() |>
    filter(H_sum_class_weights_subject == max_class_weight_sum_subject) |>
    group_by(D_employee_id) |>
    slice_sample(n = 1) |>
    ungroup() |>
    mutate(new_field = {{ field }}) |>
    select(D_employee_id, new_field) |>
    right_join(course_data_unified, by = "D_employee_id")
}

course_data_unified <- primary(C_course_subject, exclude_advisory = TRUE) |>
  rename(C_primary_course_subject = new_field)

course_data_unified <- primary(C_course_subject_area, exclude_advisory = TRUE) |>
  rename(C_primary_course_subject_area = new_field)

course_data_unified <- primary(C_course_credit_type, exclude_advisory = TRUE) |>
  rename(C_primary_course_credit_type = new_field)

course_data_unified <- primary(D_location_name) |>
  rename(C_primary_location_name = new_field)

course_data_unified <- primary(D_stu_grade) |>
  rename(C_primary_grade = new_field)

# Patch NAs with Advisory for primary subject and subject area
course_data_unified <- course_data_unified |>
  mutate(
    C_primary_course_subject = if_else(
      is.na(C_primary_course_subject),
      "Advisory",
      C_primary_course_subject
    ),
    C_primary_course_subject_area = if_else(
      is.na(C_primary_course_subject_area),
      "Support & Enrichment",
      C_primary_course_subject_area
    )
  )

# Check distribution of primary flags
course_data_unified |>
  group_by(C_primary_course_subject_area, C_primary_course_subject) |>
  summarise(count = n_distinct(D_employee_id))


# 3.3 Teacher utilization ----

# What:

# Check:

# Project Team Action:

# %% 3.3 Teacher utilization
# Teacher utilization metrics - first check to see number of periods and then actually calculate the metrics
course_data_unified %>%
  group_by(C_primary_location_name, C_class_id, D_period) %>%
  summarize() %>%
  group_by(C_primary_location_name, D_period) %>%
  count()

# In Springfield: schools had different bell schedules, so needed to calculate teacher utilization rates differently
teacher_utilization <- course_data_unified |> 
  filter(C_teacher_load_exclude == "Include") |>
  group_by(
    C_primary_location_name,
    D_employee_id,
    D_term,
    D_rotation,
    D_period) |> 
  summarize(
    M_teacher_class_weight = max(M_class_weight)) |> 
  group_by(
    C_primary_location_name,
    D_employee_id) |>
  summarise(
    M_num_periods = sum(M_teacher_class_weight)) |> 
  group_by(
    C_primary_location_name,
    D_employee_id) |>
  mutate(
    C_school_total_periods = case_when(
      C_primary_location_name == "PUTNAM" ~ 8,
      C_primary_location_name == "SCI TECH" ~ 5,
      TRUE ~ NA_real_),
    M_teacher_utilization = M_num_periods / C_school_total_periods)

# Merge teacher utilization back into main data set
course_data_unified <- course_data_unified |>
  left_join(
    teacher_utilization %>% 
      select(
        D_employee_id,
        M_num_periods,
        M_teacher_utilization),
    by = c(
      "D_employee_id",
      "C_primary_location_name"))

# Check teachers missing teacher utilization
course_data_unified |>
  filter(is.na(M_teacher_utilization)) |>
  group_by(
    D_employee_id,
    C_course_name) |>
  summarise(count = n_distinct(C_class_id)) |> 
  print(n = 100)


# 3.4 Teacher load ----

# What:

# Check:

# Project Team Action:

# %% 3.4 Teacher load
# Calculate teacher load as number of unique students taught across all classes that a teacher teaches
course_data_unified <- course_data_unified |>
  group_by(
    C_primary_location_name,
    D_employee_id
  ) |>
  mutate(
    M_teacher_load = n_distinct(D_stu_id[C_teacher_load_exclude == "Include"]),
    M_teacher_load_bucket = case_when(
      M_teacher_load <= 34 ~ "0 to 34",
      M_teacher_load <= 74 ~ "35 to 74",
      M_teacher_load <= 124 ~ "75 to 124",
      M_teacher_load <= 199 ~ "125 to 199",
      M_teacher_load >= 200 ~ "200+",
      TRUE ~ NA_character_
    )
  ) |>
  ungroup()

# Check distribution by teacher load bucket
course_data_unified |> 
  group_by(M_teacher_load_bucket) |> 
  summarise(count = n_distinct(D_employee_id))

# Create 9th grade teacher flag
course_data_unified <- course_data_unified |> 
  mutate(C_ninth_grade_teacher_flag = if_else(
    C_primary_grade == 9,
    "Ninth grade teacher",
    "Not Ninth grade teacher"
  ))


# 3.5 Create roll-up table ----

# What:

# Check:

# Project Team Action:

# %% 3.5 Create roll-up table
# Create teacher rollup based on all information calculated so far
teacher_rollup <- course_data_unified |> 
  filter(C_teacher_load_exclude == "Include" & C_course_subject_area != "Untracked") |> 
  group_by(D_employee_id,
           C_primary_course_subject,
           C_primary_course_subject_area,
           C_primary_course_credit_type,
           C_primary_grade,
           C_primary_location_name,
           C_percent_ell_classes,
           C_percent_swd_classes,
           C_ell_swd_teacher_flag,
           C_ell_teacher_flag,
           C_swd_teacher_flag,
           # D_empl_score,
           # D_empl_effectiveness,
           # D_teacher_years_experience,
           C_novice_teacher_indicator,
           # D_employee_race,
           # D_employee_gender,
           # D_position_title,
           # D_position_type,
           # D_position_subtype,
           # D_job_title,
           D_employee_license_type_rank, # Added for Springfield
           D_employee_license_type, # Added for Springfield
           D_employee_license_subject_concat, # Added for Springfield
           C_ninth_grade_teacher_flag,
           M_num_periods,
           M_teacher_utilization,
           M_teacher_load,
           M_teacher_load_bucket) |> 
  summarize(
    M_num_preps = n_distinct(C_course_name),
    M_meetings_in_S1 = paste0(unique(D_meeting[D_term != 3502]), collapse = " | "),
    C_teacher_course_names = paste0(unique(C_course_name), collapse = " | "),
    M_teacher_weight = 1) |>
  arrange(
    desc(C_primary_location_name),
    desc(M_teacher_utilization))

# Check
teacher_rollup |> 
  group_by(D_employee_id) |>
  count() |> 
  pull(n) |>
  max()


# 3.6 Update class roll-up with teacher-level flags ----

# %% 3.6 Update class roll-up with teacher-level flags
# Merge teacher-level flags into class roll-up
class_rollup <- class_rollup |> 
  left_join(
    teacher_rollup |> 
      ungroup() |>
      select(
        D_employee_id,
        C_ell_swd_teacher_flag,
        C_primary_location_name,
        C_primary_course_subject,
        C_primary_grade,
        D_employee_license_type,
        C_novice_teacher_indicator,
        M_num_periods,
        M_teacher_utilization,
        M_teacher_load,
        M_teacher_load_bucket,
        M_num_preps),
    by = "D_employee_id")


# Part IV: Student Roll-ups ----

# 4.1 Create student-level proficiency flags ----

# What:

# Check:

# %% 4.1 Create student-level proficiency flags
# Code column names
course_grades <- course_grades %>% 
  rename(
    D_year = `School Year`,
    D_location_name = `School Name`,
    D_stu_id = `Student ID`,
    D_stu_grade = `Grade Level`,
    D_stu_cohort = `Graduation Cohort`,
    D_course_subject = `Course Subject Area Description`,
    D_course_name = `Course Description`,
    D_course_category = `Course Category Description`,
    D_stu_course_grade = F1
  ) %>% 
  select(
    D_year,
    D_location_name,
    D_stu_id,
    D_stu_grade,
    D_stu_cohort,
    D_course_name,
    D_course_subject,
    D_course_category,
    D_stu_course_grade
  )

# Check course categories
course_grades %>% 
  group_by(D_location_name,
           D_course_category) %>% 
  summarise(count = n())

# Code by category
course_grades <- course_grades %>% 
  mutate(
    D_course_credit_recovery = if_else(
      D_course_category %in% c("Credit Recovery", "SUM Summer School", "VHS"),
      1,
      0))

# Create student level table by subject


# %% 4.1 Create student-level proficiency flags (2)
# Prep course coding data for merge
course_coding_merge <- course_coding %>% 
  group_by(
    D_course_name,
    C_course_subject,
    C_course_credit_type) %>% 
  summarise(count = n())

# Merge into course grade data
course_grades <- course_grades %>% 
  left_join(course_coding_merge,
            by = "D_course_name")

# Create filtered course_grades_rollup
course_grades_rollup <- course_grades |> 
  filter(
    D_course_category != "TRN Transfer",
    D_course_credit_recovery == 0,
    !is.na(D_stu_course_grade),
    !is.na(D_stu_id)
  ) |> 
    mutate(
    C_stu_course_grade = str_extract(D_stu_course_grade, "^[A-F]"))


# %% 4.1 Create student-level proficiency flags (3)
# Create a table of just graduation required core courses
course_grades_core <- course_grades |>
  filter(
    C_course_credit_type == "Graduation Required",
    D_course_subject %in% c("English", "Math", "Science", "Social Studies"),
    D_course_category != "TRN Transfer",
    D_course_credit_recovery == 0,
    !is.na(D_stu_course_grade),
    !is.na(D_stu_id)
  ) |>
  mutate(
    C_course_grade_letter = str_extract(D_stu_course_grade, "^[A-F]"),
    C_course_grade_rank = case_when(
      C_course_grade_letter == "A" ~ 1,
      C_course_grade_letter == "B" ~ 2,
      C_course_grade_letter == "C" ~ 3,
      C_course_grade_letter == "D" ~ 4,
      C_course_grade_letter == "F" ~ 5,
      TRUE ~ NA_real_
    ),
    C_course_subject_name = D_course_subject |>
      str_to_lower() |>
      str_replace_all(" ", "_")
  ) |>
  filter(!is.na(C_course_grade_rank)) |>
  group_by(
    D_location_name,
    D_stu_id,
    D_stu_grade,
    D_stu_cohort,
    C_course_subject_name
  ) |>
  summarise(
    C_core_grad_required_grade_rank = max(
      C_course_grade_rank,
      na.rm = TRUE
    ),
    .groups = "drop"
  ) |>
  pivot_wider(
    id_cols = c(
      D_location_name,
      D_stu_id,
      D_stu_grade,
      D_stu_cohort
    ),
    names_from = C_course_subject_name,
    values_from = C_core_grad_required_grade_rank,
    names_glue = "C_course_grade_{C_course_subject_name}"
  )


# %% 4.1 Create student-level proficiency flags (4)
# Select columns to merge back into main data set
course_grades_core <- course_grades_core |> 
  mutate(
        C_proficiency_ela = case_when(
        C_course_grade_english <= 3 ~ "Proficient",
        C_course_grade_english >= 4 ~ "Below Proficient",
        TRUE ~ NA_character_
      ),
      C_proficiency_math = case_when(
        C_course_grade_math <= 3 ~ "Proficient",
        C_course_grade_math >= 4 ~ "Below Proficient",
        TRUE ~ NA_character_
      ),
      C_proficiency = case_when(
        C_proficiency_ela == "Below Proficient" &
        C_proficiency_math == "Below Proficient" ~ "Below Proficient",
  
        C_proficiency_ela == "Below Proficient" &
        C_proficiency_math == "Proficient" ~ "Partially Proficient",
  
        C_proficiency_ela == "Proficient" &
        C_proficiency_math == "Below Proficient" ~ "Partially Proficient",
  
        C_proficiency_ela == "Below Proficient" &
        is.na(C_proficiency_math) ~ "Below Proficient",

        C_proficiency_ela == "Proficient" &
        is.na(C_proficiency_math) ~ "Proficient",
  
        is.na(C_proficiency_ela) &
        C_proficiency_math == "Below Proficient" ~ "Below Proficient",

        is.na(C_proficiency_ela) &
        C_proficiency_math == "Proficient" ~ "Proficient",
  
        C_proficiency_ela == "Proficient" &
        C_proficiency_math == "Proficient" ~ "Proficient",
        TRUE ~ NA_character_
)
    ) |> 
      select(
        D_stu_id,
        C_proficiency_ela,
        C_proficiency_math,
        C_proficiency
      )

# Merge course grade data back into main data set
course_data_unified <- course_data_unified |> 
  left_join(
    course_grades_core,
    by = "D_stu_id"
    )


# 4.2 Create student-level demographic flags ----

# What:

# Check:

# %% 4.2 Create student-level demographic flags
# Create a single SWD/ELL flag for each student
student_flags <- course_data_unified |>
  group_by(D_stu_id) |>
  summarise(
    D_stu_swd_flag = max(D_stu_swd_flag, na.rm = TRUE),
    D_stu_ell_flag = max(D_stu_ell_flag, na.rm = TRUE),
    .groups = "drop"
  ) |> 
    # Create a single combined flag for each student
    mutate(
      D_stu_demographic_flag = case_when(
        D_stu_swd_flag == 1 & D_stu_ell_flag == 1 ~ "Dual identified",
        D_stu_swd_flag == 1 & D_stu_ell_flag == 0 ~ "SWD",
        D_stu_swd_flag == 0 & D_stu_ell_flag == 1 ~ "ELL",
        D_stu_swd_flag == 0 & D_stu_ell_flag == 0 ~ "Gen ed",
        TRUE ~ NA_character_
      )
    )

# Merge back into main data set
course_data_unified <- course_data_unified |> 
  left_join(
    student_flags |> 
      select(D_stu_id, D_stu_demographic_flag),
    by = "D_stu_id"
  )

# Check distribution
course_data_unified |> 
  group_by(D_stu_demographic_flag) |> 
  summarise(count = n_distinct(D_stu_id))


# 4.3 Create roll-up table ----

# What:

# Check:

# %% 4.3 Create roll-up table
student_rollup <- course_data_unified |> 
  filter(C_course_subject_area != "Untracked") |> 
  group_by(D_location_name,
           D_stu_id,
           D_stu_grade,
           D_stu_swd_flag,
           D_stu_ell_flag,
           D_stu_poverty_flag,
           D_stu_demographic_flag,
           C_class_id,
           D_employee_id,
           D_term,
           D_rotation,
           D_period,
           C_course_time_of_day,
           C_course_subject,
           C_course_name,
           C_course_rigor,
           C_course_rigor_detail,
           C_course_subject_area,
           C_course_credit_type,
           C_course_intervention,
           D_minutes_attended,
           M_pct_ell,
           M_pct_swd,
           M_class_weight,
           M_num_stu,
           M_class_size,
           C_class_size_bucket,
           C_proficiency_ela,
           C_proficiency_math,
           C_proficiency) %>% 
  summarise(class_row = 1) |> 
  select(-class_row)


# Part V: Exports ----

# 4.1 Export data ----

# What: limiting course_data_unified only to variables of interest and then exporting to SharePoint

# Check: look at data in environment before exporting it to be sure that you feel good about the set of variables you're passing on to the next script - and once you export it check your SharePoint folder to ensure it got loaded correctly

# Project Team Action: likely change (if you have any additional variables that you want to keep at hand, add them to this code block)

# %% 4.1 Export data
# Save class data for next script
ers_write_sharepoint(
  data = class_rollup,
  folder_path = raw_data_folder_path,
  file_name_with_extension = 
    "/Processed Data/08_class_rollup.xlsx")

# Save teacher data for next script
ers_write_sharepoint(
  data = teacher_rollup,
  folder_path = raw_data_folder_path,
  file_name_with_extension = 
    "/Processed Data/08_teacher_rollup.xlsx")

# Save student data for next script
ers_write_sharepoint(
  data = student_rollup,
  folder_path = raw_data_folder_path,
  file_name_with_extension = 
    "/Processed Data/08_student_rollup.xlsx")

# Save course grades data for next script
ers_write_sharepoint(
  data = course_grades_rollup,
  folder_path = raw_data_folder_path,
  file_name_with_extension = 
    "/Processed Data/08_course_grades_rollup.xlsx")
