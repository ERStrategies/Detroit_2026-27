# 05-db-flaggings
# Converted from R Markdown to a Positron-ready R script.
# Original narrative and instructions are retained as comments.

# ==============================================================================
# Part I: Setup ----
# ==============================================================================

# This part of the script allows you to:

# -   load the functions that R will be using to work with your data (Excel automatically loads all its functions - if, countif, index, match, etc. - when you open it, but R requires manual loading of functions if you want full access to what it offers)

# -   declare the folder path that will tell R where to find your data

# -   load in the data you'll be working with throughout the rest of the script

# 1.1 Load Packages ----

# **What:** This code block loads in all relevant function packages (a.k.a. R functions you'll be needing in this script). Note that if you haven't installed the packages into your computer yet, R will print an error and you'll need to install them first by running install.packages(tidyverse) and install.packages(data.frame).

# **Check:** no error message coming out of code block

# **Project Team Action:** do not change (run as is)

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

# **What:** Now that R knows how to access the folder path, it's time to input the name of the data files you want to load. Note that your files should be in CSV format and that your file names should always end with .csv

# **Check:**

# -   the code block should print the name of the data file you're looking to load

# -   confirm that your data file did load and that the columns and row numbers look right (do this by accessing the data in the environment window in the top right of RStudio)

# **Project Team Action:** maybe change (see guidance in code block)

# %% 1.2 Load data
# Create folder path
raw_data_folder_path <- "District Partners/Detroit Public Schools/26-27 HS Redesign Implementation/1. Data & Analysis - Secure/Fall 2026 Course Schedule Analysis"

# Load processed data
course_data_exploded <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/2. Processed Data/03_course_data_post_coding_exclusions.csv")

# Move D_location_id and _name all the way left
course_data_exploded <- course_data_exploded %>%
  relocate(D_location_id, D_location_name)

# ==============================================================================
# Part II: Calculation Set-Up ----
# ==============================================================================

# 2.1 Exploded Class ID ----

# **What:** this code block creates a new class id for exploded rows based on their new term, rotation, and period values. This new id is used to group exploded rows together later in the script.

# **Check:**

# -   Remind yourself of the C_class_id column you created in 03-coding-and-exclusions.Rmd

# -   confirm that your data now includes the new column (you can confirm this by looking at your data's column names in the environment)

# -   look through the output of the first code block and confirm that the number of exploded class ids per original class id is reasonable (most rows in the output should likely have a count of around 5-20, though specific amounts depend on how many rows you expected a single row to become as a result of the term / rotation / period explosions)

# -   look at the graph created in the second code block and confirm that the number of students in exploded classes is reasonable (vast majority of classes should be at or below 40 - and don't worry if there's some classes that seem at or below 0 since they're very likely just classes with very few students that R "rounds down" to 0 or below when creating bins)

# **Project Team Action:** do not change (run as is)

# %% 2.1 Exploded Class ID
# Detroit: original C_class_id from script 03
#  mutate(C_class_id = paste(D_location_id,
#                            D_term,
#                            D_period,
#                            D_rotation,
#                            D_course_id,
#                            sep = "_"))

# Rename existing class id
course_data_exploded <- course_data_exploded %>%
  rename(C_class_id_original = C_class_id)

# Create exploded class id
course_data_exploded <- course_data_exploded %>%
   mutate(
    C_class_id = paste(
      D_location_id,
      C_term_exploded,
      C_period_exploded,
      C_rotation_exploded,
      D_course_id,
    #  D_course_section,
    #  D_course_section_id,
      sep = "_"))

# Check
course_data_exploded %>%
  group_by(D_course_name) %>%
  summarize(H_count_of_class_id_exploded = n_distinct(C_class_id)) %>%
  head(50)

# %% 2.1 Exploded Class ID (2)
# Check
course_data_exploded %>%
  filter(
    D_employee_id != "@ERR",
    C_course_time_exclude == "Include",
    C_class_size_and_teacher_load_exclude == "Include",
    M_class_weight != 0
  ) %>%
  count(C_class_id) %>%
  ggplot() +
  geom_histogram(mapping = aes(x = n), binwidth = 5) +
  xlim(0, 100) +
  labs(
    title = "Students by Class ID Exploded",
    x = "Students by class",
    y = "Number of classes"
  )

# 2.2 New Record ID ----

# **What:** This will create a row ID that will be referenced in 07_db_unification

# **Check:** Output of code block should match the number of rows in course_data_exploded (which you can see in the environment at the top right corner of RStudio)

# **Project Team Action:** Run as is

# %% 2.2 New Record ID
# Creating record id
course_data_exploded$H_new_record_id <- c(1:nrow(course_data_exploded))

# Check
max(course_data_exploded$H_new_record_id)

# Before identifying double booked rows, there are a few initial steps that need to happen:

# 1.  **Re-Calculate the class_weight based on the new term, rotation and period weights:** given that the explosions process in the previous script made the vast majority of rows (all that weren't already in "lowest expression" form) change the number of terms, periods, and/or rotations that they encompass, the class weights that you calculated based on their original values of term/period/rotation are no longer valid. You will need to update the class weights of exploded rows in a way that reflects the term weight, period weight, and rotation weight that they now have.
# 2.  **Adjust class weight when a student has multiple rows for a single course:** When a student is enrolled in two courses at the same time, it is called a double booking. But if a student is enrolled in multiple rows of the same course during a term/rotation/period, it is not considered a double booking. In this case, the weight of each row needs to be adjusted because it was counted as a separate enrollment. For example, if a student is enrolled in "Algebra I" three times in period 2 / quarter 4 / Thursday, each row would initially have a weight of 0.05, but they should be adjusted to 0.05/3.

# 2.3 Re-Calculate Class Weight ----

# **What:** this code block creates the new class weight based on the new rotation weight, new term weight, and new period weight.

# **Check:** look at the columns in course_data_exploded in the environment (top right corner of RStudio) and confirm that M_class_weight_exploded is now included

# **Project Team Action:** do not change

# %% 2.3 Re-Calculate Class Weight
# Calculates class weight
course_data_exploded <- course_data_exploded %>%
  mutate(M_class_weight_exploded = C_term_weight_exploded *
                                   C_rotation_weight_exploded *
                                   C_period_weight_exploded)

# Check
course_data_exploded %>%
  group_by(M_class_weight_exploded) %>%
  count()

# 2.4 Adjust Weight for Students with Multiple Rows per Course ----

# 2.4a Remove rows that don't need to be adjusted for double-bookings ----

# %% 2.4a Remove rows that don't need to be adjusted for double-bookings
# Filter out courses that don't need to be adjusted for double bookings
# These have already been excluded from time, class size and teacher load analyses based on the coding you did in 03-coding-and-exclusions.Rmd
# They will still be brough back in 07_db_unification.Rmd but for now we need to filter them out so that they don't mess with the logic of the double bookings flagger and adjuster
course_data_exploded <- course_data_exploded |>
  filter(
    C_course_time_exclude == "Include",
    C_class_size_and_teacher_load_exclude == "Include")

# Check what's still included
check_include <- course_data_exploded |>
  group_by(
    C_teacher_load_metric,
    C_class_size_metric,
    C_time_metric,
    C_term_exploded,
    C_period_exploded,
    C_rotation_exploded,
    M_class_weight_exploded,
    C_course_name
  ) |>
    summarise(stu_count = n_distinct(D_stu_id)) |>
    arrange(M_class_weight_exploded)

# 2.4a Count Rows per Student Per Course ----

# **What:** this code block counts the rows that a student has for the same course by creating a new table that groups by student_id and course ID

# **Check:** confirm that the table printed by the code block has the vast majority of values on 1 ("normal" cases) and some values above 1 (cases where a student has multiple rows per course)

# **Project Team Action:** do not change (run as is - and note that it can take a few minutes to run)

# %% 2.4a Count Rows per Student Per Course
# Groups data by class id and student, then counts number of rows per group
student_adjuster_same_course <- course_data_exploded %>%
  group_by(
    C_class_id,
    D_stu_id) %>%
  count() %>%
  rename(H_student_adjuster_same_course = n)

# Check
# Springfield: 0 students have 2 rows for the same C_class_id
table(student_adjuster_same_course$H_student_adjuster_same_course,
      useNA = "ifany")

# 2.4b Adjust Class Weight ----

# **What:** Now that you feel good about the row count calculated in the previous code block, it's time to bring the row count values into course_data_exploded and use them to adjust class weight (which is class_weight/H_student_adjuster_same_course).

# **Check:**

# -   The values in the first table printed by the code block should match what you would expect based on the values of your row count in the previous code block. For example, if your previous block found 500 cases of 2 rows per student per class and 100 cases of 3 rows per student per class, your table should show 1000 rows (500\*2) with class weight decreased to half and 300 rows (100\*3) with class weight decreased to a third.

# -   The values in the second table printed by the code block should match what you would expect based on the values of the first table. For example, if your first table had 1000 cases with value of 1 and 200 cases with value of 2, your table should show 1000 rows with "standard" class weight and 200 rows with "standard" class weight divided by 2.

# **Project Team Action:** do not change (run as is)

# %% 2.4b Adjust Class Weight
# Merges the row count variable back into course_data_exploded
course_data_exploded <- left_join(course_data_exploded,
                                  student_adjuster_same_course,
                                  by=NULL)

# Changes NAs (rows with C_time_exclude = "Exclude") to 1 so rows still keep number value in step below
course_data_exploded$H_student_adjuster_same_course <-
replace_na(course_data_exploded$H_student_adjuster_same_course, 1)

# Updates class weight for cases where students have multiple rows per class
course_data_exploded <- course_data_exploded %>%
  mutate(M_class_weight_adj_same_course = M_class_weight_exploded /
                                          H_student_adjuster_same_course)

# Check
table(course_data_exploded$H_student_adjuster_same_course, useNA = "ifany")

# Check
# Springfield: 0 rows
table(course_data_exploded$M_class_weight_exploded, useNA = "ifany")

# Filter to see courses where students have multiple rows per course and look at the number of students and rows for those courses
View(course_data_exploded %>%
  filter(H_student_adjuster_same_course > 1) %>%
  group_by(
    D_location_name,
    D_course_name) %>%
      summarise(stu_count = n_distinct(D_stu_id),
                row_count = n(),
                .groups = "drop"))

# ==============================================================================
# Part III: Flag Double Bookings ----
# ==============================================================================

# 3.1 Flag Student Double Bookings ----

# 3.1a Calculate Flag ----

# **What:**

# -   This code block flags students that have more than one course in a given term/rotation/period by identifying which students have above the expected class weight when you sum the class weight for all of their rows in a term/rotation/period. For example, if a student has a row for "Algebra I" and "PE" in a given term/rotation/period, each row would have your "standard" class weight and the sum across them would be your "standard" class weight times 2 (whereas for cases with only one course, the sum would still just be the "standard" class weight).

# -   Methods note: the reason M_class_weight_exploded is included in the group_by() is because M_class_weight_exploded tells you what the expected class weight is at every school for a given term/period/rotation and therefore needs to be passed through so that the actual sum of class weights in a given TPR can be compared to the expected sum

# **Check:**

# -   The first line printed by code block should have a number that is lower (but still close) to the number of rows in course_data_exploded

# -   The first table printed by the code block should have the vast majority of values at your "standard" class weight and some values above the "standard" class weight

# -   The second table printed by the code block should have all values as 0 or 1 and the number of values for 0 should exactly match the number of values at the "standard" weight value in the first table

# **Project Team Action:** do not change (run as is - and note that the code block might take a couple minutes to run)

# %% 3.1a Calculate Flag
# Creates the flagger
student_db_flagger <- function(data, most_updated_class_weight) {
  data %>%
  group_by(
    D_location_id,
    D_stu_id,
    C_term_exploded,
    C_rotation_exploded,
    C_period_exploded,
    M_class_weight_exploded) %>%
  summarize(H_class_weight_sum = sum({{most_updated_class_weight}})) %>%
  mutate(H_db_student_row = ifelse(H_class_weight_sum >
                                 M_class_weight_exploded, 1, 0))
  }

# Runs the flagger
student_db <- student_db_flagger(course_data_exploded,
                                 M_class_weight_adj_same_course)

# Check
nrow(student_db)

# Check
student_db %>%
  group_by(H_class_weight_sum) %>%
  count()

# Check
    # Springfield - 113 rows have a student double-booking
student_db %>%
  group_by(H_db_student_row) %>%
  count()

# Check number of students with at least one double booking
    # Springfield - 72 of 2,485 students (3%) have at least one double booking
student_db %>%
  group_by(D_stu_id) %>%
  summarize(H_db_student = max(H_db_student_row)) %>%
  group_by(H_db_student) %>%
  count()

# 3.1b Bring Flag Back ----

# **What:** Now that you feel good about the double bookings flag calculated in the previous code block, it's time to bring the flag into course_data_exploded and use it to mark rows that are double booked

# **Check:**

# -   Look at course_data_exploded in the environment (top right of RStudio) and confirm that the student double bookings flag is now a column in it

# -   The table printed by the code block shows the percentage of rows in course_data_exploded that have been marked as double booked for student - confirm that the results seem reasonable (as a rule of thumb, about 1-20% of rows are often double booked). Note that the number is given as percentage (so 0.01 means 1%).

# **Project Team Action:** do not change (run as is)

# %% 3.1b Bring Flag Back
# Creates merger function
student_db_merger <- function(data) {
  data %>%
  select(c(D_location_id,
           D_stu_id,
           C_term_exploded,
           C_rotation_exploded,
           C_period_exploded,
           H_db_student_row)) %>%
  right_join(course_data_exploded,
            by=NULL,
            multiple = "all")
}

# Runs the merger
course_data_exploded <- student_db_merger(student_db)

# Turns NAs to 0 (since those are students that get excluded from time analysis anyway)
course_data_exploded <- course_data_exploded %>%
  mutate(H_db_student_row = ifelse(is.na(H_db_student_row),
                                   0,
                                   H_db_student_row))

# Check
     # Springfield = 0.4% of rows have been flagged as double-booked for students
mean(course_data_exploded$H_db_student_row)

# 3.1c Explore Flagged Students ----

# **What:** Now that you have your double bookings flag up and running, it's worth looking at a double booked student in depth to ensure that the double bookings flag is working as expected

# **Check:**

# -   The output of the first table printed by the code block shows the number of students with at least one double booked row (1 is double booked, 0 is not) - look at the result and make sure that the distribution across the two values feels reasonable (again, about 1-20% of values with 1 is likely what you'd expect to see)
# -   For the sample student shown in the second table printed by the code block, look at the student rows and confirm that both double booked and non double booked rows are marked as you would expect. Of note, the table is sorted to have the double booked rows first and show "related" rows (rows with similar values for term/rotation/period) close to each other, which should make it easier to realize if each row has the value you'd expect in the double bookings flag.

# **Project Team Action:** do not change (run as is)

# %% 3.1c Explore Flagged Students
# Creates list of students with at least one double booked row
db_students_check <- course_data_exploded %>%
  group_by(D_stu_id) %>%
  summarize(H_db_student = max(H_db_student_row))

# Check
    # Springfield - 72 of 2,485 students have at least one double-booked row (3%)
table(db_students_check$H_db_student, useNA = "ifany")

# Identifies a double booked student
H_db_student_checked <- db_students_check %>%
  filter(H_db_student == 1) %>%
  slice_sample(n = 1) %>%
  pull(D_stu_id)

# Filters to rows of the identified student in course_data_exploded and limits the table to columns that let you know if double bookings flag works as expected
course_data_exploded %>%
  filter(D_stu_id == H_db_student_checked) %>%
  select(D_stu_id,
         C_class_id,
         C_term_exploded,
         C_rotation_exploded,
         C_period_exploded,
         D_course_name,
         C_course_subject,
         H_db_student_row) %>%
  arrange(desc(H_db_student_row),
          C_term_exploded,
          C_rotation_exploded,
          C_period_exploded)

# 3.2 Flag Teacher Double Bookings ----

# 3.2a Calculate Flag ----

# **What:** This code block follows a similar logic as student double bookings, but a key difference is that for teachers it doesn't matter if they have multiple rows per course (in fact, it's expected since they'd have a row per student in their class). Therefore, the first step is to create a single row per teacher per course - and only then see if a teacher has multiple rows for a given term/period/rotation.

# **Check:**

# -   The first line printed by code block should have a number that is significantly lower than the number of rows in course_data_exploded

# -   The first table printed by the code block should have the vast majority of values at your "standard" class weight and some values above the "standard" class weight

# -   The second table printed by the code block should have all values as 0 or 1 and the number of values for 0 should exactly match the number of values at the "standard" weight value in the first table

# **Project Team Action:** Do not change (run as is - and note that the code block might take a couple minutes to run)

# %% 3.2a Calculate Flag
# Action: update the values of class_weight and H_class_weight_sum_expected with your "standard" class weight
    # For Springfield:
        # Term has been exploded to semester (weight = 0.5)
        # Rotation: both schools have an A/B rotation (weight = 0.5)
        # Period: and all periods have the same weight (weight = 1)
        # So, expected class weight per teacher-term-period-rotation is 0.5 * 0.5 * 1 = 0.25.

teacher_db_flagger <- function(data, H_most_updated_class_weight) {
  data %>%
  filter(C_class_size_and_teacher_load_exclude == "Include") %>%
  group_by(D_location_id,
           D_employee_id,
           C_class_id,
           C_period_exploded,
           C_term_exploded,
           C_rotation_exploded,
           M_class_weight_exploded) %>%
  summarize(H_weight_per_class = max({{H_most_updated_class_weight}})) %>%
  group_by(D_location_id,
           D_employee_id,
           C_period_exploded,
           C_term_exploded,
           C_rotation_exploded,
           M_class_weight_exploded) %>%
  summarize(H_class_weight_sum_actual = sum(H_weight_per_class)) %>%
  mutate(H_db_teacher_row = ifelse(H_class_weight_sum_actual >
                                M_class_weight_exploded, 1, 0))
}

# Runs the flagger
teacher_db <- teacher_db_flagger(course_data_exploded,
                                 M_class_weight_adj_same_course)

# Check number of unique teacher-term-period-rotation (TPR) combinations
nrow(teacher_db)

# Check number of teacher-TPRs with different weights when summed
    # Springfield - because we expect each teacher-TPR to have a class weight of 0.25, anything higher than that is double-booked.
teacher_db %>%
  group_by(H_class_weight_sum_actual) %>%
  count()

# Check number of teacher-TPRs flagged as double-booked.
    # Springfield: 785 of 4,183 teacher-TPRs (19%) are double booked
teacher_db %>%
  group_by(H_db_teacher_row) %>%
  count()

# To manually check a specific teacher, add their employee id here and run
db_teacher_checked <- 16418

teacher_db %>%
  filter(D_employee_id == db_teacher_checked)

course_data_exploded %>%
  filter(D_employee_id == db_teacher_checked) %>%
  group_by(D_employee_id,
           C_term_exploded,
           C_rotation_exploded,
           C_period_exploded,
           C_class_id) %>%
  summarise(n())

# 3.2b Bring Flag Back ----

# **What:** Now that you feel good about the double bookings flag calculated in the previous code block, it's time to bring the flag into course_data_exploded and use it to mark rows that are double booked

# **Check:**

# -   Look at course_data_exploded in the environment (top right of RStudio) and confirm that the teacher double bookings flag is now a column in it

# -   The table printed by the code block shows the percentage of rows in course_data_exploded that have been marked as double booked for teacher - confirm that the results seem reasonable (as a rule of thumb, about 1-20% of rows are often double booked). Note that the number is shown as a percentage (e.g. 0.01 means 1%)

# **Project Team Action:** do not change (run as is)

# %% 3.2b Bring Flag Back
# Creates the merger
teacher_db_merger <- function(data) {
  data %>%
  select(c(D_location_id,
           D_employee_id,
           C_term_exploded,
           C_rotation_exploded,
           C_period_exploded,
           H_db_teacher_row)) %>%
  right_join(course_data_exploded,
            by = NULL,
            multiple = "all")
}

# Runs the merger
course_data_exploded <- teacher_db_merger(teacher_db)

# Turns NAs into 0 (since they get taken out of analysis anyway from teacher load exclusion)
course_data_exploded <- course_data_exploded %>%
  mutate(H_db_teacher_row = ifelse(is.na(H_db_teacher_row),
                                   0,
                                   H_db_teacher_row))

# Check % of rows that are flagged as double-booked for teachers
    # Springfield - 23%
mean(course_data_exploded$H_db_teacher_row)

# Check courses and periods with double booked teachers
course_data_exploded %>%
  filter(H_db_teacher_row == 1) %>%
  group_by(
    D_location_name,
    D_course_name,
    C_period_exploded,
    C_period_type) %>%
  count() %>%
  arrange(desc(n))

# 3.2c Explore Flagged Teachers ----

# **What:** now that you have your double bookings flag up and running, it's worth looking at a double booked teacher in depth to ensure that the double bookings flag is working as expected

# **Check:**

# -   The output of the first table printed by the code block shows the number of teachers with at least one double booked row (1 is double booked, 0 is not) - look at the result and make sure that the distribution across the two values feels reasonable (again, about 1-20% of values with 1 is likely what you'd expect to see)
# -   For the sample teacher shown in the second table printed by the code block, look at the rows and confirm that both double booked and non double booked rows are marked as you would expect. Of note, the table is sorted to have the double booked rows first and show "related" rows (rows with similar values for term/rotation/period) close to each other, which should make it easier to realize if each row has the value you'd expect in the double bookings flag.

# **Project Team Action:** do not change (run as is)

# %% 3.2c Explore Flagged Teachers
# Creates list of teachers with at least one double booked row
db_teachers_check <- course_data_exploded %>%
  filter(C_class_size_and_teacher_load_exclude == "Include") %>%
  group_by(D_employee_id) %>%
  summarize(H_db_teacher = max(H_db_teacher_row))

# Checks number of teachers that have at least one double-booked row
    # Springfield - 103 of 226 of teachers (46%)
table(db_teachers_check$H_db_teacher, useNA = "ifany")

# Identifies a double booked teacher
H_db_teacher_checked <- db_teachers_check %>%
  filter(H_db_teacher == 1) %>%
  slice_sample(n = 1) %>%
  pull(D_employee_id)

# You can also manually check a teacher by including this code instead
H_db_teacher_checked <- 16418

# Filters to rows of the identified teacher in course_data_exploded and limits the table to columns that let you know if double bookings flag works as expected
course_data_exploded %>%
  filter(D_employee_id == H_db_teacher_checked) %>%
  group_by(D_employee_id,
           C_class_id,
           D_course_name,
           C_term_exploded,
           C_period_exploded,
           C_rotation_exploded,
           H_db_teacher_row) %>%
  summarise(stu_count = n_distinct(D_stu_id)) %>%
  arrange(desc(H_db_teacher_row),
          C_term_exploded,
          C_period_exploded,
          C_rotation_exploded)

# 3.3 Filter File ----

# Now that the data has student and teacher double bookings flagged, the rows for classes with at least one student or teacher double booking will be kept and the rest of the rows will be taken out. This removal of rows for classes with no double bookings is done to make the file lighter and does not mean that data is being lost - in the "Step 08: DB Re-merge" script these rows will reappear in their un-exploded version and appended to the exploded double booked rows.

# 3.3a Flag for classes with double bookings ----

# **What:** The first step is to create the flag for class with double booking based on the student and teacher db flags that were created in the previous parts of this script

# **Check:** The outputs of the code block show how many rows in the data are for a class with a teacher double booking (first output), a student double booking (second output), and for either a teacher or student double booking (third output).

# **Project Team Action:** Do not change (run as is)

# %% 3.3a Flag for classes with double bookings
# Creates flag for class with double booking
course_data_exploded <- course_data_exploded %>%
  group_by(C_class_id) %>%
  mutate(H_class_with_teacher_db = max(H_db_teacher_row),
         H_class_with_student_db = max(H_db_student_row),
         H_class_with_db = max(H_class_with_teacher_db,
                               H_class_with_student_db))

# Check how many rows have a teacher with any teacher-level double-bookings
    # Springfield - 23%
course_data_exploded %>%
  group_by(H_class_with_teacher_db) %>%
  summarize(count_rows = length(H_new_record_id)) %>%
  ungroup() %>%
  mutate(percent_rows = count_rows / sum(count_rows))

# Check how many rows have a student with any student-level double-bookings
    # Springfield - 4%
course_data_exploded %>%
  group_by(H_class_with_student_db) %>%
  summarize(count_rows = length(H_new_record_id)) %>%
  ungroup() %>%
  mutate(percent_rows = count_rows / sum(count_rows))

# Check how many rows have a student with a student-level double booking OR a teacher with a teacher-level double booking
    # Springfield - 26%
course_data_exploded %>%
  group_by(H_class_with_db) %>%
  summarize(count_rows = length(H_new_record_id)) %>%
  ungroup() %>%
  mutate(percent_rows = count_rows / sum(count_rows))

# 3.3b Creating the Filtered File ----

# **What:** Now that you feel good about the flag for class with double bookings, your data will be split into the classes with double bookings and the classes without double bookings. The classes with double bookings will be passed through the full resolutions script whereas the classes without double bookings will be brought back in their un-exploded version in the "Step 07: Unification" script.

# **Check:** Look at course_schedule_exploded and ensure that the number of rows has shrunk significantly (and more precisely, that it matches the number of rows flagged as a class with double booking in the previous code block)

# **Project Team Action:** Do not change (run as is)

# %% 3.3b Creating the Filtered File
# Filter for rows with a double-booking at either the student- or teacher-level
# CHECK that this equlas the number of rows flagged as a class w/ double-booking in the previous code block
course_data_exploded_filter <- course_data_exploded %>%
  filter(H_class_with_db == 1)

# 3.3c Final Pass at % DB for student and teacher ----

# **What:** For ease of comparison with the results of the resolutions script, this code block prints the percent of rows in the data passed to the resolutions script that are marked as DB for student or teacher

# **Check:** Keep these numbers in mind since they'll help make meaning of results in the resolutions script

# %% 3.3c Final Pass at % DB for student and teacher
# This checks the percentage of rows with each type of double-booking AFTER you've filtered for rows that have EITHER student- or teacher-level double-bookings
  # Student DB %
  # Springfield: 1%
mean(course_data_exploded_filter$H_db_student_row)

  # Teacher DB %
  # Springfield: 89%
mean(course_data_exploded_filter$H_db_teacher_row)

# 3.4 Export Data ----

# **What:** This code block creates the CSV with the course schedule data in its flagged state. Bring this CSV to team for review.

# **Check:** Go to your SharePoint project folder and make sure you have the CSV

# **Project Team Action:** Maybe change (see guidance in code block)

# %% 3.4 Export Data
# Save course schedule data for next script
ers_write_sharepoint(
  data = course_data_exploded_filter,
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/2. Processed Data/05_cs_data_post_db_flaggings.csv")

# ==============================================================================
# Bonus: Celebrate! ----
# ==============================================================================

# And with that, you're done with this script!!! Here's a rabbit to celebrate :)

# %% Bonus: Celebrate!
## (\_/) ##
## (o.o) ##
## (___) ##
