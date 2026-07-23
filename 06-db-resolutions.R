# 06-db-resolutions
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

# Load processed data
course_data_flagged <- ers_read_sharepoint(
  folder_path = raw_data_folder_path,
  file_name_with_extension =
    "/2. Processed Data/05_cs_data_post_db_flaggings.csv")


# Part II: Class Double Bookings Resolution ----

# This part of the script allows you to flag which courses have 100% of their students double booked and adjust the class weight of each of these classes accordingly. See Tettra for more guidance.

# 2.1 Identify 100% Double Booked Classes ----

# What: The first step of class resolutions is to identify which classes have 100% of their students double booked (which occurs if all the students in a class are also enrolled for another class - or classes - in that same term/rotation/period). The code block below does exactly this. Note that if a student has multiple rows for the same course, the extra rows get counted for both the double booked count and the total count and thus do not mess up the 100% flagger. For example, if a course has 21 rows with 20 total students and one student repeating rows, the total student count would be the "incorrect" 21 but the double booked student count would also be the "incorrect" 21 - which evens out.

# Check:

# - First line of the code block output is the number of rows in class_double_bookings (which has a single row by exploded course id) divided by the number of rows in course_data_flagged. As a rule of thumb, the number should be close to 1/20 (since most classes are expected to have around 20 students in them and therefore 20 student rows).

# - Second line of the code block output is the number of classes with 100% double booked rows. This number should match the height of the 100% bar in the graph for the next check.

# - The first graph printed by the code block shows the number of classes (y axis) with a particular number of double booked students (x axis). Classes with zero students double booked happen because the class has the teacher double booked in at least one row (which is why they still show up in the double booked data). This graph serves largely to provide context for the next graph (e.g. if a ton of your classes have 20+ students double booked, it's likely that these classes have all their students double booked and will therefore appear as 100% in the second graph).

# - The second graph printed by the code block shows the distribution of classes by their percentage of double booked students. You should likely see at least some classes in the 100% bar since it's fairly common for course data to have at least some fully double booked classes (e.g. Science and Lab).

# - Lastly, confirm that H_percent_db_by_class is now included in course_data_flagged by looking at the data in the environment (top right corner of RStudio).

# Project Team Action: do not change (run as is)

# %% 2.1 Identify 100% Double Booked Classes
# Creates helper variable with value of 1 for all rows
course_data_flagged$row_count <- 1

# Marks the percentage of double booked students in a class and flags classes with 100% double booked students
# Note: the method to flag 100% DB classes is to count the number of DB students in the class and then divide by the total number of students in the class.
class_double_bookings <- course_data_flagged %>% 
  group_by(C_class_id) %>% 
  summarize(H_count_student_rows_by_class = length(H_new_record_id),   # Count of rows by class
            H_db_count_by_class = sum(H_db_student_row),   # Count of double-booked student rows by class
            H_percent_db_by_class = H_db_count_by_class / H_count_student_rows_by_class,   # Percentage of double-booked student rows by class
            H_db_100p_class = ifelse(H_percent_db_by_class == 1, 1, 0)) # Flag for 100% double booked class

# Merge
course_data_flagged <- left_join(course_data_flagged,
                                 class_double_bookings,
                                 by=NULL)

# Sanity check
nrow(class_double_bookings) / nrow(course_data_flagged)

# Check how many classes are 100% double-booked
sum(class_double_bookings$H_db_100p_class)

# Check percentage of classes that are 100% double-booked
    # Springfield: 0.6% of classes are 100% double booked
mean(class_double_bookings$H_db_100p_class)

# Check
class_double_bookings %>%
  ggplot() +
  geom_histogram(mapping = aes(x = H_db_count_by_class), binwidth = 1) +
                 labs(title = "Count of double booked students by class id",
                      x = "Count of double booked students in a given class",
                      y = "Number of classes") +
                 scale_y_continuous(labels = function(x)
                                    format(x, scientific = FALSE))

# Check
class_double_bookings %>%
  ggplot() +
  geom_histogram(mapping = aes(x = H_percent_db_by_class), binwidth = 0.01) +
                 labs(title = "Percentage of students double booked by class id",
                      x = "percentage of double booked students in a given class",
                      y = "number of classes") +
                 scale_y_continuous(labels = function(x)
                                    format(x, scientific = FALSE))


# 2.2 Count of 100% DB classes by Student ----

# What: This section moves from the class level to the student + time block level.
# - The main question is: For each student, in each term/period/rotation, how many of their classes are 100% double-booked classes?
# This matters because the next step will use that number to split the student’s weight across those classes.

# Check:

# - First and second code output show the distribution of values for the student count variable (number of 100% DB classes for student) and the TPR 100% variable (% of classes in that TPR that are 100% DB for the student).
# The second code output shows how often the student-time-block is:
# 0 = none of the classes are 100% double-booked
# 0.5 = half are 100% double-booked
# 1 = all are 100% double-booked (ie. every class this student has in that term/rotation/period is a 100% double-booked class)

# - In the third code output you should see no rows listed with 1 for both variables (since that would mean the student/TPR combo has been flagged as double booked despite having only one class show up as 100% DB AND all classes in that student/TPR combo being 100% double booked). This pattern would be unexpected so flag to Jenny if you have it (especially if you have lots of rows in it).

# - Cases in which a single term/period/rotation has a combination of both 100% DB and non-100% DB classes are tricky to deal with (see Tettra guidance for more detail on these "part 100% DB TPRs"). The code block below accounts for these edge cases by changing the value of the student counter back to 1 for any TPRs that had a mix of 100% and non-100% classes. Fourth and fifth output shows the count of rows in each value of the student counter before and after the adjustment - you should see the count in 1 go up (since the edge cases described above are shifted to value of 1) or at least stay unchanged.

# Project Team Action: do not change (run as is)

# %% 2.2 Count of 100% DB classes by Student
# Create a helper table where each row represents one student in one specific term / rotation / period (TPR) and counts the number of 100% DB classes they have in that TPR as well as the percentage of their classes in that TPR that are 100% DB
class_100p_flagger <- course_data_flagged %>%
  group_by(D_stu_id,
           C_term_exploded,
           C_rotation_exploded,
           C_period_exploded) %>% 
  summarize(H_student_count_of_100p_classes = sum(H_db_100p_class),
                # Number of 100% double-booked classes the student has in that exact time block (TPR). This answers: How many of this student’s classes in this time block are part of a fully double-booked class?
            H_tpr_100p_percent = mean(H_db_100p_class))
                # The percent of the student’s rows in that time block (TPR) that are 100% double-booked classes. This answers: Are ALL of the student’s classes in this time block 100% double-booked classes, or only some of them?

# Check the count distribution of 100% double-booked classes
class_100p_flagger %>% 
  group_by(H_student_count_of_100p_classes) %>% 
  count()

# Check the percent of the TPR that is 100% double-booked
class_100p_flagger %>% 
  group_by(H_tpr_100p_percent) %>% 
  count()

# Check
class_100p_flagger %>% 
  group_by(H_student_count_of_100p_classes, H_tpr_100p_percent) %>% 
  count()


# %% 2.2 Count of 100% DB classes by Student (2)
# Count of rows in course_data_flagged before the merge
nrow(course_data_flagged)

# Merges the 100% DB flag into course_data_flagged
course_data_flagged <- left_join(
  course_data_flagged,
  class_100p_flagger,
  by = NULL)

# Count of rows in course_data_flagged after the merge
nrow(course_data_flagged)

# Checks how many rows have each value of H_student_count_of_100p_classes before the script changes anything
course_data_flagged %>% 
  group_by(H_student_count_of_100p_classes) %>% 
  count()


# %% 2.2 Count of 100% DB classes by Student (3)
# For each row, if the student’s full term/period/rotation is not entirely made up of 100% double-booked classes, set the flag to 1. That way, we'll only split the weight when the student’s entire time block is fully 100%-double-booked. This is to account for edge cases in which a student has some classes that are 100% DB and some that are not - in these cases, we don't want to split the student's weight across the 100% DB classes since they also have non-100% DB classes in that time block.
# In summary: If this is a clean fully-100%-double-booked time block, count the number of overlapping classes. Otherwise, leave as 1 so it doesn't impact the class weight in Step 2.3.
course_data_flagged <- course_data_flagged %>% 
  mutate(H_student_count_of_100p_classes = ifelse(H_tpr_100p_percent < 1,
                                                1,
                                                H_student_count_of_100p_classes))
         
# Check - count of rows for values in counter (post adjustment)
course_data_flagged %>% 
  group_by(H_student_count_of_100p_classes) %>% 
  count()


# 2.3 Adjust Class Weight based on 100% DB classes ----

# What: This code block uses the student counter to adjust the class weight for each row

# Check: The first output shows the distribution of values in student_class_weight_post class. Check that the distribution looks reasonable based on the number of 100% DB classes outlined in section 2.1 and the original distribution of scores before the class double bookings adjustment (given in second output for reference). For example, if you had 1000 students with a class weight of 0.5 but 500 of those students had two 100% DB classes per term/period/rotation, then you should now have 500 students at 0.5 and 500 students at 0.25 (in addition to however many students you already had at 0.25 before the 100% DB class adjustment).

# Project Team Action: do not change (run as is)

# %% 2.3 Adjust Class Weight based on 100% DB classes
# Creates student adjusted class weight that adjusts for 100% double-booked classes.
# In other words: If a student has multiple fully double-booked classes in the same term/period/rotation, split the class weight across those classes.
course_data_flagged$M_class_weight_post_class <- course_data_flagged$M_class_weight_adj_same_course / course_data_flagged$H_student_count_of_100p_classes

# Check the distribution of the old/input weights: what weights existed before this class double-booking adjustment?
course_data_flagged %>% 
  group_by(M_class_weight_adj_same_course) %>% 
  count()

# Check the distribution of the new adjusted weights: what weights exist after the adjustment?
    # If the post-class weights includes lower values, that means the adjustment is working.
course_data_flagged %>% 
  group_by(M_class_weight_post_class) %>% 
  count()


# 2.4 Rerunning Student Double Bookings ----

# 2.4a Calculate Flag ----

# What: This code block reruns the student double bookings flagger developed in the flaggings script - see that script for details on how the flagger works. The question it answers is: after splitting weights for 100% double-booked classes, are students still double-booked in any term/period/rotation?

# Check:

# - The first line printed by code block should have a number that is lower (but still close) to the number of rows in course_data_exploded

# - The first table printed by the code block should have the vast majority of values at the "standard" class weight and some values above the "standard" class weight

# - The second table printed by the code block should have all values as 0 or 1 and the number of values for 0 should exactly match the number of values at the "standard" weight value in the first table

# - Lastly, the third table printed shows the number of rows at 0 and 1 for the previous version of the student DB flag and you should see that the count of 1s was higher in this flag than in the flag we now have after class resolutions (which was shown in the second table)

# Project Team Action: do not change (run as is - and note that the code block might take a couple minutes to run)

# %% 2.4a Calculate Flag
# Creates the flagger - method note: 0.0001 gets added to expected class weight so that cases that got divided into infinite decimals during the same course adjustment now don't unintendedly go above the expected class weight
student_db_flagger <- function(data,
                               most_updated_class_weight,
                               most_updated_db_student_row) {
  data %>%
#  filter(C_course_time_exclude == "Include") %>% 
  group_by(D_location_id,
           D_stu_id,
           C_term_exploded,
           C_rotation_exploded,
           C_period_exploded,
           M_class_weight_exploded,
           #note - adding variable above doesn't break the grouping since all classes within a TPR are expected to have the same class weight
           {{most_updated_db_student_row}}) %>% 
  summarize(class_weight_sum = sum({{most_updated_class_weight}})) %>%
  mutate(H_updated_db_student_row = ifelse(class_weight_sum >
                                 M_class_weight_exploded + 0.0001,
                                 1,
                                 0))
  }

# Run the flagger using the post-class weight
student_db_post_class <- student_db_flagger(course_data_flagged,
                                            M_class_weight_post_class,
                                            H_db_student_row)

# Rename the db flag so that it's identifiable as "student double-booking flag after the class-level fix" in future steps
student_db_post_class <- student_db_post_class %>% 
  rename(H_db_student_row_post_class = H_updated_db_student_row)

# Check rows in student_db_post_class, which should be lower than course_data_flagged because the flagger filters to only include rows with course time included and also groups by class weight.
nrow(student_db_post_class)

# Check # of rows by class weight
# Rows above the class weight you'd expect are still double-booked.
student_db_post_class %>% 
  group_by(class_weight_sum) %>% 
  count()

# Check # of rows with a student-level double-booking BEFORE resolving 100% DBs
student_db_post_class %>% 
  group_by(H_db_student_row) %>% 
  count()

# Check # of rows with a student-level double-booking AFTER resolving 100% DBs
# This helps answer: Did the class-level weight adjustment bring some students back down to the expected class weight?
# This number should be LOWER than the output above.
student_db_post_class %>% 
  group_by(H_db_student_row_post_class) %>% 
  count()


# 2.4b Bring Flag Back ----

# What: now that you feel good about the double bookings flag calculated in the previous code block, it's time to bring the flag into course_data_flagged and use it to mark rows that are still double booked even after the class resolutions

# Check:

# - Look at course_data_flagged in the environment (top right of RStudio) and confirm that the new student double bookings flag is now a column in it

# - The first table printed by the code block shows the percentage of rows in course_data_flagged that have been marked as double booked for student - confirm that the results seem reasonable. Note that the number is given as percentage (so 0.01 means 1%).

# - The two numbers given in the second output are the % of rows marked as a student DB before and after the class resolutions - you should see that the % went down at least somewhat.

# Project Team Action: Do not change (run as is)

# %% 2.4b Bring Flag Back
# Creates merger function to bring the student double-booking result from the student_db_post_class summarized table and merges it back onto the full row-level dataset.
student_db_merger <- function(data,
                              H_db_student_row_from_flagger_above) {
  data %>% 
  select(c(D_location_id,
           D_stu_id,
           C_term_exploded,
           C_rotation_exploded,
           C_period_exploded,
           {{H_db_student_row_from_flagger_above}})) %>% 
  right_join(course_data_flagged,
             by = NULL,
             multiple = "all")
}

# Runs the merger
course_data_flagged <- student_db_merger(student_db_post_class,
                                         H_db_student_row_post_class)

# Turns the NAs into 0s (since they're due to exclusion from time analysis)
course_data_flagged <- course_data_flagged %>% 
  mutate(H_db_student_row_post_class = ifelse(is.na(H_db_student_row_post_class),
                                              0,
                                              H_db_student_row_post_class))
  
# Check original student double-booking rate.
mean(course_data_flagged$H_db_student_row)

# Check post-class student double-booking rate, which should be LOWER than the original mean.
mean(course_data_flagged$H_db_student_row_post_class)


# 2.5 Rerunning Teacher Double Bookings ----

# 2.5a Calculate Flag ----

# What: This code block runs the teacher DB flagger - refer to the flaggings script for a fuller description of the code and process. After the class-level fixes, is a teacher still assigned to more than one class’s worth of teaching in the same term/period/rotation?
# - For students, the script summed student class weights within a time block.
# - For teachers, it is slightly different because each class has many student rows. The script does not want to sum every student row, or it would wildly overcount the teacher’s load.

# So the teacher logic is:
# - Get one weight per teacher/class/time block.
# - Sum those class weights by teacher/time block.
# - If that total is greater than the expected class weight, flag the teacher as double-booked.

# Check:

# - The first line printed by code block should have a number that is significantly lower than the number of rows in course_data_flagged

# - The first table printed by the code block should have the vast majority of values at your "standard" class weight and some values above the "standard" class weight

# - The second table printed by the code block should have all values as 0 or 1 and the number of values for 0 should exactly match the number of values at the "standard" weight value in the first table

# - Lastly, the third table printed shows the number of rows at 0 and 1 for the previous version of the teacher DB flag and you should see that the count of 1s was higher in this flag than in the flag we now have after class resolutions (which was shown in the second table)

# Project Team Action: Do not change (run as is - and note that the code block might take a couple minutes to run)

# %% 2.5a Calculate Flag
# Creates teacher flagger function - method note: 0.0001 gets added to expected class weight so that cases that got divided into infinite decimals during the same course adjustment now don't unintendedly go above the expected class weight
teacher_db_flagger <- function(data,
                               most_updated_class_id,
                               most_updated_class_weight,
                               most_updated_db_teacher_row) {
  data %>% 
  filter(
    C_class_size_and_teacher_load_exclude == "Include") %>% 
  group_by(D_location_id,
           D_employee_id,
           {{most_updated_class_id}},
           C_period_exploded,
           C_term_exploded,
           C_rotation_exploded,
           M_class_weight_exploded,
           {{most_updated_db_teacher_row}}) %>% 
  summarize(weight_per_class = max({{most_updated_class_weight}})) %>% # Take one class weight for each teacher/class/time block
  group_by(D_location_id,
           D_employee_id,
           C_period_exploded,
           C_term_exploded,
           C_rotation_exploded,
           M_class_weight_exploded,
           {{most_updated_db_teacher_row}}) %>% 
  summarize(class_weight_sum_actual = sum(weight_per_class)) %>% # Sum the class weights for that teacher/time block
  mutate(H_updated_db_teacher_row = ifelse(class_weight_sum_actual >
                                     M_class_weight_exploded + 0.0001,
                                     1,
                                     0))
}

# Runs the flagger
teacher_db_post_class <- teacher_db_flagger(course_data_flagged,
                                            C_class_id,
                                            M_class_weight_post_class,
                                            H_db_teacher_row)

# Renames the db flag to identify it as the "teacher double-booking flag after class-level resolution" in future scripts
teacher_db_post_class <- teacher_db_post_class %>% 
  rename(H_db_teacher_row_post_class = H_updated_db_teacher_row)

# Changes NAs to 0 since it's for cases that would get excluded from teacher load analysis anyway
teacher_db_post_class <- teacher_db_post_class %>% 
  mutate(H_db_teacher_row_post_class = ifelse(is.na(H_db_teacher_row_post_class),
                                              0,
                                              H_db_teacher_row_post_class))

# Check number of teacher-time blocks. This should be much fewer than course_data_flagged, because this is no longer student-row-level data.
nrow(teacher_db_post_class)

# Checks number of teacher-TPRs by class weight. You should see the majority of values at your "standard" class weight and some values above that.
teacher_db_post_class %>% 
  group_by(class_weight_sum_actual) %>% 
  count()

# Check number of teacher-TPRs flagged as double-booked BEFORE class resolutions.
teacher_db_post_class %>% 
  group_by(H_db_teacher_row) %>% 
  count

# Check number of teacher-TPRs flagged as double-booked AFTER class resolutions.
# The number of 0 should exactly match the number with your "standard" class weight from the class_weight_sum_actual summary above.
teacher_db_post_class %>% 
  group_by(H_db_teacher_row_post_class) %>% 
  count()


# 2.5b Bring Flag Back ----

# What: Now that you feel good about the double bookings flag calculated in the previous code block, it's time to bring the flag into course_data_flagged and use it to mark rows that are double booked

# Check:

# - Look at course_data_flagged in the environment (top right of RStudio) and confirm that the teacher double bookings flag is now a column in it

# - The table printed by the code block shows the percentage of rows in course_data_exploded that have been marked as double booked for teacher - confirm that the results seem reasonable. Note that the number is shown as a percentage (e.g. 0.01 means 1%)

# - The two numbers given in the second output are the % of rows marked as a teacher DB before and after the class resolutions - you should see that the % went down at least somewhat

# Project Team Action: Do not change (run as is)

# %% 2.5b Bring Flag Back
# Creates the merger
teacher_db_merger <- function(data,
                              H_db_teacher_row_from_flagger_above) {
  data %>% 
  select(c(D_location_id,
           D_employee_id,
           C_term_exploded,
           C_rotation_exploded,
           C_period_exploded,
           {{H_db_teacher_row_from_flagger_above}})) %>% 
  right_join(course_data_flagged,
            by = NULL,
            multiple = "all")
}

# Runs the merger
course_data_flagged <- teacher_db_merger(teacher_db_post_class,
                                         H_db_teacher_row_post_class)

# Turns into 0 the cases where Teacher Load Exclude == "Exclude"
course_data_flagged <- course_data_flagged %>%
  mutate(H_db_teacher_row_post_class = ifelse(is.na(H_db_teacher_row_post_class),
                                              0,
                                              H_db_teacher_row_post_class))

# Check - % rows with teacher double-booking BEFORE the class resolutions.
mean(course_data_flagged$H_db_teacher_row)

# Check - % rows with teacher double-booking AFTER the class resolutions; you should see the percentage go DOWN at least somewhat.
mean(course_data_flagged$H_db_teacher_row_post_class)


# Part III: Teacher Double Bookings Resolution ----

# 3.1 Marking PE/Art and Class Size ----

# What: The first step of teacher DB is to identify classes as PE/Art vs Other and calculate count of students since these two pieces of information will be used to decide if teachers get combined sections or need to be removed from analysis altogether. See Tettra for a fuller explanation of the rationale behind this. This section narrows in on those remaining problem cases and summarizes them so the script can decide: Is this teacher double-booking something we can treat as a combined class, or is it too large/unrealistic?

# Check:

# - The first three outputs of the code block shows the distribution of courses as PE/Art vs Other by subject/teacher/student, by teacher/student, and by teacher alone. You would expect to see the number of PE/Art increasing from one step to the next (since a single PE/Art will turn the whole group into PE/Art).

# - Fourth output of the code block is a graph showing the student count of each teacher/TPR combo, and you would expect to see a decent number of classes with 35 students or above (since all cases would have at least two classes due to being DB in some fashion, which in turn often leads to high student counts - and do remember that the classes with exact student roster replicas were resolved in class double bookings so they don't show up here)

# Project Team Action: Do not change (run as is)

# %% 3.1 Marking PE/Art and Class Size
# Aggregate data so each teacher/student/subject combination only has one row
teachers_helper <- course_data_flagged %>% 
  filter(H_db_teacher_row_post_class == 1) %>% 
  group_by(D_location_name,
           D_stu_id,
           D_employee_id,
           C_course_subject,
           C_term_exploded,
           C_period_exploded,
           C_rotation_exploded) %>% 
  summarize()

# Mark rows that are PE or Art
teachers_helper <- teachers_helper %>% 
  mutate(PE_Art = ifelse(C_course_subject == "PE/Health" | C_course_subject == "Art/Music",
                         1,
                         0))

# Check
scales::percent(mean(teachers_helper$PE_Art, na.rm = TRUE), accuracy = 0.1)

# Take subject out of grouping to mark if student/teacher/TPR combo has Art or PE
teachers_helper <- teachers_helper %>% 
  group_by(D_location_name,
           D_stu_id,
           D_employee_id,
           C_term_exploded,
           C_period_exploded,
           C_rotation_exploded) %>% 
  summarize(row_count = 1, # Each student/teacher/TPR combo gets a count of 1 so that we can sum the number of students later
            PE_Art = max(PE_Art))

# Check
scales::percent(mean(teachers_helper$PE_Art, na.rm = TRUE), accuracy = 0.1)

# Take students out of grouping to calculate count of students per teacher/TPR combo
teachers_helper <- teachers_helper %>% 
  group_by(D_location_name,
           D_employee_id,
           C_term_exploded,
           C_period_exploded,
           C_rotation_exploded) %>% 
  summarize(student_count = sum(row_count), # Sum the number of students in each teacher/TPR combo
            PE_Art = max(PE_Art))

# Check
scales::percent(mean(teachers_helper$PE_Art, na.rm = TRUE), accuracy = 0.1)

# Check
ggplot(teachers_helper) +
  geom_histogram(mapping = aes(x = student_count))


# 3.2 Creating Consolidated Class ID and Combined Flag ----

# What: Now that the teachers helper data frame is at the level of just teacher and TPR, several classes are included within a single row. As such, these classes are said to be "consolidated" and a new class id is needed. Some of these consolidated classes will not end up consolidated by the end of the process due to being too large or not being PE/Art, but for now all teacher and TPR combos will get a distinct consolidated class id. Note that the consolidated class id is assigned starting with 1 million.
# This code is creating two things for the remaining teacher double-booking cases:
# - A new fake/consolidated class ID for each teacher + TPR combo
# - A flag for teacher/TPR combos that are probably too large to count as believable consolidated classes

# Check:

# - Check the consolidated class id variable in the environment (top right corner of RStudio) to make sure that it works as you would expect - specifically, that it starts at 1M and increases by 1 per row

# - The two numbers printed by the code block are % of teacher/TPR combos with over 35 students and % of teacher/TPR combos with over 35 students AND no PE/Art courses. This second result is the % of teacher/TPR cases that will eventually get teachers flagged for exclusion.

# Project Team Action: Do not change (run as is)

# %% 3.2 Creating Consolidated Class ID and Combined Flag
# For now, treat each teacher/TPR combo as one consolidated class and give it a new class id
teachers_helper$row_number <- c(1:nrow(teachers_helper))
teachers_helper$C_class_id_consolidated <- teachers_helper$row_number + 1000000

# Flag if student count for a teacher/TPR combo is above 35
teachers_helper <- teachers_helper %>% 
  mutate(student_count_above_35 = ifelse(student_count > 35, 1, 0))

# Check the percent above 35 students
mean(teachers_helper$student_count_above_35)

# Flag large non-PE/Art teacher/TPRs that are likely not believable combined classes and should get excluded from teacher load analysis altogether
    # Large consolidated PE/Art classes are allowed to stay. Large consolidated non-PE/Art classes are considered too unrealistic and will later be flagged for exclusion.
teachers_helper <- teachers_helper %>% 
  mutate(PE_Art_student_count_unified_flag = ifelse(student_count_above_35 == 1 &
                                             PE_Art == 0,
                                             1,
                                             0))

# Check the percent that will be excluded from teacher-load analysis
mean(teachers_helper$PE_Art_student_count_unified_flag)


# 3.3 Creating DB Flag by Teacher ----

# What: Now that every teacher/TPR combo is marked for whether they have above 35 students and don't have PE/Art, the next step is to take out TPR from the grouping so that we know if a teacher has either of these criteria in any of their TPRs. A single TPR that fulfills both of these criteria will make the teacher get excluded from our analyses. Once this is done, the flag is brought back to the original data and the teacher load metric is updated accordingly.

# Check:

# - First output shows proportion of teachers that got fixed - the 0s are now not double booked and the 1s will have to get excluded from analysis altogether (all their rows will get marked off in the teacher load variable)
# - New class id should have no NAs (see table printed out to confirm)
# - Teacher load exclude should have no NAs (see table printed out to confirm) and value of exclude should be roughly proportional than the number of teachers with value of 1 in the teacher DB table printed in the first output
# - Look at course_data_flagged in the environment (top right corner of RStudio) and make sure that the columns from teacher_helper are now included in it

# Project Team Action: Do not change (run as is)

# %% 3.3 Creating DB Flag by Teacher
# Take TPR out of grouping and mark double booked teachers.
# The logic is: If a teacher has even one problematic TPR, mark the whole teacher for exclusion.
teachers_exclusion_value <- teachers_helper %>% 
  group_by(D_employee_id) %>% 
  summarize(H_teachers_to_exclude = max(PE_Art_student_count_unified_flag))

# CHECK 1: number of teachers to exclude
    # Springfield - only 1 teacher to exclude
teachers_exclusion_value %>% 
  group_by(H_teachers_to_exclude) %>% 
  count()

# Bring flag back to teacher_helper
teachers_helper <- teachers_helper %>% 
  left_join(teachers_exclusion_value,
            by = NULL)

# Limit teacher_helper to just matching variables, new class id, and DB flag
teachers_helper <- teachers_helper %>% 
  select(D_employee_id,
         C_term_exploded,
         C_rotation_exploded,
         C_period_exploded,
         C_class_id_consolidated,
         H_teachers_to_exclude)

# Bring flag back to course_data_flagged and mark NAs as 0 for the teacher DB flag
course_data_flagged <- course_data_flagged %>% 
  left_join(teachers_helper,
            by = NULL) %>% 
  mutate(
    H_teachers_to_exclude = ifelse(
      is.na(H_teachers_to_exclude), # If this row was NOT part of the problematic teacher double-booking helper table, do not exclude it.
      0,
      H_teachers_to_exclude))
         
# CHECK 2: Counts rows to exclude based on new H_teachers_to_exclude flag
course_data_flagged %>% 
  group_by(H_teachers_to_exclude) %>% 
  count()

# Create the final post-teacher class id (consolidated for relevant rows, old one for others)
course_data_flagged <- course_data_flagged %>% 
  mutate(C_class_id_post_teachers = ifelse(is.na(C_class_id_consolidated),
                                         C_class_id,
                                         as.character(C_class_id_consolidated)))

# CHECK 3: Check whether post-teacher class ids are missing. This should be ZERO since all classes should either have an original class id or a consolidated class id.
sum(is.na(course_data_flagged$C_class_id_post_teachers))

# CHECK 4: Count rows flagged for exclusion BEFORE new load-exclusion flag.
course_data_flagged %>% 
  group_by(C_class_size_and_teacher_load_exclude) %>% 
  count()

# Update the teacher load metric.
course_data_flagged <- course_data_flagged %>% 
  mutate(C_class_size_and_teacher_load_exclude = ifelse(C_class_size_and_teacher_load_exclude == "Exclude" | H_teachers_to_exclude == 1,
                                                        "Exclude",
                                                        "Include"))
# Move new class ids all the way to the left.
course_data_flagged <- course_data_flagged %>% 
  select(C_class_id_post_teachers,
         C_class_id_consolidated,
         C_class_id,
         everything())

# CHECK 5: Count rows flagged to be excluded AFTER new load-exclusion flag; this should INCREASE from the previous step by the number of rows from CHECK 2.
course_data_flagged %>% 
  group_by(C_class_size_and_teacher_load_exclude) %>% 
  count()

# CHECK 6: Check the types of courses that are getting excluded based on the new teacher load exclusion flag. You should see that the vast majority of them are non-PE/Art courses since that's the main criteria for exclusion.
course_data_flagged %>%
  filter(H_teachers_to_exclude == 1) %>% 
  group_by(
    D_location_name,
    C_course_subject,
    D_course_name) %>% 
  summarize(
    class_count_post_teachers = n_distinct(C_class_id_post_teachers),
    count = n()) %>% 
  arrange(desc(count))

# CHECK 7: Check distribution of original class ids that are included in consolidated class id.
course_data_flagged %>% 
  group_by(C_class_id_post_teachers) %>% 
  summarise(class_id_count = n_distinct(C_class_id),
            consolidated_id_count = n_distinct(C_class_id_consolidated)) %>% 
  group_by(class_id_count,
           consolidated_id_count) %>% 
  summarise(post_teachers_count = n_distinct(C_class_id_post_teachers))


# 3.4 Recalculate Class Weight ----

# What: Now that classes with the same teachers have been consolidated, the next step is to adjust the student class weight based on the number of classes that a student saw consolidated for them during the consolidation stage within any given TPR. The code below calculates the student adjuster (a.k.a. the number of classes a student used to have in a given consolidated class) and divides a student's class weight by this adjuster.
# The key idea is: If multiple original classes got collapsed into one consolidated class, and a student appeared in more than one of those original classes, that student should not be counted multiple times inside the consolidated class.

# Starting point:
# By this point, each row has:
# - C_class_id: the original "exploded" class id
# - C_class_id_post_teachers: the updated class id after teacher consolidation
# - Original "exploded" class id if no consolidation happened
# - Consolidated class id if teacher double-booking rows were combined
# - M_class_weight_post_class: the class weight after the earlier class-level double-booking fix
# Now the script needs to adjust weights again for the teacher consolidation step.

# Check:

# - Max of student_adjuster_post_teachers is expected to be above 1 (since some classes got consolidated and at least a couple students are expected to have a couple cases where they appeared in both classes that got consolidated together) and min should be 1 (since some classes didn't have any consolidation) - if either of these two checks don't pass, dig into your consolidated classes and see what original classes are making them up to see if your max and min results make sense within your district's context of teacher DB classes

# - Look at the raw data and find a couple students with consolidated class - be sure that their class weights after this adjustment look as you would expect

# Project Team Action: No change (run as is)

# %% 3.4 Recalculate Class Weight
# Group the data at the level of one student in one post-teacher class during one term/period/rotation
# This matters because C_class_id_post_teachers may now represent multiple original classes.
# "student_adjuster_post_teachers" counts how many original classes a student had within that post-teacher class id, so that the class weight can be adjusted accordingly in the next step.
course_data_flagged <- course_data_flagged %>%
  group_by(D_stu_id,  
           C_class_id_post_teachers,
           C_rotation_exploded,
           C_term_exploded, 
           C_period_exploded) %>%
  summarize(student_adjuster_post_teachers = n_distinct(C_class_id)) %>%
  right_join(course_data_flagged)

# Check adjuster range
# If min AND max are only 1, then the consolidation did not actually create any student-level duplicates to adjust.
    # Most likely reason: the consolidated classes do NOT share students: even if a teacher had multiple classes consolidated, each individual student may only appear in one of those original classes.
max(course_data_flagged$student_adjuster_post_teachers)
min(course_data_flagged$student_adjuster_post_teachers)

# Review original class IDs under post-teacher class IDs
course_data_flagged %>% 
  group_by(C_class_id_post_teachers, C_class_id) %>%
  summarize()

# Calculates new class weight
# It divides the current class weight by the number of original classes the student had inside the consolidated class.
# So if a student appeared in TWO original classes that are now treated as ONE consolidated class, each row gets HALF weight.
course_data_flagged$M_class_weight_post_teachers <- course_data_flagged$M_class_weight_post_class / course_data_flagged$student_adjuster_post_teachers


# 3.5 Rerunning Student Double Bookings ----

# 3.5a Calculate Flag ----

# What: This code block reruns the student DB flagger

# Check:

# - The first line printed by code block should have a number that is lower (but still close) to the number of rows in course_data_flagged

# - The first table printed by the code block should have the vast majority of values at the "standard" class weight and some values above the "standard" class weight

# - The second table printed by the code block should have all values as 0 or 1 and the number of values for 0 should exactly match the number of values at the "standard" weight value in the first table

# - Lastly, the third table printed shows the number of rows at 0 and 1 for the previous version of the student DB flag and you should expect to see that the count of 1s was higher in this flag than in the flag we now have after teachers resolutions (which was shown in the second table) - if it's equal that means that none of your students were across two classes that got consolidated (surprising result) and if it's lower then you definitely have an error somewhere since you now have more double booked students than earlier in the process

# Project Team Action: Do not change (run as is - and note that the code block might take a couple minutes to run)

# %% 3.5a Calculate Flag
# Runs the flagger
student_db_post_teachers <- student_db_flagger(course_data_flagged,
                                            M_class_weight_post_teachers,
                                            H_db_student_row_post_class)

# Rename the db flag so that it's identifiable in the next steps as "student double-booking flag after teacher-level fix"
student_db_post_teachers <- student_db_post_teachers %>% 
  rename(H_db_student_row_post_teachers = H_updated_db_student_row)

# Check number of rows, which should be lower than course_data_flagged since the flagger filters and also groups by class weight.
nrow(student_db_post_teachers)

# Check number of rows by class weight - you should see the majority of values at your "standard" class weight and some values above that.
student_db_post_teachers %>% 
  group_by(class_weight_sum) %>% 
  count()

# Check number of rows with a student-level double-booking BEFORE teacher resolutions.
student_db_post_teachers %>% 
  group_by(H_db_student_row_post_class) %>% 
  count()

# Check number of rows with a student-level double-booking AFTER teacher resolutions.
# The number of 1s may go down IF students appeared in multiple classes that were consolidated.
# If the count is unchanged, check whether student_adjuster_post_teachers min and max were both 1 in the previous block.
student_db_post_teachers %>% 
  group_by(H_db_student_row_post_teachers) %>% 
  count()


# 3.5b Bring Flag Back ----

# What: Now that you feel good about the double bookings flag calculated in the previous code block, it's time to bring the flag into course_data_flagged and use it to mark rows that are still double booked even after the teachers resolutions

# Check:

# - Look at course_data_flagged in the environment (top right of RStudio) and confirm that the new student double bookings flag is now a column in it

# - The first table printed by the code block shows the percentage of rows in course_data_flagged that have been marked as double booked for student - confirm that the results seem reasonable. Note that the number is given as percentage (so 0.01 means 1%).

# - The two numbers printed show the percent of rows marked as student double-booked after class resolutions and after teacher resolutions. The second number may go down if teacher consolidation reduced student-level double-bookings; it may stay the same if student weights did not change during teacher consolidation.

# Project Team Action: Do not change (run as is)

# %% 3.5b Bring Flag Back
# Merge the new student DB flag back into the main data
# H_db_student_row_post_teachers: the new flag that says whether the student is STILL double-booked after teacher-level fixes.
course_data_flagged <- student_db_merger(student_db_post_teachers,
                                         H_db_student_row_post_teachers)
  
# Changing the NAs to 0 (since they get excluded anyway)
course_data_flagged <- course_data_flagged %>% 
  mutate(H_db_student_row_post_teachers = ifelse(is.na(H_db_student_row_post_teachers),
                                                 0,
                                                 H_db_student_row_post_teachers))

# Check the student DB rate BEFORE teacher fixes
mean(course_data_flagged$H_db_student_row_post_class)

# Check the student DB rate AFTER teacher fixes
# If it goes down, teacher consolidation helped resolve some student double-bookings.
# If it stays the same, that usually means teacher consolidation did not change student-level weights (i.e., students were not in multiple classes that got consolidated together).
mean(course_data_flagged$H_db_student_row_post_teachers)


# 3.6 Rerunning Teacher Double Bookings ----

# 3.6a Calculate Flag ----

# What: This code block runs the teacher DB flagger. The main question is: After consolidating teacher/TPR classes and excluding problematic teachers, are there still any teacher double-bookings left?

# Check:

# - The first line printed by code block should have a number that is significantly lower than the number of rows in course_data_exploded

# - The first table printed by the code block should have the vast majority of values at your "standard" class weight and some values above the "standard" class weight

# - The second table printed by the code block should have all values as 0 or 1 and the number of values for 0 should exactly match the number of values at or below the "standard" weight value in the first table

# - Lastly, the third table printed shows the number of rows at 0 and 1 after the teacher resolutions step and all rows should now appear with a value of 0 (since they're no longer double booked at the teacher level)

# Project Team Action: Do not change (run as is - and note that the code block might take a couple minutes to run)

# %% 3.6a Calculate Flag
# Runs the flagger
teacher_db_post_teachers <- teacher_db_flagger(
  course_data_flagged,
  C_class_id_post_teachers,
  M_class_weight_post_teachers,
  H_db_teacher_row_post_class)

# Renames the db flag to identify it with its step in the process
teacher_db_post_teachers <- teacher_db_post_teachers %>% 
  rename(H_db_teacher_row_post_teachers = H_updated_db_teacher_row)

# Check number of summarized teacher/TPR rows
nrow(teacher_db_post_teachers)

# Check class weight sums AFTER teacher resolutions
teacher_db_post_teachers %>% 
  group_by(class_weight_sum_actual) %>% 
  count()

# Check teacher double-bookings BEFORE teacher resolutions
teacher_db_post_class %>% 
  group_by(H_db_teacher_row_post_class) %>% 
  count()

# Check teacher double-bookings AFTER teacher resolutions; ALL SHOULD BE ZERO!
teacher_db_post_teachers %>% 
  group_by(H_db_teacher_row_post_teachers) %>% 
  count()


# 3.6b Bring Flag Back ----

# What: Now that you feel good about the double bookings flag calculated in the previous code block, it's time to bring the flag into course_data_flagged and use it to mark rows that are double booked

# Check:

# - Look at course_data_flagged in the environment (top right of RStudio) and confirm that the teacher double bookings flag is now a column in it

# - The table printed by the code block shows the percentage of rows in course_data_exploded that have been marked as double booked for teacher - confirm that the results seem reasonable. Note that the number is shown as a percentage (e.g. 0.01 means 1%)

# - The two numbers given in the second output are the % of rows marked as a student DB before and after the class resolutions - you should see that the % went down at least somewhat

# - The last output shows the percent of teachers still double booked after we pull out the teachers with C_class_size_and_teacher_load_exclude value of exclude - the number here should be 0 since all teachers are expected to be resolved by now

# Project Team Action: do not change (run as is)

# %% 3.6b Bring Flag Back
# Runs the merger
course_data_flagged <- teacher_db_merger(teacher_db_post_teachers,
                                         H_db_teacher_row_post_teachers)

# Changing the NAs to 0 (since they get excluded in teacher load analysis anyway)
course_data_flagged <- course_data_flagged %>% 
  mutate(H_db_teacher_row_post_teachers = ifelse(is.na(H_db_teacher_row_post_teachers),
                                                 0,
                                                 H_db_teacher_row_post_teachers))

# Check
mean(course_data_flagged$H_db_teacher_row_post_class)

# Check
mean(course_data_flagged$H_db_teacher_row_post_teachers)

# Check - ALL SHOULD BE ZERO since all teacher double-bookings should be resolved among rows included in both teacher-load analysis and time-based analysis.
course_data_flagged %>% 
  filter(C_class_size_and_teacher_load_exclude == "Include") %>% 
  pull(H_db_teacher_row_post_teachers) %>% 
  mean()


# Part IV: Student Double Bookings Resolution ----

# 4.1 Adjusting Class Weight by Student ----

# What: This code block backfills the NAs in C_class_id with the previous class id and then groups the data by student and TPR to count how many classes a student has in that TPR - this count is the student adjuster that will be used in student DB. Importantly, the cases of all classes in a TPR being 100% double booked were already adjusted in the class double bookings phase so I take those classes out of this step (they have already been fixed).

# This section is the final student-level double-booking resolution. At this point, the script has already tried to fix double-bookings through:
# - Class-level fixes for 100% double-booked classes
# - Teacher-level fixes for teacher double-bookings / consolidated classes

# Now it is saying: For any remaining student double-bookings, split the student’s weight across the classes they STILL appear in during the same term/period/rotation.

# Check:

# - First check shows how many NAs are present in the backfilled class id (C_class_id_post_teachers) - result should be 0

# - Second check shows the distribution of the student adjuster calculated for student DB

# Project Team Action:

# %% 4.1 Adjusting Class Weight by Student
# Backfill missing post-teacher class IDs
course_data_flagged <- course_data_flagged %>% 
  mutate(C_class_id_post_teachers = ifelse(is.na(C_class_id_post_teachers),
                                         C_class_id,
                                         as.character(C_class_id_post_teachers)))

# Check that no post-teacher class IDs are missing - the result should be ZERO
sum(is.na(course_data_flagged$C_class_id_post_teachers))

# Calculate student adjuster
# The goal is to count:how many distinct classes does this student still have in the same TPR?
student_helper <- course_data_flagged %>% 
  filter(H_tpr_100p_percent != 1) %>% 
  group_by(D_stu_id,
           C_period_exploded,
           C_term_exploded,
           C_rotation_exploded,
           row_count) %>% 
  summarize(student_adjuster_post_student = n_distinct(C_class_id_post_teachers))

# Check distribution of student adjuster - this shows how many student/TPR combinations have 1 class, 2 classes, 3 classes, etc.
student_helper %>% 
  group_by(student_adjuster_post_student) %>% 
  count()

# Bring student adjuster back to original data
course_data_flagged <- left_join(course_data_flagged,
                                 student_helper)

# Create final student-adjusted class weight
course_data_flagged <- course_data_flagged %>% 
  mutate(M_class_weight_post_student = ifelse(is.na(student_adjuster_post_student),
                                                  M_class_weight_post_teachers,
                                                  M_class_weight_post_teachers /
                                                  student_adjuster_post_student))


# 4.2 Rerunning Student Double Bookings ----

# 4.2a Calculate Flag ----

# What: This code block reruns the student DB flagger. The main question is: After the final student-level weight adjustment, are any students still double-booked? And the answer should be NO.

# Check:

# - The first line printed by code block should have a number that is lower (but still close) to the number of rows in course_data_exploded

# - The first table printed by the code block should have the vast majority of values at the "standard" class weight and some values above the "standard" class weight

# - The second table printed by the code block should have all values as 0

# Project Team Action: do not change (run as is - and note that the code block might take a couple minutes to run)

# %% 4.2a Calculate Flag
# Runs the flagger
student_db_post_student <- student_db_flagger(course_data_flagged,
                                            M_class_weight_post_student,
                                            H_db_student_row_post_teachers)

# Rename the db flag so that it's identifiable with step in process
student_db_post_student <- student_db_post_student %>% 
  rename(H_db_student_row_post_student = H_updated_db_student_row)

# Check the number of summarized rows - should be lower than course_data_flagged since the flagger filters and also groups into student/term/period/rotation summaries.
nrow(student_db_post_student)

# Check total class weight by student/time block
# Some values BELOW the standard weight can be okay, depending on term/rotation/period weights. Values ABOVE the expected weight are what would trigger the DB flag.
student_db_post_student %>% 
  group_by(class_weight_sum) %>% 
  count()

# Check final student double-booking flag - ALL SHOULD BE ZERO!
student_db_post_student %>% 
  group_by(H_db_student_row_post_student) %>% 
  count()


# 4.2b Bring Flag Back ----

# What: Now that you feel good about the double bookings flag calculated in the previous code block, it's time to bring the flag into course_data_flagged

# Check:

# - Look at course_data_flagged in the environment (top right of RStudio) and confirm that the new student double bookings flag is now a column in it

# - The first table printed by the code block shows the percentage of rows in course_data_flagged that have been marked as double booked for student - confirm that the results seem reasonable. Note that the number is given as percentage (so 0.01 means 1%).

# - The two numbers given in the second output are the % of rows marked as a student DB before and after the student resolutions - the number here should be 0 since all students are expected to be resolved by now

# Project Team Action: do not change (run as is)

# %% 4.2b Bring Flag Back
# Runs the merger
course_data_flagged <- student_db_merger(student_db_post_student,
                                         H_db_student_row_post_student)

# Turns the NAs into 0s (since they're due to exclusion from time analysis)
course_data_flagged <- course_data_flagged %>% 
  mutate(H_db_student_row_post_student = ifelse(is.na(H_db_student_row_post_student),
                                                0,
                                                H_db_student_row_post_student))
  
# Check - this represents student double-bookings after teacher resolutions, but BEFORE the final student-level weight split.
mean(course_data_flagged$H_db_student_row_post_teachers)

# Check - this is AFTER the final student-level weight split. SHOULD BE ZERO
mean(course_data_flagged$H_db_student_row_post_student)


# %% 4.2b Bring Flag Back (2)
# Check weight values by employee
# It's trying to answer: For each teacher, how much class weight did they have before and after each adjustment step?
class_weight_check <- course_data_flagged %>% 
  distinct(
    D_employee_id,
    C_class_id_post_teachers,
    M_class_weight,
    M_class_weight_exploded,
    M_class_weight_post_class,
    M_class_weight_post_teachers,
    M_class_weight_post_student
  )

class_weight_summary <- class_weight_check %>% 
  group_by(D_employee_id) %>% 
  summarise(
    sum_cw = sum(M_class_weight, na.rm = TRUE),
    sum_cw_exploded = sum(M_class_weight_exploded, na.rm = TRUE),
    sum_cw_post_class = sum(M_class_weight_post_class, na.rm = TRUE),
    sum_cw_post_teachers = sum(M_class_weight_post_teachers, na.rm = TRUE),
    sum_cw_post_student = sum(M_class_weight_post_student, na.rm = TRUE),
    .groups = "drop"
  ) %>% 
  mutate(
    change_class_resolution = sum_cw_post_class - sum_cw_exploded,
    change_teacher_resolution = sum_cw_post_teachers - sum_cw_post_class,
    change_student_resolution = sum_cw_post_student - sum_cw_post_teachers
  )

# Check weight values by C_class_id_post_teachers
class_weight_by_class <- course_data_flagged %>% 
  group_by(
    C_class_id_post_teachers) |> 
      summarise(
        course_names = paste0(unique(C_course_name), collapse = " / "),
        max_class_weight = max(M_class_weight_post_student, na.rm = TRUE))

n_distinct(course_data_flagged$C_class_id_post_teachers)


# Part IV: Export ----

# 1. Trim and export data ----

# What: Limiting course_data_final only to variables of interest and then exporting to SharePoint

# Check: Look at data in environment before exporting it to be sure that you feel good about the set of variables you're passing on to the next script - and once you export it check your SharePoint folder to ensure it got loaded correctly

# Project Team Action: Likely change (if you have any additional variables that you want to keep at hand, add them to this code block)

# %% 1. Trim and export data
# Team action: add in any variables you want to keep that aren't already listed
course_data_flagged <- course_data_flagged |>
  select(
    # Core IDs / student / teacher / school
    D_location_id,
    D_location_name,
    D_stu_id,
    D_stu_grade,
    D_employee_id,

    # Course / section fields
    D_ccid,
    D_users_dcid,
    D_course_id,
    D_course_name,
    C_course_name,
    D_course_section,
    D_course_ell_flag,
    D_course_swd_flag,
    D_year_id,
    D_term,
    D_expression,
    D_period,
    D_rotation,

    # ERS course coding fields
    C_location_type,
    C_stu_snapshot,
    C_sced_code,
    C_sced_subject_code,
    C_sced_subject_name,
    C_course_time,
    C_course_subject_area,
    C_course_subject,
    C_course_credit_type,
    C_course_rigor,
    C_course_rigor_detail,
    C_course_intervention,
    C_course_time_of_day,
    C_course_format,
    C_course_pathway,
    C_teacher_load_metric,
    C_class_size_metric,
    C_time_metric,
    C_course_time_exclude,
    C_class_size_and_teacher_load_exclude,
    C_teacher_load_metric_lgl,
    C_class_size_metric_lgl,
    C_time_metric_lgl,

    # Class IDs
    C_class_id_original,
    C_class_id,
    C_class_id_consolidated,
    C_class_id_post_teachers,

    # Original weights
    C_term_weight,
    C_rotation_weight,
    C_period_weight,
    M_class_weight,

    # Explosion fields
    H_row_id_before_explosions,
    H_row_id_after_explosions,
    H_new_record_id,
    C_term_exploded,
    C_rotation_exploded,
    C_period_exploded,
    C_term_weight_exploded,
    C_rotation_weight_exploded,
    C_period_weight_exploded,
    M_class_weight_exploded,

    # Same-course adjustment
    H_student_adjuster_same_course,
    M_class_weight_adj_same_course,

    # Original DB flags
    H_class_with_teacher_db,
    H_class_with_student_db,
    H_class_with_db,
    H_db_student_row,
    H_db_teacher_row,

    # Class-level DB resolution
    row_count,
    H_count_student_rows_by_class,
    H_db_count_by_class,
    H_percent_db_by_class,
    H_db_100p_class,
    H_student_count_of_100p_classes,
    H_tpr_100p_percent,
    M_class_weight_post_class,
    H_db_student_row_post_class,
    H_db_teacher_row_post_class,

    # Teacher-level DB resolution
    H_teachers_to_exclude,
    student_adjuster_post_teachers,
    M_class_weight_post_teachers,
    H_db_student_row_post_teachers,
    H_db_teacher_row_post_teachers,

    # Anything else not listed above
    everything()
  ) |>
  arrange(
    C_class_id_post_teachers,
    C_class_id
  )


# Export to SharePoint

# %% 1. Trim and export data (2)
# Save course schedule data for next script
ers_write_sharepoint(
  data = course_data_flagged,
  folder_path = raw_data_folder_path,
  file_name_with_extension = 
    "/2. Processed Data/06_cs_data_post_db_resolutions.csv")


# Bonus: Celebrate! ----

# And with that, you're done with this script!!! Here's a bear to celebrate :D

# %% Bonus: Celebrate!
##           ##
##   o   o   ##
##  ( *.* )  ##
##  (") (")  ##
##  (  O  )  ##
##  (") (")  ##
##           ##
##   YOU'RE  ##
##  AWESOME  ##
##           ##
