###############################################################################################
#### LOAD PACKAGES #### -----------------------------------------------------------------------
###############################################################################################

library(readxl)
library(ggplot2)
library(dplyr)
#install.packages("openxlsx", repos="https://cran.rstudio.com/", dependencies = TRUE)
library(openxlsx)

###############################################################################################
#### IMPORT DATASET #### ----------------------------------------------------------------------
###############################################################################################

PATH="/Users/marieleapouliquen/Desktop/Review/Data/"
meta_data <- read_excel(file.path(PATH, "meta.xlsx"), sheet = "techno") 


###############################################################################################
#### STATISTICS BY MODE #### ------------------------------------------------------------------
###############################################################################################


# Get unique groups to avoid double counting participants (including all modes)
unique_groups_complete <- meta_data %>%
  select(StudyID, GroupID, N_Pre, Nature, Experience, Approach, Orientation, Participant, Purpose, Duration, Frequency, Location, HNC) %>%
  distinct(StudyID, GroupID, .keep_all = TRUE)

# Calculate statistics by each mode
nature_stats <- unique_groups_complete %>%
  group_by(Nature) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

experience_stats <- unique_groups_complete %>%
  group_by(Experience) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

approach_stats <- unique_groups_complete %>%
  group_by(Approach) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

orientation_stats <- unique_groups_complete %>%
  group_by(Orientation) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

participant_stats <- unique_groups_complete %>%
  group_by(Participant) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

purpose_stats <- unique_groups_complete %>%
  group_by(Purpose) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

duration_stats <- unique_groups_complete %>%
  group_by(Duration) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

frequency_stats <- unique_groups_complete %>%
  group_by(Frequency) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

location_stats <- unique_groups_complete %>%
  mutate(Location = ifelse(is.na(Location), "Not Specified", Location)) %>%
  group_by(Location) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

hnc_stats <- unique_groups_complete %>%
  group_by(HNC) %>%
  summarise(N_Studies = n_distinct(StudyID), N_Groups = n(), N_Participants = sum(N_Pre, na.rm = TRUE), .groups = 'drop')

# Get measure counts for all modes
nature_measures <- meta_data %>% count(Nature, name = "N_Measures")
experience_measures <- meta_data %>% count(Experience, name = "N_Measures")
approach_measures <- meta_data %>% count(Approach, name = "N_Measures")
orientation_measures <- meta_data %>% count(Orientation, name = "N_Measures")
participant_measures <- meta_data %>% count(Participant, name = "N_Measures")
purpose_measures <- meta_data %>% count(Purpose, name = "N_Measures")
duration_measures <- meta_data %>% count(Duration, name = "N_Measures")
frequency_measures <- meta_data %>% count(Frequency, name = "N_Measures")
location_measures <- meta_data %>% 
  mutate(Location = ifelse(is.na(Location), "Not Specified", Location)) %>%
  count(Location, name = "N_Measures")
hnc_measures <- meta_data %>% count(HNC, name = "N_Measures")

# Create the final complete comprehensive statistics table with ALL 9 modes
final_comprehensive_stats <- rbind(
  # Overall
  data.frame(Mode = "Overall", Subgroup = "All Studies", 
             N_Studies = n_distinct(unique_groups_complete$StudyID),
             N_Groups = nrow(unique_groups_complete),
             N_Measures = nrow(meta_data),
             N_Participants = sum(unique_groups_complete$N_Pre, na.rm = TRUE)),
  
  # Nature
  data.frame(Mode = "Nature", Subgroup = nature_stats$Nature,
             N_Studies = nature_stats$N_Studies,
             N_Groups = nature_stats$N_Groups,
             N_Measures = nature_measures$N_Measures,
             N_Participants = nature_stats$N_Participants),
  
  # Experience
  data.frame(Mode = "Experience", Subgroup = experience_stats$Experience,
             N_Studies = experience_stats$N_Studies,
             N_Groups = experience_stats$N_Groups,
             N_Measures = experience_measures$N_Measures,
             N_Participants = experience_stats$N_Participants),
  
  # Approach
  data.frame(Mode = "Approach", Subgroup = approach_stats$Approach,
             N_Studies = approach_stats$N_Studies,
             N_Groups = approach_stats$N_Groups,
             N_Measures = approach_measures$N_Measures,
             N_Participants = approach_stats$N_Participants),
  
  # Orientation
  data.frame(Mode = "Orientation", Subgroup = orientation_stats$Orientation,
             N_Studies = orientation_stats$N_Studies,
             N_Groups = orientation_stats$N_Groups,
             N_Measures = orientation_measures$N_Measures,
             N_Participants = orientation_stats$N_Participants),
  
  # Participant
  data.frame(Mode = "Participant", Subgroup = participant_stats$Participant,
             N_Studies = participant_stats$N_Studies,
             N_Groups = participant_stats$N_Groups,
             N_Measures = participant_measures$N_Measures,
             N_Participants = participant_stats$N_Participants),
  
  # Purpose
  data.frame(Mode = "Purpose", Subgroup = purpose_stats$Purpose,
             N_Studies = purpose_stats$N_Studies,
             N_Groups = purpose_stats$N_Groups,
             N_Measures = purpose_measures$N_Measures,
             N_Participants = purpose_stats$N_Participants),
  
  # Duration
  data.frame(Mode = "Duration", Subgroup = duration_stats$Duration,
             N_Studies = duration_stats$N_Studies,
             N_Groups = duration_stats$N_Groups,
             N_Measures = duration_measures$N_Measures,
             N_Participants = duration_stats$N_Participants),
  
  # Frequency
  data.frame(Mode = "Frequency", Subgroup = frequency_stats$Frequency,
             N_Studies = frequency_stats$N_Studies,
             N_Groups = frequency_stats$N_Groups,
             N_Measures = frequency_measures$N_Measures,
             N_Participants = frequency_stats$N_Participants),
  
  # Setting (Location)
  data.frame(Mode = "Setting", Subgroup = location_stats$Location,
             N_Studies = location_stats$N_Studies,
             N_Groups = location_stats$N_Groups,
             N_Measures = location_measures$N_Measures,
             N_Participants = location_stats$N_Participants),
  
  # HNC (Human-Nature Connection measures)
  data.frame(Mode = "HNC", Subgroup = hnc_stats$HNC,
             N_Studies = hnc_stats$N_Studies,
             N_Groups = hnc_stats$N_Groups,
             N_Measures = hnc_measures$N_Measures,
             N_Participants = hnc_stats$N_Participants)
)


###############################################################################################
#### CREATE OUTPUTS #### ----------------------------------------------------------------------
###############################################################################################

file1 <- paste0(PATH, "metaStats.xlsx")
write.xlsx(final_comprehensive_stats, file1, rowNames = FALSE)


# Create formatted Excel workbook
#file2 <- paste0(PATH, "comprehensive_statistics_formatted.xlsx")
#wb <- createWorkbook()
#addWorksheet(wb, "Comprehensive Statistics")

# Write the data
#writeData(wb, "Comprehensive Statistics", final_comprehensive_stats, startRow = 1, startCol = 1)

# Add formatting
#headerStyle <- createStyle(textDecoration = "bold", fgFill = "#4F81BD", fontColour = "white")
#addStyle(wb, "Comprehensive Statistics", headerStyle, rows = 1, cols = 1:6)

# Auto-size columns
#setColWidths(wb, "Comprehensive Statistics", cols = 1:6, widths = "auto")

# Save the formatted workbook with custom path
#saveWorkbook(wb, file2, overwrite = TRUE)



