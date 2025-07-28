###############################################################################################
#### LOAD PACKAGES #### -----------------------------------------------------------------------
###############################################################################################

library(readxl)
library(ggplot2)
library(dplyr)
library(scales)


###############################################################################################
#### IMPORT AND CLEAN DATASET #### ------------------------------------------------------------
###############################################################################################

# LOAD DATA -----------------------------------------------------------------------------------

PATH="/Users/marieleapouliquen/Desktop/Review/Data/"
data0 <- read_excel(file.path(PATH, "all-records.xlsx"), sheet = "Final Selection")

# GROUP DATA BY TECHNO CATEGORIES -------------------------------------------------------------
data <- data0 %>%  
  mutate(TechnoFamily = case_when(  
    !is.na(TechnoFamily) ~ TechnoFamily,  
    #TechnoCategory == "Mobile Apps" ~ "Mobile Applications",  
    TechnoCategory == "Environmental Apps" ~ "Mobile Applications",   
    TechnoCategory == "Well-Being Apps" ~ "Mobile Applications",  
    TechnoCategory == "Environmental Sensors" ~ "Instruments",  
    TechnoCategory == "Responsive Instruments" ~ "Instruments",  
    TechnoCategory == "Nature Photography" ~ "Instruments",  
    TechnoCategory == "Nature Recordings" ~ "Nature Recordings",
    TechnoCategory == "Virtual Reality" ~ "Virtual Reality",
    TechnoCategory == "Video Games" ~ "Video Games",
    TRUE ~ "Other"  
  )) 


data1 <- data %>%
  mutate(TechnoFamily = case_when(
    #TechnoCategory == "Mobile Apps" ~ "Mobile Applications",  
    TechnoCategory == "Environmental Apps" ~ "Mobile Applications", 
    TechnoCategory == "Well-Being Apps" ~ "Mobile Applications",
    TechnoCategory == "Environmental Sensors" ~ "Instruments",
    TechnoCategory == "Responsive Instruments" ~ "Instruments",
    TechnoCategory == "Nature Photography" ~ "Instruments",
    TechnoCategory == "Nature Recordings" ~ "Nature Recordings",
    TechnoCategory == "Virtual Reality" ~ "Virtual Reality",
    TechnoCategory == "Video Games" ~ "Video Games",
    TRUE ~ "Other"
  )) %>%
  
  count(TechnoCategory, TechnoFamily, name = "Count") %>%
  mutate(Percentage = round(Count/sum(Count) * 100, 1)) %>%
  arrange(desc(Count))


###############################################################################################
#### CREATE VISUALIZATIONS #### ---------------------------------------------------------------
###############################################################################################


# TEMPORAL DISTRIBUTION OF STUDIES ------------------------------------------------------------

year_counts <- data0 %>%
  count(Date, name = "Count") %>%
  arrange(Date)

p1 <- ggplot(data, aes(x = Date)) +
  geom_histogram(binwidth = 1, fill = "steelblue", color = "white", alpha = 0.7) +
  labs(title = "Temporal Distribution of Studies in Scoping Review",
       subtitle = paste("Total studies:", nrow(data)),
       x = "Publication Year",
       y = "Number of Studies") +
  theme_minimal() +
  theme(plot.title = element_text(size = 14, face = "bold"),
        plot.subtitle = element_text(size = 12)) +
  scale_x_continuous(breaks = seq(2004, 2025, by = 2))

print(p1)


# TEMPORAL DISTRIBUTION OF STUDIES PER TECHNO FAMILY ---------------------------------------------------

p2 <- ggplot(data, aes(x = Date, fill = TechnoFamily)) +
  geom_bar(color = "white", alpha = 0.8, width = 0.8) +
  labs(title = "Temporal Distribution of Technology Types",
       subtitle = paste("Total studies:", nrow(data), "| Years:", min(data$Date), "-", max(data$Date)),
       x = "Publication Year",
       y = "Number of Studies",
       fill = "Technology Types :") +
  theme_minimal() +
  theme(plot.title = element_text(size = 14, face = "bold"),
        plot.subtitle = element_text(size = 12),
        axis.text.x = element_text(angle = 45, hjust = 1),
        legend.position = "bottom",
        legend.title = element_text(face = "bold")) +
  scale_fill_hue(h = c(0, 360), c = 50, l = 80) +  # Using husl with pale tones (low chroma, high lightness)
  scale_x_continuous(breaks = seq(2004, 2025, by = 2)) +
  guides(fill = guide_legend(nrow = 2))
  
print(p2)  



# PIE CHART OF TECHNO FAMILY ---------------------------------------------------------------------------

# Count number of studies per TechnoFamily

 count_techno_family <- data %>%  
  count(TechnoFamily, name = "Count") %>%  
  mutate(Percentage = round(Count/sum(Count) * 100, 1)) %>%  
  arrange(desc(Count))  

# Print the count

print("Complete TechnoFamily distribution :")  
print(count_techno_family)  

# Plot the corresponding pie chart
 
p3 <- ggplot(count_techno_family, aes(x = "", y = Count, fill = TechnoFamily)) +  
  geom_bar(stat = "identity", width = 1, color = "white") +  
  coord_polar("y", start = 0) +  
  labs(title = "Global Repartition of Technology Types",  
       subtitle = paste("Total studies:", sum(count_techno_family$Count)),  
       fill = "Technology Types") +  
  theme_void() +  
  theme(plot.title = element_text(size = 14, face = "bold", hjust = 0.5),  
        plot.subtitle = element_text(size = 12, hjust = 0.5),  
        legend.position = "right",  
        legend.text = element_text(size = 10),  
        legend.title = element_text(size = 12, face = "bold")) +  
  scale_fill_hue(h = c(0, 360), c = 50, l = 80) +  # Using pale husl colors  
  geom_text(aes(label = paste0(Count, "\n(", Percentage, "%)")),   
            position = position_stack(vjust = 0.5),  
            size = 3.5, fontface = "bold", color = "black")  

print(p3)  


# PIE CHART OF TECHNO CATEGORY ---------------------------------------------------------------------------

# Count number of studies per TechnoCategory

count_techno_category <- data %>%  
  count(TechnoCategory, name = "Count") %>%  
  mutate(Percentage = round(Count/sum(Count) * 100, 1)) %>%  
  arrange(desc(Count))  

# Plot the corresponding pie chart

p3 <- ggplot(count_techno_category, aes(x = "", y = Count, fill = TechnoCategory)) +  
  geom_bar(stat = "identity", width = 2, color = "white") +  
  coord_polar("y", start = 0) +  
  labs(title = "A) Proportions of Technology Types \n in Scoping Review ",  
       subtitle = paste("(N=",sum(count_techno_category$Count),"studies)"),  
       fill = "Technology Types") +  
  theme_void() +  
  theme(plot.title = element_text(size = 12, face = "bold", hjust = 0.5, vjust = -8),  
        plot.subtitle = element_text(size = 10, hjust = 0.5,vjust = -10),  
        legend.position = "right",  
        legend.text = element_text(size = 10),  
        legend.title = element_text(size = 12, face = "bold")) +  
  scale_fill_hue(h = c(-270, 60), c = 60, l = 80) +  # Using pale husl colors  
  geom_text(aes(label = paste0(Count, " (", Percentage, "%)")),   
            position = position_stack(vjust = 0.5),  
            size = 3.8, fontface = "bold", color = "black")  

print(p3)  




# BAR CHART OF TECHNO CATEGORIES -------------------------------------------------------------

# Coloring

unique_families <- sort(unique(data1$TechnoCategory))
family_colors_husl <- hue_pal(h = c(-270, 60), c = 60, l = 80)(length(unique_families))
names(family_colors_husl) <- unique_families

# Create the final bar chart

p4 <- ggplot(data1,aes(x = reorder(TechnoCategory, Count), y = Count, fill = TechnoCategory)) +
  geom_bar(stat = "identity", color = "white", alpha = 0.8) +
  coord_flip() +
  labs(title = "Distribution of Technology Categories",
       subtitle = paste("Total studies:", sum(data1$Count)),
       x = "",
       y = "Number of Studies",
       fill = "Types of Technology: ") +
  theme_minimal() +
  theme(plot.title = element_text(size = 14, face = "bold"),
        plot.subtitle = element_text(size = 11),
        axis.text.y = element_text(size = 9),
        legend.position = "right",
        legend.title = element_text(face = "bold")) +
  scale_fill_manual(values = family_colors_husl) +
  geom_text(aes(label = paste0(Count, " (", Percentage, "%)")), 
            hjust = 1, size = 2.6, fontface = "bold") +
  guides(fill = guide_legend(nrow = 7))

print(p4)



# TEMPORAL DISTRIBUTION OF STUDIES PER TECHNO CATEGORY ---------------------------------------------------

p5 <- ggplot(data, aes(x = Date, fill = TechnoCategory)) +
  geom_bar(color = "white", alpha = 1, width = 0.8) +
  labs(title = "B) Temporal Distribution of Technology Used to Mediate Nature Experience",
       subtitle = paste("Total studies:", nrow(data), "| Years:", min(data$Date), "-", max(data$Date)),
       x = "Publication Year",
       y = "Number of Studies",
       fill = "Technology:") +
  theme_minimal(base_size = 10) +
  theme(plot.title = element_text(size = 12, face = "bold"),
        plot.subtitle = element_text(size = 10),
        axis.text.x = element_text(angle = 0, hjust = 0.5),
        legend.position = "bottom",
        legend.title = element_text(face = "bold",size = 9.5)) +
  scale_fill_hue(h = c(-270, 60), c = 60, l = 80) +  # Using husl with pale tones (low chroma, high lightness)
  scale_x_continuous(breaks = seq(2004, 2025, by = 2)) +
  guides(fill = guide_legend(nrow = 2))

print(p5)  



