# Complete code to create the technology mode boxplot
# Load required libraries
library(readxl)
library(dplyr)
library(ggplot2)

# Read the Excel file
PATH="/Users/marieleapouliquen/Desktop/Review/Data/"
data <- read_excel(file.path(PATH, "meta.xlsx"), sheet = "techno") 

# Convert columns to numeric
data$Mean_Pre <- as.numeric(data$Mean_Pre)
data$SD_Pre <- as.numeric(data$SD_Pre)
data$Mean_Post <- as.numeric(data$Mean_Post)
data$SD_Post <- as.numeric(data$SD_Post)

# Function to calculate Cohen's d
calculate_cohens_d <- function(mean_pre, sd_pre, n_pre, mean_post, sd_post, n_post) {
  pooled_sd <- sqrt(((n_pre - 1) * sd_pre^2 + (n_post - 1) * sd_post^2) / (n_pre + n_post - 2))
  d <- (mean_post - mean_pre) / pooled_sd
  return(d)
}

# Calculate effect sizes
data$d_cohen <- mapply(calculate_cohens_d, 
                                 data$Mean_Pre, data$SD_Pre, data$N_Pre,
                                 data$Mean_Post, data$SD_Post, data$N_Post)

##################################################################################################
################################# MODE : TECHNOLOGY ##############################################
##################################################################################################

# Create simplified Mode column
data$Mode <- case_when(
  grepl('Immersive Video|Immersive Role-Play', data$TechnoType) ~ 'Virtual Reality',
  grepl('Mobile App for Attention', data$TechnoType) ~ 'Well-Being App',
  grepl('Mobile App for Enviro', data$TechnoType) ~ 'Citizen Science App',
  grepl('Photography', data$TechnoType) ~ 'Photography',
  grepl('Responsive', data$TechnoType) ~ 'Responsive Instrument',
  TRUE ~ data$TechnoType
)


# Filter modes with >= 2 studies
mode_counts <- table(data$Mode)
modes_to_keep <- names(mode_counts[mode_counts >= 2])
df_filtered <- data[data$Mode %in% modes_to_keep, ]

# Calculate mode statistics
mode_stats_filtered <- df_filtered %>%
  group_by(Mode) %>%
  summarise(
    n_studies = n(),
    mean_effect = mean(d_cohen, na.rm = TRUE),
    se_effect = sd(d_cohen, na.rm = TRUE) / sqrt(n()),
    ci_lower = mean_effect - 1.96 * se_effect,
    ci_upper = mean_effect + 1.96 * se_effect,
    .groups = 'drop'
  ) %>%
  arrange(desc(mean_effect))

# Order modes by effect size
mode_order_filtered <- mode_stats_filtered$Mode
df_filtered$Mode <- factor(df_filtered$Mode, levels = mode_order_filtered)

# Define colors
tech_colors_fill_filtered <- c(
  'Virtual Reality' = '#FE9AAA',
  'Citizen Science App' = '#BAC55D',
  'Well-Being App' = "#E1B566"
)

tech_colors_points_filtered <- c(
  'Virtual Reality' = '#FE9AAA',
  'Citizen Science App' = '#BAC55D',
  'Well-Being App' = "#E1B566"
)


# Create the boxplot
boxplot_updated <- ggplot(df_filtered, aes(x = Mode, y = d_cohen, fill = Mode)) +
  
  # Boxplots
  geom_boxplot(alpha = 0.6, outlier.shape = NA, width = 0.6) +
  
  # Individual study points
  geom_jitter(aes(color = Mode), width = 0.15, size = 2.5, alpha = 0.8) +
  
  # Mean diamonds
  geom_point(data = mode_stats_filtered, aes(x = Mode, y = mean_effect), 
             shape = 18, size = 4, color = 'gray30', inherit.aes = FALSE) +
  
  # Confidence intervals
  geom_errorbar(data = mode_stats_filtered, 
                aes(x = Mode, ymin = ci_lower, ymax = ci_upper), 
                width = 0.2, linewidth = 1, color = 'gray30', inherit.aes = FALSE) +
  
  # Reference line at 0
  geom_hline(yintercept = 0, linetype = 'dashed', color = 'gray50', alpha = 0.8) +
  
  # Study count annotations
  geom_text(data = mode_stats_filtered, 
            aes(x = Mode, y = 1.2, label = paste0('n=', n_studies, " groups")), 
            size = 3.5, fontface = 'bold', inherit.aes = FALSE) +
  
  # Apply colors
  scale_fill_manual(values = tech_colors_fill_filtered) +
  scale_color_manual(values = tech_colors_points_filtered) +
  
  # Labels and theme
  labs(title = 'Effect of the Type of Technology on Nature Connectedness',
       subtitle = 'Modes with n ≥ 2 experimental groups only',
       x = '',
       y = 'Effect Size (Cohen\'s d)') +
  
  theme_classic() +
  theme(
    plot.title = element_text(size = 12, face = 'bold', hjust = 0.5),
    plot.subtitle = element_text(size = 10, hjust = 0.5, color = 'gray40'),
    axis.title = element_text(size = 10, face = 'bold'),
    axis.text = element_text(size = 8),
    axis.text.x = element_text(angle = 0, hjust = 0.5),
    legend.position = 'bottom',
    panel.grid.major.y = element_line(color = 'gray90', linewidth = 0.3),
    panel.background = element_rect(fill = 'white'),
    plot.background = element_rect(fill = 'white')
  ) +
  
  scale_y_continuous(limits = c(-1.0, 1.2), breaks = seq(-1.0, 1.0, 0.2))

# Display and save the plot
print(boxplot_updated)
ggsave('technology_mode_boxplot.png', boxplot_updated, 
       width = 10, height = 6, dpi = 600, bg = 'white')


##################################################################################################
################################# MODE : TECHNOLOGY ##############################################
##################################################################################################



