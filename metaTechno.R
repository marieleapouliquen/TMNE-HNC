###############################################################################################
#### LOAD PACKAGES #### -----------------------------------------------------------------------
###############################################################################################

library(ggplot2)
library(metafor)
library(dplyr)
library(readxl)

###############################################################################################
#### IMPORT DATASET #### ----------------------------------------------------------------------
###############################################################################################

PATH="/Users/marieleapouliquen/Desktop/Review/Data/"
data <- read_excel(file.path(PATH, "meta.xlsx"), sheet = "techno") 

# Convertir les colonnes en numérique
data$Mean_Pre <- as.numeric(data$Mean_Pre)
data$SD_Pre <- as.numeric(data$SD_Pre)
data$Mean_Post <- as.numeric(data$Mean_Post)
data$SD_Post <- as.numeric(data$SD_Post)
data$N_Pre <- as.numeric(data$N_Pre)
data$N_Post <- as.numeric(data$N_Post)

# Calculer manuellement Cohen's d pour repeated measures et sa variance
data$d_manual <- (data$Mean_Post - data$Mean_Pre) / 
  sqrt((data$SD_Pre^2 + data$SD_Post^2) / 2)

# Calculer la variance approximative pour Cohen's d (repeated measures)
# Formule: V = (1/n) + (d^2 / (2*n))
data$var_d <- (1/data$N_Pre) + (data$d_manual^2 / (2*data$N_Pre))

# Effectuer la méta-analyse avec REML
# Méta-analyse globale
meta_global <- rma(yi = d_manual, vi = var_d, data = data, method = "REML")

# Intervalle de prédiction
pred_int <- predict(meta_global)


############################################################################################
######################################## MODE ##############################################
############################################################################################

# Créer une colonne Mode simplifiée
data$Mode <- case_when(
  grepl('Immersive Video|Immersive Role-Play', data$TechnoType) ~ 'Virtual Reality',
  grepl('Mobile App for Attention', data$TechnoType) ~ 'Well-Being App',
  grepl('Mobile App for Enviro', data$TechnoType) ~ 'Citizen Science App',
  grepl('Photography', data$TechnoType) ~ 'Photography',
  grepl('Responsive', data$TechnoType) ~ 'Responsive Instrument',
  TRUE ~ data$TechnoType
)

# Filtrer les modes avec au moins 2 études
mode_counts <- table(data$Mode)
modes_to_keep <- names(mode_counts[mode_counts >= 2])
df_mode_filtered <- data[data$Mode %in% modes_to_keep, ]

print("Modes avec ≥2 études:")
print(mode_counts[mode_counts >= 2])
print("")
print(paste("Nombre d'études après filtrage:", nrow(df_mode_filtered)))


# Effectuer des méta-analyses REML par mode technologique
print("=== MÉTA-ANALYSES PAR MODE TECHNOLOGIQUE (REML) ===")

mode_results <- list()
mode_stats <- data.frame()

for(mode in modes_to_keep) {
  subset_data <- df_mode_filtered[df_mode_filtered$Mode == mode, ]
  
  if(nrow(subset_data) >= 2) {
    meta_mode <- rma(yi = d_manual, vi = var_d, data = subset_data, method = "REML")
    mode_results[[mode]] <- meta_mode
    
    # Calculer les statistiques pour le boxplot
    mode_stat <- data.frame(
      Mode = mode,
      n_studies = nrow(subset_data),
      mean_effect = meta_mode$beta[1],
      se_effect = meta_mode$se,
      ci_lower = meta_mode$ci.lb,
      ci_upper = meta_mode$ci.ub,
      pval = meta_mode$pval,
      I2 = meta_mode$I2,
      tau2 = meta_mode$tau2
    )
    
    mode_stats <- rbind(mode_stats, mode_stat)
    
    print(paste("--- MODE:", mode, "(n =", nrow(subset_data), "études) ---"))
    print(paste("Effect size poolé:", round(meta_mode$beta, 3)))
    print(paste("IC 95%: [", round(meta_mode$ci.lb, 3), ";", round(meta_mode$ci.ub, 3), "]"))
    print(paste("p-value:", round(meta_mode$pval, 4)))
    print(paste("I² =", round(meta_mode$I2, 1), "%"))
    print(paste("τ² =", round(meta_mode$tau2, 4)))
    print("")
  }
}

print("Statistiques des modes:")
print(mode_stats)


############################################################################################
######################################## PLOT ##############################################
############################################################################################

# Ordonner les modes par effet moyen décroissant
mode_order <- mode_stats$Mode[order(mode_stats$mean_effect, decreasing = TRUE)]
df_mode_filtered$Mode <- factor(df_mode_filtered$Mode, levels = mode_order)

# Définir les couleurs pour les modes technologiques
mode_colors_fill <- c(
  'Virtual Reality' = '#FE9AAA',
  'Citizen Science App' = '#BAC55D',
  'Well-Being App' = "#E1B566"
)

mode_colors_points <- c(
  'Virtual Reality' = '#FE9AAA',
  'Citizen Science App' = '#BAC55D',
  'Well-Being App' = "#E1B566"
)

# Créer le boxplot pour les modes technologiques
boxplot_mode <- ggplot(df_mode_filtered, aes(x = Mode, y = d_manual, fill = Mode)) +
  
  # Boxplots
  geom_boxplot(alpha = 0.6, outlier.shape = NA, width = 0.6) +
  
  # Points individuels des études
  geom_jitter(aes(color = Mode), width = 0.15, size = 2.5, alpha = 0.8) +
  
  # Moyennes avec diamants noirs (utiliser les résultats REML)
  geom_point(data = mode_stats, aes(x = Mode, y = mean_effect), 
             shape = 18, size = 4, color = "gray30", inherit.aes = FALSE) +
  
  # Intervalles de confiance pour les moyennes (REML)
  geom_errorbar(data = mode_stats, 
                aes(x = Mode, ymin = ci_lower, ymax = ci_upper), 
                width = 0.2, linewidth = 1, color = "gray30", inherit.aes = FALSE) +
  
  # Ligne de référence à 0
  geom_hline(yintercept = 0, linetype = "dashed", color = "gray50", alpha = 0.8) +
  
  # Annotations avec nombre d'études
  geom_text(data = mode_stats, 
            aes(x = Mode, y = 1.1, label = paste0("n=", n_studies," groups")), 
            size = 3, inherit.aes = FALSE) +
  
  # Annotations avec d et IC 95%
  geom_text(data = mode_stats, 
            aes(x = Mode, y = 1.2, 
                label = paste0("d=", sprintf("%.3f", mean_effect))), 
            size = 3, fontface = "bold", inherit.aes = FALSE) +
  
  geom_text(data = mode_stats, 
            aes(x = Mode, y = 1.15, 
                label = paste0("[", sprintf("%.3f", ci_lower), ";", sprintf("%.3f", ci_upper), "]")), 
            size = 2.8, inherit.aes = FALSE) +
  
  
  # Appliquer les couleurs
  scale_fill_manual(values = mode_colors_fill) +
  scale_color_manual(values = mode_colors_points) +
  
  labs(title = "B) Effect of Technology Type",
       subtitle = "REML meta-analysis results by technology type (≥2 groups)",
       x = "Technology",
       y = "Effect Size (Cohen's d)") +
  
  theme_classic() +
  theme(
    plot.title = element_text(size = 14, face = "bold", hjust = 0.5),
    plot.subtitle = element_text(size = 11, hjust = 0.5, color = "gray40"),
    axis.title = element_text(size = 12, face = "bold"),
    axis.text = element_text(size = 10),
    axis.text.x = element_text(angle = 0, hjust = 0.5),
    legend.position = "none",
    panel.grid.major.y = element_line(color = "gray90", linewidth = 0.3),
    panel.background = element_rect(fill = "white"),
    plot.background = element_rect(fill = "white")
  ) +
  
  scale_y_continuous(limits = c(-1.0, 1.2), breaks = seq(-1.0, 1.0, 0.2))

print(boxplot_mode)

# Sauvegarder le graphique
ggsave("technology_mode_boxplot_reml.png", boxplot_mode, 
       width = 10, height = 6, dpi = 600, bg = "white")

print("")
print("Graphique Mode Technologique sauvegardé: technology_mode_boxplot_reml.png")

