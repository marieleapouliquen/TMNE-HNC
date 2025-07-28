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

# Calculer Cohen's d et sa variance avec metafor
#data <- escalc(measure="SMCR", 
#                     m1i=Mean_Post, sd1i=SD_Post, n1i=N_Post,
#                     m2i=Mean_Pre, sd2i=SD_Pre, n2i=N_Pre,
#                     data=data)

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

# Créer les sous-groupes par échelle HNC
hnc_scales <- unique(data$HNC)
hnc_results <- list()

for(scale in hnc_scales) {
  if(sum(data$HNC == scale, na.rm = TRUE) >= 2) {  # Au moins 2 études
    subset_data <- data[data$HNC == scale & !is.na(data$HNC), ]
    
    if(nrow(subset_data) >= 2) {
      meta_hnc <- rma(yi = d_manual, vi = var_d, data = subset_data, method = "REML")
      hnc_results[[scale]] <- meta_hnc
      
      print(paste("--- ÉCHELLE:", scale, "(n =", nrow(subset_data), "études) ---"))
      print(paste("Effect size poolé:", round(meta_hnc$beta, 3)))
      print(paste("IC 95%: [", round(meta_hnc$ci.lb, 3), ";", round(meta_hnc$ci.ub, 3), "]"))
      print(paste("p-value:", round(meta_hnc$pval, 4)))
      print(paste("I² =", round(meta_hnc$I2, 1), "%"))
      print(paste("τ² =", round(meta_hnc$tau2, 4)))
      print("")
    }
  }
}


# Filtrer les échelles HNC avec au moins 2 études
hnc_counts <- table(data$HNC)
hnc_to_keep <- names(hnc_counts[hnc_counts >= 2])
df_hnc_filtered <- data[data$HNC %in% hnc_to_keep, ]


# Créer le dataframe des statistiques HNC à partir des résultats REML
hnc_stats <- data.frame(
  HNC = c("CNS", "INS", "NRS", "NR-6"),
  n_studies = c(6, 7, 4, 6),
  mean_effect = c(0.399, 0.355, 0.451, 0.004),
  ci_lower = c(0.272, 0.135, 0.07, -0.311),
  ci_upper = c(0.527, 0.575, 0.833, 0.318),
  pval = c(0.0000, 0.0016, 0.0203, 0.9823),
  I2 = c(0.0, 89.4, 59.4, 94.4),
  tau2 = c(0.0000, 0.0733, 0.0893, 0.1373)
)


############################################################################################
######################################## PLOT ##############################################
############################################################################################


# Ordonner les échelles HNC par effet moyen décroissant
hnc_order <- hnc_stats$HNC[order(hnc_stats$mean_effect, decreasing = TRUE)]
df_hnc_filtered$HNC <- factor(df_hnc_filtered$HNC, levels = hnc_order)

# Définir les couleurs pour les échelles HNC
hnc_colors_fill <- c(
  "CNS" = '#4daf8c',
  "INS" = '#729ece', 
  "NRS" = '#ff9d6c',
  "NR-6" = '#F6CD03'
)

hnc_colors_points <- c(
  "CNS" = '#4daf8c',
  "INS" = '#729ece',
  "NRS" = '#ff9d6c', 
  "NR-6" = '#F6CD03'
)


# Créer le boxplot pour les échelles HNC avec résultats REML
boxplot_hnc <- ggplot(df_hnc_filtered, aes(x = HNC, y = d_manual, fill = HNC)) +
  
  # Boxplots
  geom_boxplot(alpha = 0.6, outlier.shape = NA, width = 0.6) +
  
  # Points individuels des études
  geom_jitter(aes(color = HNC), width = 0.15, size = 2.5, alpha = 0.8) +
  
  # Moyennes avec diamants noirs (résultats REML)
  geom_point(data = hnc_stats, aes(x = HNC, y = mean_effect), 
             shape = 18, size = 4, color = "gray30", inherit.aes = FALSE) +
  
  # Intervalles de confiance pour les moyennes (REML)
  geom_errorbar(data = hnc_stats, 
                aes(x = HNC, ymin = ci_lower, ymax = ci_upper), 
                width = 0.2, linewidth = 1, color = "gray30", inherit.aes = FALSE) +
  
  # Ligne de référence à 0
  geom_hline(yintercept = 0, linetype = "dashed", color = "gray50", alpha = 0.8) +
  
  # Annotations avec nombre d'études
  geom_text(data = hnc_stats, 
            aes(x = HNC, y = 1.1, label = paste0("n=", n_studies," groups")), 
            size = 3, inherit.aes = FALSE) +
  
  # Annotations avec d et IC 95%
  geom_text(data = hnc_stats, 
            aes(x = HNC, y = 1.2, 
                label = paste0("d=", sprintf("%.3f", mean_effect))), 
            size = 3, fontface = "bold", inherit.aes = FALSE) +
  
  geom_text(data = hnc_stats, 
            aes(x = HNC, y = 1.15, 
                label = paste0("[", sprintf("%.3f", ci_lower), ";", sprintf("%.3f", ci_upper), "]")), 
            size = 2.8, inherit.aes = FALSE) +
  
  # Appliquer les couleurs
  scale_fill_manual(values = hnc_colors_fill) +
  scale_color_manual(values = hnc_colors_points) +
  
  labs(title = "B) Effect of Nature Connectedness Scale",
       subtitle = "REML meta-analysis results by measurement scale (≥2 groups)",
       x = "Nature Connectedness Scale",
       y = "Effect Size (Cohen's d)") +
  
  theme_classic() +
  theme(
    plot.title = element_text(size = 14, face = "bold", hjust = 0.5),
    plot.subtitle = element_text(size = 11, hjust = 0.5, color = "gray40"),
    axis.title = element_text(size = 12, face = "bold"),
    axis.text = element_text(size = 10),
    legend.position = "none",
    panel.grid.major.y = element_line(color = "gray90", linewidth = 0.3),
    panel.background = element_rect(fill = "white"),
    plot.background = element_rect(fill = "white")
  ) +
  
  scale_y_continuous(limits = c(-1.0, 1.2), breaks = seq(-1.0, 1.0, 0.2))

print(boxplot_hnc)

# Sauvegarder le graphique
ggsave("hncBoxplot.png", boxplot_hnc, 
       width = 10, height = 6, dpi = 600, bg = "white")







###########################################################################################
################################# FOREST HNC ##############################################
###########################################################################################


# Préparer les données pour le forest plot
forest_data <- data.frame(
  study = paste(data$Citation, data$GroupID, sep = " - "),
  effect = data$d_manual,
  lower = data$d_manual - 1.96 * sqrt(data$var_d),
  upper = data$d_manual + 1.96 * sqrt(data$var_d),
  hnc_scale = data$HNC,
  weight = 1/data$var_d
)

# Ordonner par échelle HNC et effet
forest_data <- forest_data[order(forest_data$hnc_scale, forest_data$effect), ]
forest_data$study_order <- 1:nrow(forest_data)

# Créer le forest plot
forest_plot <- ggplot(forest_data, aes(x = effect, y = study_order)) +
  
  # Points et intervalles de confiance
  geom_point(aes(size = weight, color = hnc_scale), alpha = 0.7) +
  geom_errorbarh(aes(xmin = lower, xmax = upper, color = hnc_scale), 
                 height = 0.3, alpha = 0.7) +
  
  # Ligne de référence à 0
  geom_vline(xintercept = 0, linetype = "dashed", color = "gray50") +
  
  # Effet global
  geom_vline(xintercept = meta_global$beta, linetype = "solid", 
             color = "dodgerblue", size = 1.2, alpha = 0.8) +
  
  # Étiquettes des études
  scale_y_continuous(breaks = forest_data$study_order,
                     labels = forest_data$study,
                     expand = c(0.02, 0.02)) +
  
  # Couleurs par échelle
  scale_color_manual(values = c("CNS" = '#4daf8c', "NR-6" = '#F6CD03', 
                                "INS" = '#729ece', "NRS" = '#ff9d6c')) +
  
  
  labs(title = "A) Effect Sizes of 23 Experimental Groups\nordered by Nature Connectedness Scale",
       subtitle = paste("Pooled Effect : d =", round(meta_global$beta, 2), 
                        "95% CI [", round(meta_global$ci.lb, 2), ";", 
                        round(meta_global$ci.ub, 2), "]"),
       x = "Effect Size (Cohen's d)",
       y = "Studies",
       color = "HNC Scale",
       size = "Weight") +
  
  theme_classic() +
  theme(
    plot.title = element_text(size = 12, face = "bold"),
    plot.subtitle = element_text(size = 11, color = "gray40"),
    axis.text.y = element_text(size = 8),
    axis.text.x = element_text(size = 10),
    axis.title = element_text(size = 12, face = "bold"),
    legend.position = "bottom"
  ) +
  
  annotate("text", x = 1.3, y =25, 
           label = "Pooled Effect: d=0.29", hjust = 1, size = 3, fontface = "bold", color="dodgerblue") +
  
  xlim(-1.5, 1.5) +
  
  guides(
    color = guide_legend(nrow = 2, byrow = TRUE),
    size = guide_legend(nrow = 2, byrow = TRUE)
  )

print(forest_plot)

# Sauvegarder
ggsave("forestHNC.png", forest_plot, 
       width = 14, height = 10, dpi = 300, bg = "white")



