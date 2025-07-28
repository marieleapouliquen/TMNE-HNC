###############################################################################################
#### LOAD PACKAGES #### -----------------------------------------------------------------------
###############################################################################################

library(ggplot2)
library(metafor)
library(dplyr)

###############################################################################################
#### IMPORT DATASET #### ----------------------------------------------------------------------
###############################################################################################

# Charger les données (remplacez par votre méthode de chargement)
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

# Effectuer la méta-analyse REML pour chaque type de Nature
nature_types <- c("Real", "Simulated")
nature_results <- list()

for(nature_type in nature_types) {
  # Filtrer les données pour ce type
  data_subset <- data[data$Nature == nature_type, ]
  
  print(paste("=== Analyse REML pour Nature:", nature_type, "==="))
  print(paste("Nombre d'études:", nrow(data_subset)))
  
  if(nrow(data_subset) >= 2) {
    # Effectuer la méta-analyse REML
    res <- rma(yi = d_manual, vi = var_d, data = data_subset, method = "REML")
    
    # Stocker les résultats
    nature_results[[nature_type]] <- list(
      n_studies = nrow(data_subset),
      mean_effect = as.numeric(res$beta),
      se = as.numeric(res$se),
      ci_lower = as.numeric(res$ci.lb),
      ci_upper = as.numeric(res$ci.ub),
      pval = as.numeric(res$pval),
      I2 = as.numeric(res$I2),
      tau2 = as.numeric(res$tau2)
    )
    
    print(paste("Effet moyen (d):", round(as.numeric(res$beta), 3)))
    print(paste("IC 95%: [", round(as.numeric(res$ci.lb), 3), ";", round(as.numeric(res$ci.ub), 3), "]"))
    print(paste("p-value:", round(as.numeric(res$pval), 4)))
    print(paste("I²:", round(as.numeric(res$I2), 1), "%"))
    print(paste("τ²:", round(as.numeric(res$tau2), 4)))
  } else {
    print("Pas assez d'études pour la méta-analyse")
  }
  print("")
}

# Créer le dataframe des statistiques Nature à partir des résultats REML
nature_stats <- data.frame(
  Nature = c("Simulated", "Real"),  # Ordonné par effet décroissant
  n_studies = c(nature_results[["Simulated"]]$n_studies, nature_results[["Real"]]$n_studies),
  mean_effect = c(nature_results[["Simulated"]]$mean_effect, nature_results[["Real"]]$mean_effect),
  ci_lower = c(nature_results[["Simulated"]]$ci_lower, nature_results[["Real"]]$ci_lower),
  ci_upper = c(nature_results[["Simulated"]]$ci_upper, nature_results[["Real"]]$ci_upper),
  pval = c(nature_results[["Simulated"]]$pval, nature_results[["Real"]]$pval),
  I2 = c(nature_results[["Simulated"]]$I2, nature_results[["Real"]]$I2),
  tau2 = c(nature_results[["Simulated"]]$tau2, nature_results[["Real"]]$tau2)
)

print("Statistiques REML pour les types de Nature:")
print(nature_stats)


############################################################################################
######################################## PLOT ##############################################
############################################################################################

# Filtrer les données pour le graphique
df_nature_filtered <- data[data$Nature %in% c("Real", "Simulated"), ]

# Ordonner par effet décroissant (Simulated > Real)
#df_nature_filtered$Nature <- factor(df_nature_filtered$Nature, levels = c("Simulated", "Real"))

# Définir les couleurs pour les types de Nature
nature_colors_fill <- c(
  "Simulated" = "skyblue",
  "Real" = "forestgreen"
)

nature_colors_points <- c(
  "Simulated" = "darkblue",
  "Real" = "darkgreen"
)

# Créer le boxplot pour les types de Nature avec résultats REML
boxplot_nature <- ggplot(df_nature_filtered, aes(x = Nature, y = d_manual, fill = Nature)) +
  
  # Boxplots
  geom_boxplot(alpha = 0.6, outlier.shape = NA, width = 0.6) +
  
  # Points individuels des études
  geom_jitter(aes(color = Nature), width = 0.15, size = 2.5, alpha = 0.8) +
  
  # Moyennes avec diamants noirs (résultats REML)
  geom_point(data = nature_stats, aes(x = Nature, y = mean_effect), 
             shape = 18, size = 4, color = "gray30", inherit.aes = FALSE) +
  
  # Intervalles de confiance pour les moyennes (REML)
  geom_errorbar(data = nature_stats, 
                aes(x = Nature, ymin = ci_lower, ymax = ci_upper), 
                width = 0.2, linewidth = 1, color = "gray30", inherit.aes = FALSE) +
  
  # Ligne de référence à 0
  geom_hline(yintercept = 0, linetype = "dashed", color = "gray50", alpha = 0.8) +
  
  # Annotations avec nombre d'études
  geom_text(data = nature_stats, 
            aes(x = Nature, y = 1.1, label = paste0("n=", n_studies,"groups")), 
            size = 3, inherit.aes = FALSE) +
  
  # Annotations avec d et IC 95%
  geom_text(data = nature_stats, 
            aes(x = Nature, y = 1.2, 
                label = paste0("d=", sprintf("%.3f", mean_effect))), 
            size = 3, fontface = "bold", inherit.aes = FALSE) +
  
  geom_text(data = nature_stats, 
            aes(x = Nature, y = 1.15, 
                label = paste0("[", sprintf("%.3f", ci_lower), ";", sprintf("%.3f", ci_upper), "]")), 
            size = 2.8, inherit.aes = FALSE) +
  
  # Appliquer les couleurs
  scale_fill_manual(values = nature_colors_fill) +
  scale_color_manual(values = nature_colors_points) +
  
  labs(title = "A) Effect of Nature Ontological Status",
       subtitle = "REML meta-analysis results by nature ontological status (≥2 studies)",
       x = "Nature Ontological Status",
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

print(boxplot_nature)

# Sauvegarder le graphique
ggsave("nature_type_boxplot_reml.png", boxplot_nature, 
       width = 10, height = 6, dpi = 600, bg = "white")

print("Graphique sauvegardé: nature_type_boxplot_reml.png")
