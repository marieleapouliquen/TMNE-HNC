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

print("=== MÉTA-ANALYSE REML POUR LES TYPES DE POPULATION ===")
print("")

population_types <- c("University Studients", "General Public")
population_results <- list()

for(pop_type in population_types) {
  # Filtrer les données pour ce type
  data_subset <- data[data$Population == pop_type, ]
  
  print(paste("=== Analyse REML pour Population:", pop_type, "==="))
  print(paste("Nombre d'études:", nrow(data_subset)))
  
  if(nrow(data_subset) >= 2) {
    # Effectuer la méta-analyse REML
    res <- rma(yi = d_manual, vi = var_d, data = data_subset, method = "REML")
    
    # Stocker les résultats
    population_results[[pop_type]] <- list(
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

# Créer le dataframe des statistiques Population à partir des résultats REML
population_stats <- data.frame(
  Population = c("General Public", "University Studients"),  # Ordonné par effet décroissant
  n_studies = c(population_results[["General Public"]]$n_studies, population_results[["University Studients"]]$n_studies),
  mean_effect = c(population_results[["General Public"]]$mean_effect, population_results[["University Studients"]]$mean_effect),
  ci_lower = c(population_results[["General Public"]]$ci_lower, population_results[["University Studients"]]$ci_lower),
  ci_upper = c(population_results[["General Public"]]$ci_upper, population_results[["University Studients"]]$ci_upper),
  pval = c(population_results[["General Public"]]$pval, population_results[["University Studients"]]$pval),
  I2 = c(population_results[["General Public"]]$I2, population_results[["University Studients"]]$I2),
  tau2 = c(population_results[["General Public"]]$tau2, population_results[["University Studients"]]$tau2)
)

print("Statistiques REML pour les types de Population:")
print(population_stats)


############################################################################################
######################################## PLOT ##############################################
############################################################################################

# Créer le boxplot pour les types de Population
# Filtrer les données pour le graphique
df_population_filtered <- df_updated[df_updated$Population %in% c("General Public", "University Studients"), ]

# Ordonner par effet décroissant (General Public > University Students)
df_population_filtered$Population <- factor(df_population_filtered$Population, 
                                            levels = c("General Public", "University Studients"))

# Définir les couleurs pour les types de Population
population_colors_fill <- c(
  "General Public" = "lightcoral",
  "University Studients" = "sandybrown"
)

population_colors_points <- c(
  "General Public" = "brown",
  "University Studients" = "saddlebrown"
)

# Créer le boxplot pour les types de Population avec résultats REML
boxplot_population <- ggplot(df_population_filtered, aes(x = Population, y = d_manual, fill = Population)) +
  
  # Boxplots
  geom_boxplot(alpha = 0.6, outlier.shape = NA, width = 0.6) +
  
  # Points individuels des études
  geom_jitter(aes(color = Population), width = 0.15, size = 2.5, alpha = 0.8) +
  
  # Moyennes avec diamants noirs (résultats REML)
  geom_point(data = population_stats, aes(x = Population, y = mean_effect), 
             shape = 18, size = 4, color = "black", inherit.aes = FALSE) +
  
  # Intervalles de confiance pour les moyennes (REML)
  geom_errorbar(data = population_stats, 
                aes(x = Population, ymin = ci_lower, ymax = ci_upper), 
                width = 0.2, linewidth = 1, color = "black", inherit.aes = FALSE) +
  
  # Ligne de référence à 0
  geom_hline(yintercept = 0, linetype = "dashed", color = "gray50", alpha = 0.8) +
  
  # Annotations avec nombre d'études
  geom_text(data = population_stats, 
            aes(x = Population, y = 1.1, label = paste0("n=", n_studies,"groups")), 
            size = 3, inherit.aes = FALSE) +
  
  # Annotations avec d et IC 95%
  geom_text(data = population_stats, 
            aes(x = Population, y = 1.2, 
                label = paste0("d=", sprintf("%.3f", mean_effect))), 
            size = 3, fontface = "bold", inherit.aes = FALSE) +
  
  geom_text(data = population_stats, 
            aes(x = Population, y = 1.15, 
                label = paste0("[", sprintf("%.3f", ci_lower), ";", sprintf("%.3f", ci_upper), "]")), 
            size = 2.8, inherit.aes = FALSE) +
  
  # Appliquer les couleurs
  scale_fill_manual(values = population_colors_fill) +
  scale_color_manual(values = population_colors_points) +
  
  labs(title = "D) Effect of Target Population",
       subtitle = "REML meta-analysis results by target population (≥2 studies)",
       x = "Population",
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

print(boxplot_population)