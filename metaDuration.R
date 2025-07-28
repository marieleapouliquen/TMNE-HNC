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


# Créer des groupes d'interventions basés sur la durée
print("=== REGROUPEMENT DES INTERVENTIONS PAR TYPE DE DURÉE ===")
print("")


# Créer les groupes d'interventions basés sur la durée
data$Duration_Group <- NA
data$Duration_Group[data$Duration %in% c("10 min", "5-7 min")] <- "Short Single (~10 min)"
data$Duration_Group[data$Duration %in% c("< 1d", "< 1-2h")] <- "Single Day (<1d)"
data$Duration_Group[data$Duration %in% c("1/d x 7d", "20 min/d x 7d", "7 min/d x 7d", "4d/wk x 3wk")] <- "Repeated"

# Effectuer la méta-analyse REML pour chaque groupe
duration_group_types <- c("Short Single (~10 min)", "Single Day (<1d)", "Repeated")
duration_group_results <- list()

for(group_type in duration_group_types) {
  data_subset <- data[data$Duration_Group == group_type & !is.na(data$Duration_Group), ]
  
  if(nrow(data_subset) >= 2) {
    res <- rma(yi = d_manual, vi = var_d, data = data_subset, method = "REML")
    
    duration_group_results[[group_type]] <- list(
      n_studies = nrow(data_subset),
      mean_effect = as.numeric(res$beta),
      se = as.numeric(res$se),
      ci_lower = as.numeric(res$ci.lb),
      ci_upper = as.numeric(res$ci.ub),
      pval = as.numeric(res$pval),
      I2 = as.numeric(res$I2),
      tau2 = as.numeric(res$tau2)
    )
  }
}

# Créer le dataframe des statistiques
duration_group_stats <- data.frame(
  Duration_Group = c("Single Day (<1d)", "Short Single (~10 min)", "Repeated"),
  n_studies = c(
    duration_group_results[["Single Day (<1d)"]]$n_studies,
    duration_group_results[["Short Single (~10 min)"]]$n_studies,
    duration_group_results[["Repeated"]]$n_studies
  ),
  mean_effect = c(
    duration_group_results[["Single Day (<1d)"]]$mean_effect,
    duration_group_results[["Short Single (~10 min)"]]$mean_effect,
    duration_group_results[["Repeated"]]$mean_effect
  ),
  ci_lower = c(
    duration_group_results[["Single Day (<1d)"]]$ci_lower,
    duration_group_results[["Short Single (~10 min)"]]$ci_lower,
    duration_group_results[["Repeated"]]$ci_lower
  ),
  ci_upper = c(
    duration_group_results[["Single Day (<1d)"]]$ci_upper,
    duration_group_results[["Short Single (~10 min)"]]$ci_upper,
    duration_group_results[["Repeated"]]$ci_upper
  ),
  pval = c(
    duration_group_results[["Single Day (<1d)"]]$pval,
    duration_group_results[["Short Single (~10 min)"]]$pval,
    duration_group_results[["Repeated"]]$pval
  ),
  I2 = c(
    duration_group_results[["Single Day (<1d)"]]$I2,
    duration_group_results[["Short Single (~10 min)"]]$I2,
    duration_group_results[["Repeated"]]$I2
  ),
  tau2 = c(
    duration_group_results[["Single Day (<1d)"]]$tau2,
    duration_group_results[["Short Single (~10 min)"]]$tau2,
    duration_group_results[["Repeated"]]$tau2
  )
)


############################################################################################
######################################## PLOT ##############################################
############################################################################################


# Préparer les données pour le graphique
df_duration_group_filtered <- data[!is.na(data$Duration_Group), ]
df_duration_group_filtered$Duration_Group <- factor(df_duration_group_filtered$Duration_Group, 
                                                    levels = c("Single Day (<1d)", "Short Single (~10 min)", "Repeated"))

# Définir les couleurs
duration_group_colors_fill <- c(
  "Single Day (<1d)" = "deepskyblue",
  "Short Single (~10 min)" = "lightskyblue", 
  "Repeated" = "powderblue"
)

duration_group_colors_points <- c(
  "Single Day (<1d)" = "skyblue",
  "Short Single (~10 min)" = "steelblue",
  "Repeated" = "cadetblue"
)

# Créer le boxplot
boxplot_duration_groups <- ggplot(df_duration_group_filtered, aes(x = Duration_Group, y = d_manual, fill = Duration_Group)) +
  geom_boxplot(alpha = 0.6, outlier.shape = NA, width = 0.6) +
  
  geom_jitter(aes(color = Duration_Group), width = 0.15, size = 2.5, alpha = 0.8) +
  
  geom_point(data = duration_group_stats, aes(x = Duration_Group, y = mean_effect), 
             shape = 18, size = 4, color = "grey30", inherit.aes = FALSE) +
  
  geom_errorbar(data = duration_group_stats, 
                aes(x = Duration_Group, ymin = ci_lower, ymax = ci_upper), 
                width = 0.2, linewidth = 1, color = "grey30", inherit.aes = FALSE) +
  
  geom_hline(yintercept = 0, linetype = "dashed", color = "gray50", alpha = 0.8) +
  
  geom_text(data = duration_group_stats, 
            aes(x = Duration_Group, y = 1.1, label = paste0("n=", n_studies, " groups")), 
            size = 3, inherit.aes = FALSE) +
  
  geom_text(data = duration_group_stats, 
            aes(x = Duration_Group, y = 1.2, 
                label = paste0("d=", sprintf("%.3f", mean_effect))), 
            size = 3, fontface = "bold", inherit.aes = FALSE) +
  
  geom_text(data = duration_group_stats, 
            aes(x = Duration_Group, y = 1.15, 
                label = paste0("[", sprintf("%.3f", ci_lower), ";", sprintf("%.3f", ci_upper), "]")), 
            size = 2.8, inherit.aes = FALSE) +
  
  scale_fill_manual(values = duration_group_colors_fill) +
  
  scale_color_manual(values = duration_group_colors_points) +
  
  labs(title = "C) Effect of Duration and Frequency",
       subtitle = "REML meta-analysis results by duration/frequency (≥2 studies)",
       x = "Intervention Duration and Frequency",
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

# Afficher et sauvegarder le graphique
print(boxplot_duration_groups)
ggsave("durationboxplot_reml.png", boxplot_duration_groups, 
       width = 10, height = 6, dpi = 600, bg = "white")

print("Code complet exécuté avec succès!")

