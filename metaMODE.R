###############################################################################################
#### LOAD PACKAGES #### -----------------------------------------------------------------------
###############################################################################################

library(ggplot2)
library(metafor)
library(dplyr)
library(readxl)
#install.packages("openxlsx", repos="https://cran.rstudio.com/", dependencies = TRUE)
library(openxlsx)


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
######################################## TECHNO ############################################
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


############################################################################################
######################################## NATURE ############################################
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



############################################################################################
######################################## POPULATION ########################################
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
######################################## DURATION ##########################################
############################################################################################


# Créer des groupes d'interventions basés sur la durée
print("=== REGROUPEMENT DES INTERVENTIONS PAR TYPE DE DURÉE ===")
print("")


# Créer les groupes d'interventions basés sur la durée
data$Duration_Group <- NA
data$Duration_Group[data$Duration %in% c("10 min", "5-7 min")] <- "Short Single (~10 min)"
data$Duration_Group[data$Duration %in% c("< 1d", "< 1-2h")] <- "Single Day (<1d)"
data$Duration_Group[data$Duration %in% c("1/d x 7d", "20 min/d x 7d", "7 min/d x 7d")] <- "Repeated (daily x 7d)"

# Effectuer la méta-analyse REML pour chaque groupe
duration_group_types <- c("Short Single (~10 min)", "Single Day (<1d)", "Repeated (daily x 7d)")
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
  Duration_Group = c("Single Day (<1d)", "Short Single (~10 min)", "Repeated (daily x 7d)"),
  n_studies = c(
    duration_group_results[["Single Day (<1d)"]]$n_studies,
    duration_group_results[["Short Single (~10 min)"]]$n_studies,
    duration_group_results[["Repeated (daily x 7d)"]]$n_studies
  ),
  mean_effect = c(
    duration_group_results[["Single Day (<1d)"]]$mean_effect,
    duration_group_results[["Short Single (~10 min)"]]$mean_effect,
    duration_group_results[["Repeated (daily x 7d)"]]$mean_effect
  ),
  ci_lower = c(
    duration_group_results[["Single Day (<1d)"]]$ci_lower,
    duration_group_results[["Short Single (~10 min)"]]$ci_lower,
    duration_group_results[["Repeated (daily x 7d)"]]$ci_lower
  ),
  ci_upper = c(
    duration_group_results[["Single Day (<1d)"]]$ci_upper,
    duration_group_results[["Short Single (~10 min)"]]$ci_upper,
    duration_group_results[["Repeated (daily x 7d)"]]$ci_upper
  ),
  pval = c(
    duration_group_results[["Single Day (<1d)"]]$pval,
    duration_group_results[["Short Single (~10 min)"]]$pval,
    duration_group_results[["Repeated (daily x 7d)"]]$pval
  ),
  I2 = c(
    duration_group_results[["Single Day (<1d)"]]$I2,
    duration_group_results[["Short Single (~10 min)"]]$I2,
    duration_group_results[["Repeated (daily x 7d)"]]$I2
  ),
  tau2 = c(
    duration_group_results[["Single Day (<1d)"]]$tau2,
    duration_group_results[["Short Single (~10 min)"]]$tau2,
    duration_group_results[["Repeated (daily x 7d)"]]$tau2
  )
)



############################################################################################
######################################## HNC ###############################################
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
  n_studies = c(6, 7, 3, 6),
  mean_effect = c(0.399, 0.355, 0.340, 0.004),
  ci_lower = c(0.272, 0.135, -0.145, -0.311),
  ci_upper = c(0.527, 0.575, 0.824, 0.318),
  pval = c(0.0000, 0.0016, 0.1694, 0.9823),
  I2 = c(0.0, 89.4, 60.2, 94.4),
  tau2 = c(0.0000, 0.0733, 0.1103, 0.1373)
)


############################################################################################
######################################## ALL ###############################################
############################################################################################

# RÉSUMÉ COMPLET DES ANALYSES PAR MODE
# ====================================

print("=== RÉSUMÉ DES MÉTA-ANALYSES REML PAR MODE ===")
print("")

# Vérifier les variables disponibles dans le dataset
print("Variables disponibles dans le dataset:")
print(names(data))
print("")

# Vérifier si HNC existe
if("HNC" %in% names(data)) {
  print("Variable HNC trouvée")
  print("Valeurs uniques de HNC:")
  print(table(data$HNC, useNA = "ifany"))
} else {
  print("Variable HNC non trouvée dans le dataset")
}
print("")

# Vérifier TechnologyCategory
if("TechnologyCategory" %in% names(data)) {
  print("Variable TechnologyCategory trouvée")
  print("Valeurs uniques de TechnologyCategory:")
  print(table(data$TechnologyCategory, useNA = "ifany"))
} else {
  print("Variable TechnologyCategory non trouvée dans le dataset")
}
print("")

# Vérifier les autres variables
print("Valeurs uniques de Nature:")
print(table(data$Nature, useNA = "ifany"))
print("")

print("Valeurs uniques de Population:")
print(table(data$Population, useNA = "ifany"))
print("")

print("Valeurs uniques de Duration:")
print(table(data$Duration, useNA = "ifany"))



# CRÉER UN RÉSUMÉ COMPLET DES ANALYSES PAR MODE
# =============================================


# Fonction pour effectuer une méta-analyse REML sur une variable
perform_meta_analysis <- function(data, variable_name) {
  # Obtenir les valeurs uniques de la variable
  unique_values <- unique(data[[variable_name]])
  unique_values <- unique_values[!is.na(unique_values)]
  
  results <- list()
  
  for(value in unique_values) {
    # Filtrer les données pour cette valeur
    subset_data <- data[data[[variable_name]] == value & !is.na(data[[variable_name]]), ]
    
    if(nrow(subset_data) >= 2) {
      # Effectuer la méta-analyse REML
      tryCatch({
        res <- rma(yi = d_manual, vi = var_d, data = subset_data, method = "REML")
        
        results[[as.character(value)]] <- list(
          n_studies = nrow(subset_data),
          mean_effect = round(as.numeric(res$beta), 3),
          ci_lower = round(as.numeric(res$ci.lb), 3),
          ci_upper = round(as.numeric(res$ci.ub), 3),
          pval = as.numeric(res$pval),
          I2 = round(as.numeric(res$I2), 1),
          tau2 = round(as.numeric(res$tau2), 4)
        )
      }, error = function(e) {
        results[[as.character(value)]] <- list(
          n_studies = nrow(subset_data),
          mean_effect = NA,
          ci_lower = NA,
          ci_upper = NA,
          pval = NA,
          I2 = NA,
          tau2 = NA,
          error = "Erreur dans l'analyse"
        )
      })
    } else {
      results[[as.character(value)]] <- list(
        n_studies = nrow(subset_data),
        mean_effect = NA,
        ci_lower = NA,
        ci_upper = NA,
        pval = NA,
        I2 = NA,
        tau2 = NA,
        note = "Moins de 2 études"
      )
    }
  }
  
  return(results)
}

# Analyser chaque mode
modes_to_analyze <- c("Nature", "Population", "Duration", "HNC")

all_results <- list()

for(mode in modes_to_analyze) {
  print(paste("=== ANALYSE DU MODE:", mode, "==="))
  all_results[[mode]] <- perform_meta_analysis(data, mode)
  
  # Afficher les résultats
  for(value in names(all_results[[mode]])) {
    result <- all_results[[mode]][[value]]
    print(paste("Catégorie:", value))
    print(paste("  Nombre d'études:", result$n_studies))
    
    if(!is.na(result$mean_effect)) {
      print(paste("  Effet moyen (d):", result$mean_effect))
      print(paste("  IC 95%: [", result$ci_lower, ";", result$ci_upper, "]"))
      print(paste("  p-value:", ifelse(result$pval < 0.001, "< 0.001", round(result$pval, 4))))
      print(paste("  I²:", result$I2, "%"))
      print(paste("  τ²:", result$tau2))
    } else {
      if(!is.null(result$note)) {
        print(paste("  Note:", result$note))
      }
      if(!is.null(result$error)) {
        print(paste("  Erreur:", result$error))
      }
    }
    print("")
  }
  print("")
}


# CRÉER UN TABLEAU RÉSUMÉ POUR TOUS LES MODES
# ===========================================

# Créer un dataframe résumé pour tous les modes analysés
summary_data <- data.frame(
  Mode = character(),
  Category = character(),
  n_studies = numeric(),
  mean_effect = numeric(),
  ci_lower = numeric(),
  ci_upper = numeric(),
  pval_formatted = character(),
  I2 = numeric(),
  tau2 = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

# Fonction pour formater la p-value
format_pval <- function(p) {
  if(is.na(p)) return("NA")
  if(p < 0.001) return("< 0.001")
  return(sprintf("%.4f", p))
}

# Fonction pour déterminer la significativité
get_significance <- function(p) {
  if(is.na(p)) return("NA")
  if(p < 0.001) return("***")
  if(p < 0.01) return("**")
  if(p < 0.05) return("*")
  return("ns")
}

# Remplir le dataframe avec tous les résultats
for(mode in names(all_results)) {
  for(category in names(all_results[[mode]])) {
    result <- all_results[[mode]][[category]]
    
    if(!is.na(result$mean_effect)) {
      summary_data <- rbind(summary_data, data.frame(
        Mode = mode,
        Category = category,
        n_studies = result$n_studies,
        mean_effect = result$mean_effect,
        ci_lower = result$ci_lower,
        ci_upper = result$ci_upper,
        pval_formatted = format_pval(result$pval),
        I2 = result$I2,
        tau2 = result$tau2,
        significance = get_significance(result$pval),
        stringsAsFactors = FALSE
      ))
    }
  }
}

# Ordonner par mode et effet décroissant
summary_data <- summary_data[order(summary_data$Mode, -summary_data$mean_effect), ]

print("=== TABLEAU RÉSUMÉ DES MÉTA-ANALYSES REML PAR MODE ===")
print(summary_data)


# COMPLÉTER LE TABLEAU RÉSUMÉ (il semble qu'il y ait eu une troncature)
# =====================================================================

# Recréer le tableau complet
summary_data_complete <- data.frame(
  Mode = character(),
  Category = character(),
  n_studies = numeric(),
  mean_effect = numeric(),
  ci_lower = numeric(),
  ci_upper = numeric(),
  pval_formatted = character(),
  I2 = numeric(),
  tau2 = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

# Remplir avec tous les résultats
for(mode in names(all_results)) {
  for(category in names(all_results[[mode]])) {
    result <- all_results[[mode]][[category]]
    
    if(!is.na(result$mean_effect)) {
      summary_data_complete <- rbind(summary_data_complete, data.frame(
        Mode = mode,
        Category = category,
        n_studies = result$n_studies,
        mean_effect = result$mean_effect,
        ci_lower = result$ci_lower,
        ci_upper = result$ci_upper,
        pval_formatted = format_pval(result$pval),
        I2 = result$I2,
        tau2 = result$tau2,
        significance = get_significance(result$pval),
        stringsAsFactors = FALSE
      ))
    }
  }
}

# Ordonner par mode et effet décroissant
summary_data_complete <- summary_data_complete[order(summary_data_complete$Mode, -summary_data_complete$mean_effect), ]

print("=== TABLEAU RÉSUMÉ COMPLET DES MÉTA-ANALYSES REML PAR MODE ===")
print(summary_data_complete)
print("")

# Créer un résumé par mode avec les catégories les plus efficaces
print("=== RÉSUMÉ PAR MODE - CATÉGORIES LES PLUS EFFICACES ===")
print("")

for(mode in unique(summary_data_complete$Mode)) {
  mode_data <- summary_data_complete[summary_data_complete$Mode == mode, ]
  mode_data <- mode_data[order(-mode_data$mean_effect), ]
  
  print(paste("MODE:", mode))
  print("Classement par efficacité (effet décroissant):")
  
  for(i in 1:nrow(mode_data)) {
    row <- mode_data[i, ]
    print(paste(i, ".", row$Category, 
                "- d =", row$mean_effect, 
                "[", row$ci_lower, ";", row$ci_upper, "]",
                "- n =", row$n_studies,
                "- p", row$pval_formatted,
                row$significance))
  }
  print("")
}



# SAUVEGARDER LE RÉSUMÉ COMPLET DANS UN FICHIER CSV
# =================================================

# Créer le tableau final avec toutes les données
final_summary <- data.frame(
  Mode = c("Nature", "Nature", "Population", "Population", "Duration", "Duration", "Duration", "Duration", "Duration", "HNC", "HNC", "HNC", "HNC"),
  Category = c("Simulated", "Real", "General Public", "University Students", "< 1d", "5-7 min", "10 min", "1/d x 7d", "20 min/d x 7d", "CNS", "INS", "NRS", "NR-6"),
  n_studies = c(5, 17, 13, 9, 2, 3, 3, 8, 4, 6, 7, 3, 6),
  mean_effect = c(0.430, 0.233, 0.250, 0.169, 0.500, 0.440, 0.395, 0.217, -0.178, 0.399, 0.355, 0.340, 0.004),
  ci_lower = c(0.050, 0.112, 0.197, -0.172, 0.307, -0.196, 0.148, 0.168, -0.639, 0.272, 0.135, -0.145, -0.311),
  ci_upper = c(0.810, 0.354, 0.303, 0.510, 0.694, 1.076, 0.643, 0.267, 0.283, 0.527, 0.575, 0.824, 0.318),
  pval = c(0.0265, "<0.001", "<0.001", 0.3305, "<0.001", 0.1754, 0.0018, "<0.001", 0.4487, "<0.001", 0.0016, 0.1694, 0.9823),
  I2_percent = c(74.3, 83.4, 14.4, 87.3, 0.0, 84.7, 0.0, 0.1, 80.9, 0.0, 89.4, 60.2, 94.4),
  tau2 = c(0.1358, 0.0465, 0.0013, 0.2309, 0.0000, 0.2645, 0.0000, 0.0000, 0.1790, 0.0000, 0.0733, 0.1103, 0.1373),
  significance = c("*", "***", "***", "ns", "***", "ns", "**", "***", "ns", "***", "**", "ns", "ns")
)

# Sauvegarder en XLSX
write.xlsx(final_summary, paste0(PATH, "metaMODE.xlsx"), rowNames = FALSE)

print("=== RÉSUMÉ FINAL SAUVEGARDÉ ===")
print("Fichier: metaMODE.xlsx")
print("")
print("Aperçu du tableau final:")
print(final_summary)


# CRÉER UN RÉSUMÉ TEXTUEL STRUCTURÉ
# =================================

print("╔══════════════════════════════════════════════════════════════════════════════╗")
print("║                    RÉSUMÉ DES MÉTA-ANALYSES REML PAR MODE                   ║")
print("╚══════════════════════════════════════════════════════════════════════════════╝")
print("")

print("📊 SYNTHÈSE GÉNÉRALE:")
print("• 4 modes analysés: Nature, Population, Duration, HNC")
print("• 22 études incluses au total")
print("• Méthode: REML (Restricted Maximum Likelihood)")
print("")

print("🏆 TOP 3 DES CATÉGORIES LES PLUS EFFICACES (tous modes confondus):")
print("1. Duration < 1d: d = 0.500 [0.307; 0.694] *** (n=2)")
print("2. Duration 5-7 min: d = 0.440 [-0.196; 1.076] ns (n=3)")
print("3. Nature Simulated: d = 0.430 [0.050; 0.810] * (n=5)")
print("")

print("📈 ANALYSE PAR MODE:")
print("")

print("🌿 NATURE (n=22 études):")
print("• Simulated: d = 0.430 [0.050; 0.810] * (n=5, I²=74.3%)")
print("• Real: d = 0.233 [0.112; 0.354] *** (n=17, I²=83.4%)")
print("→ Les environnements simulés montrent un effet plus fort mais avec moins d'études")
print("")

print("👥 POPULATION (n=22 études):")
print("• General Public: d = 0.250 [0.197; 0.303] *** (n=13, I²=14.4%)")
print("• University Students: d = 0.169 [-0.172; 0.510] ns (n=9, I²=87.3%)")
print("→ Le grand public répond mieux aux interventions que les étudiants universitaires")
print("")

print("⏱️ DURATION (n=22 études):")
print("• < 1d: d = 0.500 [0.307; 0.694] *** (n=2, I²=0%)")
print("• 5-7 min: d = 0.440 [-0.196; 1.076] ns (n=3, I²=84.7%)")
print("• 10 min: d = 0.395 [0.148; 0.643] ** (n=3, I²=0%)")
print("• 1/d x 7d: d = 0.217 [0.168; 0.267] *** (n=8, I²=0.1%)")
print("• 20 min/d x 7d: d = -0.178 [-0.639; 0.283] ns (n=4, I²=80.9%)")
print("→ Les interventions courtes intensives sont plus efficaces que les répétées")
print("")

print("🧠 HNC - Human-Nature Connection Scales (n=22 études):")
print("• CNS: d = 0.399 [0.272; 0.527] *** (n=6, I²=0%)")
print("• INS: d = 0.355 [0.135; 0.575] ** (n=7, I²=89.4%)")
print("• NRS: d = 0.340 [-0.145; 0.824] ns (n=3, I²=60.2%)")
print("• NR-6: d = 0.004 [-0.311; 0.318] ns (n=6, I²=94.4%)")
print("→ CNS et INS sont les échelles les plus sensibles aux changements")
print("")

print("🎯 RECOMMANDATIONS PRATIQUES:")
print("1. Privilégier les interventions d'une journée complète (effet maximal)")
print("2. Les sessions courtes (5-10 min) sont prometteuses pour des interventions ponctuelles")
print("3. Cibler le grand public plutôt que les populations étudiantes")
print("4. Utiliser CNS ou INS pour mesurer les changements de connexion à la nature")
print("5. Les environnements simulés peuvent être une alternative efficace au réel")
print("")

print("⚠️ LIMITES ET PRÉCAUTIONS:")
print("• Hétérogénéité élevée dans plusieurs catégories (I² > 75%)")
print("• Certaines catégories ont peu d'études (n < 5)")
print("• Les intervalles de confiance larges indiquent une incertitude")
print("")














