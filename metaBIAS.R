# ================================================================================
# SCRIPT COMPLET: ANALYSE DES BIAIS DE PUBLICATION EN MÉTA-ANALYSE
# Avec calculs réels de Cohen's d pour mesures répétées
# ================================================================================

# ÉTAPE 1: INSTALLATION ET CHARGEMENT DES PACKAGES
# ================================================
# Installation (décommentez si nécessaire)
# install.packages(c("readxl", "metafor", "ggplot2", "dplyr"), 
#                  repos="https://cran.rstudio.com/", dependencies = TRUE)

library(readxl)    # Pour lire les fichiers Excel
library(metafor)   # Pour les méta-analyses et tests de biais
library(ggplot2)   # Pour les graphiques
library(dplyr)     # Pour la manipulation des données
library(openxlsx)  # Pour l'écriture du fichier final

# ÉTAPE 2: LECTURE DES DONNÉES
# ============================
PATH <- "/Users/marieleapouliquen/Desktop/Review/Data/"
data <- read_excel(file.path(PATH, "meta.xlsx"), sheet = "techno")

print(paste("Données chargées:", nrow(data), "études"))

# ÉTAPE 3: CALCUL DES TAILLES D'EFFET (COHEN'S D POUR MESURES RÉPÉTÉES)
# =======================================================================

# Convertir les colonnes en numérique
data$Mean_Pre <- as.numeric(data$Mean_Pre)
data$SD_Pre <- as.numeric(data$SD_Pre)
data$Mean_Post <- as.numeric(data$Mean_Post)
data$SD_Post <- as.numeric(data$SD_Post)
data$N_Pre <- as.numeric(data$N_Pre)
data$N_Post <- as.numeric(data$N_Post)

# Calculer manuellement Cohen's d pour repeated measures
data$d_manual <- (data$Mean_Post - data$Mean_Pre) / 
  sqrt((data$SD_Pre^2 + data$SD_Post^2) / 2)

# Calculer la variance approximative pour Cohen's d (repeated measures)
# Formule: V = (1/n) + (d^2 / (2*n))
data$var_d <- (1/data$N_Pre) + (data$d_manual^2 / (2*data$N_Pre))

# Filtrer les données valides
data_clean <- data %>%
  filter(!is.na(d_manual) & !is.na(var_d) & is.finite(d_manual) & is.finite(var_d))

print(paste("Données traitées:", nrow(data_clean), "études valides"))
print("Statistiques des tailles d'effet:")
print(summary(data_clean$d_manual))

# ÉTAPE 4: MÉTA-ANALYSE GLOBALE
# =============================
meta_global <- rma(yi = d_manual, vi = var_d, data = data_clean, method = "REML")
print("Méta-analyse globale:")
print(meta_global)

# ÉTAPE 5: TESTS DE BIAIS DE PUBLICATION
# ======================================

# Test d'Egger (régression de l'asymétrie)
egger_test <- regtest(meta_global, model = "lm")
print("Test d'Egger:")
print(egger_test)
egger_interpretation <- ifelse(egger_test$pval < 0.05, "Biais détecté", "Pas de biais")
print(paste("Interprétation:", egger_interpretation))

# Test de Begg (corrélation de rang)
begg_test <- ranktest(meta_global)
print("Test de Begg:")
print(begg_test)
begg_interpretation <- ifelse(begg_test$pval < 0.05, "Biais détecté", "Pas de biais")
print(paste("Interprétation:", begg_interpretation))

# Méthode Trim-and-Fill
trimfill_result <- trimfill(meta_global)
print("Méthode Trim-and-Fill:")
print(trimfill_result)
trimfill_interpretation <- ifelse(trimfill_result$k0 > 0, 
                                  paste("Biais possible -", trimfill_result$k0, "études manquantes"), 
                                  "Pas de biais")
print(paste("Interprétation:", trimfill_interpretation))

# ÉTAPE 6: GRAPHIQUES EN ENTONNOIR
# ================================

# Graphique en entonnoir standard avec ligne verticale pour la taille d'effet globale
png("funnel_plot_standard.png", width = 2000, height = 1500, res = 300)  # haute résolution

funnel(meta_global, 
       main = "Publication Biases (n=11 studies)",
       cex.main = 1,
       xlab = "Cohen's d", 
       ylab = "Standard Error",  # ← plus standard que "Standard Deviation"
       col = "darkblue",
       pch = 21,
       back = "gray95", 
       shade = "white", 
       hlines = "white")

# Ajouter la ligne verticale au centre (taille d'effet estimée)
abline(v = meta_global$b[1], lty = 2, col = "gray40")

dev.off()
print("✓ Graphique standard sauvegardé: funnel_plot_standard.png")


# Graphique Trim-and-Fill (avec études imputées)
png("funnel_plot_trimfill.png", width = 800, height = 600, res = 100)
funnel(trimfill_result, 
       main = "Graphique Trim-and-Fill",
       xlab = "Taille d'effet (d de Cohen)", 
       ylab = "Erreur Standard",
       back = "white", 
       shade = "white", 
       hlines = "white")
dev.off()
print("✓ Graphique Trim-and-Fill sauvegardé: funnel_plot_trimfill.png")

# ÉTAPE 7: RÉSUMÉ DES RÉSULTATS
# =============================

# Créer un tableau de résumé complet
bias_summary <- data.frame(
  Test = c("Test d'Egger", "Test de Begg", "Trim-and-Fill"),
  Statistique = c(
    paste("t =", round(egger_test$zval, 3)),
    paste("τ =", round(begg_test$tau, 4)),
    paste(trimfill_result$k0, "études manquantes estimées")
  ),
  P_value = c(
    round(egger_test$pval, 4),
    round(begg_test$pval, 4),
    "N/A"
  ),
  Interprétation = c(egger_interpretation, begg_interpretation, trimfill_interpretation),
  stringsAsFactors = FALSE
)

print("=== RÉSUMÉ DES TESTS DE BIAIS ===")
print(bias_summary)

# Sauvegarder le résumé
write.xlsx(bias_summary, paste0(PATH, "publication_bias_summary.xlsx"), rowNames = FALSE)
print("✓ Résumé sauvegardé: publication_bias_summary.xlsx")


# ÉTAPE 8: CONCLUSION FINALE
# ==========================
print("=== CONCLUSION FINALE ===")
print(paste("Nombre d'études analysées:", meta_global$k))
print(paste("Taille d'effet globale: d =", round(meta_global$beta, 3)))
print(paste("Intervalle de confiance 95%: [", round(meta_global$ci.lb, 3), ";", round(meta_global$ci.ub, 3), "]"))
print(paste("Hétérogénéité (I²):", round(meta_global$I2, 1), "%"))
print("")
print("Tests de biais de publication:")
print(paste("- Test d'Egger: p =", round(egger_test$pval, 4), "→", egger_interpretation))
print(paste("- Test de Begg: p =", round(begg_test$pval, 4), "→", begg_interpretation))
print(paste("- Trim-and-Fill:", trimfill_interpretation))
print("")
print("FICHIERS GÉNÉRÉS:")
print("- publication_bias_summary.csv")
print("- funnel_plot_standard.png")
print("- funnel_plot_trimfill.png")