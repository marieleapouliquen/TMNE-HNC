#rm(list = ls())#clean workspace
#fold="~/Desktop/Marie-Lea/"#folder name
fold_fig="~/Desktop/"#folder name


#load packages
#-------------
library(ggplot2)
library(metafor)
library(dplyr)
library(readxl)


#Load data 
#------------
PATH="/Users/marieleapouliquen/Desktop/Review/Data/"
d <- read_excel(file.path(PATH, "meta.xlsx"), sheet = "techno") 

#remame variables

d$Mean.Control <- as.numeric(d$Mean_Pre)
d$Mean.Treatment <- as.numeric(d$Mean_Post) 
d$sd.control <-as.numeric(d$SD_Pre)
d$sd.treatment <-as.numeric(d$SD_Post)
d$Nature.n <-as.numeric(d$N_Post)
Control.n <- as.numeric(d$N_Pre)

#------------------------------------------------------------------------
#make multivariate meta-analysis for experimental data
#------------------------------------------------------------------------
#matrix of variances
dat2 <- escalc(measure="SMD", m1i = Mean.Treatment, m2i= Mean.Control, sd1i=sd.treatment, sd2i=sd.control, n1i=Nature.n, n2i=Control.n,  data=d, slab=paste(Citation)) #subset=d$LargeMETA %in% c('ExpNC'), 
dat2$estid <- 1:nrow(dat2)


#meta-regression
# Clean up variables first
dat2$HNC <- relevel(as.factor(dat2$HNC), ref = "INS")
dat2$Nature <- relevel(as.factor(dat2$Nature), ref = "Real")
dat2$Techno_reco <- relevel(as.factor(dat2$Techno_reco), ref = "Citizen Science App")
dat2$Dur_Freq_reco <- relevel(as.factor(dat2$Dur_Freq_reco), ref = "short_repeated")
dat2$Population <- relevel(as.factor(dat2$Population), ref = "General Public")


# Run models AFTER reference levels are set
res <- rma.mv(yi, vi, mods = ~ HNC + Techno_reco +  Dur_Freq_reco +  Population,
              random = ~1 | Citation, data = dat2)

res <- rma.mv(yi, vi, mods = ~ HNC + Nature +  Dur_Freq_reco +  Population,
              random = ~1 | Citation, data = dat2)

res <- rma.mv(yi, vi, mods = ~ Nature ,
              random = ~1 | Citation, data = dat2)
summary(res)

library(broom)
tidy(res)



