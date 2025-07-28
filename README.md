# TMNE-HNC

This GitHub repository contains the data and code for a meta-analysis investigating the impact of technology-mediated nature experiences on nature connectedness.

### Project Overview

This project is a comprehensive research effort that includes a scoping review, a meta-analysis, and moderator analyses to synthesize the existing literature on the topic. The primary goal is to understand the overall effect of technology-mediated nature experiences on individuals' connection to nature and to identify key factors that may moderate this effect.

### Repository Structure

The repository is organized into the following directories:

* **1-scoping review**: Contains the R script used for the statistical analysis of the scoping review data.
* **2-meta analysis**: Includes a Jupyter Notebook and R scripts for the main meta-analysis and publication bias assessment.
* **3-moderators**: Contains R scripts for analyzing the influence of various moderators on the effect size.
* **data**: Includes all the datasets used in the analyses, from the initial literature search to the final curated dataset for the meta-analysis.

### Key Files and Descriptions

#### Scoping Review

* `scopingSTATS.R`: An R script that performs a statistical analysis of the studies included in the scoping review. It loads the dataset from an Excel file, categorizes different technologies, and generates visualizations, such as the temporal distribution of studies.

#### Meta-Analysis

* `metaPy.ipynb`: The main Jupyter Notebook for the meta-analysis. It uses a combination of Python and R (via `rpy2` and the `metafor` package) to calculate effect sizes (Cohen's d) and their variances for repeated measures designs. The notebook generates a comprehensive forest plot visualizing the effect sizes of all included studies and the overall pooled effect.
* `metaANALYSIS.R`: An R script that likely performs the core meta-analysis calculations.
* `metaBIAS.R`: An R script dedicated to assessing publication bias in the meta-analysis, using methods such as Egger's test and Begg's test.
* `metaSTATS.R`: An R script for generating descriptive statistics and summaries of the meta-analysis data.

#### Moderator Analysis

This directory contains several R scripts, each focusing on a specific moderator:

* `metaHNC.R`: Analyzes the moderating effect of different Human-Nature Connectedness (HNC) scales.
* `metaPopulation.R`: Examines the influence of different population types (e.g., university students vs. general public).
* `metaDuration.R`: Investigates the moderating effect of the duration of the nature experience.
* `metaTechno.R`: Analyzes the impact of different types of technology.
* `metaNature.R`: Explores the difference between real and simulated nature experiences.
* `Metareg.R`: Likely performs a meta-regression analysis to explore the combined effect of multiple moderators.
* `metaMODE.R`: Analyzes the moderating effect of different modes of engagement.

#### Data

The `data` directory contains a series of CSV files that document the literature screening process:

* `all-records - Query1.csv`, `all-records - Query2.csv`, `all-records - Query1+2.csv`: Initial search results from different queries.
* `all-records - Titles+Abstract Screening.csv`: Data after the title and abstract screening phase.
* `all-records - Full-Text Screening.csv`: Data after the full-text screening phase.
* `all-records - Final Selection.csv`: The final dataset of studies included in the meta-analysis.
* `meta - techno.csv`: A curated dataset used for the technology moderator analysis.

### How to Run the Analysis

1.  **Scoping Review**: Run the `scopingSTATS.R` script in RStudio. Make sure to have the necessary R packages installed (`readxl`, `ggplot2`, `dplyr`, `scales`).
2.  **Meta-Analysis**: Open and run the `metaPy.ipynb` notebook in a Jupyter environment. The notebook handles the installation of the required R package `metafor`. Alternatively, you can run the individual R scripts (`metaANALYSIS.R`, `metaBIAS.R`, `metaSTATS.R`) in RStudio.
3.  **Moderator Analysis**: Run the R scripts in the `3-moderators` directory to perform the respective moderator analyses.

### Tools and Packages

* **R**: The primary tool for statistical analysis and visualization.
    * **Packages**: `readxl`, `ggplot2`, `dplyr`, `scales`, `metafor`, `openxlsx`
* **Python**: Used in the Jupyter notebook for data manipulation and visualization.
    * **Packages**: `pandas`, `numpy`, `matplotlib`, `seaborn`, `scipy`, `openpyxl`, `rpy2`

  
