library(readr)
library(dplyr)
library(tidyr)
library(ggplot2)
library(hms) 
library(ggpubr)

# Load your data
data <- read_csv("~/Downloads/Important Stats - Varun Oct Surg and Dec Surg.csv")

# Add Stroke Severity
data <- data %>%
  mutate(
    Stroke_Severity = case_when(
      grepl("Ex", Outcomes, ignore.case = TRUE) ~ 0,
      grepl("Inc", Outcomes, ignore.case = TRUE) ~ 0.5,
      grepl("Dead", Outcomes, ignore.case = TRUE) ~ 1,
      TRUE ~ NA_real_
    ),
    Time_to_Reached = as.numeric(gsub(" sec", "", Time_to_Reached))
  )
# Set variables to plot
vars <- c(
  "Time_to_Reached", "Perc_In_Range", "Avg_Abs_Change", "Post_Area_Total",
  "Max_Temp", "Avg_Temp", "Min_Temp",
  "Perc_Out_Total", "Perc_Out_Post",
  "Time_Above", "Time_Below",
  "Overall_Area_Above", "Overall_Area_Below", "Overall_Area_Total",
  "Post_Area_Above", "Post_Area_Below"
)

# Variable name dictionary
var_labels <- c(
  Time_to_Reached       = "Time to Reach Preferred Temp (sec)",
  Perc_In_Range         = "% Time Within Preferred Temp",
  Perc_Out_Total        = "% Time Outside Preferred Temp (Overall)",
  Perc_Out_Post         = "% Time Outside Temp (After Entry)",
  Time_Above            = "Time Above Preferred Temp (sec)",
  Time_Below            = "Time Below Preferred Temp (sec)",
  Avg_Abs_Change        = "Avg Absolute Temp Change (°C/sec)",
  Max_Temp              = "Max Temp (°C)",
  Avg_Temp              = "Avg Temp (°C)",
  Min_Temp              = "Min Temp (°C)",
  Overall_Area_Above    = "Area Above Preferred Temp (°C·sec)",
  Overall_Area_Below    = "Area Below Preferred Temp (°C·sec)",
  Overall_Area_Total    = "Total Area Outside Temp (°C·sec)",
  Post_Area_Above       = "Post-Entry Area Above (°C·sec)",
  Post_Area_Below       = "Post-Entry Area Below (°C·sec)",
  Post_Area_Total       = "Total Post-Entry Area (°C·sec)"
)

# Correlations
cor_data <- data %>% select(Stroke_Severity, all_of(vars)) %>% na.omit()
cor_values <- sapply(vars, function(v) cor(cor_data[[v]], cor_data$Stroke_Severity, use = "complete.obs"))
cor_df <- data.frame(
  Variable = names(cor_values),
  Correlation = round(cor_values, 2),
  Label = var_labels[names(cor_values)]
) %>%
  arrange(desc(abs(Correlation)))

# ─────────────────────────
# 🌈 1. Horizontal Correlation Plot
# ─────────────────────────
ggplot(cor_df, aes(x = reorder(Label, Correlation), y = Correlation, fill = Correlation)) +
  geom_col() +
  coord_flip() +
  scale_fill_gradient2(low = "#72B4E6", mid = "white", high = "#E67C73", midpoint = 0) +
  geom_text(aes(label = Correlation), hjust = ifelse(cor_df$Correlation > 0, -0.2, 1.1), size = 3.5) +
  theme_minimal(base_size = 14) +
  labs(title = "Correlation with Stroke Severity", x = "", y = "Pearson Correlation") +
  theme(legend.position = "none")

# ─────────────────────────
# 📦 Prep for Boxplots
# ─────────────────────────
long_data <- data %>%
  select(Stroke_Severity, all_of(vars)) %>%
  pivot_longer(cols = all_of(vars), names_to = "Variable", values_to = "Value") %>%
  left_join(cor_df, by = c("Variable")) %>%
  mutate(Stroke_Severity = as.factor(Stroke_Severity))

# ─────────────────────────
# 🌸 3. Individual Faceted Boxplots
# ─────────────────────────
ggplot(long_data, aes(x = factor(Stroke_Severity,
                                 levels = c(0, 0.5, 1),
                                 labels = c("Excluded", "Included", "Died")),
                      y = Value,
                      fill = factor(Stroke_Severity,
                                    levels = c(0, 0.5, 1),
                                    labels = c("Excluded", "Included", "Died")))) +
  geom_boxplot(outlier.alpha = 0.1, alpha = 0.8) +
  stat_compare_means(comparisons = list(c("Excluded", "Included"),
                                        c("Included", "Died"),
                                        c("Excluded", "Died")),
                     method = "wilcox.test", label = "p.signif", size = 3) +
  facet_wrap(~ paste0(Label, "\n(r = ", Correlation, ")"), scales = "free_y", ncol = 4) +
  theme_minimal(base_size = 14) +
  scale_fill_manual(values = c("Excluded" = "#66C2A5", "Included" = "#FC8D62", "Died" = "#8DA0CB")) +
  labs(title = "Boxplots of Each Variable by Stroke Severity",
       x = "Stroke Severity", y = "Value", fill = "Outcome")
