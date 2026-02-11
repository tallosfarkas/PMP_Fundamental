# ==============================================================================
# PROFESSIONAL BLACK-LITTERMAN STRATEGY (GCC FOCUS)
# ==============================================================================
# 1. LOAD PACKAGES
if(!require("pacman")) install.packages("pacman")
pacman::p_load(tidyverse, readxl, zoo, xts, quadprog, PerformanceAnalytics, corrplot, corpcor, lubridate)

# ==============================================================================
# STEP 1: LOAD & CLEAN PRICE DATA
# ==============================================================================
price_data <- readRDS("bloomberg_funda.rds")

# Helper: Convert list to DF
convert_to_df <- function(ts_obj, ticker_name) {
  if (is.xts(ts_obj) || is.zoo(ts_obj)) {
    df <- data.frame(Date = index(ts_obj), Price = as.numeric(coredata(ts_obj)))
  } else if (is.data.frame(ts_obj)) {
    df <- ts_obj
    if (ncol(df) >= 2) names(df)[1:2] <- c("Date", "Price")
  } else { return(NULL) }
  df$Ticker <- ticker_name
  df$Date <- as.Date(df$Date)
  return(df)
}

# Convert & Bind
price_df_list <- lapply(seq_along(price_data), function(i) {
  convert_to_df(price_data[[i]], names(price_data)[i])
})
price_df_list <- price_df_list[!sapply(price_df_list, is.null)]
price_df_long <- bind_rows(price_df_list)

# Pivot to Wide
price_df_wide <- price_df_long %>%
  select(Date, Ticker, Price) %>%
  pivot_wider(names_from = Ticker, values_from = Price) %>%
  arrange(Date)

# ==============================================================================
# STEP 2: LOAD & CLEAN FUNDAMENTAL DATA
# ==============================================================================
# Use the CSV version for reliability
funda_raw <- read_xlsx("code/sector_industry_filtered_stocks.xlsx")

# Clean column names
names(funda_raw) <- names(funda_raw) %>% 
  str_replace_all(" ", "_") %>% 
  str_replace_all("/", "_") %>% 
  str_replace_all("-", "_") %>%
  str_replace_all("%", "Pct")

# ==============================================================================
# STEP 3: ROBUST SECTOR MAPPING
# ==============================================================================
ai_subinds <- c(
  "Electric Utilities", "Independent Power Producers & Energy Traders",
  "Renewable Electricity", "Electrical Components & Equipment",
  "Electrical Equipment & Instruments", "Electronic Equipment & Instruments",
  "Electronic Equipment, Instruments & Components", "Application Software"
)
logistics_pattern <- "Logistics|Marine|Freight|Trucking|Airport|Rail|Transport|Shipping|Ports|Storage"

sector_mapping_clean <- funda_raw %>%
  mutate(
    Ticker = Ticker, # Ensure Ticker column exists
    Market_Cap_Num = as.numeric(Market_Cap),
    PE = as.numeric(P_E),
    ROIC = as.numeric(ROIC_LF),
    DivYield = as.numeric(Dvd_Ind_Yld),
    SalesGrowth = as.numeric(Net_Sales___5_Yr_Geo_Gr_LF),
    
    Sector_Group = case_when(
      GICS_SubInd_Name %in% ai_subinds ~ "AI",
      str_detect(GICS_SubInd_Name, regex(logistics_pattern, ignore_case = TRUE)) ~ "Logistics",
      str_detect(GICS_SubInd_Name, "Insurance|Reinsurance|Insurance Brokers") ~ "Insurance",
      str_detect(GICS_SubInd_Name, "Banks") ~ "Banks",
      str_detect(GICS_Sector, "Health") ~ "Healthcare",
      str_detect(GICS_Sector, "Consumer Discretionary") ~ "Consumer Discretionary",
      TRUE ~ "Other"
    )
  )

target_sectors <- c("Banks", "Insurance", "Logistics", "AI", "Healthcare", "Consumer Discretionary")
sector_mapping_final <- sector_mapping_clean %>% filter(Sector_Group %in% target_sectors)
target_tickers <- sector_mapping_final$Ticker

# ==============================================================================
# STEP 4: PREPARE WEEKLY DATA
# ==============================================================================
matched_tickers <- intersect(names(price_df_wide), target_tickers)
cutoff_date <- max(price_df_wide$Date, na.rm = TRUE) - years(10)

price_filtered <- price_df_wide %>%
  filter(Date >= cutoff_date) %>%
  select(Date, all_of(matched_tickers)) %>%
  arrange(Date)

# Fill NAs Daily then Convert to Weekly
price_filled <- price_filtered %>%
  mutate(across(-Date, ~ zoo::na.locf(., na.rm = FALSE))) %>%
  mutate(across(-Date, ~ zoo::na.locf(., fromLast = TRUE)))

price_weekly <- price_filled %>%
  mutate(Week_End = ceiling_date(Date, "week")) %>%
  group_by(Week_End) %>%
  summarise(across(everything(), last), .groups = 'drop') %>%
  select(-Date) %>%
  rename(Date = Week_End) %>%
  slice(1:(n()-1))

# Filter for liquid tickers
missing_pct <- colSums(is.na(select(price_weekly, -Date))) / nrow(price_weekly)
valid_tickers <- names(missing_pct[missing_pct < 0.60]) 
price_final <- price_weekly %>% select(Date, all_of(valid_tickers))

# Update mapping
sector_mapping_final <- sector_mapping_final %>% filter(Ticker %in% valid_tickers)

# ==============================================================================
# STEP 5: CALCULATE RETURNS & INDICES
# ==============================================================================
returns_df <- price_final %>%
  arrange(Date) %>%
  mutate(across(-Date, ~ c(NA, diff(log(.))))) %>%
  slice(-1) %>%
  mutate(across(-Date, ~ replace_na(., 0)))

returns_long <- returns_df %>%
  pivot_longer(-Date, names_to = "Ticker", values_to = "Return") %>%
  left_join(sector_mapping_final %>% select(Ticker, Sector_Group, Market_Cap_Num), by = "Ticker") %>%
  filter(!is.na(Sector_Group), Market_Cap_Num > 0)

sector_indices <- returns_long %>%
  group_by(Date, Sector_Group) %>%
  summarise(
    IndexReturn = weighted.mean(Return, w = Market_Cap_Num, na.rm = TRUE),
    TotalMarketCap = sum(Market_Cap_Num, na.rm = TRUE), 
    .groups = 'drop'
  )

returns_wide <- sector_indices %>%
  select(Date, Sector_Group, IndexReturn) %>%
  pivot_wider(names_from = Sector_Group, values_from = IndexReturn) %>%
  arrange(Date)

# Check for NAs (e.g. if a sector had no returns for a week)
returns_wide[is.na(returns_wide)] <- 0
returns_xts <- xts(returns_wide[,-1], order.by = returns_wide$Date)

# ==============================================================================
# STEP 6: MARKET PRIOR (Shrinkage)
# ==============================================================================
freq <- 52 
Sigma_shrink <- cov.shrink(returns_xts, verbose = FALSE)
Sigma <- as.matrix(Sigma_shrink) * freq
colnames(Sigma) <- rownames(Sigma) <- colnames(returns_xts)

# Market Weights
sector_mcap <- sector_indices %>%
  group_by(Sector_Group) %>%
  summarise(MCap = mean(TotalMarketCap, na.rm = TRUE)) 
w_mkt <- setNames(sector_mcap$MCap / sum(sector_mcap$MCap), sector_mcap$Sector_Group)
w_mkt <- w_mkt[colnames(returns_xts)]

# Prior (Pi)
lambda <- 2.5
Pi <- as.vector(lambda * Sigma %*% w_mkt)
names(Pi) <- colnames(returns_xts)

# ==============================================================================
# STEP 7: QUANT FACTOR SCORES
# ==============================================================================
sectors <- colnames(returns_xts)

# 1. Fundamental Scores (Current Snapshot)
sector_factors <- sector_mapping_final %>%
  group_by(Sector_Group) %>%
  summarise(
    Median_Yield = median(DivYield, na.rm = TRUE),
    Median_PE = median(PE, na.rm = TRUE),
    Median_ROIC = median(ROIC, na.rm = TRUE),
    Median_Growth = median(SalesGrowth, na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  mutate(across(where(is.numeric), ~replace_na(., 0))) %>%
  mutate(Median_PE = ifelse(Median_PE == 0, 20, Median_PE))

# Z-Score Normalization
sector_scores <- sector_factors %>%
  mutate(
    # Value: Low PE (-), High Yield (+)
    z_Value = as.vector(scale(Median_Yield) - scale(Median_PE)), 
    z_Quality = as.vector(scale(Median_ROIC)),
    z_Growth = as.vector(scale(Median_Growth)),
    Final_Funda_Score = (0.3 * z_Value) + (0.4 * z_Quality) + (0.3 * z_Growth)
  )

# Align Scores
funda_z <- sector_scores$Final_Funda_Score[match(sectors, sector_scores$Sector_Group)]
funda_z[is.na(funda_z)] <- 0

# 2. Momentum Scores (Historical 12M Trend)
last_52_weeks <- tail(returns_xts, 52)
mom_raw <- apply(last_52_weeks, 2, function(x) prod(1 + x) - 1)
mom_z <- as.vector(scale(mom_raw))

# 3. Comparison Table (As requested)
comparison_table <- data.frame(
  Sector = sectors,
  Funda_Score_Z = round(funda_z, 2),
  Momentum_Score_Z = round(mom_z, 2),
  Interpretation = ifelse(funda_z > 0 & mom_z > 0, "Strong Buy (Double Confirm)",
                          ifelse(funda_z < 0 & mom_z < 0, "Avoid (Double Negative)", "Mixed Signal"))
)
print("--- FACTOR COMPARISON: CURRENT HEALTH VS HISTORICAL TREND ---")
print(comparison_table)

# ==============================================================================
# STEP 8: BLEND VIEWS
# ==============================================================================
# Scaling: 1 Z-Score unit = 2.5% Alpha
view_scaling <- 0.025

# Combine: 50% Fundamental, 50% Momentum
alpha_combined <- (0.5 * funda_z * view_scaling) + (0.5 * mom_z * view_scaling)
Q_Final <- Pi + alpha_combined

# Matrices
n_sectors <- length(sectors)
P <- diag(n_sectors)
Q <- as.vector(Q_Final)

# Confidence: High if both signals agree
agreement <- sign(funda_z) == sign(mom_z)
conf_scaler <- ifelse(agreement, 0.5, 0.8) 
omega_diag <- (conf_scaler * sqrt(diag(Sigma)))^2
Omega <- diag(omega_diag)

# ==============================================================================
# STEP 9: OPTIMIZATION WITH CAPS
# ==============================================================================
# Posterior
tau <- 0.025
inv_tau_Sigma <- solve(tau * Sigma)
inv_Omega <- solve(Omega)
posterior_Sigma <- solve(inv_tau_Sigma + t(P) %*% inv_Omega %*% P)
posterior_mu <- as.vector(posterior_Sigma %*% (inv_tau_Sigma %*% Pi + t(P) %*% inv_Omega %*% Q))

# Optimization (Max 35% per sector)
max_weight <- 0.35
Amat <- cbind(rep(1, n_sectors), diag(n_sectors), -diag(n_sectors))
bvec <- c(1, rep(0, n_sectors), rep(-max_weight, n_sectors))

solution <- solve.QP(
  Dmat = 2 * (lambda * posterior_Sigma), 
  dvec = posterior_mu,                   
  Amat = Amat,
  bvec = bvec,
  meq = 1
)

optimal_weights <- solution$solution
names(optimal_weights) <- sectors

# ==============================================================================
# STEP 10: FINAL RESULTS
# ==============================================================================
results <- data.frame(
  Sector = sectors,
  Market_Weight = round(w_mkt * 100, 1),
  Optimal_Weight = round(optimal_weights * 100, 1),
  Post_Return = round(posterior_mu * 100, 1)
)

print("--- FINAL ALLOCATION ---")
print(results)

# Visualization
sector_colors <- c("Banks"="#1f77b4", "Insurance"="#00cccc", "Logistics"="#ff7f0e",
                   "AI"="#9467bd", "Healthcare"="#2ca02c", "Consumer Discretionary"="#d62728")

results %>%
  pivot_longer(c(Market_Weight, Optimal_Weight), names_to = "Type", values_to = "Weight") %>%
  ggplot(aes(x = Sector, y = Weight, fill = Sector)) +
  geom_bar(stat = "identity", position = "dodge") +
  facet_wrap(~Type) +
  scale_fill_manual(values = sector_colors) +
  theme_minimal() +
  labs(title = "Final Leveled-Up Strategy", y = "Weight %") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))