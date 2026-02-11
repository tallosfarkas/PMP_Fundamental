# ==============================================================================
# GCC GARRISON STATE STRATEGY: UNIFIED AUDIT & BLACK-LITTERMAN MODEL
# ==============================================================================
# Focus: Bayesian Factor Overlays & Data-Aware Confidence
# ==============================================================================

if(!require("pacman")) install.packages("pacman")
pacman::p_load(tidyverse, readxl, zoo, xts, quadprog, corpcor, lubridate, ggrepel, viridis, scales, patchwork)

# ------------------------------------------------------------------------------
# 1. ROBUST DATA CONSOLIDATION & CLEANING
# ------------------------------------------------------------------------------
# Load price data and convert to long format
price_data <- readRDS("bloomberg_funda.rds")
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
price_df_long <- bind_rows(lapply(seq_along(price_data), function(i) convert_to_df(price_data[[i]], names(price_data)[i])))

# Load & Clean Fundamentals
funda_raw <- read_xlsx("code/sector_industry_filtered_stocks.xlsx")
names(funda_raw) <- make.names(names(funda_raw)) %>% 
  str_replace_all("\\.", "_") %>% str_replace_all("_+", "_") %>% str_remove("_$")

# Define Strategy Sectors (Capturing the "Garrison Artery")
logistics_pattern <- "Logistics|Marine|Freight|Transport|Shipping|Ports|Storage"
funda_clean <- funda_raw %>%
  mutate(
    Sector_Group = case_when(
      str_detect(GICS_SubInd_Name, regex(logistics_pattern, ignore_case = TRUE)) ~ "Logistics",
      str_detect(GICS_SubInd_Name, "Banks") ~ "Banks",
      str_detect(GICS_SubInd_Name, "Insurance") ~ "Insurance",
      str_detect(GICS_Sector, "Health") ~ "Healthcare",
      str_detect(GICS_Sector, "Consumer Discretionary") ~ "Consumer Discretionary",
      TRUE ~ "Other"
    ),
    ROIC = as.numeric(ROIC_LF),
    Hist_Growth = as.numeric(Net_Sales_5_Yr_Geo_Gr_LF),
    PE = as.numeric(P_E),
    MCap = as.numeric(Market_Cap)
  ) %>% filter(Sector_Group != "Other")

# ------------------------------------------------------------------------------
# 2. DEEP AUDIT: DATA QUALITY & VALUATION EFFICIENCY
# ------------------------------------------------------------------------------
# Sectoral Health Summary
dq_report <- funda_clean %>%
  group_by(Sector_Group) %>%
  summarise(
    N_Stocks = n(),
    Data_Reliability = round(sum(!is.na(PE) & !is.na(ROIC)) / n() * 100, 1),
    Median_Current_PE = round(median(PE, na.rm=T), 2),
    Median_Hist_Growth = round(median(Hist_Growth, na.rm=T), 2),
    Median_ROIC = round(median(ROIC, na.rm=T), 2)
  ) %>% mutate(Value_Score = round(Median_Hist_Growth / Median_Current_PE, 2))

print("--- INSTITUTIONAL SECTOR AUDIT ---")
print(dq_report)

# Plot: Current Health vs. Historical Trend
# Identifies 'Alpha' sectors where current ROIC exceeds historical growth
p_quadrant <- ggplot(dq_report, aes(x = Median_Hist_Growth, y = Median_ROIC, label = Sector_Group)) +
  geom_point(aes(size = N_Stocks, color = Data_Reliability)) +
  geom_text_repel(fontface = "bold") +
  scale_color_viridis_c(option = "mako", name="Data Quality %") +
  theme_minimal() +
  labs(title = "GCC Sector Selection: Historic Growth vs Current Efficiency",
       subtitle = "Top-Right: Garrison Champions | Bottom-Left: Stagnant Sectors",
       x = "Historical 5Y Sales Growth (%)", y = "Current Median ROIC (%)")

# ------------------------------------------------------------------------------
# 3. CONSTRUCT SECTOR INDICES (FOR BLACK-LITTERMAN)
# ------------------------------------------------------------------------------
matched_tickers <- intersect(price_df_long$Ticker, funda_clean$Ticker)

sector_returns <- price_df_long %>%
  filter(Ticker %in% matched_tickers) %>%
  left_join(funda_clean %>% select(Ticker, Sector_Group, MCap), by = "Ticker") %>%
  arrange(Ticker, Date) %>%
  group_by(Ticker) %>%
  mutate(Ret = c(NA, diff(log(Price)))) %>%
  filter(!is.na(Ret)) %>%
  group_by(Date, Sector_Group) %>%
  summarise(IndexRet = weighted.mean(Ret, MCap, na.rm=T), .groups="drop") %>%
  pivot_wider(names_from = Sector_Group, values_from = IndexRet) %>%
  filter(complete.cases(.))

# Ensure Matrix is PURELY Numeric (Fixes the Error)
ret_matrix <- sector_returns %>% select(-Date) %>% as.matrix()
storage.mode(ret_matrix) <- "numeric" # Hard-cast to numeric
returns_xts <- xts(ret_matrix, order.by = sector_returns$Date)

# ------------------------------------------------------------------------------
# 4. BLACK-LITTERMAN WITH THE "LOGISTICS PUSH"
# ------------------------------------------------------------------------------
# Robust Prior (Shrinkage Covariance)
Sigma <- as.matrix(cov.shrink(ret_matrix, verbose = FALSE)) * 52
w_mkt <- dq_report$N_Stocks / sum(dq_report$N_Stocks) # Sectoral Proxy Weights
Pi <- as.vector(2.5 * Sigma %*% (w_mkt / sum(w_mkt)))
names(Pi) <- colnames(ret_matrix)

# THE VIEW PUSH: Integrating Macro Thesis
# Adding +3.5% Alpha to Logistics (The Artery) and +1.5% to Banks (Sovereign Core)
Q_Leveled <- Pi
Q_Leveled["Logistics"] <- Q_Leveled["Logistics"] + 0.035 
Q_Leveled["Banks"] <- Q_Leveled["Banks"] + 0.015

# Bayesian Confidence: Penalize sectors with < 50% Data Reliability
base_vol <- sqrt(diag(Sigma))
conf_scaler <- dq_report$Data_Reliability[match(names(Pi), dq_report$Sector_Group)] / 100
Omega_diag <- ( (1 - conf_scaler) * base_vol )^2
# Override: Set Logistics confidence to High (0.2x Vol) manually
Omega_diag["Logistics"] <- (0.2 * base_vol["Logistics"])^2
Omega <- diag(as.numeric(Omega_diag))

# ------------------------------------------------------------------------------
# 5. POSTERIOR & OPTIMIZATION
# ------------------------------------------------------------------------------
tau <- 0.025
P <- diag(length(Pi))
inv_tau_Sigma <- solve(tau * Sigma)
inv_Omega <- solve(Omega)
post_Sigma <- solve(inv_tau_Sigma + t(P) %*% inv_Omega %*% P)
post_mu <- as.vector(post_Sigma %*% (inv_tau_Sigma %*% Pi + t(P) %*% inv_Omega %*% Q_Leveled))

# Constraints: Max 35% Cap, No Shorting
n_sec <- length(Pi)
Amat <- cbind(1, diag(n_sec), -diag(n_sec))
bvec <- c(1, rep(0, n_sec), rep(-0.35, n_sec))
sol <- solve.QP(2 * (2.5 * post_Sigma), post_mu, Amat, bvec, meq = 1)

# Final Result Table
final_res <- data.frame(
  Sector = names(Pi),
  Prior_Expected_Ret = round(Pi * 100, 2),
  Posterior_Expected_Ret = round(post_mu * 100, 2),
  Benchmark_W = round(w_mkt * 100, 1),
  Garrison_Strategy_W = round(sol$solution * 100, 1)
)

print("--- FINAL OPTIMIZED ALLOCATION ---")
print(final_res)

# Plot Results
p_weights <- final_res %>%
  pivot_longer(cols = ends_with("_W"), names_to = "Type", values_to = "Weight") %>%
  ggplot(aes(x = Sector, y = Weight, fill = Type)) +
  geom_bar(stat = "identity", position = "dodge") +
  theme_minimal() + labs(title = "Benchmark vs. Garrison Strategy Weights", y = "Weight %")

(p_quadrant / p_weights)