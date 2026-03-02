# ==============================================================================
# Load Required Packages
# ==============================================================================
if (!require("pacman")) {
  install.packages("pacman")
}
pacman::p_load(
  tidyverse,
  readxl,
  zoo,
  xts,
  quadprog,
  PerformanceAnalytics,
  corrplot,
  corpcor
)


# ==============================================================================
# STEP 1: Load Data
# ==============================================================================

# Load price/returns data from .rds file
# Each column = company ticker, each row = date
price_data <- readRDS("bloomberg_funda.rds")

head(price_data[[1]])

# Convert to data frame if it's a matrix
if (is.matrix(price_data)) {
  dates <- rownames(price_data)
  price_data <- as.data.frame(price_data)
  price_data$Date <- as.Date(dates)
} else if ("Date" %in% names(price_data)) {
  price_data$Date <- as.Date(price_data$Date)
} else {
  # If no date column, create one (adjust as needed)
  price_data$Date <- as.Date(rownames(price_data))
}

# Load ticker-sector mapping from Excel
sector_mapping <- read_excel("Code/stock.xlsx")

# Clean column names (remove spaces)
names(sector_mapping) <- gsub(" ", "_", names(sector_mapping))


print("Price data structure:")
print(str(price_data, max.level = 1))
print("\nFirst element of price_data:")
print(head(price_data[[1]]))

print("\nSector mapping structure:")
print(str(sector_mapping))
print(head(sector_mapping))
print("\nColumn names in sector_mapping:")
print(names(sector_mapping))


# ==============================================================================
# STEP 2: Convert price_data List to DataFrame
# ==============================================================================

# Extract ticker names from list
ticker_names <- names(price_data)

# Check if first element is xts, zoo, or data.frame
first_element <- price_data[[1]]
print(paste("\nFirst element class:", class(first_element)[1]))

# Function to convert each time series to data frame
convert_to_df <- function(ts_obj, ticker_name) {
  if (is.xts(ts_obj) || is.zoo(ts_obj)) {
    df <- data.frame(
      Date = index(ts_obj),
      Price = as.numeric(coredata(ts_obj))
    )
  } else if (is.data.frame(ts_obj)) {
    df <- ts_obj
    if (ncol(df) >= 2) {
      names(df)[1:2] <- c("Date", "Price")
    }
  } else if (is.numeric(ts_obj) || is.matrix(ts_obj)) {
    # Numeric vector or matrix - create dates
    df <- data.frame(
      Date = seq.Date(
        from = as.Date("2005-01-01"),
        by = "day",
        length.out = length(ts_obj)
      ),
      Price = as.numeric(ts_obj)
    )
  } else {
    return(NULL)
  }

  df$Ticker <- ticker_name
  df$Date <- as.Date(df$Date)
  return(df)
}

# Convert all elements
print("\nConverting price data to data frame...")
price_df_list <- lapply(seq_along(price_data), function(i) {
  ticker <- ifelse(
    !is.null(names(price_data)[i]),
    names(price_data)[i],
    paste0("Stock_", i)
  )
  convert_to_df(price_data[[i]], ticker)
})

# Remove NULL elements
price_df_list <- price_df_list[!sapply(price_df_list, is.null)]

# Combine into long format
price_df_long <- bind_rows(price_df_list)

print(paste("Total rows in price data:", nrow(price_df_long)))
print(paste("Unique tickers:", length(unique(price_df_long$Ticker))))
print(paste(
  "Date range:",
  min(price_df_long$Date, na.rm = TRUE),
  "to",
  max(price_df_long$Date, na.rm = TRUE)
))

# Convert to wide format
price_df_wide <- price_df_long %>%
  select(Date, Ticker, Price) %>%
  pivot_wider(names_from = Ticker, values_from = Price) %>%
  arrange(Date)

print("\nPrice data (wide format) dimensions:")
print(dim(price_df_wide))

# ==============================================================================
# STEP 3: Clean Sector Mapping with Exact Column Names
# ==============================================================================

sapply(
  sector_mapping[c("GICS_Sector", "GICS_Ind_Name", "GICS_SubInd_Name")],
  unique
)

# ==============================================================================
# STEP 3: Clean Sector Mapping (ROBUST VERSION)
# ==============================================================================

# 1. Define AI subindustries (Keep as is)
ai_subinds <- c(
  "Electric Utilities",
  "Independent Power Producers & Energy Traders",
  "Renewable Electricity",
  "Electrical Components & Equipment",
  "Electrical Equipment & Instruments",
  "Electronic Equipment & Instruments",
  "Electronic Equipment, Instruments & Components",
  "Application Software"
)

# 2. Define Logistics Keywords (Use patterns, not exact strings)
# This captures "Marine Transportation", "Air Freight", "Transportation Inf", etc.
logistics_pattern <- "Logistics|Marine|Freight|Trucking|Airport|Rail|Transport|Shipping|Ports|Storage"

sector_mapping_clean <- sector_mapping %>%
  mutate(
    Market_Cap_Num = as.numeric(Market_Cap),

    # Build custom Sector_Group (Order matters!)
    Sector_Group = case_when(
      # 1. AI (High priority)
      GICS_SubInd_Name %in% ai_subinds ~ "AI",

      # 2. Logistics (Using Pattern Matching)
      # This catches "Marine Transportation", "Air Freight & Logistics", etc.
      str_detect(
        GICS_SubInd_Name,
        regex(logistics_pattern, ignore_case = TRUE)
      ) ~ "Logistics",

      # 3. Financials Split
      str_detect(
        GICS_SubInd_Name,
        "Insurance|Reinsurance|Insurance Brokers"
      ) ~ "Insurance",
      str_detect(GICS_SubInd_Name, "Banks") ~ "Banks",

      # 4. Standard Sectors
      str_detect(GICS_Sector, "Health") ~ "Healthcare",
      str_detect(
        GICS_Sector,
        "Consumer Discretionary"
      ) ~ "Consumer Discretionary",

      # Fallback
      TRUE ~ "Other"
    )
  )

# --- DEBUGGING STEP ---
# Run this immediately after Step 3 to see if it worked
print("Check Logistics Count:")
print(table(sector_mapping_clean$Sector_Group))

# Check exactly WHICH companies were caught in Logistics
print("Companies in Logistics:")
sector_mapping_clean %>%
  filter(Sector_Group == "Logistics") %>%
  select(Ticker, Name, GICS_SubInd_Name) %>%
  print()

print(table(sector_mapping_clean$Sector_Group))


print("\nSector distribution:")
print(table(sector_mapping_clean$Sector_Group))

print("\nCountry distribution:")
print(table(sector_mapping_clean$Cntry_Terrtry_Fl_Name))

# ==============================================================================
# STEP 4: Match Tickers Between Datasets
# ==============================================================================

price_tickers <- setdiff(names(price_df_wide), "Date")
sector_tickers <- sector_mapping_clean$Ticker

print(paste("\nTickers in price data:", length(price_tickers)))
print(paste("Tickers in sector mapping:", length(sector_tickers)))

# Direct match
matched_tickers <- intersect(price_tickers, sector_tickers)
print(paste("Directly matched tickers:", length(matched_tickers)))


# Filter data
price_df_matched <- price_df_wide %>%
  select(Date, all_of(matched_tickers))

sector_mapping_matched <- sector_mapping_clean %>%
  filter(Ticker %in% matched_tickers)

print(paste("\nFinal matched tickers:", length(matched_tickers)))

# ==============================================================================
# STEP 5: Define Target Sectors and Filter
# ==============================================================================

target_sectors <- c(
  "Banks",
  "Insurance", # <--- Ensuring Insurance is here
  "Logistics", # <--- Ensuring Logistics is here
  "AI",
  "Healthcare",
  "Consumer Discretionary"
)

sector_mapping_filtered <- sector_mapping_matched %>%
  filter(Sector_Group %in% target_sectors)

target_tickers <- sector_mapping_filtered$Ticker

print("\nCompanies by target sector:")
print(table(sector_mapping_filtered$Sector_Group))
# ==============================================================================
# STEP 6: Filter, Fill NAs, and Convert to Weekly (CORRECTED)
# ==============================================================================
library(lubridate)

# 1. Filter Date Range (Last 10 Years)
cutoff_date <- max(price_df_matched$Date, na.rm = TRUE) - years(10)
price_filtered <- price_df_matched %>%
  filter(Date >= cutoff_date) %>%
  select(Date, all_of(target_tickers)) %>%
  arrange(Date)

last_non_na <- function(x) {
  x <- x[!is.na(x)]
  if (length(x) == 0) NA_real_ else tail(x, 1)
}

price_weekly <- price_filtered %>%
  mutate(Week_End = ceiling_date(Date, "week")) %>%
  group_by(Week_End) %>%
  summarise(across(-Date, last_non_na), .groups = "drop") %>%
  rename(Date = Week_End) %>%
  slice(1:(n() - 1))

# 4. Filter for valid tickers (Relaxed Threshold)
# Allow up to 60% missing data to capture recent IPOs (like ADNOC LS)
# Or just ensure we have *some* data.
missing_pct <- colSums(is.na(select(price_weekly, -Date))) / nrow(price_weekly)
valid_tickers <- names(missing_pct[missing_pct < 0.60])

# If valid_tickers is empty, force keep the top 50 by market cap
if (length(valid_tickers) < 5) {
  print(
    "Warning: Strict filtering removed too many stocks. Selecting top stocks by Market Cap instead."
  )
  top_stocks <- sector_mapping_filtered %>%
    arrange(desc(Market_Cap_Num)) %>%
    head(50) %>%
    pull(Ticker)
  valid_tickers <- intersect(names(price_weekly), top_stocks)
}

price_final <- price_weekly %>%
  select(Date, all_of(valid_tickers))

print(paste("Frequency: WEEKLY"))
print(paste("Valid Tickers Remaining:", length(valid_tickers)))
print(paste("Observations:", nrow(price_final)))

# Update sector mapping to match valid tickers
sector_mapping_final <- sector_mapping_filtered %>%
  filter(Ticker %in% valid_tickers)

# ==============================================================================
# STEP 7: Calculate Returns (Weekly)
# ==============================================================================

returns_df <- price_final %>%
  arrange(Date) %>%
  mutate(across(-Date, ~ c(NA, diff(log(.))))) %>% # log returns
  slice(-1) %>%
  mutate(across(-Date, ~ ifelse(is.infinite(.), NA, .)))

# ==============================================================================
# STEP 8: Create Market-Cap Weighted Sector Indices
# ==============================================================================

returns_long <- returns_df %>%
  pivot_longer(-Date, names_to = "Ticker", values_to = "Return") %>%
  left_join(
    sector_mapping_final %>% select(Ticker, Sector_Group, Market_Cap_Num),
    by = "Ticker"
  ) %>%
  filter(!is.na(Sector_Group), Market_Cap_Num > 0, !is.na(Return))

sector_indices <- returns_long %>%
  group_by(Date, Sector_Group) %>%
  summarise(
    IndexReturn = weighted.mean(Return, w = Market_Cap_Num, na.rm = TRUE),
    TotalMarketCap = sum(Market_Cap_Num, na.rm = TRUE),
    .groups = 'drop'
  )

# Convert to xts for Black-Litterman
returns_wide <- sector_indices %>%
  select(Date, Sector_Group, IndexReturn) %>%
  pivot_wider(names_from = Sector_Group, values_from = IndexReturn) %>%
  arrange(Date)

# Check for NAs in sectors and fill with 0 (flat week)
returns_xts <- xts(returns_wide[, -1], order.by = returns_wide$Date)

# Keep only complete rows for covariance estimation
returns_xts_cc <- returns_xts[complete.cases(returns_xts), ]

# ==============================================================================
# STEP 9: Market Statistics (Prior) with Robust Covariance (Corpcor)
# ==============================================================================
freq <- 52 # Weekly frequency

# 1. Historical Returns (Mu)
mu_hist <- colMeans(returns_xts_cc) * freq

# 2. Robust Covariance (Ledoit-Wolf Shrinkage via Corpcor)
# This estimates the covariance matrix by shrinking towards constant correlation
# It is extremely fast and stable
Sigma_shrink <- cov.shrink(returns_xts_cc, verbose = FALSE)

# Convert to standard matrix and annualize
Sigma <- as.matrix(Sigma_shrink) * freq

# Ensure column names are preserved
colnames(Sigma) <- colnames(returns_xts_cc)
rownames(Sigma) <- colnames(returns_xts_cc)

# 3. Calculate Market Weights
dates_cc <- as.Date(index(returns_xts_cc))

sector_mcap <- sector_indices %>%
  filter(Date %in% dates_cc) %>%
  group_by(Sector_Group) %>%
  summarise(MCap = mean(TotalMarketCap, na.rm = TRUE), .groups = "drop")

w_mkt <- setNames(
  sector_mcap$MCap / sum(sector_mcap$MCap),
  sector_mcap$Sector_Group
)
w_mkt <- w_mkt[colnames(returns_xts_cc)] # Ensure order matches


# 4. Risk Aversion (deterministic) & Implied Equilibrium
# ------------------------------------------------------

# Choose an annual risk-free rate (in decimal).
rf_annual <- 0.02

# Convert to weekly (geometric). Good practice even if rf_annual = 0.
rf_weekly <- (1 + rf_annual)^(1 / freq) - 1

# Market portfolio weekly returns from your sector indices
r_mkt_weekly <- as.numeric(returns_xts_cc %*% w_mkt)

# Estimate market risk aversion lambda:
# lambda = (E[Rm] - Rf) / Var(Rm)  with annualized mean/var
mu_mkt_annual <- mean(r_mkt_weekly - rf_weekly, na.rm = TRUE) * freq
var_mkt_annual <- var(r_mkt_weekly, na.rm = TRUE) * freq

lambda_mkt <- mu_mkt_annual / var_mkt_annual

# Sanity checks
if (!is.finite(lambda_mkt) || lambda_mkt <= 0) {
  stop(
    "Estimated lambda_mkt is not finite/positive. Check rf_annual, returns, and data cleaning."
  )
}

# Implied equilibrium (EXCESS) returns
Pi <- as.vector(lambda_mkt * Sigma %*% w_mkt)
names(Pi) <- colnames(returns_xts_cc)

cat("\n--- Market-implied risk aversion (lambda_mkt) ---\n")
print(lambda_mkt)

cat("\n--- Implied Equilibrium EXCESS Returns (Pi) ---\n")
print(round(Pi * 100, 2))

# ==============================================================================
# STEP 10: Black-Litterman ABSOLUTE Views
# ==============================================================================

sectors <- colnames(returns_xts_cc)
n_sectors <- length(sectors)

# Absolute Return Views (Annualized)
view_returns <- c(
  "Insurance" = 0.15,
  "Logistics" = 0.13,
  "AI" = 0.11,
  "Healthcare" = 0.085,
  "Consumer Discretionary" = 0.10,
  "Banks" = 0.09
)

# Filter views
valid_views <- intersect(names(view_returns), sectors)
view_returns <- view_returns[valid_views]
n_views <- length(valid_views)

# Build P Matrix
P <- matrix(0, nrow = n_views, ncol = n_sectors)
colnames(P) <- sectors
rownames(P) <- valid_views

for (i in seq_along(valid_views)) {
  P[i, valid_views[i]] <- 1
}

# Build Q Vector
Q <- as.numeric(view_returns) - rf_annual

# Build Omega (Uncertainty)
tau <- 0.025
view_conf <- c(
  # Lower number = Higher confidence
  "Insurance" = 0.6,
  "Logistics" = 0.4,
  "AI" = 0.8,
  "Healthcare" = 0.6,
  "Consumer Discretionary" = 0.6,
  "Banks" = 0.5
)
# Ensure view_conf matches valid_views
conf_vec <- view_conf[valid_views]
# Use heuristics if sector missing from conf vector
conf_vec[is.na(conf_vec)] <- 0.8

omega_diag <- (conf_vec * sqrt(diag(Sigma))[valid_views])^2
Omega <- diag(omega_diag)


# ============================================================
# STRESS TEST v2: vary tau (prior uncertainty) and kappa (view uncertainty)
# ============================================================

delta <- lambda_mkt # investor risk aversion in optimization
ub <- 1 # max weight per sector (prevents 100% corners)

# Constraints: sum(w)=1, 0<=w<=ub
Amat_base <- cbind(rep(1, n_sectors), diag(n_sectors), -diag(n_sectors))
bvec_base <- c(1, rep(0, n_sectors), rep(-ub, n_sectors))

tau_grid <- c(0.005, 0.01, 0.025, 0.05, 0.10)
kappa_grid <- c(0.25, 0.5, 1, 2, 4) # smaller = more confident in views

make_omega <- function(tau, kappa, Sigma, P, conf_vec) {
  # Base uncertainty of each view from covariance
  base_diag <- diag(P %*% Sigma %*% t(P)) # NOTE: no tau here

  # Scale by confidence and global kappa
  # (lower conf_vec = higher confidence -> smaller Omega)
  omega_diag <- kappa * base_diag * (conf_vec^2)

  diag(as.numeric(omega_diag))
}

run_bl <- function(tau, kappa) {
  Omega <- make_omega(
    tau = tau,
    kappa = kappa,
    Sigma = Sigma,
    P = P,
    conf_vec = conf_vec
  )

  inv_tau_Sigma <- solve(tau * Sigma)
  inv_Omega <- solve(Omega)

  posterior_Sigma <- solve(inv_tau_Sigma + t(P) %*% inv_Omega %*% P)
  posterior_mu <- as.vector(
    posterior_Sigma %*% (inv_tau_Sigma %*% Pi + t(P) %*% inv_Omega %*% Q)
  )
  names(posterior_mu) <- sectors

  # MV optimization (long-only + bounds)
  sol <- solve.QP(
    Dmat = 2 * (delta * Sigma), # <-- use return covariance
    dvec = posterior_mu,
    Amat = Amat_base,
    bvec = bvec_base,
    meq = 1
  )

  w_opt <- sol$solution
  names(w_opt) <- sectors

  list(
    posterior_mu = posterior_mu,
    posterior_Sigma = posterior_Sigma,
    weights = w_opt
  )
}

stress_grid <- tidyr::expand_grid(tau = tau_grid, kappa = kappa_grid) %>%
  mutate(out = purrr::map2(tau, kappa, run_bl))

weights_tbl <- stress_grid %>%
  mutate(weights = purrr::map(out, "weights")) %>%
  select(tau, kappa, weights) %>%
  tidyr::unnest_wider(weights) %>%
  pivot_longer(
    cols = all_of(sectors),
    names_to = "Sector",
    values_to = "Weight"
  )

# Plot
ggplot(
  weights_tbl,
  aes(x = kappa, y = Weight, group = Sector, color = Sector)
) +
  geom_line() +
  geom_point() +
  facet_wrap(~tau) +
  scale_x_log10() +
  theme_minimal() +
  labs(
    title = "Stress test: Optimal weights vs kappa (log scale), faceted by tau",
    x = "kappa (log scale)  (smaller = more confident in views)",
    y = "Optimal weight"
  )

# ============================================================
# Summarize each (tau, kappa) scenario with portfolio diagnostics
# ============================================================

summarize_out <- function(out, w_mkt, sectors, ub, Sigma) {
  w <- out$weights[sectors]
  mu <- out$posterior_mu[sectors]

  # IMPORTANT: use return covariance for vol (annualized)
  S <- Sigma[sectors, sectors]

  exp_ret <- sum(w * mu) # annualized excess return
  vol <- sqrt(as.numeric(t(w) %*% S %*% w)) # annualized vol
  sharpe <- exp_ret / vol

  tibble::tibble(
    exp_ret = exp_ret,
    vol = vol,
    sharpe = sharpe,
    l1_to_mkt = sum(abs(w - w_mkt[sectors])),
    hhi = sum(w^2),
    n_active = sum(w > 1e-4),
    hits_ub = any(abs(w - ub) < 1e-6)
  )
}

scenario_tbl <- stress_grid %>%
  mutate(
    metrics = purrr::map(
      out,
      summarize_out,
      w_mkt = w_mkt,
      sectors = sectors,
      ub = ub,
      Sigma = Sigma
    )
  ) %>%
  tidyr::unnest(metrics) %>%
  arrange(desc(sharpe)) %>%
  print(n = 25)


# ============================================================
# Pick a scenario (tau_star, kappa_star) and build final results
# ============================================================
tau_star <- 0.05
kappa_star <- 0.5

chosen_out <- stress_grid %>%
  filter(tau == tau_star, kappa == kappa_star) %>%
  pull(out) %>%
  .[[1]]

optimal_weights <- chosen_out$weights
posterior_mu <- chosen_out$posterior_mu

results <- tibble::tibble(
  Sector = sectors,
  Market_Weight = round(100 * w_mkt[sectors], 1),
  Optimal_Weight = round(100 * optimal_weights[sectors], 1),
  Equilibrium_Excess = round(100 * Pi[sectors], 1),
  Posterior_Excess = round(100 * posterior_mu[sectors], 1)
)

print(results)

results %>%
  pivot_longer(
    c(Market_Weight, Optimal_Weight),
    names_to = "Type",
    values_to = "Weight"
  ) %>%
  ggplot(aes(x = Sector, y = Weight, fill = Type)) +
  geom_col(position = "dodge") +
  theme_minimal() +
  labs(
    title = paste0("BL Weights (tau=", tau_star, ", kappa=", kappa_star, ")"),
    y = "Weight (%)"
  )

# Logistics bucket is 4 stocks, all Oil & Gas Storage & Transportation. Hence, is more
# “midstream/shipping energy logistics” than broad logistics.
