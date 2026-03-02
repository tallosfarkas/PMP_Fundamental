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

# --------------------------------------------------------------------------
# FIX: Ensure Sector_Group is present in sector_mapping_final
# (prevents "Unknown or uninitialised column: Sector_Group")
# --------------------------------------------------------------------------
if (!"Sector_Group" %in% names(sector_mapping_final)) {
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

  logistics_pattern <- "Logistics|Marine|Freight|Trucking|Airport|Rail|Transport|Shipping|Ports|Storage"

  sector_mapping_final <- sector_mapping_final %>%
    mutate(
      Sector_Group = case_when(
        GICS_SubInd_Name %in% ai_subinds ~ "AI",
        str_detect(
          GICS_SubInd_Name,
          regex(logistics_pattern, ignore_case = TRUE)
        ) ~ "Logistics",
        str_detect(
          GICS_SubInd_Name,
          "Insurance|Reinsurance|Insurance Brokers"
        ) ~ "Insurance",
        str_detect(GICS_SubInd_Name, "Banks") ~ "Banks",
        str_detect(GICS_Sector, "Health") ~ "Healthcare",
        str_detect(
          GICS_Sector,
          "Consumer Discretionary"
        ) ~ "Consumer Discretionary",
        TRUE ~ "Other"
      )
    )
}

# Sanity check (optional)
print("Sector_Group present?")
print("Sector_Group" %in% names(sector_mapping_final))
print(table(sector_mapping_final$Sector_Group))

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

delta <- lambda_mkt * 1.2 # investor risk aversion in optimization
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

# =========================
# HARD SCREENS + SECTOR-AWARE SCORING & RANKING
# =========================

# --- helpers ---
num <- function(x) readr::parse_number(as.character(x))

# Convert "percent-like" ratios (e.g., 95, 127, 283) to ratio units (0.95, 1.27, 2.83)
# Works for Debt/Assets and Loan/Deposit fields in your data.
as_ratio01 <- function(x) {
  x <- as.numeric(x)
  ifelse(is.finite(x) & x > 2, x / 100, x)
}

# Ensure percent fields are in 0..100 scale
# If already 0..100 it stays; if 0..1 it becomes 0..100
as_pct100 <- function(x) {
  x <- as.numeric(x)
  ifelse(is.finite(x) & x > 0 & x <= 1, x * 100, x)
}

# Robust z-score: never returns NA; if too few data points/zero variance => neutral (0)
z_safe <- function(x) {
  x <- as.numeric(x)
  ok <- is.finite(x)
  if (sum(ok) < 2) {
    return(rep(0, length(x)))
  }
  s <- sd(x[ok])
  if (!is.finite(s) || s == 0) {
    return(rep(0, length(x)))
  }
  out <- rep(0, length(x))
  out[ok] <- (x[ok] - mean(x[ok])) / s
  out
}

# -------------------------
# 1) HARD SCREEN PARAMETERS
# -------------------------
# Investability (keep "hard")
min_liquidity_value <- 1e6 # Avg $ value traded (20D)
min_free_float_pct <- 10 # %

# Non-financial red flags (relaxed to avoid wiping sectors; tune later)
max_netdebt_ebitda <- 10
min_int_cov <- 1.0

# Debt/Assets is percent-like in your file -> after as_ratio01 it becomes ratio
max_debt_assets_nf <- 1.50 # 150% debt/assets hard red flag for non-financials
max_debt_assets_ins <- 2.00 # insurance often looks higher; keep slightly looser

# Banks
banks_max_npl_pct <- 15 # NPL in percent (many are NA; we only filter when present)
banks_min_cet1_buffer <- -5 # buffer definition varies; allow slightly negative
banks_max_loan_deposit <- 1.30 # Loan/Deposit ratio (after conversion)

# Insurance
ins_max_fin_lev <- 30

# Optional: tail-based screen for Debt/Assets within sector (less brittle than hard cutoff)
use_debt_assets_quantile <- FALSE
debt_assets_q <- 0.95 # if enabled, drop worst 5% within each sector (only on non-NA)

# -----------------------------------
# 2) BUILD A CLEAN FUNDAMENTAL TABLE
# -----------------------------------
funda_base <- sector_mapping_final %>%
  mutate(
    Sector_Group = Sector_Group,
    MarketCap = num(Market_Cap),

    # investability
    Liquidity = num(`Avg_D_Val_Traded_20D:D-20`),
    FreeFloat = num(`Free_Float_%`),

    # valuation / yield
    PE = num(`P/E`),
    PB = num(`P/B`),
    DivY = num(`Dvd_Ind_Yld`),
    FCFY = num(`FCF_Yld`),
    EVEbitdaY = num(`EBITDA_/_EV_Yld_Adj`),

    # profitability / growth (broad)
    ROIC = num(`ROIC_LF`),
    SalesG = num(`Net_Sales_-_5_Yr_Geo_Gr_LF`),
    ROA_ROE = num(`ROA_to_ROE_LF`), # fallback for financials when ROIC missing

    # margin proxy (fallback chain to avoid NA-heavy fields)
    Margin = dplyr::coalesce(
      num(`Operating_Margin_/_EBITDA_Margin`),
      num(`EBITDA_to_Net_Sales:Q`),
      num(`EBIT/Net_Sales:Y`),
      num(`NI_Mrgn_Adj_LF`)
    ),

    # balance sheet risk (non-financials)
    NetDebtEBITDA = num(`Net_Debt_to_EBITDA_LF`),
    DebtAssets_raw = num(`Debt/Assets_LF`),
    IntCov = num(`Net_Int_Cov`),

    # banks
    CET1Buf = num(`CET1_Bffr_Pct`),
    NIM = num(`Annualized_Net_Interest_Margin`),
    NPL_raw = num(`NPL_to_Tot_Lns`),
    Prov = num(`Provision_for_Loan_Losses_T12M`),
    LoanDep_raw = num(`Tot_Ln_to_Tot_Dep_LF`),

    # insurance-ish
    FinLev = num(`Finl_Lev_LF`)
  ) %>%
  mutate(
    # Unit normalization
    DebtAssets = as_ratio01(DebtAssets_raw), # 95 -> 0.95, 154 -> 1.54, etc.
    LoanDep = as_ratio01(LoanDep_raw), # 95 -> 0.95, 127 -> 1.27, 283 -> 2.83
    NPL = as_pct100(NPL_raw) # 0.03 -> 3, 7 stays 7
  )

# -----------------------------------
# 3) INVESTABILITY SCREENS (HARD)
# -----------------------------------
funda_investable <- funda_base %>%
  filter(is.na(Liquidity) | Liquidity >= min_liquidity_value) %>%
  filter(is.na(FreeFloat) | FreeFloat >= min_free_float_pct) %>%
  filter(is.na(MarketCap) | MarketCap > 0)

# -----------------------------------
# 4) RED-FLAG SCREENS (SECTOR-AWARE)
#    IMPORTANT: apply only to rows where the relevant metric is NON-NA/finite.
#    Also: if a red-flag filter wipes a whole sector, revert to investable names for that sector.
# -----------------------------------
# -----------------------------------
# 4) RED-FLAG SCREENS (sector-aware)  ✅ FIXED FOR group_modify
# -----------------------------------

apply_redflags <- function(df, key) {
  sec <- key$Sector_Group # <-- FIX: get group name from key (.y), not df

  out <- df

  if (sec %in% c("AI", "Logistics", "Healthcare", "Consumer Discretionary")) {
    debt_cut <- max_debt_assets_nf
    if (use_debt_assets_quantile) {
      da_ok <- is.finite(out$DebtAssets)
      if (sum(da_ok) >= 5) {
        debt_cut <- stats::quantile(
          out$DebtAssets[da_ok],
          debt_assets_q,
          na.rm = TRUE
        )
      }
    }

    out <- out %>%
      filter(
        (!is.finite(NetDebtEBITDA) | NetDebtEBITDA <= max_netdebt_ebitda) &
          (!is.finite(DebtAssets) | DebtAssets <= debt_cut) &
          (!is.finite(IntCov) | IntCov >= min_int_cov)
      )
  } else if (sec == "Banks") {
    out <- out %>%
      filter(
        (!is.finite(NPL) | NPL <= banks_max_npl_pct) &
          (!is.finite(CET1Buf) | CET1Buf >= banks_min_cet1_buffer) &
          (!is.finite(LoanDep) | LoanDep <= banks_max_loan_deposit)
      )
  } else if (sec == "Insurance") {
    out <- out %>%
      filter(
        (!is.finite(DebtAssets) | DebtAssets <= max_debt_assets_ins) &
          (!is.finite(FinLev) | FinLev <= ins_max_fin_lev)
      )
  }

  # safeguard: if wiped out, revert to pre-filtered df for this sector
  if (nrow(out) == 0) df else out
}

funda_screened <- funda_investable %>%
  group_by(Sector_Group) %>%
  group_modify(apply_redflags) %>% # <-- FIX: pass function directly
  ungroup()

cat(
  "\n--- Survivors after screens (investability + red flags w/ safeguard) ---\n"
)
print(table(funda_screened$Sector_Group))

# -----------------------------------
# 5) SCORE & RANK WITH SECTOR METRICS
# -----------------------------------
funda_scored <- funda_screened %>%
  mutate(Profitability = dplyr::coalesce(ROIC, ROA_ROE)) %>%
  group_by(Sector_Group) %>%
  mutate(
    # robust z-scores
    z_Prof = z_safe(Profitability),
    z_Margin = z_safe(Margin),
    z_FCFY = z_safe(FCFY),
    z_EVy = z_safe(EVEbitdaY),
    z_PE = z_safe(PE),
    z_PB = -z_safe(PB), # lower P/B better
    z_DivY = z_safe(DivY),
    z_Growth = z_safe(SalesG),

    z_Lev_nf = -z_safe(NetDebtEBITDA),
    z_DebtA = -z_safe(DebtAssets),
    z_IntCov = z_safe(IntCov),

    z_CET1 = z_safe(CET1Buf),
    z_NIM = z_safe(NIM),
    z_NPL = -z_safe(NPL),
    z_Prov = -z_safe(Prov),
    z_LoanDep = -z_safe(LoanDep),

    z_FinLev = -z_safe(FinLev),

    Score = case_when(
      Sector_Group == "Banks" ~
        0.20 *
        z_CET1 +
        0.20 * z_NIM +
        0.20 * z_NPL +
        0.10 * z_Prov +
        0.10 * z_LoanDep +
        0.10 * z_PB +
        0.10 * z_DivY,

      Sector_Group == "Insurance" ~
        0.30 *
        z_Prof +
        0.15 * z_Margin +
        0.20 * z_PB +
        0.15 * z_DivY +
        0.10 * z_Growth +
        0.10 * (z_DebtA + z_FinLev) / 2,

      TRUE ~
        0.30 *
        ((z_Prof + z_Margin) / 2) +
        0.30 * ((z_FCFY + z_EVy - z_PE) / 3) +
        0.20 * z_Growth +
        0.20 * z_Lev_nf
    ),

    Rank_in_Sector = dplyr::dense_rank(dplyr::desc(Score))
  ) %>%
  ungroup() %>%
  arrange(Sector_Group, Rank_in_Sector)

# =========================
# PICK EXACTLY total_stocks (3–5) WITH THRESHOLD + REDISTRIBUTION (UPDATED)
# =========================

total_stocks <- 4
threshold <- 0.15
temperature <- 1.0

# IMPORTANT FIX: only allocate to sectors that actually have candidates after screening
available_sectors <- intersect(sectors, unique(funda_scored$Sector_Group))

sector_w <- optimal_weights[available_sectors]
sector_w <- sector_w[sector_w > 1e-10]
sector_w <- sector_w / sum(sector_w)

# Keep only sectors above threshold; fallback to top sectors if none qualify
keep_sectors <- names(sector_w)[sector_w > threshold]
if (length(keep_sectors) == 0) {
  keep_sectors <- names(sort(sector_w, decreasing = TRUE))[
    1:min(total_stocks, length(sector_w))
  ]
}

# Redistribute excluded sector weights proportionally among kept sectors
sector_w_keep <- sector_w[keep_sectors]
sector_w_keep <- sector_w_keep / sum(sector_w_keep)

# Allocate stock slots per sector so total = total_stocks
slots_raw <- sector_w_keep * total_stocks
slots <- floor(slots_raw)
slots[slots < 1] <- 1

# Largest remainder adjustment
while (sum(slots) < total_stocks) {
  frac <- slots_raw - floor(slots_raw)
  add_to <- names(which.max(frac))
  slots[add_to] <- slots[add_to] + 1
}
while (sum(slots) > total_stocks) {
  candidates <- names(slots)[slots > 1]
  if (length(candidates) == 0) {
    break
  }
  rm_from <- candidates[which.min(sector_w_keep[candidates])]
  slots[rm_from] <- slots[rm_from] - 1
}

# Pick top N stocks per included sector
funda_pick_base <- funda_scored %>%
  mutate(Score_clean = ifelse(is.finite(Score), Score, NA_real_))

pick_list <- lapply(names(slots), function(sec) {
  n <- slots[[sec]]
  df <- funda_pick_base %>%
    filter(Sector_Group == sec) %>%
    arrange(desc(Score_clean))
  df %>% slice_head(n = n)
})

picks <- bind_rows(pick_list) %>%
  mutate(Sector_Weight = sector_w_keep[Sector_Group])

# Softmax weights within sector based on Score
softmax <- function(x, temp = 1) {
  if (all(is.na(x))) {
    return(rep(1 / length(x), length(x)))
  }
  x2 <- x
  x2[is.na(x2)] <- min(x2, na.rm = TRUE)
  x2 <- x2 / temp
  ex <- exp(x2 - max(x2))
  ex / sum(ex)
}

picks <- picks %>%
  group_by(Sector_Group) %>%
  mutate(
    within_sector_w = softmax(Score_clean, temp = temperature),
    Stock_Weight = Sector_Weight * within_sector_w
  ) %>%
  ungroup() %>%
  select(Sector_Group, Ticker, Name, Score, Sector_Weight, Stock_Weight) %>%
  arrange(desc(Sector_Weight), desc(Score))

print(picks)
cat("\nIncluded sectors:", paste(names(sector_w_keep), collapse = ", "), "\n")
cat("Slots per sector:\n")
print(slots)
cat("Sum stock weights:", round(sum(picks$Stock_Weight), 6), "\n")


# =========================
# 1) SCORE BREAKDOWN FOR PICKS  (UPDATED + ROBUST)
# =========================

# Helper: safe relative difference
rel_diff <- function(x, y) abs(x - y) / pmax(abs(y), 1e-6)

# Join picks back to funda_scored to grab all the factor columns
# NOTE: We intentionally rename the funda_scored Name/Score to avoid .x/.y confusion.
picks_detail <- picks %>%
  left_join(
    funda_scored %>%
      transmute(
        Sector_Group,
        Ticker,
        Name_funda = Name,
        Score_funda = Score,

        # raw metrics
        Profitability,
        ROIC,
        ROA_ROE,
        Margin,
        FCFY,
        EVEbitdaY,
        PE,
        PB,
        DivY,
        SalesG,
        NetDebtEBITDA,
        DebtAssets,
        IntCov,
        CET1Buf,
        NIM,
        NPL,
        Prov,
        LoanDep,
        FinLev,

        # z-scores
        dplyr::across(dplyr::starts_with("z_"), identity)
      ),
    by = c("Sector_Group", "Ticker")
  ) %>%
  mutate(
    # Standardize Name/Score into single columns
    Name = dplyr::coalesce(Name, Name_funda),
    Score = dplyr::coalesce(Score, Score_funda)
  )

# Compute component contributions (must mirror your Score formulas)
picks_breakdown <- picks_detail %>%
  mutate(
    # Banks
    c_z_CET1 = ifelse(Sector_Group == "Banks", 0.20 * z_CET1, 0),
    c_z_NIM = ifelse(Sector_Group == "Banks", 0.20 * z_NIM, 0),
    c_z_NPL = ifelse(Sector_Group == "Banks", 0.20 * z_NPL, 0),
    c_z_Prov = ifelse(Sector_Group == "Banks", 0.10 * z_Prov, 0),
    c_z_LoanDep = ifelse(Sector_Group == "Banks", 0.10 * z_LoanDep, 0),
    c_z_PB_b = ifelse(Sector_Group == "Banks", 0.10 * z_PB, 0),
    c_z_DivY_b = ifelse(Sector_Group == "Banks", 0.10 * z_DivY, 0),

    # Insurance
    c_z_Prof_i = ifelse(Sector_Group == "Insurance", 0.30 * z_Prof, 0),
    c_z_Margin_i = ifelse(Sector_Group == "Insurance", 0.15 * z_Margin, 0),
    c_z_PB_i = ifelse(Sector_Group == "Insurance", 0.20 * z_PB, 0),
    c_z_DivY_i = ifelse(Sector_Group == "Insurance", 0.15 * z_DivY, 0),
    c_z_Growth_i = ifelse(Sector_Group == "Insurance", 0.10 * z_Growth, 0),
    c_z_DebtA_i = ifelse(Sector_Group == "Insurance", 0.05 * z_DebtA, 0),
    c_z_FinLev_i = ifelse(Sector_Group == "Insurance", 0.05 * z_FinLev, 0),

    # Non-financials (AI, Logistics, Healthcare, Consumer Discretionary)
    c_z_Prof_nf = ifelse(
      !(Sector_Group %in% c("Banks", "Insurance")),
      0.15 * z_Prof,
      0
    ),
    c_z_Margin_nf = ifelse(
      !(Sector_Group %in% c("Banks", "Insurance")),
      0.15 * z_Margin,
      0
    ),
    c_z_FCFY_nf = ifelse(
      !(Sector_Group %in% c("Banks", "Insurance")),
      0.10 * z_FCFY,
      0
    ),
    c_z_EVy_nf = ifelse(
      !(Sector_Group %in% c("Banks", "Insurance")),
      0.10 * z_EVy,
      0
    ),
    c_z_PE_nf = ifelse(
      !(Sector_Group %in% c("Banks", "Insurance")),
      -0.10 * z_PE,
      0
    ),
    c_z_Growth_nf = ifelse(
      !(Sector_Group %in% c("Banks", "Insurance")),
      0.20 * z_Growth,
      0
    ),
    c_z_Lev_nf = ifelse(
      !(Sector_Group %in% c("Banks", "Insurance")),
      0.20 * z_Lev_nf,
      0
    )
  )

# Long table: one row per component per picked stock
picks_contrib_long <- picks_breakdown %>%
  select(Sector_Group, Ticker, Name, Score, starts_with("c_")) %>%
  pivot_longer(
    cols = starts_with("c_"),
    names_to = "Component",
    values_to = "Contribution"
  ) %>%
  filter(abs(Contribution) > 1e-10) %>%
  arrange(Sector_Group, Ticker, desc(abs(Contribution)))

# Sanity check: contributions sum back to Score
picks_contrib_check <- picks_contrib_long %>%
  group_by(Sector_Group, Ticker) %>%
  summarise(
    Score = first(Score),
    Sum_Contrib = sum(Contribution),
    Diff = Score - Sum_Contrib,
    .groups = "drop"
  )

cat("\n--- PICKS: CONTRIBUTION CHECK (Score vs Sum of components) ---\n")
print(picks_contrib_check)

cat("\n--- PICKS: COMPONENT CONTRIBUTIONS (long) ---\n")
print(picks_contrib_long)

cat("\n--- PICKS: RAW METRICS + Z-SCORES (for context) ---\n")
print(
  picks_detail %>%
    select(
      Sector_Group,
      Ticker,
      Name,
      Score,
      Profitability,
      ROIC,
      ROA_ROE,
      Margin,
      FCFY,
      EVEbitdaY,
      PE,
      PB,
      DivY,
      SalesG,
      NetDebtEBITDA,
      DebtAssets,
      IntCov,
      CET1Buf,
      NIM,
      NPL,
      Prov,
      LoanDep,
      FinLev,
      z_Prof,
      z_Margin,
      z_FCFY,
      z_EVy,
      z_PE,
      z_PB,
      z_DivY,
      z_Growth,
      z_Lev_nf,
      z_DebtA,
      z_CET1,
      z_NIM,
      z_NPL,
      z_Prov,
      z_LoanDep,
      z_FinLev
    )
)

# =========================
# 2) STOCKS WITHIN ±20% SCORE OF EACH PICK (PER PICK, NO MANY-TO-MANY)
# =========================

score_band <- 0.20

pick_ref <- picks_detail %>%
  distinct(Sector_Group, Pick_Ticker = Ticker, Pick_Score = Score, Name) %>%
  select(Sector_Group, Pick_Ticker, Pick_Score)

within_20pct <- purrr::map_dfr(seq_len(nrow(pick_ref)), function(i) {
  sec <- pick_ref$Sector_Group[i]
  pt <- pick_ref$Pick_Ticker[i]
  ps <- pick_ref$Pick_Score[i]

  funda_scored %>%
    filter(Sector_Group == sec) %>%
    mutate(
      Pick_Ticker = pt,
      Pick_Score = ps,
      RelDiff = rel_diff(Score, ps),
      Is_Pick = (Ticker == pt)
    ) %>%
    filter(RelDiff <= score_band) %>%
    arrange(RelDiff, desc(Score))
})

cat(
  "\n--- ALL STOCKS WITHIN ±20% OF EACH PICK'S SCORE (same sector, per-pick) ---\n"
)
print(
  within_20pct %>%
    select(
      Sector_Group,
      Ticker,
      Name,
      Score,
      Pick_Ticker,
      Pick_Score,
      RelDiff,
      Is_Pick
    )
)

cat("\n--- CLOSEST ALTERNATIVES (exclude the picked stock) ---\n")
print(
  within_20pct %>%
    filter(!Is_Pick) %>%
    select(Sector_Group, Ticker, Name, Score, Pick_Ticker, Pick_Score, RelDiff)
)
