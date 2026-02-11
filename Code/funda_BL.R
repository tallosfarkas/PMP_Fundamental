# ==============================================================================
# Load Required Packages
# ==============================================================================
if(!require("pacman")) {
  install.packages("pacman")
}
pacman::p_load(tidyverse, readxl, zoo, xts, quadprog, PerformanceAnalytics, corrplot, corpcor)


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

sapply(sector_mapping[c("GICS_Sector", "GICS_Ind_Name", "GICS_SubInd_Name")], unique)

# ==============================================================================
# STEP 3: Clean Sector Mapping (ROBUST VERSION)
# ==============================================================================

# 1. Define AI subindustries (Keep as is)
ai_subinds <- c(
  "Electric Utilities", "Independent Power Producers & Energy Traders",
  "Renewable Electricity", "Electrical Components & Equipment",
  "Electrical Equipment & Instruments", "Electronic Equipment & Instruments",
  "Electronic Equipment, Instruments & Components", "Application Software"
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
      str_detect(GICS_SubInd_Name, regex(logistics_pattern, ignore_case = TRUE)) ~ "Logistics",
      
      # 3. Financials Split
      str_detect(GICS_SubInd_Name, "Insurance|Reinsurance|Insurance Brokers") ~ "Insurance",
      str_detect(GICS_SubInd_Name, "Banks") ~ "Banks",
      
      # 4. Standard Sectors
      str_detect(GICS_Sector, "Health") ~ "Healthcare",
      str_detect(GICS_Sector, "Consumer Discretionary") ~ "Consumer Discretionary",
      
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
  "Insurance",              # <--- Ensuring Insurance is here
  "Logistics",              # <--- Ensuring Logistics is here
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

# 2. CRITICAL FIX: Fill NAs *Daily* before aggregating to Weekly
# This prevents "last()" from picking an NA if the specific Friday/Sunday is missing
price_filled <- price_filtered %>%
  mutate(across(-Date, ~ zoo::na.locf(., na.rm = FALSE))) %>%
  mutate(across(-Date, ~ zoo::na.locf(., fromLast = TRUE))) # Fill leading NAs too

# 3. Convert to WEEKLY data
price_weekly <- price_filled %>%
  mutate(Week_End = ceiling_date(Date, "week")) %>%
  group_by(Week_End) %>%
  summarise(across(everything(), last), .groups = 'drop') %>%
  select(-Date) %>%
  rename(Date = Week_End) %>%
  slice(1:(n()-1)) # Remove incomplete last week

# 4. Filter for valid tickers (Relaxed Threshold)
# Allow up to 60% missing data to capture recent IPOs (like ADNOC LS)
# Or just ensure we have *some* data.
missing_pct <- colSums(is.na(select(price_weekly, -Date))) / nrow(price_weekly)
valid_tickers <- names(missing_pct[missing_pct < 0.60]) 

# If valid_tickers is empty, force keep the top 50 by market cap
if(length(valid_tickers) < 5) {
  print("Warning: Strict filtering removed too many stocks. Selecting top stocks by Market Cap instead.")
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
  mutate(across(-Date, ~ c(NA, diff(log(.))))) %>% # Log returns
  slice(-1) %>%
  mutate(across(-Date, ~ ifelse(is.infinite(.), NA, .))) %>%
  # Zero fill remaining NAs (e.g., if stock didn't exist yet, return is 0)
  # This is necessary for Black-Litterman covariance matrix
  mutate(across(-Date, ~ replace_na(., 0))) 

# ==============================================================================
# STEP 8: Create Market-Cap Weighted Sector Indices
# ==============================================================================

returns_long <- returns_df %>%
  pivot_longer(-Date, names_to = "Ticker", values_to = "Return") %>%
  left_join(sector_mapping_final %>% select(Ticker, Sector_Group, Market_Cap_Num), by = "Ticker") %>%
  filter(!is.na(Sector_Group), Market_Cap_Num > 0) # Removed !is.na(Return) to keep 0s

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
returns_wide[is.na(returns_wide)] <- 0

returns_xts <- xts(returns_wide[,-1], order.by = returns_wide$Date)

# ==============================================================================
# STEP 9: Market Statistics (Prior) with Robust Covariance (Corpcor)
# ==============================================================================

# Install lightweight shrinkage package
if(!require("corpcor")) install.packages("corpcor")
library(corpcor)

freq <- 52 # Weekly frequency

# 1. Historical Returns (Mu)
mu_hist <- colMeans(returns_xts) * freq

# 2. Robust Covariance (Ledoit-Wolf Shrinkage via Corpcor)
# This estimates the covariance matrix by shrinking towards constant correlation
# It is extremely fast and stable
Sigma_shrink <- cov.shrink(returns_xts, verbose = FALSE) 

# Convert to standard matrix and annualize
Sigma <- as.matrix(Sigma_shrink) * freq

# Ensure column names are preserved
colnames(Sigma) <- colnames(returns_xts)
rownames(Sigma) <- colnames(returns_xts)

# 3. Calculate Market Weights
sector_mcap <- sector_indices %>%
  group_by(Sector_Group) %>%
  summarise(MCap = mean(TotalMarketCap, na.rm = TRUE)) 

w_mkt <- setNames(sector_mcap$MCap / sum(sector_mcap$MCap), sector_mcap$Sector_Group)
w_mkt <- w_mkt[colnames(returns_xts)] # Ensure order matches

# 4. Risk Aversion & Implied Equilibrium
lambda <- 2.5
Pi <- as.vector(lambda * Sigma %*% w_mkt)
names(Pi) <- colnames(returns_xts)

cat("\n--- Implied Equilibrium Returns (Pi) ---\n")
print(round(Pi * 100, 2))

# ==============================================================================
# STEP 10: Black-Litterman ABSOLUTE Views
# ==============================================================================

sectors <- colnames(returns_xts)
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

for(i in seq_along(valid_views)) {
  P[i, valid_views[i]] <- 1
}

# Build Q Vector
Q <- as.numeric(view_returns)

# Build Omega (Uncertainty)
tau <- 0.025
view_conf <- c( # Lower number = Higher confidence
  "Insurance" = 0.6, "Logistics" = 0.4, "AI" = 0.8,
  "Healthcare" = 0.6, "Consumer Discretionary" = 0.6, "Banks" = 0.5
)
# Ensure view_conf matches valid_views
conf_vec <- view_conf[valid_views]
# Use heuristics if sector missing from conf vector
conf_vec[is.na(conf_vec)] <- 0.8 

omega_diag <- (conf_vec * sqrt(diag(Sigma))[valid_views])^2
Omega <- diag(omega_diag)

# ==============================================================================
# STEP 11: Posterior Calculation
# ==============================================================================

inv_tau_Sigma <- solve(tau * Sigma)
inv_Omega <- solve(Omega)
posterior_Sigma <- solve(inv_tau_Sigma + t(P) %*% inv_Omega %*% P)
posterior_mu <- as.vector(posterior_Sigma %*% (inv_tau_Sigma %*% Pi + t(P) %*% inv_Omega %*% Q))
names(posterior_mu) <- sectors

# ==============================================================================
# STEP 12: Optimization
# ==============================================================================

Amat <- cbind(rep(1, n_sectors), diag(n_sectors))
bvec <- c(1, rep(0, n_sectors))

# Maximize: mu*w - (lambda/2)*w*S*w
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
# STEP 13 & 15: Summary & Plot
# ==============================================================================

results <- data.frame(
  Sector = sectors,
  Market_Weight = round(w_mkt * 100, 1),
  Optimal_Weight = round(optimal_weights * 100, 1),
  Equilibrium_Ret = round(Pi * 100, 1),
  Posterior_Ret = round(posterior_mu * 100, 1)
)

print("--- Final Results ---")
print(results)

# Define Colors
sector_colors <- c(
  "Banks" = "#1f77b4", "Insurance" = "#00cccc", "Logistics" = "#ff7f0e",
  "AI" = "#9467bd", "Healthcare" = "#2ca02c", "Consumer Discretionary" = "#d62728"
)

# Plot
results %>%
  pivot_longer(c(Market_Weight, Optimal_Weight), names_to = "Type", values_to = "Weight") %>%
  ggplot(aes(x = Sector, y = Weight, fill = Sector)) +
  geom_bar(stat = "identity", position = "dodge") +
  facet_wrap(~Type) +
  scale_fill_manual(values = sector_colors) +
  theme_minimal() +
  labs(title = "BL Optimization: Weekly Data & Absolute Views", y = "Weight %") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))