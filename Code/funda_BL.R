# ==============================================================================
# Load Required Packages
# ==============================================================================
install.packages(c("tidyverse", "readxl", "zoo", "xts", "quadprog", 
                   "PerformanceAnalytics", "corrplot"))

library(tidyverse)
library(readxl)
library(zoo)
library(xts)
library(quadprog)
library(PerformanceAnalytics)
library(corrplot)

# ==============================================================================
# STEP 1: Load Data
# ==============================================================================

# Load price/returns data from .rds file
# Each column = company ticker, each row = date
price_data <- readRDS("bloomberg_funda.rds")

# Convert to data frame if it's a matrix
if(is.matrix(price_data)) {
  dates <- rownames(price_data)
  price_data <- as.data.frame(price_data)
  price_data$Date <- as.Date(dates)
} else if("Date" %in% names(price_data)) {
  price_data$Date <- as.Date(price_data$Date)
} else {
  # If no date column, create one (adjust as needed)
  price_data$Date <- as.Date(rownames(price_data))
}

# Load ticker-sector mapping from Excel
sector_mapping <- read_excel("stock.xlsx")

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
  if(is.xts(ts_obj) || is.zoo(ts_obj)) {
    df <- data.frame(
      Date = index(ts_obj),
      Price = as.numeric(coredata(ts_obj))
    )
  } else if(is.data.frame(ts_obj)) {
    df <- ts_obj
    if(ncol(df) >= 2) {
      names(df)[1:2] <- c("Date", "Price")
    }
  } else if(is.numeric(ts_obj) || is.matrix(ts_obj)) {
    # Numeric vector or matrix - create dates
    df <- data.frame(
      Date = seq.Date(from = as.Date("2005-01-01"), 
                      by = "day", 
                      length.out = length(ts_obj)),
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
  ticker <- ifelse(!is.null(names(price_data)[i]), 
                   names(price_data)[i], 
                   paste0("Stock_", i))
  convert_to_df(price_data[[i]], ticker)
})

# Remove NULL elements
price_df_list <- price_df_list[!sapply(price_df_list, is.null)]

# Combine into long format
price_df_long <- bind_rows(price_df_list)

print(paste("Total rows in price data:", nrow(price_df_long)))
print(paste("Unique tickers:", length(unique(price_df_long$Ticker))))
print(paste("Date range:", min(price_df_long$Date, na.rm = TRUE), 
            "to", max(price_df_long$Date, na.rm = TRUE)))

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


sector_mapping_clean <- sector_mapping %>%
  mutate(
    # Standardize sector names
    Sector_Group = case_when(
      str_detect(GICS_Sector, "Financial") ~ "Financials",
      str_detect(GICS_Sector, "Real Estate") ~ "Real Estate",
      str_detect(GICS_Sector, "Energy") ~ "Energy",
      str_detect(GICS_Sector, "Industrial") ~ "Industrials",
      str_detect(GICS_Sector, "Utilities") ~ "Utilities",
      str_detect(GICS_Sector, "Consumer") ~ "Consumer",
      str_detect(GICS_Sector, "Technology") ~ "Technology",
      str_detect(GICS_Sector, "Communication") ~ "Communication",
      str_detect(GICS_Sector, "Materials") ~ "Materials",
      str_detect(GICS_Sector, "Health") ~ "Healthcare",
      TRUE ~ GICS_Sector
    ),
    
    # Clean market cap (convert to numeric if it's character)
    Market_Cap_Num = as.numeric(Market_Cap)
  )

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

# Financials (Banking/Insurance), Logistics, Energy (Renewables), Real Estate

target_sectors <- c("Financials", "Real Estate", "Energy")

# Note: Logistics might be under "Industrials" 
# Check if we need to add it
if("Industrials" %in% sector_mapping_matched$Sector_Group) {
  # Check for logistics companies
  logistics_companies <- sector_mapping_matched %>%
    filter(str_detect(GICS_Ind_Name, "Logistic|Transport|Freight") |
             str_detect(GICS_SubInd_Name, "Logistic|Transport|Freight"))
  
  if(nrow(logistics_companies) > 0) {
    print(paste("\nFound", nrow(logistics_companies), "logistics companies"))
    sector_mapping_matched <- sector_mapping_matched %>%
      mutate(Sector_Group = ifelse(
        str_detect(GICS_Ind_Name, "Logistic|Transport|Freight") |
          str_detect(GICS_SubInd_Name, "Logistic|Transport|Freight"),
        "Logistics",
        Sector_Group
      ))
    target_sectors <- c(target_sectors, "Logistics")
  }
}


# Filter for target sectors
sector_mapping_filtered <- sector_mapping_matched %>%
  filter(Sector_Group %in% target_sectors)

target_tickers <- sector_mapping_filtered$Ticker

print("\nCompanies by target sector:")
print(table(sector_mapping_filtered$Sector_Group))

# ==============================================================================
# STEP 6: Filter for Last 10 Years
# ==============================================================================

cutoff_date <- max(price_df_matched$Date, na.rm = TRUE) - years(10)

price_df_final <- price_df_matched %>%
  filter(Date >= cutoff_date) %>%
  select(Date, all_of(target_tickers))

# Remove tickers with >20% missing data
missing_pct <- colSums(is.na(select(price_df_final, -Date))) / nrow(price_df_final)
valid_tickers <- names(missing_pct[missing_pct < 0.2])

price_df_final <- price_df_final %>%
  select(Date, all_of(valid_tickers))

print(paste("\nCompanies after filtering:", length(valid_tickers)))
print(paste("Date range:", min(price_df_final$Date), "to", max(price_df_final$Date)))
print(paste("Observations:", nrow(price_df_final)))

# Update sector mapping
sector_mapping_final <- sector_mapping_filtered %>%
  filter(Ticker %in% valid_tickers)

print("\nFinal distribution by sector:")
print(table(sector_mapping_final$Sector_Group))

print("\nFinal distribution by country:")
print(table(sector_mapping_final$Cntry_Terrtry_Fl_Name))

# ==============================================================================
# STEP 7: Calculate Returns
# ==============================================================================

returns_df <- price_df_final %>%
  arrange(Date) %>%
  mutate(across(-Date, ~c(NA, diff(log(.)))))  # Log returns

# Remove first row
returns_df <- returns_df[-1, ]

# Handle any infinite values
returns_df <- returns_df %>%
  mutate(across(-Date, ~ifelse(is.infinite(.), NA, .)))

# Fill NAs using last observation carried forward
returns_df <- returns_df %>%
  mutate(across(-Date, ~zoo::na.locf(., na.rm = FALSE)))

# Remove remaining NAs
returns_df <- na.omit(returns_df)

print("\nReturns data:")
print(paste("Dimensions:", nrow(returns_df), "x", ncol(returns_df)))

# ==============================================================================
# STEP 8: Create Market-Cap Weighted Sector Indices
# ==============================================================================

# Convert to long format and join with sector information
returns_long <- returns_df %>%
  pivot_longer(-Date, names_to = "Ticker", values_to = "Return") %>%
  left_join(
    sector_mapping_final %>% 
      select(Ticker, Sector_Group, Market_Cap_Num, Name, 
             Cntry_Terrtry_Fl_Name, GICS_SubInd_Name),
    by = "Ticker"
  ) %>%
  filter(!is.na(Return), !is.na(Sector_Group), !is.na(Market_Cap_Num), Market_Cap_Num > 0)

print(paste("  Total observations:", nrow(returns_long)))
print(paste("  Unique dates:", length(unique(returns_long$Date))))
print(paste("  Unique tickers:", length(unique(returns_long$Ticker))))

# Calculate market-cap weighted sector indices
# Market_Cap_Num is ALREADY in returns_long, no need for another join!
sector_indices <- returns_long %>%
  group_by(Date, Sector_Group) %>%
  summarise(
    # Market-cap weighted return
    IndexReturn = weighted.mean(Return, w = Market_Cap_Num, na.rm = TRUE),
    # Alternative manual calculation (same result):
    # IndexReturn = sum(Return * Market_Cap_Num, na.rm = TRUE) / sum(Market_Cap_Num, na.rm = TRUE),
    NumCompanies = n(),
    TotalMarketCap = sum(Market_Cap_Num, na.rm = TRUE),
    .groups = 'drop'
  )

print("\nSector indices summary:")
sector_summary <- sector_indices %>% 
  group_by(Sector_Group) %>% 
  summarise(
    AvgReturn = mean(IndexReturn, na.rm = TRUE),
    StdDev = sd(IndexReturn, na.rm = TRUE),
    NumObs = n(),
    AvgNumCompanies = mean(NumCompanies),
    AvgMarketCapBn = mean(TotalMarketCap) / 1e9,
    .groups = 'drop'
  )
print(sector_summary)

# Convert to wide format
returns_wide <- sector_indices %>%
  select(Date, Sector_Group, IndexReturn) %>%
  pivot_wider(names_from = Sector_Group, values_from = IndexReturn) %>%
  arrange(Date)


# Convert to xts
dates_vec <- returns_wide$Date
returns_matrix <- returns_wide %>% select(-Date)
returns_xts <- xts(returns_matrix, order.by = dates_vec)

# Clean data
returns_xts <- na.omit(returns_xts)

print(paste("Dimensions:", nrow(returns_xts), "x", ncol(returns_xts)))
print(paste("Date range:", min(index(returns_xts)), "to", max(index(returns_xts))))

# Verify no NAs remain
if(any(is.na(returns_xts))) {
  warning("WARNING: NAs still present in returns_xts")
  print(colSums(is.na(returns_xts)))
}

# ==============================================================================
# STEP 9: Calculate Statistics for Black-Litterman
# ==============================================================================

# Determine frequency
date_diff <- as.numeric(diff(index(returns_xts)))
avg_diff <- median(date_diff, na.rm = TRUE)

if(avg_diff <= 2) {
  freq <- 252
  freq_name <- "Daily"
} else if(avg_diff <= 10) {
  freq <- 52
  freq_name <- "Weekly"
} else {
  freq <- 12
  freq_name <- "Monthly"
}

print(paste("\nData frequency:", freq_name))
print(paste("Average days between observations:", round(avg_diff, 1)))

# Annualized statistics
mu_hist <- colMeans(returns_xts, na.rm = TRUE) * freq
Sigma <- cov(returns_xts, use = "complete.obs") * freq

# Market-cap based weights
sector_market_caps <- sector_indices %>%
  group_by(Sector_Group) %>%
  summarise(TotalMarketCap = mean(TotalMarketCap, na.rm = TRUE), .groups = 'drop') %>%
  arrange(match(Sector_Group, colnames(returns_xts)))

w_mkt <- sector_market_caps$TotalMarketCap / sum(sector_market_caps$TotalMarketCap)
names(w_mkt) <- colnames(returns_xts)

# Risk aversion parameter
lambda <- 2.5

# Implied equilibrium returns
Pi <- as.vector(lambda * Sigma %*% w_mkt)
names(Pi) <- colnames(returns_xts)

cat("\n", rep("=", 90), "\n", sep = "")
cat("MARKET STATISTICS\n")
cat(rep("=", 90), "\n", sep = "")
print(data.frame(
  Sector = names(mu_hist),
  Historical_Return_Pct = round(mu_hist * 100, 2),
  Volatility_Pct = round(sqrt(diag(Sigma)) * 100, 2),
  Market_Weight_Pct = round(w_mkt * 100, 2),
  Market_Cap_Bn = round(sector_market_caps$TotalMarketCap / 1e9, 2),
  Equilibrium_Return_Pct = round(Pi * 100, 2)
))
cat(rep("=", 90), "\n", sep = "")

# ==============================================================================
# STEP 10: Black-Litterman Views
# ==============================================================================

sectors <- colnames(returns_xts)
n_sectors <- length(sectors)

# Define your views - CUSTOMIZE THESE!

views <- list()

# View 1: Energy sector (Renewables) will outperform by 8%
if("Energy" %in% sectors) {
  views$energy <- list(
    P = ifelse(sectors == "Energy", 1, 0),
    Q = 0.08,
    description = "Energy (renewables focus) outperforms by 8%"
  )
}

# View 2: Real Estate (Tourism & Data Centers) outperforms Financials by 4%
if("Real Estate" %in% sectors && "Financials" %in% sectors) {
  views$realestate_vs_financials <- list(
    P = ifelse(sectors == "Real Estate", 1, 
               ifelse(sectors == "Financials", -1, 0)),
    Q = 0.04,
    description = "Real Estate outperforms Financials by 4%"
  )
}

# View 3: If you have Logistics, you can add a view
if("Logistics" %in% sectors) {
  views$logistics <- list(
    P = ifelse(sectors == "Logistics", 1, 0),
    Q = 0.06,
    description = "Logistics sector outperforms by 6%"
  )
}

# Construct P and Q matrices
n_views <- length(views)
P <- matrix(0, nrow = n_views, ncol = n_sectors)
Q <- numeric(n_views)
view_names <- character(n_views)

for(i in seq_along(views)) {
  P[i, ] <- views[[i]]$P
  Q[i] <- views[[i]]$Q
  view_names[i] <- views[[i]]$description
}

colnames(P) <- sectors
rownames(P) <- view_names

cat("\n", rep("=", 90), "\n", sep = "")
cat("YOUR INVESTMENT VIEWS\n")
cat(rep("=", 90), "\n", sep = "")
print("P Matrix (View specification):")
print(P)
print("\nQ Vector (Expected outperformance):")
print(data.frame(View = view_names, Expected_Return_Pct = Q * 100))
cat(rep("=", 90), "\n", sep = "")

# Confidence in views (lower = more confident)
tau <- 0.025
view_confidence <- rep(0.5, n_views)  # Adjust per view if needed
Omega <- diag(tau * view_confidence)

# ==============================================================================
# STEP 11: Black-Litterman Posterior Returns
# ==============================================================================

inv_tau_Sigma <- solve(tau * Sigma)
inv_Omega <- solve(Omega)

posterior_precision <- inv_tau_Sigma + t(P) %*% inv_Omega %*% P
posterior_Sigma <- solve(posterior_precision)
posterior_mu <- as.vector(posterior_Sigma %*% 
                            (inv_tau_Sigma %*% Pi + t(P) %*% inv_Omega %*% Q))
names(posterior_mu) <- sectors

cat("\n", rep("=", 90), "\n", sep = "")
cat("BLACK-LITTERMAN POSTERIOR RETURNS\n")
cat(rep("=", 90), "\n", sep = "")
print(data.frame(
  Sector = sectors,
  Equilibrium_Pct = round(Pi * 100, 2),
  Posterior_Pct = round(posterior_mu * 100, 2),
  Change_Pct = round((posterior_mu - Pi) * 100, 2)
))
cat(rep("=", 90), "\n", sep = "")

# ==============================================================================
# STEP 12: Portfolio Optimization
# ==============================================================================

target_return <- mean(posterior_mu)
n <- length(posterior_mu)

Amat <- cbind(
  rep(1, n),           # weights sum to 1
  posterior_mu,        # return constraint
  diag(n)             # no short selling
)

bvec <- c(1, target_return, rep(0, n))

solution <- solve.QP(
  Dmat = 2 * posterior_Sigma,
  dvec = rep(0, n),
  Amat = Amat,
  bvec = bvec,
  meq = 1
)

optimal_weights <- solution$solution
names(optimal_weights) <- sectors

# ==============================================================================
# STEP 13: Results Summary
# ==============================================================================

results_summary <- data.frame(
  Sector = sectors,
  Num_Companies = as.vector(table(sector_mapping_final$Sector_Group)[sectors]),
  Market_Cap_Bn = round(sector_market_caps$TotalMarketCap / 1e9, 2),
  Historical_Return_Pct = round(mu_hist * 100, 2),
  Volatility_Pct = round(sqrt(diag(Sigma)) * 100, 2),
  Equilibrium_Return_Pct = round(Pi * 100, 2),
  Posterior_Return_Pct = round(posterior_mu * 100, 2),
  Market_Weight_Pct = round(w_mkt * 100, 2),
  Optimal_Weight_Pct = round(optimal_weights * 100, 2),
  Weight_Change_Pct = round((optimal_weights - w_mkt) * 100, 2)
) %>%
  mutate(
    Sharpe_Ratio = round(Posterior_Return_Pct / Volatility_Pct, 3),
    Return_Contribution_Pct = round(Optimal_Weight_Pct * Posterior_Return_Pct / 100, 2)
  )

cat("\n", rep("=", 100), "\n", sep = "")
cat("BLACK-LITTERMAN OPTIMIZATION RESULTS\n")
cat(rep("=", 100), "\n", sep = "")
print(results_summary)
cat(rep("=", 100), "\n", sep = "")

# ==============================================================================
# STEP 14: Stock Selection Within Sectors
# ==============================================================================

stock_allocation <- sector_mapping_final %>%
  select(Ticker, Name, Sector_Group, Cntry_Terrtry_Fl_Name, 
         Market_Cap_Num, GICS_SubInd_Name) %>%
  left_join(
    results_summary %>% select(Sector, Optimal_Weight_Pct),
    by = c("Sector_Group" = "Sector")
  ) %>%
  # Market-cap weight within each sector
  group_by(Sector_Group) %>%
  mutate(
    Sector_Total_MCap = sum(Market_Cap_Num, na.rm = TRUE),
    Weight_In_Sector = Market_Cap_Num / Sector_Total_MCap,
    Stock_Weight_Pct = (Optimal_Weight_Pct / 100) * Weight_In_Sector * 100,
    Market_Cap_Bn = round(Market_Cap_Num / 1e9, 2)
  ) %>%
  ungroup() %>%
  arrange(desc(Stock_Weight_Pct)) %>%
  select(Ticker, Name, Sector_Group, Cntry_Terrtry_Fl_Name, 
         GICS_SubInd_Name, Market_Cap_Bn, Stock_Weight_Pct)

cat("\n", rep("=", 100), "\n", sep = "")
cat("TOP 20 STOCK ALLOCATIONS\n")
cat(rep("=", 100), "\n", sep = "")
print(stock_allocation %>% head(20))
cat(rep("=", 100), "\n", sep = "")

# ==============================================================================
# STEP 15: Visualizations
# ==============================================================================

library(ggplot2)

# 1. Weight comparison
p1 <- results_summary %>%
  select(Sector, Market_Weight_Pct, Optimal_Weight_Pct) %>%
  pivot_longer(-Sector, names_to = "Type", values_to = "Weight") %>%
  mutate(Type = recode(Type, 
                       Market_Weight_Pct = "Market Weight",
                       Optimal_Weight_Pct = "Optimal Weight")) %>%
  ggplot(aes(x = Sector, y = Weight, fill = Type)) +
  geom_bar(stat = "identity", position = "dodge", width = 0.7) +
  geom_text(aes(label = paste0(round(Weight, 1), "%")), 
            position = position_dodge(width = 0.7), 
            vjust = -0.5, size = 3) +
  theme_minimal() +
  labs(title = "Market vs. Optimal Portfolio Weights",
       subtitle = "Black-Litterman Model - Middle East Investment Strategy",
       y = "Weight (%)", x = "", fill = "") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        legend.position = "top")

print(p1)

# 2. Returns comparison
p2 <- results_summary %>%
  select(Sector, Historical_Return_Pct, Equilibrium_Return_Pct, Posterior_Return_Pct) %>%
  pivot_longer(-Sector, names_to = "Type", values_to = "Return") %>%
  mutate(Type = recode(Type,
                       Historical_Return_Pct = "Historical",
                       Equilibrium_Return_Pct = "Equilibrium",
                       Posterior_Return_Pct = "BL Posterior")) %>%
  ggplot(aes(x = Sector, y = Return, fill = Type)) +
  geom_bar(stat = "identity", position = "dodge", width = 0.7) +
  theme_minimal() +
  labs(title = "Expected Returns Comparison",
       subtitle = "Historical vs. Equilibrium vs. Black-Litterman Posterior",
       y = "Annualized Return (%)", x = "", fill = "") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        legend.position = "top")

print(p2)

# 3. Risk-Return scatter
p3 <- results_summary %>%
  ggplot(aes(x = Volatility_Pct, y = Posterior_Return_Pct, 
             size = Optimal_Weight_Pct, color = Sector, label = Sector)) +
  geom_point(alpha = 0.7) +
  geom_text(vjust = -1, size = 3) +
  theme_minimal() +
  labs(title = "Risk-Return Profile with Optimal Weights",
       subtitle = "Bubble size represents optimal portfolio weight",
       x = "Volatility (% p.a.)", 
       y = "Expected Return (% p.a.)",
       size = "Weight (%)",
       color = "Sector") +
  theme(legend.position = "right")

print(p3)

# 4. Correlation matrix
corrplot(cor(returns_xts), method = "color", type = "upper",
         addCoef.col = "black", number.cex = 0.8,
         tl.col = "black", tl.srt = 45,
         title = "Sector Correlation Matrix",
         mar = c(0,0,2,0))

# 5. Stock allocation by sector
p5 <- stock_allocation %>%
  ggplot(aes(x = reorder(Sector_Group, -Stock_Weight_Pct), 
             y = Stock_Weight_Pct, 
             fill = Cntry_Terrtry_Fl_Name)) +
  geom_bar(stat = "identity") +
  theme_minimal() +
  labs(title = "Stock Allocation by Sector and Country",
       x = "Sector", y = "Portfolio Weight (%)", fill = "Country") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        legend.position = "top")

print(p5)

# ==============================================================================
# STEP 16: Export Results
# ==============================================================================

export_list <- list(
  "1_Summary" = results_summary,
  "2_Stock_Allocation" = stock_allocation,
  "3_Sector_Returns" = as.data.frame(returns_xts) %>% 
    rownames_to_column("Date") %>%
    mutate(Date = as.Date(Date)),
  "4_Company_Details" = sector_mapping_final %>%
    select(Ticker, Name, Sector_Group, Cntry_Terrtry_Fl_Name,
           GICS_SubInd_Name, Market_Cap_Num),
  "5_Correlation" = round(cor(returns_xts), 3),
  "6_Views" = data.frame(View = view_names, Expected_Return_Pct = Q * 100),
  "7_Country_Allocation" = stock_allocation %>%
    group_by(Cntry_Terrtry_Fl_Name) %>%
    summarise(Total_Weight_Pct = sum(Stock_Weight_Pct),
              Num_Companies = n(),
              .groups = 'drop') %>%
    arrange(desc(Total_Weight_Pct))
)

# Also save as CSV for easy viewing
write.csv(results_summary, "sector_allocation_summary.csv", row.names = FALSE)
write.csv(stock_allocation, "stock_allocation_detailed.csv", row.names = FALSE)

cat("\n", rep("=", 100), "\n", sep = "")
cat("RESULTS EXPORTED\n")
cat(rep("=", 100), "\n", sep = "")
cat("✓ Excel file: BlackLitterman_Middle_East_Results.xlsx\n")
cat("✓ CSV files: sector_allocation_summary.csv, stock_allocation_detailed.csv\n")
cat(rep("=", 100), "\n", sep = "")

# ==============================================================================
# STEP 17: Portfolio Performance Metrics
# ==============================================================================

portfolio_return <- sum(optimal_weights * posterior_mu)
portfolio_sd <- sqrt(as.numeric(t(optimal_weights) %*% posterior_Sigma %*% optimal_weights))
sharpe <- portfolio_return / portfolio_sd

# Calculate diversification ratio
weighted_vol <- sum(optimal_weights * sqrt(diag(Sigma)))
diversification_ratio <- weighted_vol / portfolio_sd

cat("\n", rep("=", 100), "\n", sep = "")
cat("PORTFOLIO PERFORMANCE METRICS\n")
cat(rep("=", 100), "\n", sep = "")
cat(sprintf("Expected Annual Return:       %.2f%%\n", portfolio_return * 100))
cat(sprintf("Expected Annual Volatility:   %.2f%%\n", portfolio_sd * 100))
cat(sprintf("Expected Sharpe Ratio:        %.2f\n", sharpe))
cat(sprintf("Diversification Ratio:        %.2f\n", diversification_ratio))
cat(sprintf("\nNumber of Stocks:             %d\n", nrow(stock_allocation)))
cat(sprintf("Target Sectors:               %s\n", paste(target_sectors, collapse = ", ")))
cat(sprintf("Countries:                    %s\n", 
            paste(unique(stock_allocation$Cntry_Terrtry_Fl_Name), collapse = ", ")))
cat(rep("=", 100), "\n", sep = "")


