library(readxl)
library(openxlsx)

# Import the details of the stocks
stock_details <- read_excel("~/Library/CloudStorage/GoogleDrive-pmp.zzgroup@gmail.com/My Drive/ZZ Team 24 25/2025 - New/8. Personal/Luuuuuuuuuuk/Funda Analysis/stock details.xlsx")

# Remove the first two rows
stock_details <- stock_details[-c(1,2), ]

# Import the first set of metrics
stocks_metrics_1 <- read_excel("~/Library/CloudStorage/GoogleDrive-pmp.zzgroup@gmail.com/My Drive/ZZ Team 24 25/2025 - New/8. Personal/Luuuuuuuuuuk/Funda Analysis/stocks metrics 1.xlsx")

# Remove the first two rows
stocks_metrics_1 <- stocks_metrics_1[-c(1,2), ]

# Remove the "Name" column
stocks_metrics_1$Name <- NULL

# Import the second set of metrics
stocks_metrics_2_banking_ <- read_excel("~/Library/CloudStorage/GoogleDrive-pmp.zzgroup@gmail.com/My Drive/ZZ Team 24 25/2025 - New/8. Personal/Luuuuuuuuuuk/Funda Analysis/stocks metrics 2 (banking).xlsx")

# Remove the first two rows
stocks_metrics_2_banking_ <- stocks_metrics_2_banking_[-c(1,2), ]

# Remove the "Name" column
stocks_metrics_2_banking_$Name <- NULL

# Import the third set of metrics
stocks_metrics_3_health_and_energy_ <- read_excel("~/Library/CloudStorage/GoogleDrive-pmp.zzgroup@gmail.com/My Drive/ZZ Team 24 25/2025 - New/8. Personal/Luuuuuuuuuuk/Funda Analysis/stocks metrics 3 (health and energy).xlsx")

# Remove the first two rows
stocks_metrics_3_health_and_energy_ <- stocks_metrics_3_health_and_energy_[-c(1,2), ]

# Remove the "Name" column
stocks_metrics_3_health_and_energy_$Name <- NULL

# Add the "GICS SubInd Name" column to stock_details from stocks_metrics_3_health_and_energy_ by Ticker
stock_details <- merge(stock_details, 
                       stocks_metrics_3_health_and_energy_[, c("Ticker", "GICS SubInd Name")], 
                       by = "Ticker", 
                       all.x = TRUE)

# Remove the "GICS SubInd Name" column from stocks_metrics_3_health_and_energy_
stocks_metrics_3_health_and_energy_$`GICS SubInd Name` <- NULL
                                                    
# Merge all data frames by "Ticker", make sure that stock_details is the first data frame
merged_data <- Reduce(function(x, y) merge(x, y, by = "Ticker",
                                         all.x = TRUE), 
                      list(stock_details, stocks_metrics_1, 
                           stocks_metrics_2_banking_, 
                           stocks_metrics_3_health_and_energy_))

# Give the unique list of possible values in the "Cntry Terrtry FI Name", "GICS Sector" and "GICS Ind Name" columns
unique_countries <- sort(unique(merged_data$`Cntry Terrtry Fl Name`))
unique_sectors <- sort(unique(merged_data$`GICS Sector`))
unique_industries <- sort(unique(merged_data$`GICS Ind Name`))
unique_subindustries <- sort(unique(merged_data$`GICS SubInd Name`))

print(unique_countries)
print(unique_sectors)
print(unique_industries)
print(unique_subindustries)

# Give the amount of stocks for each industry and sort them on amount
industry_counts <- sort(table(merged_data$`GICS Ind Name`), decreasing = TRUE)
print(industry_counts)

# Create a dataframe with the the Energy, Financials, Health Care, and Real Estate sectors only
filtered_data <- merged_data[merged_data$`GICS Sector` %in% c("Energy", "Financials", "Health Care", "Real Estate"), ]

# Create a dataframe with the Banks, Capital Markets, Construction Materials, Consumer FInance, Diversified REITs, Electric Utilities, Electrical Equipment, Electronic Equipment, Instruments & Components, Energy Equipment & Services, Entertainment, Financial Services, Hotel & Resort REITs, Independent Power and Renwable Electricity Producers, Insurance, Interactive Media & Services, Real Estate Management & Development, Technology Hardware, Storage & Peripherals
detailed_filtered_data <- merged_data[merged_data$`GICS Ind Name` %in% c("Banks", "Capital Markets", "Construction Materials", 
                                                                        "Consumer Finance", "Diversified REITs", 
                                                                        "Electric Utilities", "Electrical Equipment", 
                                                                        "Electronic Equipment, Instruments & Components", 
                                                                        "Energy Equipment & Services", "Entertainment", 
                                                                        "Financial Services", "Hotel & Resort REITs", 
                                                                        "Independent Power and Renewable Electricity Producers", 
                                                                        "Insurance", "Interactive Media & Services", 
                                                                        "Real Estate Management & Development", "Software",
                                                                        "Hotels, Restaurants & Leisure", 
                                                                        "Technology Hardware, Storage & Peripherals", "Metals & Mining",
                                                                        "Semiconductors & Semiconductor Equipment"), ]

# Combine both filtered dataframes and remove duplicates
final_filtered_data <- unique(rbind(filtered_data, detailed_filtered_data))

# Export the final filtered data to an Excel file
write.xlsx(final_filtered_data, file = "sector_industry_filtered_stocks.xlsx", rowNames = FALSE)

# Filter on countries: Qatar, UAE, Saudi Arabia
country_filtered_data <- final_filtered_data[final_filtered_data$`Cntry Terrtry Fl Name` %in% c("QATAR", "UAE", "SAUDI ARABIA"), ]

# Export the country filtered data to an Excel file
write.xlsx(country_filtered_data, file = "country_filtered_stocks.xlsx", rowNames = FALSE)
