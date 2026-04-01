library(readxl)
library(dplyr)
library(purrr)
library(stringr)
library(janitor)
library(tibble)
library(tidyr)
library(haven)    
library(ranger)
library(Metrics)

### Part 1: Data Cleaning ###
### Target Variables: Annual change in log carbon intensity ###
### Calculating Scope 1 Carbon Emissions ###
dir_path <- "D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/target_variable"

# List all Emission Inventories Excel files in the folder
files <- list.files(
  dir_path,
  pattern = "Emission_Inventories.*\\.xlsx$",
  full.names = TRUE,
  ignore.case = TRUE
)
cat("Files found:", length(files), "\n")

# Extract province name from sheet name by removing trailing 4-digit year
prov_from_sheet <- function(x) {
  x <- gsub("\\s+", "", x)
  sub("\\d{4}$", "", x)
}

# For each file, iterate over province sheets and extract Scope 1 Total Emissions
extract_scope1 <- function(file) {
  year <- as.integer(str_extract(basename(file), "\\d{4}"))
  sheets <- excel_sheets(file)
  
  # Skip the NOTE sheet (metadata only)
  sheets <- sheets[sheets != "NOTE"]
  
  map_dfr(sheets, function(sh) {
    tryCatch({
      df <- read_excel(file, sheet = sh, col_names = TRUE)
      df <- clean_names(df)
      
      # Ensure first column is named emission_inventory
      if (!"emission_inventory" %in% names(df)) names(df)[1] <- "emission_inventory"
      
      # Find the TotalEmissions row (strip spaces and lowercase for robust matching)
      row_idx <- which(
        tolower(gsub("\\s+", "", as.character(df$emission_inventory))) == "totalemissions"
      )
      
      # Find the Scope_1_Total column; fall back to regex if exact match fails
      cand <- which(tolower(names(df)) == "scope_1_total")
      if (length(cand) == 0)
        cand <- grep("scope.*1.*total", tolower(names(df)), perl = TRUE)
      
      # Skip sheet if either the row or column cannot be located
      if (length(cand) == 0 || length(row_idx) == 0) return(NULL)
      
      tibble(
        province     = prov_from_sheet(sh),
        year         = year,
        scope1_total = suppressWarnings(as.numeric(df[[cand[1]]][row_idx[1]]))
      )
    }, error = function(e) {
      message("Error in sheet ", sh, ": ", e$message)
      NULL
    })
  })
}

# Apply extraction across all files and combine into one data frame
final_scope1 <- map_dfr(files, extract_scope1) %>%
  filter(!is.na(scope1_total)) %>%
  arrange(province, year)

# Save output to CSV
out_file <- file.path(dir_path, "Scope1_Total_Emissions_2008_2022.csv")
write.csv(final_scope1, out_file, row.names = FALSE)

# Preview results
print(head(final_scope1, 10))

# Check number of provinces per year (should be ~30 each year)
final_scope1 %>%
  count(year) %>%
  arrange(year) %>%
  print(n = 50)

### Carbon Intensity Calculation ###
emissions <- final_scope1 

# Load Provincial GDP
gdp_path <- "D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/economic_structure/provincial_gdp_2008_2021.xlsx"

gdp_wide <- read_excel(gdp_path, skip = 2) 

# Reshape from wide to long (province, year, gdp)
gdp_long <- gdp_wide %>%
  rename(province = 1) %>%                   
  filter(!is.na(province)) %>%               
  pivot_longer(
    cols      = -province,
    names_to  = "year",
    values_to = "gdp_100m_yuan"              # Unit: 100 million yuan
  ) %>%
  mutate(year = as.integer(year))

# Harmonize province names between the two datasets
# Check for mismatches before joining
emissions_provs <- sort(unique(emissions$province))
gdp_provs       <- sort(unique(gdp_long$province))

cat("Provinces only in emissions:\n")
print(setdiff(emissions_provs, gdp_provs))
cat("Provinces only in GDP:\n")
print(setdiff(gdp_provs, emissions_provs))

# Fix the known mismatch
emissions <- emissions %>%
  mutate(province = recode(province, "InnerMongolia" = "Inner Mongolia"))

# Merge and calculate carbon intensity
carbon_intensity <- emissions %>%
  left_join(gdp_long, by = c("province", "year")) %>%
  filter(!is.na(gdp_100m_yuan)) %>%
  mutate(
    # Carbon intensity: Mt CO2 per 100 million yuan GDP
    ci       = scope1_total / gdp_100m_yuan,
    
    # Log carbon intensity
    ln_ci    = log(ci),
    
    # Annual change in log carbon intensity (target variable)
    d_ln_ci  = ln_ci - lag(ln_ci, 1)
  ) %>%
  
  # lag() must be applied within each province
  group_by(province) %>%
  arrange(year, .by_group = TRUE) %>%
  mutate(d_ln_ci = ln_ci - lag(ln_ci, 1)) %>%
  ungroup()

print(head(carbon_intensity, 10))

# Check coverage
cat("\nRows:", nrow(carbon_intensity), "\n")
carbon_intensity %>% count(year) %>% arrange(year) %>% print(n = 20)

# Save
econ_path <- "D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/target_variable"
out_ci <- file.path(econ_path, "carbon_intensity_2008_2021.csv")
write.csv(carbon_intensity, out_ci, row.names = FALSE)
cat("Saved to:", out_ci, "\n")

### Control Variables ### 
### GDP per Capita Calculation ###
econ_path <- "D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/economic_structure"

# Load GDP data
gdp_wide <- read_excel(
  file.path(econ_path, "provincial_gdp_2008_2021.xlsx"),
  skip = 2  
)

gdp_long <- gdp_wide %>%
  rename(province = 1) %>%
  filter(!is.na(province), province != "Province") %>%  
  pivot_longer(
    cols      = -province,
    names_to  = "year",
    values_to = "gdp_100m_yuan"
  ) %>%
  mutate(year = as.integer(year))

# Load Population data (unit: 10,000 persons) 
province_order <- c(
  "Beijing", "Tianjin", "Hebei", "Shanxi", "Inner Mongolia",
  "Liaoning", "Jilin", "Heilongjiang",
  "Shanghai", "Jiangsu", "Zhejiang", "Anhui", "Fujian", "Jiangxi", "Shandong",
  "Henan", "Hubei", "Hunan", "Guangdong", "Guangxi", "Hainan",
  "Chongqing", "Sichuan", "Guizhou", "Yunnan", "Tibet",
  "Shaanxi", "Gansu", "Qinghai", "Ningxia", "Xinjiang"
)

pop_wide <- read_xls(
  file.path(econ_path, "provincial_population_2008_2021.xls"),
  skip = 0   # Keep row 0 as header (Region, 2021, 2020, ..., 2008)
)

pop_long <- pop_wide %>%
  select(-1) %>%                            # Drop the "Region" placeholder column
  mutate(province = province_order) %>%     # Assign province names by row position
  pivot_longer(
    cols      = -province,
    names_to  = "year",
    values_to = "population_10k"
  ) %>%
  mutate(year = as.integer(year))

# Calculate GDP per capita and log GDP per capita
gdp_percap <- gdp_long %>%
  left_join(pop_long, by = c("province", "year")) %>%
  filter(!is.na(population_10k), province != "Tibet") %>% 
  mutate(
    # GDP per capita in yuan: (100m yuan × 1e8) / (10k persons × 1e4)
    gdp_percap    = gdp_100m_yuan * 10000 / population_10k,
    ln_gdp_percap = log(gdp_percap)
  ) %>%
  arrange(province, year)

print(head(gdp_percap, 10))
cat("\nRows:", nrow(gdp_percap), "\n")  

out_gdp <- file.path(econ_path, "gdp_percap_2008_2021.csv")
write.csv(gdp_percap, out_gdp, row.names = FALSE)
cat("Saved to:", out_gdp, "\n")


### Part 2: Merge All Data and Run Models ###
# Define paths
econ_path   <- "D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/economic_structure"
energy_path <- "D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/energy_structure"
policy_path <- "D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/policy"
target_path <- "D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/target_variable"

# Load each dataset

# Target variable
ci <- read.csv(file.path(target_path, "carbon_intensity_2008_2021.csv")) %>%
  select(province, year, ln_ci, d_ln_ci)

# GDP per capita
gdp_pc <- read.csv(file.path(econ_path, "gdp_percap_2008_2021.csv")) %>%
  select(province, year, ln_gdp_percap)

# Urban population rate
urban <- read_excel(file.path(econ_path, "urban_population_rate.xlsx"), skip = 2) %>%
  rename(province = 1) %>%
  filter(!is.na(province)) %>%
  pivot_longer(-province, names_to = "year", values_to = "urban_rate") %>%
  mutate(year = as.integer(year))

# Secondary industry share
sh_sec <- read_excel(
  file.path(econ_path, "industry_shares_2008_2021.xlsx"),
  sheet = "Secondary Industry Share", skip = 2
) %>%
  rename(province = 1) %>%
  filter(!is.na(province)) %>%
  pivot_longer(-province, names_to = "year", values_to = "sh_secondary") %>%
  mutate(year = as.integer(year))

# Tertiary industry share
sh_ter <- read_excel(
  file.path(econ_path, "industry_shares_2008_2021.xlsx"),
  sheet = "Tertiary Industry Share", skip = 2
) %>%
  rename(province = 1) %>%
  filter(!is.na(province)) %>%
  pivot_longer(-province, names_to = "year", values_to = "sh_tertiary") %>%
  mutate(year = as.integer(year))

# Coal share
coal <- read.csv(file.path(energy_path, "coal_share.csv")) %>%
  select(province, year, coal_share) %>%
  mutate(province = recode(province, "Inner mongolia" = "Inner Mongolia"))

# Policy intensity — filter to 2008-2021, drop Xizang
policy <- read_dta(file.path(policy_path, "Provincial_Policy_Intensity.dta")) %>%
  filter(year >= 2008, year <= 2021, province != "Xizang") %>%
  mutate(year = as.integer(year))

# Merge all into one panel
panel <- ci %>%
  left_join(gdp_pc,  by = c("province", "year")) %>%
  left_join(urban,   by = c("province", "year")) %>%
  left_join(sh_sec,  by = c("province", "year")) %>%
  left_join(sh_ter,  by = c("province", "year")) %>%
  left_join(coal,    by = c("province", "year")) %>%
  left_join(policy,  by = c("province", "year"))

cat("Panel rows after merge:", nrow(panel), "\n")
print(head(panel, 5))

# Construct lagged features (t-1 predicts t)
features <- c("ln_gdp_percap", "urban_rate", "sh_secondary",
              "sh_tertiary", "coal_share", "PI_provincial", "ln_ci")

panel_lagged <- panel %>%
  group_by(province) %>%
  arrange(year, .by_group = TRUE) %>%
  mutate(across(all_of(features), ~ lag(., 1), .names = "{.col}_lag")) %>%
  ungroup() %>%
  filter(!is.na(d_ln_ci), !is.na(ln_gdp_percap_lag))  

cat("Rows after lagging (model-ready):", nrow(panel_lagged), "\n")

# Train / test split
lag_features <- paste0(features, "_lag")

train <- panel_lagged %>% filter(year <= 2017)
test  <- panel_lagged %>% filter(year >= 2018)

cat("Train rows:", nrow(train), " | Test rows:", nrow(test), "\n")

X_train <- train[, lag_features]
y_train <- train$d_ln_ci
X_test  <- test[, lag_features]
y_test  <- test$d_ln_ci

### Model 1: OLS ###
train_df <- cbind(d_ln_ci = y_train, X_train)
ols_formula <- as.formula(paste("d_ln_ci ~", paste(lag_features, collapse = " + ")))
ols_model <- lm(ols_formula, data = train_df)

ols_pred <- predict(ols_model, newdata = X_test)

cat("\n--- OLS Results ---\n")
cat("MAE: ",  round(mae(y_test, ols_pred), 4), "\n")
cat("RMSE:", round(rmse(y_test, ols_pred), 4), "\n")
cat("R²:  ",  round(cor(y_test, ols_pred)^2, 4), "\n")

# Show OLS predictions
results_ols <- test %>%
  select(province, year, d_ln_ci) %>%
  mutate(ols_pred = ols_pred)
print(head(results_ols, 10))

### Model 2: Random Forest ###
train_rf <- cbind(d_ln_ci = y_train, X_train)

rf_model <- ranger(
  formula    = d_ln_ci ~ .,
  data       = train_rf,
  num.trees  = 500,
  importance = "permutation",
  seed       = 42
)

rf_pred <- predict(rf_model, data = X_test)$predictions

cat("\n--- Random Forest Results ---\n")
cat("MAE: ",  round(mae(y_test, rf_pred), 4), "\n")
cat("RMSE:", round(rmse(y_test, rf_pred), 4), "\n")
cat("R²:  ",  round(cor(y_test, rf_pred)^2, 4), "\n")

# Save predictions
results <- test %>%
  select(province, year, d_ln_ci) %>%
  mutate(
    ols_pred = ols_pred,
    rf_pred  = rf_pred
  )
print(head(results, 10))