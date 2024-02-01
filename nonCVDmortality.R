#------------------------------------------------------------------------------#
#---------------------------Created by Luke Barry------------------------------#
#-----------------------------Date:10/10/2023----------------------------------#
#---------------------Purpose: MSM CEA Medicaid-CVD----------------------------#
#---------------------------Non-CVD mortality rates----------------------------#
#------------------------------------------------------------------------------#

#### Predict Non-CVD mortality rates ----

# Predict Non-CVD Death Rates for those age 85+ from those aged 18-84
# See "CDC_non-CVD_DeathRates_age85plus.R" for detailed description

# Read the Excel file with CDC Non-CVD death rates calculated for those aged 18-84
nonCVD_mort <- read.xlsx(here(data_domain, "AllCause-CVD_Deaths_Rates_2013-14.xlsx"), sheet = "ImportR_nCVD")
data_18_84 <- data.frame(age = nonCVD_mort$Age, female = factor(nonCVD_mort$Gender), rate = nonCVD_mort$rate, se = nonCVD_mort$se)

# Linear regression model on ages 18-84 with gender and age-gender interaction
lm_model <- lm(rate ~ age + I(age^2) + I(age^3) + I(age^4) + I(age^5) + female + age * female, data = data_18_84)

# Generate predictions and predicted standard deviations for ages 85+
age_85_plus <- 85:150
female_85_plus <- rep("Female", length(age_85_plus)) # Assuming all are females for simplicity
predicted_rate_females <- predict(lm_model, newdata = data.frame(age = age_85_plus, female = factor(female_85_plus)))
predicted_sd_females <- predict(lm_model, newdata = data.frame(age = age_85_plus, female = factor(female_85_plus)), se.fit = TRUE)$se.fit

# Simulated data for ages 85+ for females
simulated_data_85_plus_females <- data.frame(age = age_85_plus, female = factor(female_85_plus), rate = predicted_rate_females, se = predicted_sd_females)

# Generate predictions and predicted standard deviations for ages 85+ and males
female_85_plus <- rep("Male", length(age_85_plus)) # Assuming all are males for simplicity
predicted_rate_males <- predict(lm_model, newdata = data.frame(age = age_85_plus, female = factor(female_85_plus)))
predicted_sd_males <- predict(lm_model, newdata = data.frame(age = age_85_plus, female = factor(female_85_plus)), se.fit = TRUE)$se.fit

# Simulated data for ages 85+ for males
simulated_data_85_plus_males <- data.frame(age = age_85_plus, female = factor(female_85_plus), rate = predicted_rate_males, se = predicted_sd_males)

# Combine data for ages 18-100+ for males and females
simulated_data_nCVD <- bind_rows(data_18_84, simulated_data_85_plus_males, simulated_data_85_plus_females)
simulated_data_nCVD <- simulated_data_nCVD %>%
  mutate(female = ifelse(female == "Female", 1, ifelse(female == "Male", 0, NA))) 

rm(simulated_data_85_plus_females,
   simulated_data_85_plus_males,
   lm_model,
   data_18_84,
   nonCVD_mort,
   age_85_plus,
   female_85_plus,
   predicted_rate_females,
   predicted_sd_females,
   predicted_rate_males,
   predicted_sd_males)