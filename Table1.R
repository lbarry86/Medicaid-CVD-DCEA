#------------------------------------------------------------------------------#
#---------------------------Created by Luke Barry------------------------------#
#-----------------------------Date:10/17/2023----------------------------------#
#---------------------Purpose: MSM CEA Medicaid-CVD----------------------------#
#--------------------------------Table 1---------------------------------------#
#------------------------------------------------------------------------------#

#### Table 1 ----
{
  #---------------------------------Setting up data------------------------
  variable_names <- c("smoking", "stroke_dx", "heart_attack_dx", "chd_dx",
                      "cong_heart_fail_dx", "angina_dx", "meds_hyperten", 
                      "diabetes_dx", "diabetes", "history_cvd")
  
  # remove those older and younger
  nhanes_gtsummary <- nhanes_orig %>%
    filter(# no missing weights
      !is.na(WTSAF)
      # non-elderly adults
      , age_years <= .stop_age 
      , age_years >= .start_age) %>%
    mutate(insurance = case_when(health_ins_private == "1"  ~ 1,          
                                 medicare           == "1"  ~ 2,
                                 medicaid           == "1"  ~ 3,
                                 health_ins_cover   == "1"  ~ 4,
                                 health_ins_cover   == "0"  ~ 0,
                                 TRUE                       ~ NA),
           medicaid_exp   = ifelse(insurance  == 0 & ratio_poverty_inc < .fpl & age_years < 65, 1, 0),
           # create single insurance variable to correspond to MEPS cost and utility calculations and for cleaner distinction of treatment scenario
           education = case_when((edu_lev_5cat == "1" | edu_lev_5cat == "2")  ~ 0
                                 , edu_lev_5cat == "3"  ~ 1
                                 , edu_lev_5cat == "4"  ~ 2
                                 , edu_lev_5cat == "5"  ~ 3
                                 , TRUE                 ~ NA),
           # create single race variable to correspond to MEPS cost and utility calculations and for cleaner distinction of treatment scenario
           race = case_when(race_5cat == "1"    ~ 0                    
                            , race_5cat == "2"  ~ 1
                            , race_5cat == "3"  ~ 2
                            , race_5cat == "4"  ~ 3
                            , race_5cat == "5"  ~ 4
                            , TRUE              ~ NA),   
           # create "current smoker" (smoking) variable to correspond to MEPS cost and utility calculations and ASCVD calculation
           smoking = case_when(smoke_3cat         == "1"  ~ 1,                  
                               smoke_3cat         == "2"  ~ 0,
                               smoke_3cat         == "0"  ~ 0,
                               TRUE                       ~ NA),
           # create a family income variable to correspond to MEPS cost and utility calculations and according to the poverty line [see, https://meps.ahrq.gov/survey_comp/hc_technical_notes.shtml]
           fam_income = case_when(ratio_poverty_inc <= 1     ~ 0                          
                                  , ratio_poverty_inc <= 1.25  ~ 1
                                  , ratio_poverty_inc <= 2     ~ 2
                                  , ratio_poverty_inc <= 4     ~ 3
                                  , ratio_poverty_inc  > 4     ~ 4
                                  , TRUE         ~ NA),
           # same for bmi categories as family income
           bmi_cat = case_when(BMXBMI < 18.5                 ~ 0
                               , (BMXBMI >= 18.5 & BMXBMI < 25) ~ 1
                               , (BMXBMI >= 25 & BMXBMI < 30)   ~ 2
                               , BMXBMI >= 30                ~ 3
                               , TRUE              ~ NA),
           diabetes = ifelse((diabetes_dx == 1 | (hba1c >= 6.5 & diabetes_dx == 0)), 1, 0),
           wave_year = case_when(  wave_year == "G_11" ~ 0
                                   , wave_year == "H_13" ~ 1
                                   , wave_year == "I_15" ~ 2
                                   , wave_year == "J_17" ~ 3
                                   , TRUE ~ NA),
           history_cvd = ifelse((heart_attack_dx == 1 | stroke_dx == 1 | chd_dx == 1 | cong_heart_fail_dx == 1 | angina_dx == 1),
                                1, 0)
    ) %>%
    set_variable_labels(age_years = "Age in years"
                        , ratio_poverty_inc = "Ratio of family income to FPL"
                        , hdl_chol_mgdl = "HDL-Cholesterol (mg/dL)"
                        , total_chol_mgdl = "Total Cholesterol (mg/dL)"
                        , BMXBMI = "Body Mass Index (kg/m^2)"
                        , bp_systolic = "Systolic blood pressure (mm Hg)"
                        , hba1c = "Glycohemoglobin (%); HbA1c"
                        , heart_attack_dx = "Myocardial infarction"
                        , stroke_dx = "Stroke"
                        , cong_heart_fail_dx = "Congestive heart failure"
                        , chd_dx = "Coronary heart disease"
                        , angina_dx = "Angina pectoris"
                        , diabetes_dx = "Diabetes (doctor diagnosed)"
                        , diabetes = "Diabetes (doctor diagnosed OR HbA1c >= 6.5%)"
                        , meds_hyperten = "Currently taking hypertension medication"
                        , male = "Gender"
                        , insurance = "Insurance Status"
                        , education = "Highest education"
                        , race = "Race/Ethnicity"
                        , smoking = "Current smoker"
                        , medicaid_exp = "Medicaid eligibility under expansion"
                        , fam_income = "Family income"
                        , bmi_cat = "Body Mass Index (kg/m^2) category"
                        , history_cvd = "History of CVD*") %>%
    set_value_labels(
      insurance = c("Uninsured" = 0, "Private" = 1, "Medicare" = 2, "Medicaid" = 3, "Other plan" = 4),
      fam_income = c("Poor" = 0, "Near poor" = 1, "Low" = 2, "Medium" = 3, "High" = 4),
      bmi_cat = c("Underweight" = 0, "Normal weight" = 1, "Overweight" = 2, "Obese" = 3),
      education = c("No degree" = 0, "GED/HS" = 1, "Associate/Bachelor" = 2, "Master/Doctorate" = 3),
      race = c("Hispanic" = 0, "White" = 1, "Black" = 2, "Asian" = 3, "Other race" = 4),
      male = c("Male" = 1, "Female" = 0),
      medicaid_exp = c("Eligible" = 1, "Ineligible" = 0),
      wave_year = c("2011/12" = 0, "2013/14" = 1, "2015/16" = 2, "2017/18" = 3)
    ) %>%
    mutate_at(vars(all_of(variable_names)), ~ set_value_labels(., c("No" = 0, "Yes" = 1))) %>%
    select(-WTSAF
           , -WTINT2YR
           , -WTINT
           , -WTMEC
           , -WTMEC2YR
           , -WTSAF2YR
           , -race_5cat
           , -edu_lev_5cat
           , -ratio_poverty_inc
           , -health_ins_cover
           , -health_ins_private
           , -medicare
           , -medicaid
           , -smoke_3cat
           , -BMXBMI
           , -hba1c
           , -diabetes_dx) 
  
  #---------------------------------Creating functions------------------------
  
  # Define a function to check if a variable is categorical
  is_categorical_var <- function(var){
    # Check if the variable is a factor or has fewer than 10 unique non-missing values
    return(is.factor(var) || (length(unique(var[!is.na(var)])) <= 10))
  }
  
  # Define a function to check if a variable is binary
  is_binary_var <- function(var){
    # Check if the variable is a factor or has exactly 2 unique non-missing values
    return(is.factor(var) || (length(unique(var[!is.na(var)])) == 2))
  }
  
  # Define a function to check if a variable is continuous.
  is_continous_var <- function(var){
    # Check if the variable is labelled (assuming the `is.labelled` function is used) and convert it to a factor if so.
    if (is.labelled(var)) {
      var = as.factor(var)
    }
    # Check if the number of distinct non-missing values is not less than or equal to 2 and if the variable is not a factor.
    return(!(n_distinct(var[!is.na(var)]) <= 2 | is.factor(var)))
  }
  
  # Define a function to replace missing values with a specified value and label them.
  to_label_missing <- function(var, value = 999, label_missing = "N missing") {
    val_label(var, value) <- label_missing
    return(var)
  }
  
  # Define a function to replace missing values with a specified value.
  missing_to_value <- function(variable, value = 0) {
    newx <- replace(variable, which(variable %in% NA), value)
  }
  
  # Define a function to add missing value categories based on variable type.
  add_missing_categories <- function(.data) {
    .data %>% 
      # Replace missing values with 999 and label them for categorical variables.
      mutate_if(is_categorical_var, missing_to_value, value = 999) %>% 
      mutate_if(is_categorical_var, to_label_missing) %>% 
      # Replace missing values with 999 and label them for binary variables.
      mutate_if(is_binary_var, missing_to_value, value = 999) %>% 
      mutate_if(is_binary_var, to_label_missing)
  }
  
  
  #---------------------------------Creating table 1------------------------------
  
  #Without missing in gtsummary----
  total_summary_nomiss <- nhanes_gtsummary %>% 
    select(-SEQN) %>%
    modify_if(is.labelled, to_factor) %>% 
    tbl_summary(missing = "no", 
                statistic = list(
                  all_continuous() ~ c("{mean} ({sd})"),
                  all_categorical() ~ "{n} ({p}%)")) %>%
    modify_header(label = "**Variable**") %>%
    add_stat_label(location = "row") %>%
    modify_spanning_header(starts_with("stat_") ~ "Table 1") %>% 
    bold_labels()
  
  # Create separate columns for each value of "wave_year"
  # Create the summary table
  wave_summary_nomiss <- nhanes_gtsummary %>%
    select(-SEQN) %>%
    modify_if(is.labelled, to_factor) %>%
    tbl_summary(
      by = wave_year,  # Create separate columns for each value of "wave_year"
      missing = "no",
      statistic = list(
        all_continuous() ~ c("{mean} ({sd})"),
        all_categorical() ~ "{n} ({p}%)"
      )
    ) %>%
    modify_header(label = "**Variable**") %>%
    add_stat_label(location = "row") %>%
    modify_spanning_header(starts_with("stat_") ~ "Table 1") %>% 
    bold_labels() %>%
    # Add a final column that combines all columns
    add_overall() %>%
    modify_footnote(everything() ~ NA) # delete footnote
  
  print(wave_summary_nomiss)
  
  #With missing in gtsummary----
  total_summary_miss <- nhanes_gtsummary %>% 
    select(-SEQN) %>%
    add_missing_categories() %>% 
    modify_if(is.labelled, to_factor) %>% 
    tbl_summary(missing = "no", 
                type = list(
                  all_continuous() ~ "continuous2"),
                statistic = list(
                  all_continuous() ~ c("{mean} ({sd})",
                                       "{N_miss} ({p_miss}%)"),
                  all_categorical() ~ "{n} ({p}%)")) %>%
    modify_header(label = "**Variable**") %>%
    add_stat_label(location = "row") %>%
    modify_spanning_header(starts_with("stat_") ~ "Table 1") %>% 
    bold_labels()
  
  # to get table across waves working, need to separate wave_year from data to reuse
  nhanes_wave_year <- nhanes_gtsummary %>%
    select(SEQN,
           wave_year)
  
  # Update the table to include a column for each "wave_year"
  wave_summary_miss <- nhanes_gtsummary %>% 
    select(-wave_year) %>%
    add_missing_categories() %>% 
    modify_if(is.labelled, to_factor) %>% 
    left_join(nhanes_wave_year, by = join_by(SEQN)) %>%
    select(-SEQN) %>%
    tbl_summary(
      missing = "no", 
      by = wave_year,  # Split the table by "wave_year" variable
      type = list(
        all_continuous() ~ "continuous2"),
      statistic = list(
        all_continuous() ~ c("{mean} ({sd})",
                             "{N_miss} ({p_miss}%)"),
        all_categorical() ~ "{n} ({p}%)")
    ) %>%
    modify_header(label = "**Variable**") %>%
    add_stat_label(location = "row") %>%
    modify_spanning_header(starts_with("stat_") ~ "Table 1") %>% 
    bold_labels()
  
  print(wave_summary_miss)
  
  
  #---------------------------------Saving the table in word/pptx------------------------------
  
  wave_summary_miss %>% as_flex_table() %>% 
    flextable::save_as_docx(path=here(table_domain, "table1_miss.docx"))
  
  wave_summary_miss %>% as_flex_table() %>% 
    flextable::save_as_pptx(path=here(table_domain, "table1_miss.pptx"))
  
  wave_summary_nomiss %>% as_flex_table() %>% 
    flextable::save_as_docx(path=here(table_domain, "table1_nomiss.docx"))
  
  wave_summary_nomiss %>% as_flex_table() %>% 
    flextable::save_as_pptx(path=here(table_domain, "table1_nomiss.pptx"))
}

rm(nhanes_gtsummary
   , nhanes_wave_year)
