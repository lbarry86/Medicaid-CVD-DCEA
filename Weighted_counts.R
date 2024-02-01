#------------------------------------------------------------------------------#
#---------------------------Created by Luke Barry------------------------------#
#-----------------------------Date:10/17/2023----------------------------------#
#---------------------Purpose: MSM CEA Medicaid-CVD----------------------------#
#-----------------------------Weighted counts----------------------------------#
#------------------------------------------------------------------------------#

#### Weighted counts ----

.weight <- "WTSAF2YR"
.psu    <- "SDMVPSU"
.strata <- "SDMVSTRA"

#For more information about how to use the srvyr package, please
#check out this website: 
#https://cran.r-project.org/web/packages/srvyr/vignettes/srvyr-vs-survey.html
#--------------------Survey option--------------------------

options(survey.lonely.psu = "adjust") #very important for the survey packages
#Without this option, the numbers would be incorrect. 
#Please do not forget THIS!!!!

#--------------------Creating functions--------------------------
#--Create function to display point estimate and confidence interval
ci_fct <- function(pe, low, high, multiplier=1, digits=1, scale="%", count=FALSE){
  if(count==F){
    res <- paste(multiplier*formattable::digits(pe,digits),scale, " " ,"(", 
                 multiplier*formattable::digits(low,digits),",", 
                 multiplier*formattable::digits(high,digits), ")", sep="")
  } else {
    res <- paste(formattable::comma(pe, digits=digits), " " ,"(", 
                 formattable::comma(low, digits=digits),", ", 
                 formattable::comma(high, digits=digits), ")", sep="")
  }
  return(res)
}

#--Test the function
ci_fct(pe=2, low=1, high=3, multiplier=100, scale="%")
ci_fct(pe=200000, low=200000, high=200000, multiplier=NULL, digits=0, scale=NULL, count=T)

#--Create function to calculate the count

count_function <- function(.data, variable){
  .data %>% 
    count(!!sym(variable)) %>% 
    rename(value = variable,
           unweighted_count=n) %>% 
    mutate(variables=variable) %>% 
    relocate(variables, value, unweighted_count)
}

#to create a binary variable
d <- function(expr){
  as.numeric(expr)
}

#--------------------Obtaining weighted and unweighted counts-------------------

##data prep
nhanes_wt <- nhanes_count %>%
  filter(!is.na(.data[[.weight]])
         , age <= .stop_age
         , age >= .start_age
         , wave_year == "H_13") %>%
  # create a family income variable to correspond to MEPS cost and utility calculations and according to the poverty line [see, https://meps.ahrq.gov/survey_comp/hc_technical_notes.shtml]
  mutate(age_cat = cut(age,                                            
                       breaks = c(0, 24, 44, 64, Inf), 
                       labels = c('<25', '25-44', '45-64', '65+'))
         , medicaid_exp = ifelse(insurance  == "uninsured" & fpl < .fpl & age < 65, 1, 0)
         , fam_income = case_when(fpl <= 1     ~ "poor",                          
                                fpl <= 1.25  ~ "near_poor",
                                fpl <= 2     ~ "low",
                                fpl <= 4     ~ "medium",
                                fpl  > 4     ~ "high",
                                TRUE         ~ NA)) %>%
  #na.omit() %>%
  filter(medicaid_exp == 1 | .ATT != 1) %>%
  select(  .data[[.weight]]
         , .data[[.psu]]
         , .data[[.strata]]
         , fam_income
         , education
         , race
         , age_cat
         , female) %>%
  mutate(  total = 1
         , inc_poor               = d(fam_income == "poor")
         , inc_near_poor          = d(fam_income == "near_poor")
         , inc_lower              = d(fam_income == "low")
         , inc_medium             = d(fam_income == "medium")
         , inc_high               = d(fam_income == "high")
         , age_24minus            = d(age_cat == "<25")
         , age_25to44             = d(age_cat == "25-44")
         , age_45to64             = d(age_cat == "45-64")
         , race_white             = d(race == "white")
         , race_black             = d(race == "black")
         , race_hispanic          = d(race == "hispanic")
         , race_asian             = d(race == "asian")
         , race_other             = d(race == "other_race")
         , edu_no_degree          = d(education == "no_degree")
         , edu_ged_hs             = d(education == "ged_hs")
         , edu_associate_bachelor = d(education == "associate_bachelor")
         , edu_master_doctorate   = d(education == "master_doctorate")) %>%
select(-age_cat
       , -education
       , -race
       , -female
       , -fam_income)


##Estimating weighted weighted_count----
weighted_count_data_ci <- nhanes_wt %>% 
  as_survey_design(weights = .data[[.weight]]
                   , strata = .data[[.strata]]
                   , ids = .data[[.psu]]
                   , nest = TRUE
                   ) %>% 
  summarize_all(survey_total, vartype = "ci", na.rm=T) %>% 
  t() %>%
  data.frame(weighted_count=.) %>%
  rownames_to_column(var = "variables")

weighted_count_data_low <- weighted_count_data_ci %>% 
  filter(str_detect(variables, '_low')) %>% 
  mutate(variables=str_replace_all(variables,"_low", "")) %>%
  rename(weighted_count_low=weighted_count)

weighted_count_data_high <- weighted_count_data_ci %>% 
  filter(str_detect(variables, '_upp')) %>% 
  mutate(variables=str_replace_all(variables,"_upp", "")) %>%
  rename(weighted_count_high=weighted_count)

weighted_count_data_pe <- weighted_count_data_ci %>% 
  filter(variables %in% c(
    "total"
    , "inc_poor"
    , "inc_near_poor"
    , "inc_lower"
    , "inc_medium"
    , "inc_high"
    , "age_24minus"
    , "age_25to44"
    , "age_45to64"
    , "race_white"
    , "race_black"
    , "race_hispanic"
    , "race_asian"
    , "race_other"
    , "edu_no_degree"
    , "edu_ged_hs"
    , "edu_associate_bachelor"
    , "edu_master_doctorate"
  )) %>% 
  rename(weighted_count_pe=weighted_count)

weighted_count_data <- weighted_count_data_pe %>% 
  full_join(weighted_count_data_low, by=c("variables")) %>% 
  full_join(weighted_count_data_high, by=c("variables")) 

table1_us_weighted_count <- weighted_count_data %>% 
  transmute(
    `Pop. characteristic` = case_when(variables== "total"                     ~ "Total"
                                      , variables== "inc_poor"                ~ "Family income - Poor"
                                      , variables== "inc_near_poor"           ~ "Family income - Near poor"
                                      , variables== "inc_lower"               ~ "Family income - Low"
                                      , variables== "inc_medium"              ~ "Family income - Medium"
                                      , variables== "inc_high"                ~ "Family income - High"
                                      , variables== "age_24minus"             ~ "Age - <25 years"
                                      , variables== "age_25to44"              ~ "Age - 25 to 44 years"
                                      , variables== "age_45to64"              ~ "Age - 45-64 years"
                                      , variables== "race_white"              ~ "Race - White"
                                      , variables== "race_black"              ~ "Race - Black"
                                      , variables== "race_hispanic"           ~ "Race - Hispanic"
                                      , variables== "race_asian"              ~ "Race - Asian"
                                      , variables== "race_other"              ~ "Race - Other"
                                      , variables== "edu_no_degree"           ~ "Education - No degree"
                                      , variables== "edu_ged_hs"              ~ "Education - GED/HS"
                                      , variables== "edu_associate_bachelor"  ~ "Education - Associate/Bachelor"
                                      , variables== "edu_master_doctorate"    ~ "Education - Master/Doctorate"
                                      , TRUE	~	NA_character_),
    `NHANES weighted_count`=ci_fct(pe=weighted_count_pe, low=weighted_count_low, 
                                  high=weighted_count_high, multiplier=NULL, digits=0, scale=NULL, count=T),
    `NHANES weighted_count pe`= paste(formattable::comma(weighted_count_pe, digits = 0), sep=""))



##Outputing table 1 - weighted ----
table1_us_weighted_count %>% 
  as_hux()


write_xlsx(table1_us_weighted_count, 
           here(table_domain, "table1_us_overall_weighted_count.xlsx"))

##Unweighted count----

unweighted_count_data_us <- map_dfr(c("inc_poor"
                                      , "inc_near_poor"
                                      , "inc_lower"
                                      , "inc_medium"
                                      , "inc_high"
                                      , "age_24minus"
                                      , "age_25to44"
                                      , "age_45to64"
                                      , "race_white"
                                      , "race_black"
                                      , "race_hispanic"
                                      , "race_asian"
                                      , "race_other"
                                      , "edu_no_degree"
                                      , "edu_ged_hs"
                                      , "edu_associate_bachelor"
                                      , "edu_master_doctorate") ,
                                     count_function, 
                                    .data=nhanes_wt) %>% 
  filter(value==1) %>% 
  select(-value)


table1_us_unweighted_count <- unweighted_count_data_us %>% 
  transmute(
    `Pop. characteristic`= case_when(  variables== "inc_poor"       ~ "Family income - Poor"
                                     , variables== "inc_near_poor"           ~ "Family income - Near poor"
                                     , variables== "inc_lower"               ~ "Family income - Low"
                                     , variables== "inc_medium"              ~ "Family income - Medium"
                                     , variables== "inc_high"                ~ "Family income - High"
                                     , variables== "age_24minus"             ~ "Age - <25 years"
                                     , variables== "age_25to44"              ~ "Age - 25 to 44 years"
                                     , variables== "age_45to64"              ~ "Age - 45-64 years"
                                     , variables== "race_white"              ~ "Race - White"
                                     , variables== "race_black"              ~ "Race - Black"
                                     , variables== "race_hispanic"           ~ "Race - Hispanic"
                                     , variables== "race_asian"              ~ "Race - Asian"
                                     , variables== "race_other"              ~ "Race - Other"
                                     , variables== "edu_no_degree"           ~ "Education - No degree"
                                     , variables== "edu_ged_hs"              ~ "Education - GED/HS"
                                     , variables== "edu_associate_bachelor"  ~ "Education - Associate/Bachelor"
                                     , variables== "edu_master_doctorate"    ~ "Education - Master/Doctorate"
                                     , TRUE	~	NA_character_),
    `NHANES unweighted_count`= paste(formattable::comma(unweighted_count, digits = 0), sep=""))

##Outputing table 1 - unweighted----
table1_us_unweighted_count %>% 
  as_hux()

write_xlsx(table1_us_unweighted_count, 
           here(table_domain, "table1_us_overall_unweighted_count.xlsx"))

rm(nhanes_wt)