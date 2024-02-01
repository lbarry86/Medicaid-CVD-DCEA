#------------------------------------------------------------------------------#
#---------------------------Created by Luke Barry------------------------------#
#-----------------------------Date:07/18/2023----------------------------------#
#---------------------Purpose: MSM CEA Medicaid-CVD----------------------------#
#------------------------------MicroSim Model----------------------------------#
#------------------------------------------------------------------------------#

##### Install packages ----

'
#install these packages if not already installed
install.packages("dplyr")
install.packages("magrittr")
install.packages("tidyr")
install.packages("mise")
install.packages("heemod")
install.packages("pacman")
install.packages("labelled")
install.packages("officer")
install.packages("flextable")
install.packages("abind")
install.packages("ineq")
install.packages("openxlsx")
install.packages("gridExtra")
install.packages("pptx")
install.packages("DataExplorer")
install.packages("nhanesA")
install.packages("viridis")
install.packages("survey")
install.packages("miceRanger")
install.packages("formattable")
install.packages("huxtable")
install.packages("writexl")
install.packages("srvyr")
install.packages("gtsummary")
install.packages("hesim")
install.packages("here")
install.packages("Amelia")
install.packages("VIM")
install.packages("naniar")
install.packages("patchwork")
install.packages("haven")
'

##### Clear environment ----
mise::mise()

##### Set directory ----
wd <- "/Users/lukebarry_1/Library/CloudStorage/Dropbox/Research/Projects/CVD/Simulation/docs/Submissions/HA/Code"

##### Load packages ----
library(pacman)
p_load(dplyr
       , magrittr
       , tidyr
       , mise
       , tidyverse
       , here
       , heemod
       , labelled
       , officer
       , flextable
       , abind
       , ineq
       , openxlsx
       , gridExtra
       , ggplot2
       , DataExplorer
       , nhanesA
       , colourpicker
       , viridis
       , survey
       , huxtable
       , miceRanger
       , writexl
       , srvyr
       , gtsummary
       , hesim
       , stringr
       , Amelia
       , VIM
       , naniar
       , patchwork
       , haven)

##### PARAMETERS ----
{
  #### Scenario parameters ----
  
  # Set perspective  = "healthcare" or "societal" for ATT only, always use societal perspective for non-ATT as examining distributional impacts
  .perspective      <- "societal"
  # Set .Deterministic == 1 for .Deterministic and anything else for probabilistic
  .Deterministic    <- 0                                                         
  # Set .one_way == 1 to conduct one-way sensitivity analysis (Must also set .Deterministic, .fix_rcat & .fix_slice_sample == 1)
  .one_way          <- 0 
  # Set .fix_rcat == 1 in order to have the same random draws from the markov transition matrix between treatment and control (in the absence of a treatment effect)
  .fix_rcat         <- 0     
  # Set .fix_slice_sample == 1 in order to have the same random sample to be drawn across bootstrap replications
  .fix_slice_sample <- 0 
  # Set .raw_data == 1 in order to store individual level data
  .raw_data         <- 0
  
  # Set .ATT == 1 to examine just those receiving treatment (Medicaid expansion) and treatment effects (otherwise whole NHANES sample, not just those who receive medicaid - NB option for distributional analysis)
  .ATT              <- 0 
  # Change the strata by which the ASCVD weights are calculated
  .ascvd_cat        <- c("age", "female")                     
  # List of variables for which to calculate summaries
  .grouping_vars    <- c("age_cat", "race", "education", "fam_income")
  # List of equity-relevant variables
  .equity_vars      <- c("Race/Ethnicity", "Education", "Family Income")
  # Set .stop_age (default == 64) according to the maximum starting age of individuals in the model
  # https://www.medicaid.gov/medicaid/eligibility/index.html
  .stop_age         <- 64
  # Set .start_age (default == 19) according to the minimum starting age of individuals in the model
  .start_age        <- 19
  # Set .trunc_age (default == 85) according to the age at which the model should end per person
  .trunc_age        <- 85
  # Set .fpl == 1.38 to match threshold for Medicaid expansion qualification for family income to federal poverty ratio
  .fpl              <- 1.38  
  # Set .fpl == 1.5 to reallocate non-oop costs for those receiving medicaid in the medicaid scenario to those with income >= 1.5 of the FPL
  .fpl_reallocate   <- 1.5                                                       
  
  # Set .half == 1 to examine the effect of the half-cycle correction; trapezoidal method http://dx.doi.org/10.1007/s40273-015-0337-0 
  .half             <- 1
  # Set .positive_dist == 1 to allocate non-opp costs for those receiving medicaid across those with incomes >= 1.5 fpl POSITIVELY (i.e. proportionate to income); 
  # default is to spread (not proportionate to income, for those with incomes over 150% FPL)
  .positive_dist    <- 0
  # Set .rate == 1 to calculate rate of MI, stroke etc instead of counts
  .rate             <- 1
  # Set .ir_py == 100000 to set the denominator for the incidence rate 100,000 person-years
  .ir_py            <- 100000
  # Set .excl_death == 1 to exclude deaths from calculation of person_years
  .excl_death       <- 1

  # Set .medicaid_cost == 1 to examine the effect of Medicaid on healthcare costs in the treatment scenario
  .medicaid_cost    <- 1
  # Set .gc == 1 to examine the effect government cost of administering Medicaid in the treatment scenario
  .gc               <- 1                                                          
  # Set .c_absent == 1 to include the cost of workplace absenteeism from an MI or Stroke
  .c_absent         <- 1
  # Set .c_disab == 1 to include the cost of short-term disability from an MI or Stroke
  .c_disab          <- 1
  # Set .prem_mort == 1 to include the cost of premature death in terms of lost annual earnings
  .c_prem_mort      <- 1
  
  # Set .hba1c == 1 for those who newly receive medicaid as part of the medicaid expansion to have lower hba1c and thus diabetes
  .hba1c            <- 1                                                          
  # Set .bp == 1 for those who newly receive medicaid as part of the medicaid expansion to have lower blood pressure
  .bp               <- 1    
  # Set .nCVDmort == 1 according to examine effect of Medicaid expansion on non-CVD mortality for those who receive medicaid in the treatment scenario
  .nCVDmort         <- 0                                                           
  
  # Re-run Multiple Imputation (.mi == 1); otherwise use previously saved MI datasets (MI datasets are used in both cases)
  .mi               <- 0
  # Re-run nhanesA package to create NHANES dataset
  .nhanesA          <- 0
  
  # Cannot use government and lost productivity costs when examining non-societal perspective 
  if (.perspective != "societal") {
    .gc             <- 0                                                          
    .c_absent       <- 0
    .c_disab        <- 0
    .c_prem_mort    <- 0
  }
  
  #### Sensitivity analysis parameters  ----
  
  # Define the % swing in baseline parameter values for one-way sensitivity analysis
  delta_one_way <- 0.20  # 20% change
  # Define range of epsilon (inequality aversion) values over which to perform the distributional cost-effectiveness analysis
  epsilon_values  <- seq(0, 10, by = 0.3)
  # Define range of WTP values over which to perform the distributional cost-effectiveness analysis
  wtp_values <- seq(50000, 250000, by = 10000)
    
  #### Progress Display parameters ----
  
  # Set .track_prob == 1 to track whether cycle probabilities sum to 1
  .track_prob       <- 0
  # Set .track_cycle == 1 to track simulation progress as % cycles completed
  .track_cycle      <- 0
  # Set .track_boot == 1 to track bootstrapping progress as % replications completed
  .track_boot       <- 1
  # Set .track_dsa == 1 to track deterministic sensitivity analysis progress as % variables completed
  .track_dsa        <- 1
  # Set timer == 1 to track time of code execution
  .timer            <- 1
  
  #### Model parameters ----
  
  # number of simulated individuals
  n.i          <- 1000
  # number of cycles (years) [default to 67; this allows everyone in the model to turn 85, i.e. 19 + 67 = 86]
  n.t          <- 67
  # number of bootstrapped model iterations 
  bootsize     <- 1000
  # Number of imputed datasets
  m            <- 10
  # Number of iterations per imputed dataset
  m_iter       <- 20
  # set willingness-to-pay threshold [default should be $100,000 as recommended by the US Institute for Clinical and Economic Research]
  wtp          <- 100000                                                          
  seed         <- 1252
  # model state names
  v.n          <- c("No_CVD", 
                    "MI", 
                    "Stroke", 
                    "History_CVD", 
                    "CVD_death", 
                    "Non-CVD_death")  
  # number of states (6)
  n.s          <- length(v.n)  
  # discount rate for QALYs (a.k.a. annual utility); "e" refers more broadly to "effect"
  d.e          <- 0.03           
  # discount rate for annual healthcare costs
  d.c          <- 0.03       
  # strategy names
  v.Trt        <- c("no_medicaid_expansion", 
                    "medicaid_expansion") 
  # Define the duration of each cycle (assuming it's constant)
  cycle_length <- 1                 
  # Calculate the half-cycle correction factor
  half_cycle_correction <- 0.5 * cycle_length 
  # converting 95% CI's to SE's
  ci2se   <- 3.92
  # if no variability for estimate provided, use 20% of mean
  est_se  <- 0.2
  # Inequality aversion parameter (you can adjust this value)
  epsilon <- 0.5                                                                 
  
  #### Intervention parameters ----
  
  # absolute reduction in systolic blood pressure after receiving medicaid [−3.03 mmHg; 95% CI, −5.33 mmHg to − 0.73 mmHg] from DOI: 10.1007/s11606-020-06417-6
  delta_sys_bp    <- list(mean = 3.03, lci = 0.73, uci = 5.33)  
  # absolute reduction in hba1c after receiving medicaid [−0.14 percentage points [pp]; 95% CI, −0.24 pp to −0.03 pp;] from DOI: 10.1007/s11606-020-06417-6
  delta_hba1c     <- list(mean = 0.14, lci = 0.03, uci = 0.24)  
  # option to consider rate reduction in non-CVD mortality estimated as difference in all-cause mortality rate from medicaid exp and CVD mortality rate from medicaid exp 
  # [https://doi.org/10.1016/S2468-2667(21)00252-8]
  # https://www.bmj.com/about-bmj/resources-readers/publications/statistics-square-one/5-differences-between-means-type-i-an
  delta_nCVD_mort <- list(mean = 0, se = 1)
  
  #### Transition parameters ----
  
  # Main risk parameters, as a function of age, sex, bp and hba1c, among other variables, 
  # are estimated per person using the "Probs" and "average_ascvd_risk" functions below using
  # a similar appraoch to Choi and Basu, 2017 (http://dx.doi.org/10.1016/j.amepre.2016.12.013) 
  # except ascvd weighting has been updated to use Bayes formula
  
  # "history of CVD" multiplier (see Basu, 2017)
  delta_CVD_history <- list(shape = 3.842, scale = 0.521, mean = 2) 

  #### Cost parameters ----
  
  # Annual healthcare costs, as a function of CVD status among other variables, 
  # are estimated per person using the "Costs" function below using
  # the Morey et al 2021 equation (DOI: 10.1161/CIRCOUTCOMES.120.006769)
  
  # Inflation adjustment for Medicaid Admin costs (priced in 2019 USD) using Health PCE [https://apps.bea.gov/]
  pce_hc_2019      <- 1.041
  # Inflation adjustment for Morey et al (2021) healthcare costs (priced in 2017 USD) using Health PCE [https://apps.bea.gov/]
  pce_hc_2017      <- 1.074
  # Inflation adjustment for Song et al (2015) lost productivity costs (priced in 2013 USD) using Overall PCE [https://apps.bea.gov/]
  pce_all_2013     <- 1.141 
  
  # scaling factor ($29,545,128,289 / $645,708,814,950) for the government cost of administering the Medicaid (https://www.medicaid.gov/state-overviews/scorecard/annual-medicaid-chip-expenditures/index.html)
  c_gov_admin      <- 0.046  
  # Medicaid spending per newly eligible adult enrollee (https://www.kff.org/medicaid/state-indicator/medicaid-spending-per-enrollee/)
  c_enrollee       <- 5225
  # Medicaid administration spending per newly eligible adult enrollee in 2021 (https://www.kff.org/medicaid/state-indicator/medicaid-spending-per-enrollee/)
  c_admin_enrollee_mean  <- c_enrollee * c_gov_admin * pce_hc_2019 # probabilistic component, with se = mean*0.2, added in cost function
  c_admin_enrollee_se    <- c_admin_enrollee_mean * est_se
  c_admin_enrollee       <- list(mean = (c_admin_enrollee_mean), shape = ((c_admin_enrollee_mean^2)/(c_admin_enrollee_se^2)), rate = ((c_admin_enrollee_mean)/(c_admin_enrollee_se^2)))
  
  # scaling factor for mi and stroke costs in the year of the event [see 'Inputs' excel file] vs the years after the event
  delta_c_y1    <- list(mean = 0.33, shape1 = 16.5, shape2 = 34.3) 
  
  # change in costs from medicaid (oop + other = total) for US non-immigrants (assumed 2019 prices); https://jamanetwork.com/journals/jamanetworkopen/fullarticle/2809604
  delta_c_total    <- list(mean = 0.105, shape1 = 2.10, shape2 = 17.81) 
  # proportion of total costs that are covered by other (e.g. insurer), i.e. not OOP costs (Determined by total costs so deterministic); see LB's "Inputs_'date'.xls" for calcs
  non_oop_prop     <- 1.13 # no probabilistic component
  # Average out-of-pocket costs (USD 2017) per person in pre-expansion states (aged 19-64 and 138% below the FPL); http://dx.doi.org/10.1136/bmj.m40 
  c_oop            <- 429 * pce_hc_2017 # no probabilistic component

  # 2021 US average annual earnings (from OECD 2022) for premature mortality before retirement at age 65
  annual_earnings  <- ifelse(.c_prem_mort == 1, 74000, 0) # no probabilistic component
  # Annual per person absenteeism costs [See LB's inputs excel file for calculations]
  cost_absenteeism <- list(shape = 3.739, rate = 0.015, mean = 244)
  # Annual per person short-term disability costs [See LB's inputs excel file for calculations]
  cost_disability  <- list(shape = 328.882, rate = 0.367, mean = 897)  
  
  # Coefficients for odds of non-zero healthcare costs for Morey et al 2021 equation
    odds_hc <- list(
      int      = list(mean = 1.653, lci = 1.168, uci = 2.339) # Intercept for estimating odds of non-zero healthcare costs 
    , mi       = list(mean = 2.937, lci = 1.798, uci = 4.798) # non-zero healthcare costs for those with an MI from Morey et al (2021)
    , stroke   = list(mean = 0.948, lci = 0.354, uci = 2.540) # non-zero healthcare costs for those with a stroke from Morey et al (2021) - see "Costs" function below   
    , hf       = list(mean = 1.830, lci = 0.402, uci = 8.325) # heart failure
    , cd       = list(mean = 2.389, lci = 1.269, uci = 4.499) # cardiac_dysrhythmia
    , ang      = list(mean = 1.472, lci = 0.566, uci = 3.828) # angina
    , pad      = list(mean = 1.334, lci = 0.617, uci = 2.888) # peripheral_artery_disease
    , diab     = list(mean = 4.382, lci = 3.401, uci = 5.630) # diabetes
    , a_25     = list(mean = 1.171, lci = 1.036, uci = 1.323) # age_cat == "25-44" 
    , a_45     = list(mean = 2.121, lci = 1.834, uci = 2.453) # age_cat == "45-64"
    , a_65     = list(mean = 3.728, lci = 2.076, uci = 6.695) # age_cat == "65+"
    , fem      = list(mean = 2.369, lci = 2.153, uci = 2.607) # female == 1
    , race_w   = list(mean = 1.627, lci = 1.452, uci = 1.824) # race == "white"
    , race_b   = list(mean = 1.013, lci = 0.885, uci = 1.161) # race == "black"
    , race_a   = list(mean = 1.025, lci = 0.848, uci = 1.239) # race == "asian"
    , race_o   = list(mean = 1.461, lci = 1.134, uci = 1.881) # race == "other_race"
    , medicare = list(mean = 1.714, lci = 0.918, uci = 3.2  ) # medicare
    , medicaid = list(mean = 0.959, lci = 0.834, uci = 1.104) # medicaid
    , uninsur  = list(mean = 0.355, lci = 0.313, uci = 0.403) # uninsured
    , inc_np   = list(mean = 0.905, lci = 0.726, uci = 1.128) # fam_income == "near_poor"
    , inc_l    = list(mean = 0.919, lci = 0.789, uci = 1.071) # fam_income == "low"
    , inc_m    = list(mean = 0.960, lci = 0.843, uci = 1.092) # fam_income == "medium"
    , inc_h    = list(mean = 1.344, lci = 1.127, uci = 1.601) # fam_income == "high"
    , ed_gh    = list(mean = 1.178, lci = 1.029, uci = 1.348) # education == "ged_hs"      
    , ed_ab    = list(mean = 1.726, lci = 1.462, uci = 2.038) # education == "associate_bachelor"
    , ed_md    = list(mean = 2.124, lci = 1.684, uci = 2.678) # education == "master_doctorate"
    , bmi_norm = list(mean = 0.985, lci = 0.709, uci = 1.369) # bmi_cat == "normal_weight"
    , bmi_ovw  = list(mean = 1.015, lci = 0.726, uci = 1.417) # bmi_cat == "overweight"  
    , bmi_obe  = list(mean = 1.233, lci = 0.873, uci = 1.74 ) # bmi_cat == "obese"     
    , cci1     = list(mean = 2.037, lci = 1.701, uci = 2.441) # cci == 1
    , cci2     = list(mean = 3.920, lci = 2.08 , uci = 7.387) # cci == 2
    , cci3     = list(mean = 8.287, lci = 3.163, uci = 21.715) # cci > 2
  )

  # Coefficients for mean of non-zero healthcare costs for Morey et al 2021 equation
  hc_cost <- list(
      int      = list(mean = 2600.846, lci = 1902.747, uci = 3555.072) # Intercept for estimating odds of non-zero healthcare costs 
    , mi       = list(mean = 1.177, lci = 1.032, uci = 1.343) # non-zero healthcare costs for those with an MI from Morey et al (2021)
    , stroke   = list(mean = 1.016, lci = 0.871, uci = 1.186) # non-zero healthcare costs for those with a stroke from Morey et al (2021) - see "Costs" function below   
    , hf       = list(mean = 0.948, lci = 0.810, uci = 1.110) # heart failure 
    , cd       = list(mean = 1.449, lci = 1.286, uci = 1.634) # cardiac_dysrhythmia
    , ang      = list(mean = 1.442, lci = 1.172, uci = 1.774) # angina
    , pad      = list(mean = 1.425, lci = 1.23 , uci = 1.651) # peripheral_artery_disease
    , diab     = list(mean = 1.676, lci = 1.471, uci = 1.909) # diabetes
    , a_25     = list(mean = 1.403, lci = 1.213, uci = 1.622) # age_cat == "25-44" 
    , a_45     = list(mean = 2.104, lci = 1.815, uci = 2.439) # age_cat == "45-64"
    , a_65     = list(mean = 2.240, lci = 1.711, uci = 2.932) # age_cat == "65+"
    , fem      = list(mean = 1.145, lci = 1.06 , uci = 1.236) # female == 1
    , race_w   = list(mean = 1.180, lci = 1.059, uci = 1.315) # race == "white"
    , race_b   = list(mean = 1.079, lci = 0.953, uci = 1.221) # race == "black"
    , race_a   = list(mean = 1.034, lci = 0.866, uci = 1.235) # race == "asian"
    , race_o   = list(mean = 1.224, lci = 1.02 , uci = 1.469) # race == "other_race"
    , medicare = list(mean = 1.122, lci = 0.898, uci = 1.403) # medicare
    , medicaid = list(mean = 1.308, lci = 1.17 , uci = 1.462) # medicaid
    , uninsur  = list(mean = 0.619, lci = 0.511, uci = 0.749) # uninsured
    , inc_np   = list(mean = 0.958, lci = 0.847, uci = 1.084) # fam_income == "near_poor"
    , inc_l    = list(mean = 1.020, lci = 0.882, uci = 1.179) # fam_income == "low"
    , inc_m    = list(mean = 0.919, lci = 0.824, uci = 1.026) # fam_income == "medium"
    , inc_h    = list(mean = 0.960, lci = 0.853, uci = 1.08 ) # fam_income == "high"
    , ed_gh    = list(mean = 1.076, lci = 0.943, uci = 1.228) # education == "ged_hs"      
    , ed_ab    = list(mean = 1.114, lci = 0.963, uci = 1.289) # education == "associate_bachelor"
    , ed_md    = list(mean = 1.185, lci = 1.008, uci = 1.392) # education == "master_doctorate"
    , bmi_norm = list(mean = 0.822, lci = 0.653, uci = 1.035) # bmi_cat == "normal_weight"
    , bmi_ovw  = list(mean = 0.794, lci = 0.628, uci = 1.004) # bmi_cat == "overweight"  
    , bmi_obe  = list(mean = 0.922, lci = 0.73 , uci = 1.165) # bmi_cat == "obese"     
    , cci1     = list(mean = 1.558, lci = 1.358, uci = 1.788) # cci == 1
    , cci2     = list(mean = 2.582, lci = 1.94 , uci = 3.436) # cci == 2
    , cci3     = list(mean = 2.991, lci = 2.447, uci = 3.655) # cci > 2
  )
  
  #### Utility parameters ----
  
  # Annual utilities, as a function of CVD status among other variables,
  # are estimated per person using the "Effs" function below using
  # the Morey et al 2021 equation (DOI: 10.1161/CIRCOUTCOMES.120.006769)

  # Coefficients for mean of utility (SF-6D) for Morey et al 2021 equation
  u <- list(
      int      = list(mean =  0.817, lci =  0.798, uci =  0.835) # Intercept for estimating odds of non-zero healthcare costs 
    , mi       = list(mean = -0.012, lci = -0.021, uci = -0.003) # non-zero healthcare costs for those with an MI from Morey et al (2021)
    , stroke   = list(mean = -0.026, lci = -0.043, uci = -0.009) # non-zero healthcare costs for those with a stroke from Morey et al (2021) - see "Costs" function below   
    , hf       = list(mean = -0.045, lci = -0.066, uci = -0.024) # heart failure
    , cd       = list(mean = -0.024, lci = -0.035, uci = -0.012) # cardiac_dysrhythmia
    , ang      = list(mean = -0.051, lci = -0.071, uci = -0.031) # angina
    , pad      = list(mean = -0.035, lci = -0.055, uci = -0.014) # peripheral_artery_disease
    , diab     = list(mean = -0.033, lci = -0.038, uci = -0.028) # diabetes
    , a_25     = list(mean = -0.026, lci = -0.035, uci = -0.012) # age_cat == "25-44" 
    , a_45     = list(mean = -0.047, lci = -0.071, uci = -0.031) # age_cat == "45-64"
    , a_65     = list(mean = -0.052, lci = -0.055, uci = -0.014) # age_cat == "65+"
    , fem      = list(mean = -0.024, lci = -0.038, uci = -0.028) # female == 1
    , race_w   = list(mean = -0.021, lci = -0.026, uci = -0.016) # race == "white"
    , race_b   = list(mean = -0.003, lci = -0.009, uci =  0.004) # race == "black"
    , race_a   = list(mean = -0.008, lci = -0.016, uci = -0.001) # race == "asian"
    , race_o   = list(mean = -0.030, lci = -0.041, uci = -0.019) # race == "other_race"
    , medicare = list(mean = -0.007, lci = -0.022, uci =  0.007) # medicare
    , medicaid = list(mean = -0.066, lci = -0.073, uci = -0.058) # medicaid
    , uninsur  = list(mean = -0.016, lci = -0.023, uci = -0.009) # medicaid
    , inc_np   = list(mean =  0.009, lci = -0.002, uci =  0.020) # fam_income == "near_poor"
    , inc_l    = list(mean =  0.023, lci =  0.015, uci =  0.032) # fam_income == "low"
    , inc_m    = list(mean =  0.040, lci =  0.033, uci =  0.047) # fam_income == "medium"
    , inc_h    = list(mean =  0.063, lci =  0.055, uci =  0.072) # fam_income == "high"
    , ed_gh    = list(mean =  0.013, lci =  0.007, uci =  0.019) # education == "ged_hs"      
    , ed_ab    = list(mean =  0.016, lci =  0.009, uci =  0.022) # education == "associate_bachelor"
    , ed_md    = list(mean =  0.022, lci =  0.014, uci =  0.030) # education == "master_doctorate"
    , bmi_norm = list(mean =  0.007, lci = -0.009, uci =  0.023) # bmi_cat == "normal_weight"
    , bmi_ovw  = list(mean =  0.006, lci = -0.010, uci =  0.023) # bmi_cat == "overweight"  
    , bmi_obe  = list(mean = -0.015, lci = -0.031, uci =  0.001) # bmi_cat == "obese"     
    , cci1     = list(mean = -0.039, lci = -0.045, uci = -0.034) # cci == 1
    , cci2     = list(mean = -0.052, lci = -0.063, uci = -0.041) # cci == 2
    , cci3     = list(mean = -0.076, lci = -0.093, uci = -0.060) # cci > 2
  )

} # Close "PARAMETERS" section
##### FUNCTIONS ----
{
  #### Create new folders function ----
  create_folders <- function(){
    
    folder_names <- c("data", "rmd", "docs", "figures", "tables", "scripts")
    purrr::walk(folder_names, dir.create)
    
    sub_folder_data <- c(here("data", "raw_data"),
                         here("data", "output_data"))
    
    purrr::walk(sub_folder_data, dir.create)
    
    
    sub_folder_docs <- c(here("docs", "Lit"),
                         here("docs", "Manuscript")
    )
    
    purrr::walk(sub_folder_docs, dir.create)
    
  }
  
  #### Check R project exists function ----
  check_Rproject_exists <- function(){
    dplyr::if_else(length(list.files(path = here::here(), pattern = "\\.Rproj$")) > 0,
                   "You are working in an R Project. You can go on.",
                   "STOP!!! Please create an R Project before going further:\nGo to File --> New Project --> New Directory --> New Project."
    )
  }
  check_Rproject_exists()
  
  
  #### Missingness functions ----
  #to plot missing values
  gtmiss <- function(...){
    DataExplorer::profile_missing(...) %>% 
      arrange(pct_missing) %>% 
      print(n=Inf)
  }
  #e.g. gtmiss(dataset)
  
  #to list missing values
  ggmiss <- function(...){
    DataExplorer::plot_missing(...) 
  }
  #e.g. ggmiss(dataset)
  
  
  #### ASCVD weighting function ----
  # The ASCVD function predicts an individuals 10 year ASCVD risk and estimates their average risk within age, sex and race strata
  
  average_ascvd_risk <- function(mydata) {
    
    mydata <- mydata %>%
      
      # Estimate 10 year ASCVD risk
      mutate(ascvd_risk_10yr = case_when(
        # When using NHANES create "black_race" variable and apply this to black individuals as all non-black races are ascribed the white-race risk 
        # see (https://www.framinghamheartstudy.org/fhs-risk-functions/cardiovascular-disease-10-year-risk/) for details
        (race != "black" & female == 0) ~ (1 - 0.9144) * exp(12.344 * log(age) + 
                                                               0      * log(age) * log(age) + 
                                                               11.853 * log(total_cholesterol) + 
                                                               -2.664 * log(age) * log(total_cholesterol) +
                                                               -7.990 * log(hdl_cholesterol) + 
                                                               1.769  * log(age) * log(hdl_cholesterol) + 
                                                               (1.797 * log(mean_systolic_bp_orig) + 
                                                                  0      * log(age) * log(mean_systolic_bp_orig)) * treated + 
                                                               (1.764 * log(mean_systolic_bp_orig) + 
                                                                  0      * log(age) * log(mean_systolic_bp_orig)) * (1 - treated) + 
                                                               7.837  * smoking +
                                                               -1.795 * log(age) * smoking + 
                                                               0.658  * diabetes_orig + 
                                                               -(61.18)),
        (race != "black" & female == 1) ~ (1 - 0.9665) * exp(-29.799   * log(age) + 
                                                               4.884   * log(age)*log(age) + 
                                                               13.54   * log(total_cholesterol) + 
                                                               -3.114  * log(age) * log(total_cholesterol) +
                                                               -13.578 * log(hdl_cholesterol) + 
                                                               3.149   * log(age) * log(hdl_cholesterol) + 
                                                               (2.019  * log(mean_systolic_bp_orig) + 
                                                                  0       * log(age) * log(mean_systolic_bp_orig)) * treated + 
                                                               (1.957  * log(mean_systolic_bp_orig) + 
                                                                  0       * log(age) * log(mean_systolic_bp_orig)) * (1 - treated) + 
                                                               7.574   * smoking +
                                                               -1.665  * log(age) * smoking + 
                                                               0.661   * diabetes_orig + 
                                                               (-(-29.18))),
        (race == "black" & female == 0) ~ (1 - 0.8954) * exp(2.469    * log(age) + 
                                                               0      * log(age) * log(age) + 
                                                               0.302  * log(total_cholesterol) + 
                                                               0      * log(age) * log(total_cholesterol) +
                                                               -0.307 * log(hdl_cholesterol) + 
                                                               0      * log(age) * log(hdl_cholesterol) + 
                                                               (1.916 * log(mean_systolic_bp_orig) + 
                                                                  0      * log(age) * log(mean_systolic_bp_orig)) * treated + 
                                                               (1.809 * log(mean_systolic_bp_orig) + 
                                                                  0      * log(age) * log(mean_systolic_bp_orig)) * (1 - treated) + 
                                                               0.549  * smoking +
                                                               0      * log(age) * smoking + 
                                                               0.645  * diabetes_orig + 
                                                               -(19.54)),
        (race == "black" & female == 1) ~ (1 - 0.9533) * exp(17.114    * log(age) + 
                                                               0       * log(age) * log(age) + 
                                                               0.940   * log(total_cholesterol) + 
                                                               0       * log(age) * log(total_cholesterol) +
                                                               -18.920 * log(hdl_cholesterol) + 
                                                               4.475   * log(age)*log(hdl_cholesterol) + 
                                                               (29.291 * log(mean_systolic_bp_orig) +
                                                                  -6.432  * log(age) * log(mean_systolic_bp_orig)) * treated + 
                                                               (27.820 * log(mean_systolic_bp_orig) + 
                                                                  -6.087  * log(age) * log(mean_systolic_bp_orig)) * (1 - treated) + 
                                                               0.691   * smoking +
                                                               0       * log(age) * smoking + 
                                                               0.874   * diabetes_orig + 
                                                               -(86.61)),
        TRUE~NA_real_),
        
        ascvd_risk_10yr   = ifelse(ascvd_risk_10yr < 1, 
                                   # Equation can result in probabilities > 1 for very high ages; this puts a limit on it [not relevant if restricting model to age < 65]
                                   ascvd_risk_10yr, 0.999999999),               
        # Estimate individual weights to be applied to FHS MI or stroke risk within strata 
        ascvd_rate   = ((-log(1 - ascvd_risk_10yr))/(10)), 
        ascvd_risk   = 1 - exp(-(ascvd_rate) * 1)) %>%
      # Estimate average 10 year ASCVD risk within strata according to age and sex 
      group_by(across(all_of( .ascvd_cat ))) %>%
      mutate(avg_ascvd_risk = mean(ascvd_risk)) %>%
      ungroup() %>%
      select(-ascvd_risk,
             -ascvd_rate,
             -ascvd_risk_10yr)
    
    return(mydata) 
    
  }
  
  #### Probability function ----
  # The Probs function that updates the transition probabilities of every cycle is shown below.
  
  # Predict annual mi risk for those with no history of CVD from FHS monthly risk re-weighted by ASCVD event risk
  Probs <- function(mydata
                    , Trt = FALSE
                    , parameters = parameters
                    , seed_offset = seed_offset) {
    
    mydata  <- mydata %>%
      # Estimate 10 year ASCVD risk
      mutate(ascvd_risk_10yr = case_when(
        # When using NHANES create "black_race" variable and apply this to black individuals as all non-black races are ascribed the white-race risk 
        # see (https://www.framinghamheartstudy.org/fhs-risk-functions/cardiovascular-disease-10-year-risk/) for details
        (race != "black" & female == 0) ~ (1 - 0.9144) * exp(12.344 * log(age) + 
                                                               0      * log(age) * log(age) + 
                                                               11.853 * log(total_cholesterol) + 
                                                               -2.664 * log(age) * log(total_cholesterol) +
                                                               -7.990 * log(hdl_cholesterol) + 
                                                               1.769  * log(age) * log(hdl_cholesterol) + 
                                                               (1.797 * log(mean_systolic_bp) + 
                                                                  0      * log(age) * log(mean_systolic_bp)) * treated + 
                                                               (1.764 * log(mean_systolic_bp) + 
                                                                  0      * log(age) * log(mean_systolic_bp)) * (1 - treated) + 
                                                               7.837  * smoking +
                                                               -1.795 * log(age) * smoking + 
                                                               0.658  * diabetes + 
                                                               -(61.18)),
        (race != "black" & female == 1) ~ (1 - 0.9665) * exp(-29.799   * log(age) + 
                                                               4.884   * log(age)*log(age) + 
                                                               13.54   * log(total_cholesterol) + 
                                                               -3.114  * log(age) * log(total_cholesterol) +
                                                               -13.578 * log(hdl_cholesterol) + 
                                                               3.149   * log(age) * log(hdl_cholesterol) + 
                                                               (2.019  * log(mean_systolic_bp) + 
                                                                  0       * log(age) * log(mean_systolic_bp)) * treated + 
                                                               (1.957  * log(mean_systolic_bp) + 
                                                                  0       * log(age) * log(mean_systolic_bp)) * (1 - treated) + 
                                                               7.574   * smoking +
                                                               -1.665  * log(age) * smoking + 
                                                               0.661   * diabetes + 
                                                               (-(-29.18))),
        (race == "black" & female == 0) ~ (1 - 0.8954) * exp(2.469    * log(age) + 
                                                               0      * log(age) * log(age) + 
                                                               0.302  * log(total_cholesterol) + 
                                                               0      * log(age) * log(total_cholesterol) +
                                                               -0.307 * log(hdl_cholesterol) + 
                                                               0      * log(age) * log(hdl_cholesterol) + 
                                                               (1.916 * log(mean_systolic_bp) + 
                                                                  0      * log(age) * log(mean_systolic_bp)) * treated + 
                                                               (1.809 * log(mean_systolic_bp) + 
                                                                  0      * log(age) * log(mean_systolic_bp)) * (1 - treated) + 
                                                               0.549  * smoking +
                                                               0      * log(age) * smoking + 
                                                               0.645  * diabetes + 
                                                               -(19.54)),
        (race == "black" & female == 1) ~ (1 - 0.9533) * exp(17.114    * log(age) + 
                                                               0       * log(age) * log(age) + 
                                                               0.940   * log(total_cholesterol) + 
                                                               0       * log(age) * log(total_cholesterol) +
                                                               -18.920 * log(hdl_cholesterol) + 
                                                               4.475   * log(age)*log(hdl_cholesterol) + 
                                                               (29.291 * log(mean_systolic_bp) +
                                                                  -6.432  * log(age) * log(mean_systolic_bp)) * treated + 
                                                               (27.820 * log(mean_systolic_bp) + 
                                                                  -6.087  * log(age) * log(mean_systolic_bp)) * (1 - treated) + 
                                                               0.691   * smoking +
                                                               0       * log(age) * smoking + 
                                                               0.874   * diabetes + 
                                                               -(86.61)),
        TRUE~NA_real_),
        ascvd_risk_10yr   = ifelse(ascvd_risk_10yr < 1, 
                                   # Equation can result in probabilities > 1 for very high ages; this puts a limit on it [not relevant if restricting model to age < 65]
                                   ascvd_risk_10yr, 0.9999999),               
        # Estimate individual weights to be applied to FHS MI or stroke risk within strata 
        ascvd_rate   = ((-log(1 - ascvd_risk_10yr))/(10)), 
        ascvd_risk   = 1 - exp(-(ascvd_rate) * 1),
        
        ## Estimate monthly FHS MI risk, convert to annual rate, apply 10 year ASCVD weights, convert to re-weighted annual MI risk for those without a history of CVD
        # Estimate monthly MI risk as a function of age within categories of gender
        monthly_mi_risk       = ifelse(female == 0,
                                       0.0001   * exp(0.0312 * age),      
                                       0.000008 * exp(0.0599 * age)),
        
        # Equation can result in probabilities > 1 for very high ages; this puts a limit on it
        monthly_mi_risk       = ifelse(monthly_mi_risk < 1, 
                                       monthly_mi_risk, 1),  
        
        # Convert monthly risk to annual rate (assumes constant rate over time)
        annual_mi_rate        = ((-log(1 - monthly_mi_risk))/(1 / 12)),  
        
        # Convert annual rate to annual risk (assumes constant rate over time)  
        annual_mi_risk        = 1 - exp(-(annual_mi_rate) * 1),   
        
        # Using Bayes formula to estimate the probability of an MI given an ASCVD event for age and sex categories [avg_ascvd_risk grouping (age and sex) should be the same as annual_mi_risk parameters (age and sex)]
        annual_ascvd_mi_risk  = annual_mi_risk / avg_ascvd_risk, 
        
        # Use Bayes formula to then estimate the probability of an ASCVD and MI event [this is more granular as the probability of an ASCVD event uses more information, importantly systolic BP]  
        mi_risk_noCVD         = annual_ascvd_mi_risk * ascvd_risk,          
        
        ## Estimate monthly FHS stroke risk, convert to annual rate, apply 10 year ASCVD weights, convert to re-weighted annual stroke risk for those without a history of CVD
        # Estimate monthly stroke risk as a function of age within categories of gender
        monthly_stroke_risk      = ifelse(female == 0, 
                                          0.000009 * exp(0.0622 * age),             
                                          0.000003 * exp(0.0741 * age)),
        
        # Equation can result in probabilities > 1 for very high ages; this puts a limit on it
        monthly_stroke_risk      = ifelse(monthly_stroke_risk < 1, 
                                          monthly_stroke_risk, 1),  
        
        # Convert monthly risk to annual rate (assumes constant rate over time)
        annual_stroke_rate       = ((-log(1 - monthly_stroke_risk))/(1 / 12)), 
        
        # Convert annual rate to annual risk (assumes constant rate over time) 
        annual_stroke_risk       = 1 - exp(-(annual_stroke_rate) * 1),  
        
        # Using Bayes formula to estimate the probability of an MI given an ASCVD event for age and sex categories [avg_ascvd_risk grouping (age and sex) should be the same as annual_mi_risk parameters (age and sex)]
        annual_ascvd_stroke_risk = annual_stroke_risk / avg_ascvd_risk,
        
        # Use Bayes formula to then estimate the probability of an ASCVD and MI event [this is more granular as the probability of an ASCVD event uses more information, importantly systolic BP] 
        stroke_risk_noCVD        = annual_ascvd_stroke_risk * ascvd_risk,          
        
        # Estimate MI mortality risk as a function of age within categories of gender
        mi_mortality_risk     = ifelse(female == 0, 
                                       0.0289 * exp(0.0269 * age),        
                                       0.0004 * exp(0.0706 * age)), 
        
        # Equation can result in probabilities > 1 for very high ages; this puts a limit on it
        mi_mortality_risk     = ifelse(mi_mortality_risk < 1, 
                                       mi_mortality_risk, 1),                        

        # Estimate stroke mortality risk as a function of age within categories of gender
        stroke_mortality_risk     = ifelse(female == 0, 
                                           0.0003 * exp(0.0782 * age),                
                                           0.0034 * exp(0.0428 * age)),
        
        # Equation can result in probabilities > 1 for very high ages; this puts a limit on it
        stroke_mortality_risk     = ifelse(stroke_mortality_risk < 1, 
                                           stroke_mortality_risk, 1),                

        # Convert annual MI risk to rate, apply "history of CVD multiplier" (see Basu, 2017)  and convert back to prob
        mi_risk_CVD            = 1 - exp((-(((-log(1 - mi_risk_noCVD))/(1)) * parameters[["CVD_history_multiplier"]])) * 1), # Conversion assumes constant rate over time
        
        # Convert annual stroke risk to rate, apply "history of CVD multiplier" (see Basu, 2017)  and convert back to prob
        stroke_risk_CVD        = 1 - exp((-(((-log(1 - stroke_risk_noCVD))/(1)) * parameters[["CVD_history_multiplier"]])) * 1) # Conversion assumes constant rate over time
      ) 
    
    # add non-CVD mortality data
    mydata <- left_join(mydata, 
                        simulated_data_nCVD, 
                        by = c("age", "female"), 
                        relationship = "many-to-many")
    
    # These ensure the same dataset across bootstrapped replications, else (.fix_slice_sample != 1) then different samples across replications
    if (.Deterministic == 1) {
      set.seed(seed)
    } else {
      set.seed(seed + seed_offset)
    }
    
    mydata  <- mydata %>%
      # Estimate non-CVD mortality risk
      mutate(
        # calculate annual non-CVD mortality rate (per person) as difference between all-cause mortality rate from CVD mortality rate in CDC WONDER database
        noncvd_mortality_rate  = case_when(.Deterministic == 1 ~ rate,
                                           .Deterministic != 1 ~ rnorm(n(), rate, se)),
        noncvd_mortality_rate  = case_when(Trt == FALSE ~ noncvd_mortality_rate, 
                                           Trt == TRUE  ~ noncvd_mortality_rate - (ifelse(medicaid_exp == 1 & .nCVDmort == 1, 
                                                                                          parameters[["rd_nCVD_mort_ins"]], 0)),
                                           TRUE ~ noncvd_mortality_rate),
        
        # Convert non-CVD rate to annual risk
        noncvd_mortality_risk  = 1 - exp((-noncvd_mortality_rate) * 1), # Conversion assumes constant rate over time
        
        # Equation can result in probabilities > 1 for very high ages; this puts a limit on it
        noncvd_mortality_risk  = ifelse(noncvd_mortality_risk < 1, 
                                        noncvd_mortality_risk, 1)                               
      )  %>%  
      # Keep only variables for transitions matrices
      select(stroke_risk_noCVD,
             stroke_risk_CVD,
             stroke_mortality_risk,
             mi_risk_noCVD,
             mi_risk_CVD,
             mi_mortality_risk,
             noncvd_mortality_risk)
    
    return(mydata) 
    
  }
  
  
  #### Cost function ----
  # The Costs function estimates the costs at every cycle

  Costs <- function(mydata,
                    cci,
                    cardiac_dysrhythmia,      
                    peripheral_artery_disease,
                    Trt = FALSE,
                    parameters = parameters) {
    
    # Below formulas come from Morey et al (2021; DOI: 10.1161/CIRCOUTCOMES.120.006769) using MEPS 2016 data
    mydata  <- mydata %>%
      mutate( # calculate the odds of having non-zero health expenditures 
        # to avoid non-sensical values draw from log values and exponentiate after
        odds_nonzero_hc = ((parameters[["odds_nonzero_hc_int"]]) * # baseline utility                                                 
                             ifelse(mi == 1,                             parameters[["odds_nonzero_hc_mi"]], 1) *
                             ifelse(stroke == 1,                         parameters[["odds_nonzero_hc_stroke"]], 1) *
                             ifelse(heart_failure == 1,                  parameters[["odds_nonzero_hc_hf"]], 1) *
                             # do not have data in nhanes on dc and pad so set == 0 for eveyrone when using function exp
                             ifelse(cardiac_dysrhythmia == 1,            parameters[["odds_nonzero_hc_cd"]]  , 1) *      
                             ifelse(angina == 1,                         parameters[["odds_nonzero_hc_ang"]] , 1) *
                             ifelse(peripheral_artery_disease == 1,      parameters[["odds_nonzero_hc_pad"]] , 1) *
                             ifelse(diabetes == 1,                       parameters[["odds_nonzero_hc_diab"]], 1) *
                             # baseline = "18-24"
                             ifelse((age_cat == "25-44"),                parameters[["odds_nonzero_hc_a_25"]], 1) *       
                             ifelse((age_cat == "45-64"),                parameters[["odds_nonzero_hc_a_45"]], 1) *
                             ifelse((age_cat == "65+"),                  parameters[["odds_nonzero_hc_a_65"]], 1) *
                             ifelse(female == 1,                         parameters[["odds_nonzero_hc_fem"]], 1) *
                             # baseline = "hispanic"
                             ifelse((race == "white"),                   parameters[["odds_nonzero_hc_race_w"]], 1) *       
                             ifelse((race == "black"),                   parameters[["odds_nonzero_hc_race_b"]], 1) *
                             ifelse((race == "asian"),                   parameters[["odds_nonzero_hc_race_a"]], 1) *
                             ifelse((race == "other_race"),              parameters[["odds_nonzero_hc_race_o"]], 1) *
                             # baseline = "Private insurance"
                             ifelse((insurance == "medicare"),      parameters[["odds_nonzero_hc_medicare"]], 1) *      
                             # referred to as "other public" but assumed to be "medicaid" - vast majority of "other public" insurance in 2016 by type was medicaid [https://www.census.gov/content/dam/Census/library/visualizations/2017/demo/p60-260/figure1.pdf]
                             ifelse((insurance == "medicaid" | 
                                       insurance == "other plan"),  parameters[["odds_nonzero_hc_medicaid"]], 1) *       
                             ifelse((insurance == "uninsured"),     parameters[["odds_nonzero_hc_uninsur"]], 1) *
                             # baseline = "Poor"
                             ifelse((fam_income == "near_poor"),         parameters[["odds_nonzero_hc_inc_np"]], 1) *       
                             ifelse((fam_income == "low"),               parameters[["odds_nonzero_hc_inc_l"]], 1) *
                             ifelse((fam_income == "medium"),            parameters[["odds_nonzero_hc_inc_m"]], 1) *
                             ifelse((fam_income == "high"),              parameters[["odds_nonzero_hc_inc_h"]], 1) *
                             # baseline = "no degree"
                             ifelse((education == "ged_hs"),             parameters[["odds_nonzero_hc_ed_gh"]], 1) *       
                             ifelse((education == "associate_bachelor"), parameters[["odds_nonzero_hc_ed_ab"]], 1) * 
                             ifelse((education == "master_doctorate"),   parameters[["odds_nonzero_hc_ed_md"]], 1) *
                             # baseline = "Underweight" == "1" in NHANES
                             ifelse((bmi_cat == "normal_weight"),        parameters[["odds_nonzero_hc_bmi_norm"]], 1) *        
                             ifelse((bmi_cat == "overweight"),           parameters[["odds_nonzero_hc_bmi_ovw"]], 1) *
                             ifelse((bmi_cat == "obese"),                parameters[["odds_nonzero_hc_bmi_obe"]] , 1) *
                             # baseline = 0
                             ifelse((cci == 1),                          parameters[["odds_nonzero_hc_cci1"]], 1) *       
                             ifelse((cci == 2),                          parameters[["odds_nonzero_hc_cci2"]], 1) *
                             ifelse((cci > 2),                           parameters[["odds_nonzero_hc_cci3"]], 1)),

        
        # calculate the probability of having non-zero health expenditures
        prob_nonzero_hc = (odds_nonzero_hc/(1+odds_nonzero_hc)),
  
        # calculate expected healthcare expenditure 
        # baseline utility 
        nonzero_hc_cost = ((parameters[["nonzero_hc_cost_int"]]) *                                        
                             ifelse(mi == 1,                             parameters[["nonzero_hc_cost_mi"]], 1) *
                             ifelse(stroke == 1,                         parameters[["nonzero_hc_cost_stroke"]], 1) *
                             ifelse(heart_failure == 1,                  parameters[["nonzero_hc_cost_hf"]], 1) *
                             # do not have data in nhanes on dc and pad so set == 0 for eveyrone when using function
                             ifelse(cardiac_dysrhythmia == 1,            parameters[["nonzero_hc_cost_cd"]]  , 1) *            
                             ifelse(angina == 1,                         parameters[["nonzero_hc_cost_ang"]] , 1) *
                             ifelse(peripheral_artery_disease == 1,      parameters[["nonzero_hc_cost_pad"]] , 1) *
                             ifelse(diabetes == 1,                       parameters[["nonzero_hc_cost_diab"]], 1) *
                             # baseline = "18-24"
                             ifelse((age_cat == "25-44"),                parameters[["nonzero_hc_cost_a_25"]], 1) *               
                             ifelse((age_cat == "45-64"),                parameters[["nonzero_hc_cost_a_45"]], 1) *        
                             ifelse((age_cat == "65+"),                  parameters[["nonzero_hc_cost_a_65"]], 1) *        
                             ifelse(female == 1,                         parameters[["nonzero_hc_cost_fem"]], 1) * 
                             # baseline = "hispanic"
                             ifelse((race == "white"),                   parameters[["nonzero_hc_cost_race_w"]], 1) *               
                             ifelse((race == "black"),                   parameters[["nonzero_hc_cost_race_b"]], 1) *        
                             ifelse((race == "asian"),                   parameters[["nonzero_hc_cost_race_a"]], 1) *        
                             ifelse((race == "other_race"),              parameters[["nonzero_hc_cost_race_o"]], 1) *
                             # baseline = "Private insurance"
                             ifelse((insurance == "medicare"),      parameters[["nonzero_hc_cost_medicare"]], 1) *               
                             # referred to as "other public"; majority would be "medicaid" - vast majority of "other public" insurance in 2016 by type was medicaid [https://www.census.gov/content/dam/Census/library/visualizations/2017/demo/p60-260/figure1.pdf]
                             ifelse((insurance == "medicaid" | 
                                       insurance == "other_plan"),  parameters[["nonzero_hc_cost_medicaid"]], 1) *                
                             ifelse((insurance == "uninsured"),     parameters[["nonzero_hc_cost_uninsur"]], 1) *
                             # baseline = "Poor"
                             ifelse((fam_income == "near_poor"),         parameters[["nonzero_hc_cost_inc_np"]], 1) *              
                             ifelse((fam_income == "low"),               parameters[["nonzero_hc_cost_inc_l"]] , 1) *        
                             ifelse((fam_income == "medium"),            parameters[["nonzero_hc_cost_inc_m"]] , 1) *        
                             ifelse((fam_income == "high"),              parameters[["nonzero_hc_cost_inc_h"]] , 1) *
                             # baseline = "no degree"
                             ifelse((education == "ged_hs"),             parameters[["nonzero_hc_cost_ed_gh"]], 1) *              
                             ifelse((education == "associate_bachelor"), parameters[["nonzero_hc_cost_ed_ab"]], 1) *         
                             ifelse((education == "master_doctorate"),   parameters[["nonzero_hc_cost_ed_md"]], 1) *
                             # baseline = "Underweight"
                             ifelse((bmi_cat == "normal_weight"),        parameters[["nonzero_hc_cost_bmi_norm"]], 1) *               
                             ifelse((bmi_cat == "overweight"),           parameters[["nonzero_hc_cost_bmi_ovw"]] , 1) *        
                             ifelse((bmi_cat == "obese"),                parameters[["nonzero_hc_cost_bmi_obe"]] , 1) *
                             # baseline = 0
                             ifelse((cci == 1),                          parameters[["nonzero_hc_cost_cci1"]], 1) *               
                             ifelse((cci == 2),                          parameters[["nonzero_hc_cost_cci2"]], 1) *        
                             ifelse((cci > 2),                           parameters[["nonzero_hc_cost_cci3"]], 1)),
        
        # calculate annual healthcare expenditure times the probability of non-zero healthcare expenditures 
        hc_cost = (prob_nonzero_hc * nonzero_hc_cost),
        
        # inflate to 2021 prices using the Health PCE 
        hc_cost = (hc_cost * pce_hc_2017),
        
        # government administration costs as part of treatment scenario for those receiving medicaid who are under 65 years (at age 65 years everyone switches to Medicare)
        gov_admin_cost = case_when(Trt == TRUE & medicaid_exp == 1 & age < 65 & .gc == 1 ~ parameters[["cost_admin_enrollee"]],
                                   Trt == FALSE ~ 0,
                                   TRUE ~ 0),
        
        # OOP costs are estimated using average OOP costs per person and estimating what proportion these are of average healthcare costs
        mean_hc_cost = mean(hc_cost),
        # OOP costs are proportionate to hc costs
        oop_cost = (c_oop / mean_hc_cost) * hc_cost,
        
        # change in total healthcare costs following receiving medicaid in medicaid expansion
        delta_total_hc_cost = (hc_cost * parameters[["delta_total_cost"]]),
        delta_other_hc_cost = (delta_total_hc_cost * non_oop_prop),
        delta_oop_hc_cost   = (delta_total_hc_cost - delta_other_hc_cost),
      ) %>% 
      
      # Keep only variables for transitions matrices
      select(fpl
             , fam_income
             , medicaid_exp
             , hc_cost
             , oop_cost
             , gov_admin_cost
             , delta_total_hc_cost
             , delta_other_hc_cost
             , delta_oop_hc_cost)
    
    # Calculate the sum of non-oop_costs 
    sum_delta_non_oop_cost <- sum((mydata$delta_other_hc_cost[mydata$medicaid_exp == 1]))
    # Calculate the sum of gov costs 
    sum_gov_cost <- sum(mydata$gov_admin_cost[mydata$medicaid_exp == 1])
    # Calculate number of individuals above or equal to 150% of the FPL and apply non-oop costs to them; http://dx.doi.org/10.1136/bmj.m40 (Supplement, page 3)
    length_non_medicaid <- length(mydata$medicaid_exp[mydata$fpl >= .fpl_reallocate & mydata$medicaid_exp == 0])
    # amount to apportion to all those with income >= 150% FPL
    mean_delta_cost_society <- (sum_delta_non_oop_cost + sum_gov_cost) / length_non_medicaid
    # mean fpl for income >= 150% FPL
    mean_fpl <- mean(mydata$fpl[mydata$fpl >= .fpl_reallocate & mydata$medicaid_exp == 0])
    
    mydata <- mydata %>%
      # combining perspective and ATT scenarios with and without treatment
      mutate(fpl_weight = ifelse(fpl >= 1.5 & medicaid_exp == 0 & .positive_dist == 1, fpl/mean_fpl, 1),
             hc_cost = case_when(
               # for healthcare perspective, uncompensated care costs (non-OOP costs) would be compensated
               (.ATT == 1 & .perspective == "healthcare" & Trt == FALSE) ~ hc_cost,
               (.ATT == 1 & .perspective == "healthcare" & Trt == TRUE & .medicaid_cost == 0)  ~ hc_cost,
               (.ATT == 1 & .perspective == "healthcare" & (Trt == TRUE & medicaid_exp == 1) & .medicaid_cost == 1)  ~ hc_cost + delta_oop_hc_cost,
               # for healthcare perspective, costs of gov admin, disability claims, absenteeism and premature mortality are included
               (.ATT == 1 & .perspective == "societal" & Trt == FALSE) ~ hc_cost,
               (.ATT == 1 & .perspective == "societal" & Trt == TRUE & .medicaid_cost == 0) ~ hc_cost,
               (.ATT == 1 & .perspective == "societal" & (Trt == TRUE & medicaid_exp == 1) & .medicaid_cost == 1) ~ hc_cost + delta_total_hc_cost + gov_admin_cost,
               # always societal perspective when .ATT != 1 as doesn't make sense any other way
               (.ATT != 1 & medicaid_exp == 1 & Trt == TRUE & .medicaid_cost == 1) ~ (hc_cost + delta_oop_hc_cost), # individual has a reduction in their from a reduction in their out-of-pocket costs
               (.ATT != 1 & medicaid_exp == 0 & Trt == TRUE & .medicaid_cost == 1 & fpl >= .fpl_reallocate) ~ hc_cost + (mean_delta_cost_society*fpl_weight),
               TRUE ~ hc_cost)) %>%
      # Keep only variables for transitions matrices
      select(hc_cost)

    return(mydata)
    
  }
  
  #### Utilities function ---- 
  # The Effs function to update the utilities at every cycle

  Effs <- function(mydata,
                   cardiac_dysrhythmia,      
                   peripheral_artery_disease,
                   cci, 
                   Trt = FALSE,
                   parameters = parameters) {
    
    # Below formulas come from Morey et al (2021; DOI: 10.1161/CIRCOUTCOMES.120.006769) using MEPS 2016 data
    mydata  <- mydata %>%
      # calculate annual utility (and by default QALYs; cycle length = 1 year) using DOI: 10.1161/CIRCOUTCOMES.120.006769
      mutate(utility = ((parameters[["util_int"]]) +     # baseline utility                                                          
                          mi                                                     * (parameters[["util_mi"]]) +
                          stroke                                                 * (parameters[["util_stroke"]]) +
                          heart_failure                                          * (parameters[["util_hf"]]) +
                          # do not have data in nhanes on dc and pad so set == 0 for eveyrone when using function
                          cardiac_dysrhythmia                                    * (parameters[["util_cd"]]  ) +     
                          angina                                                 * (parameters[["util_ang"]] ) +
                          peripheral_artery_disease                              * (parameters[["util_pad"]] ) +
                          diabetes                                               * (parameters[["util_diab"]]) +
                          (age_cat == "<25")                                     * 0 +                # baseline = "18-24"
                          (age_cat == "25-44")                                   * (parameters[["util_a_25"]]) +
                          (age_cat == "45-64")                                   * (parameters[["util_a_45"]]) +
                          (age_cat == "65+")                                     * (parameters[["util_a_65"]]) +
                          female                                                 * (parameters[["util_fem"]]) +  
                          (race == "hispanic")                                   * 0 +                # baseline = "hispanic"
                          (race == "white")                                      * (parameters[["util_race_w"]]) + 
                          (race == "black")                                      * (parameters[["util_race_b"]]) +
                          (race == "asian")                                      * (parameters[["util_race_a"]]) +
                          (race == "other_race")                                 * (parameters[["util_race_o"]]) +
                          (insurance == "private")                          * 0 +                # baseline = "Private insurance"
                          (insurance == "medicare")                         * (parameters[["util_medicare"]]) +
                          (insurance == "medicaid" | 
                             insurance == "other_plan")                     * (parameters[["util_medicaid"]]) +  # referred to as "other public" but assumed to be "medicaid" - vast majority of "other public" insurance in 2016 by type was medicaid [https://www.census.gov/content/dam/Census/library/visualizations/2017/demo/p60-260/figure1.pdf]
                          (insurance == "uninsured")                        * (parameters[["util_uninsur"]]) +  
                          (fam_income == "poor")                                 * 0 +                # baseline = "Poor"
                          (fam_income == "near_poor")                            * (parameters[["util_inc_np"]]) + 
                          (fam_income == "low")                                  * (parameters[["util_inc_l"]] ) +
                          (fam_income == "medium")                               * (parameters[["util_inc_m"]] ) +
                          (fam_income == "high")                                 * (parameters[["util_inc_h"]] ) +
                          (education == "no_degree")                             * 0 +                # baseline = "no degree"
                          (education == "ged_hs")                                * (parameters[["util_ed_gh"]]) +
                          (education == "associate_bachelor")                    * (parameters[["util_ed_ab"]]) + 
                          (education == "master_doctorate")                      * (parameters[["util_ed_md"]]) +
                          (bmi_cat == "underweight")                             * 0 +                # baseline = "Underweight"
                          (bmi_cat == "normal_weight")                           * (parameters[["util_bmi_norm"]]) + 
                          (bmi_cat == "overweight")                              * (parameters[["util_bmi_ovw"]] ) +
                          (bmi_cat == "obese")                                   * (parameters[["util_bmi_obe"]] )  +
                          (cci == 0)                                             * 0 +                # baseline = 0
                          (cci == 1)                                             * (parameters[["util_cci1"]]) + 
                          (cci == 2)                                             * (parameters[["util_cci2"]]) +
                          (cci > 2)                                              * (parameters[["util_cci3"]]))
             ) %>% 
      
      # Keep only variables for transitions matrices
      select(utility)
    
    return(mydata) 
    
  }
  
  #### Weighted Standard Error function ----
  
  # https://www.alexstephenson.me/post/2022-04-02-weighted-variance-in-r/
    weighted.se.mean <- function(x, w, na.rm = T){
    ## Remove NAs 
    if (na.rm) {
      i <- !is.na(x)
      w <- w[i]
      x <- x[i]
    }
    
    ## Calculate effective N and correction factor
    n_eff <- (sum(w))^2/(sum(w^2))
    correction = n_eff/(n_eff-1)
    
    ## Get weighted variance 
    numerator = sum(w*(x-weighted.mean(x,w))^2)
    denominator = sum(w)
    
    ## get weighted standard error of the mean 
    se_x = sqrt((correction * (numerator/denominator))/n_eff)
    return(se_x)
    }
  
  #### Rubin's rules Standard Error function ----
  
  # Define a function to calculate the standard error using Rubin's rules
  rubin_se <- function(x, se) {
    se <- sqrt((sum(se^2) / length(x)) + ((1 + (1 / length(x))) * var(x)))
    return(se)
  }
  
  #### Transition probability rescaling function ----
  # Function to rescale non-missing values in each row to equal 1
  # This is necessary to rescale transition probabilities when their sum is greater than or equal to 1
  rescale_to_one <- function(row) {
    non_blank_values <- row[!is.na(row)]
    if ((length(non_blank_values) > 0) & (sum(non_blank_values) >= 1)) {
      row[!is.na(row)] <- non_blank_values / sum(non_blank_values)
    }
    return(row)
  }
  
  #### Transition probability residual function ----
  # Function to replace blank values with 1 minus the sum of non-blank values in each row
  replace_blanks <- function(row) {
    non_blank_values <- row[!is.na(row)]
    if (length(non_blank_values) > 0) {
      row[is.na(row)] <- 1 - sum(non_blank_values)
    }
    return(row)
  }
  
  #### Deterministic v Probabilistic values ----
  
  generate_parameters <- function(Deterministic, ci2se = 3.92) {
    parameter_list <- list()
    
    # Set-up intervention parameters
    variables <- c("int", 
                   "mi", "stroke", "hf", "cd", "ang", "pad", "diab", 
                   "a_25", "a_45", "a_65",
                   "fem", 
                   "race_w", "race_b", "race_a", "race_o", 
                   "medicare", "medicaid", "uninsur", 
                   "inc_np", "inc_l", "inc_m", "inc_h", 
                   "ed_gh", "ed_ab", "ed_md", 
                   "bmi_norm", "bmi_ovw", "bmi_obe",
                   "cci1", "cci2", "cci3")
    
      # predicting odds of non-zero costs
      for (variable in variables) {
        odds_nonzero_var <- if (Deterministic == 1) {
          odds_hc[[variable]]$mean
        } else {
          exp(rnorm(1, log(odds_hc[[variable]]$mean), ((log(odds_hc[[variable]]$uci) - log(odds_hc[[variable]]$lci)) / ci2se)))
        }
        
        parameter_list[[paste0("odds_nonzero_hc_", variable)]] <- odds_nonzero_var
      }
    
    # predicting non-zero costs
    for (variable in variables) {
      nonzero_hc_cost_var <- if (Deterministic == 1) {
        hc_cost[[variable]]$mean
      } else {
        exp(rnorm(1, log(hc_cost[[variable]]$mean), ((log(hc_cost[[variable]]$uci) - log(hc_cost[[variable]]$lci)) / ci2se)))
      }
      
      parameter_list[[paste0("nonzero_hc_cost_", variable)]] <- nonzero_hc_cost_var
    }
    
    # predicting utilities
    for (variable in variables) {
      util_var_value <- if (Deterministic == 1) {
        u[[variable]]$mean
      } else {
        (rnorm(1, (u[[variable]]$mean), (((u[[variable]]$uci) - (u[[variable]]$lci)) / ci2se)))
      }
      
      parameter_list[[paste0("util_", variable)]] <- util_var_value
    }

    # Additional parameters
    # government administration costs as part of treatment scenario for those receiving medicaid who are under 65 years (at age 65 years everyone switches to Medicare)
    parameter_list[["cost_admin_enrollee"]]    <- ifelse(Deterministic == 1, c_admin_enrollee$mean, rgamma(1, shape = c_admin_enrollee$shape, rate = c_admin_enrollee$rate))
    
    # Annual per person absenteeism costs [See LB's inputs excel file for calculations]
    parameter_list[["c_absenteeism"]]          <- ifelse(Deterministic == 1, cost_absenteeism$mean, rgamma(1, shape = cost_absenteeism$shape, rate = cost_absenteeism$rate)) * pce_all_2013
    
    # Option to switch on and off cost of absenteeism
    parameter_list[["c_absenteeism"]]          <- ifelse(.c_absent == 1, parameter_list[["c_absenteeism"]], 0) 
    
    # Annual per person short-term disability costs [See LB's inputs excel file for calculations]
    parameter_list[["c_disability"]]           <- ifelse(Deterministic == 1, cost_disability$mean, rgamma(1, shape = cost_disability$shape, rate = cost_disability$rate)) * pce_all_2013
    
    # Option to switch on and off cost of short-term disability
    parameter_list[["c_disability"]]           <- ifelse(.c_disab == 1, parameter_list[["c_disability"]] , 0)
    
    # change in total healthcare costs following receiving medicaid in medicaid expansion
    parameter_list[["delta_total_cost"]]       <- ifelse(Deterministic == 1, delta_c_total$mean, rbeta(1, delta_c_total$shape1, delta_c_total$shape2))
    
    # rate reduction in non-CVD mortality estimated as difference in all-cause mortality rate from medicaid exp and CVD mortality rate from medicaid exp [https://doi.org/10.1016/ S2468-2667(21)00252-8]
    parameter_list[["rd_nCVD_mort_ins"]]       <- ifelse(Deterministic == 1, delta_nCVD_mort$mean, rnorm(1, delta_nCVD_mort$mean, delta_nCVD_mort$se))
    
    # "history of CVD" multiplier (see Basu, 2017)
    parameter_list[["CVD_history_multiplier"]] <- ifelse(Deterministic == 1, delta_CVD_history$mean, rgamma(1, shape = delta_CVD_history$shape, rate = delta_CVD_history$scale))
    
    # reduction in systolic blood pressure after receiving medicaid [−3.03 mmHg; 95% CI, −5.33 mmHg to − 0.73 mmHg] from DOI: 10.1007/s11606-020-06417-6
    parameter_list[["sysbp_reduction"]]        <- ifelse(Deterministic == 1, delta_sys_bp$mean, rnorm(1, delta_sys_bp$mean, ((delta_sys_bp$uci - delta_sys_bp$lci) / ci2se)))
   
    # reduction in hba1c after receiving medicaid [−0.14 percentage points [pp]; 95% CI, −0.24 pp to −0.03 pp;] from DOI: 10.1007/s11606-020-06417-6
    parameter_list[["hba1c_reduction"]]        <- ifelse(Deterministic == 1, delta_hba1c$mean, rnorm(1, delta_hba1c$mean, ((delta_hba1c$uci - delta_hba1c$lci) / ci2se)))
    
    # change in total healthcare costs in year one of CVD event
    parameter_list[["delta_cost_y1"]]          <- ifelse(Deterministic == 1, delta_c_y1$mean, rbeta(1, delta_c_y1$shape1, delta_c_y1$shape2))

    return(parameter_list)
  }
  
  #### Microsimulation function ----
  # The MicroSim function for the simple microsimulation of the 'Sick-Sicker' model keeps track of what happens to each individual during each cycle. 
  # Adapted from https://doi.org/10.1177%2F0272989X18754513
  
  MicroSim <- function(original_data
                       , n.i
                       , n.t
                       , v.n
                       , d.c
                       , d.e
                       , TR.out = TRUE
                       , TS.out = TRUE
                       , Trt = FALSE
                       , parameters
                       , seed_offset) {
    # Arguments:  
    # n.i:     number of individuals
    # n.t:     total number of cycles to run the model
    # v.n:     vector of health state names
    # d.c:     discount rate for costs
    # d.e:     discount rate for health outcome (QALYs)
    # TR.out:  should the output include a Microsimulation trace? (default is TRUE)
    # TS.out:  should the output include a transition array between states? (default is TRUE)
    # Trt:     are the n.i individuals receiving treatment? (scalar with a Boolean value, default is FALSE)
    # seed:    starting seed number for random number generator (default is 1)
    # Makes use of:
    # Probs:   function for the estimation of transition probabilities
    # Costs:   function for the estimation of cost state values
    # Effs:    function for the estimation of state specific health outcomes (QALYs)
    
    # update bootstrapped NHANES data in order to execute probs, cost and effects functions each cycle
    bootdata <- original_data %>%
      # create vector of starting states per individual according to their history of a diagnosis of any CVD conditions in bootstrapped NHANES data
      mutate(cvd              = ifelse((mi == 1 | stroke == 1 | chd == 1 | heart_failure == 1 | angina == 1),
                                       "History_CVD", "No_CVD"),  
             # duplicate age variable so that is can be updated with each cycle
             age_baseline     = age,
             # categorize age into specified categories each cycle that age is updated - Same categories used for cost and utility prediction as for weighting MI/stroke risk by ASCVD risk
             age_cat          = cut(age,                                 
                                    breaks = c(0, 24, 44, 64, Inf), 
                                    labels = c('<25', '25-44', '45-64', '65+')),
             mean_systolic_bp_orig = mean_systolic_bp,
             mean_systolic_bp = case_when(Trt == FALSE ~ mean_systolic_bp, 
                                          Trt == TRUE & medicaid_exp == 0 ~ mean_systolic_bp,
                                          Trt == TRUE & medicaid_exp == 1 & .bp == 0 ~ mean_systolic_bp, 
                                          # option to turn off effect of medicaid on bp in Trt scenario
                                          Trt == TRUE & medicaid_exp == 1 & .bp == 1 ~ (mean_systolic_bp - parameters[["sysbp_reduction"]]),
                                          TRUE ~ mean_systolic_bp), 
             # diagnosed or undiagnosed diabetes https://diabetes.org/diabetes/a1c/diagnosis
             diabetes_orig    = ifelse((diabetes_ever == 1 | (hba1c >= 6.5 & diabetes_ever == 0)), 1, 0),             
             hba1c = case_when(Trt == FALSE ~ hba1c, 
                               Trt == TRUE & medicaid_exp == 0 ~ hba1c,
                               # option to turn off effect of medicaid on hba1c in Trt scenario
                               Trt == TRUE & medicaid_exp == 1 & .hba1c == 0 ~ hba1c, 
                               # apply effect of medicaid on hba1c
                               Trt == TRUE & medicaid_exp == 1 & .hba1c == 1 ~ (hba1c - parameters[["hba1c_reduction"]]),
                               TRUE ~ hba1c),
             # diagnosed or undiagnosed diabetes https://diabetes.org/diabetes/a1c/diagnosis
             diabetes = ifelse((diabetes_ever == 1 | (hba1c >= 6.5 & diabetes_ever == 0)), 1, 0)
      )                                          
    
    # calculate the cost discount weight based on the discount rate d.c [using "0:n.t" assumes costs are accumulated at the START of each cycle - NB for half-cycle correction]
    v.dwc <- 1 / (1 + d.c) ^ (0:n.t)          
    # calculate the QALY discount weight based on the discount rate d.e [using "0:n.t" assumes utilities are accumulated at the START of each cycle - NB for half-cycle correction]
    v.dwe <- 1 / (1 + d.e) ^ (0:n.t)                                              
    
    # create the matrix capturing the state name/costs/health outcomes for all individuals at each time point 
    m.M <- m.C <- m.E <-  matrix(nrow = n.i, ncol = n.t + 1, 
                                 dimnames = list(paste("ind",   1:n.i, sep = " "), 
                                                 paste("cycle", 0:n.t, sep = " ")))  
    
    # estimate costs per individual of the initial health state
    costs <- Costs(mydata  = bootdata, 
                   cardiac_dysrhythmia = 0,      
                   peripheral_artery_disease = 0,
                   cci     = 1,
                   Trt     = Trt,
                   parameters = parameters)
    
    # estimate effects/QALYs per individual of the initial health state
    effects <- Effs(mydata = bootdata, 
                    cardiac_dysrhythmia = 0,      
                    peripheral_artery_disease = 0,
                    cci    = 1,
                    Trt    = Trt,
                    parameters = parameters) 
    
    # indicate the initial health states as a function of the distribution of MI or stroke (or CHD or heart failure) history
    m.M[, 1] <- bootdata$cvd  
    # assign initial costs to the m.C matrix
    m.C[, 1] <- costs$hc_cost       
    # assign effects/QALYs to the m.E matrix
    m.E[, 1] <- effects$utility                                                  
    
    # Beginning of time point loop
    for (t in 1:n.t) {
      
      bootdata <- bootdata %>%
        select(-age_cat) %>%
        # Update age to age + t with each cycle
        mutate(age            = age_baseline + t,  
               # categorize age into specified categories each cycle that age is updated - Same categories used for cost and utility prediction as for weighting MI/stroke risk by ASCVD risk
               age_cat        = cut(age,                                          
                                    breaks = c(0, 24, 44, 64, Inf), 
                                    labels = c('<25', '25-44', '45-64', '65+')),
               # Update both insurance variables used in the model such that all individual regardless of scenario receive Medicare at age 65
               insurance      = ifelse(age < 65, insurance,      "medicare"), 
               # no more effect of medicaid after age 65 (in particular on costs)
               medicaid_exp = ifelse(age < 65, medicaid_exp, 0),
               # diagnosed or undiagnosed diabetes https://diabetes.org/diabetes/a1c/diagnosis
               diabetes = ifelse((diabetes_ever == 1 | (hba1c >= 6.5 & diabetes_ever == 0)), 1, 0),
               # update bmi categories to correspond
               bmi_cat = case_when(bmi < 18.5               ~ "underweight", 
                                   (bmi >= 18.5 & bmi < 25) ~ "normal_weight",
                                   (bmi >= 25 & bmi < 30)   ~ "overweight",  
                                   bmi >= 30                ~ "obese",
                                   TRUE              ~ bmi_cat)
               )
               
      
      # Generate ASCVD weights each cycle using values from control group so that weights do not change between treatment scenarios [Otherwise artificially distorts results when BP is adjusted in treatment scenario]
      bootdata <- average_ascvd_risk(mydata = bootdata) 

      # Calculate transition probabilities for each cycle
      probs <- Probs(mydata = bootdata,
                     Trt    = Trt,
                     parameters = parameters,
                     seed_offset = seed_offset)
      
      # create vector of state transition probabilities as a function of an individual's characteristics
      # probability of dying from a MI
      p_MI_CVDdeath       <- matrix(probs$mi_mortality_risk,     1, n.i) 
      # probability of dying from a stroke
      p_stroke_CVDdeath   <- matrix(probs$stroke_mortality_risk, 1, n.i) 
      
      # probability of remaining in the dead from other causes state
      p_death_death       <- matrix(1, 1, n.i) 
      # probability of remaining in the dead from CVD state
      p_CVDdeath_CVDdeath <- matrix(1, 1, n.i)                                
      
      # probability of dying from other causes for those with a history of CVD
      p_death             <- matrix(probs$noncvd_mortality_risk, 1, n.i)
      # probability of having MI with a history of CVD
      p_CVD_MI            <- matrix(probs$mi_risk_CVD,           1, n.i)
      # probability of having stroke with a history of CVD
      p_CVD_stroke        <- matrix(probs$stroke_risk_CVD,       1, n.i)
      # probability of having a stroke with no history of CVD
      p_nCVD_MI           <- matrix(probs$mi_risk_noCVD,         1, n.i)
      # probability of having MI with a history of CVD
      p_nCVD_stroke       <- matrix(probs$stroke_risk_noCVD,     1, n.i)      
      
      # create a transition matrix as a function of an individual's characteristics for those in the "No_CVD" state
      m.p.it_nCVD <- matrix(NA, n.s, n.i)   
      # assign names to the vector
      rownames(m.p.it_nCVD) <- v.n                                                
      m.p.it_nCVD[1,] <-  NA # Residual; see function below
      m.p.it_nCVD[2,] <-  p_nCVD_MI * (1-p_death)
      m.p.it_nCVD[3,] <-  p_nCVD_stroke * (1-p_death)
      m.p.it_nCVD[4,] <-  0
      m.p.it_nCVD[5,] <-  0
      m.p.it_nCVD[6,] <-  p_death
      m.p.it_nCVD     <-  t(m.p.it_nCVD)
      
      # create a transition matrix as a function of an individual's characteristics for those in the "MI" state
      m.p.it_MI <- matrix(NA, n.s, n.i) 
      rownames(m.p.it_MI) <- v.n            
      m.p.it_MI[1,] <-  0
      m.p.it_MI[2,] <-  0
      m.p.it_MI[3,] <-  0
      m.p.it_MI[4,] <-  NA # Residual; see function below
      m.p.it_MI[5,] <-  p_MI_CVDdeath * (1-p_death)
      m.p.it_MI[6,] <-  p_death
      m.p.it_MI     <-  t(m.p.it_MI)
      
      # create a transition matrix as a function of an individual's characteristics for those in the "Stroke" state
      m.p.it_stroke <- matrix(NA, n.s, n.i)     
      rownames(m.p.it_stroke) <- v.n          
      m.p.it_stroke[1,] <-  0
      m.p.it_stroke[2,] <-  0
      m.p.it_stroke[3,] <-  0
      m.p.it_stroke[4,] <-  NA # Residual; see function below
      m.p.it_stroke[5,] <-  p_stroke_CVDdeath * (1-p_death)
      m.p.it_stroke[6,] <-  p_death
      m.p.it_stroke     <-  t(m.p.it_stroke)
      
      # create a transition matrix as a function of an individual's characteristics for those in the "History_CVD" state
      m.p.it_CVD <- matrix(NA, n.s, n.i)     
      rownames(m.p.it_CVD) <- v.n            
      m.p.it_CVD[1,] <-  0
      m.p.it_CVD[2,] <-  p_CVD_MI * (1-p_death)
      m.p.it_CVD[3,] <-  p_CVD_stroke * (1-p_death) 
      m.p.it_CVD[4,] <-  NA # Residual; see function below
      m.p.it_CVD[5,] <-  0
      m.p.it_CVD[6,] <-  p_death
      m.p.it_CVD     <-  t(m.p.it_CVD)
      
      # create a transition matrix as a function of an individual's characteristics for those in the "CVD_death" state
      m.p.it_CVD_death <- matrix(NA, n.s, n.i)     
      rownames(m.p.it_CVD_death) <- v.n            
      m.p.it_CVD_death[1,] <-  0
      m.p.it_CVD_death[2,] <-  0
      m.p.it_CVD_death[3,] <-  0
      m.p.it_CVD_death[4,] <-  0
      m.p.it_CVD_death[5,] <-  p_CVDdeath_CVDdeath
      m.p.it_CVD_death[6,] <-  0
      m.p.it_CVD_death     <-  t(m.p.it_CVD_death)
      
      # create a transition matrix as a function of an individual's characteristics for those in the "Non-CVD_death" state
      m.p.it_death <- matrix(NA, n.s, n.i)     
      rownames(m.p.it_death) <- v.n            
      m.p.it_death[1,] <-  0
      m.p.it_death[2,] <-  0
      m.p.it_death[3,] <-  0
      m.p.it_death[4,] <-  0
      m.p.it_death[5,] <-  0
      m.p.it_death[6,] <-  p_death_death
      m.p.it_death     <-  t(m.p.it_death)
      
      # Apply functions to each row to generate the residual transition probability after rescaling transition probabilities whose sum is greater than 1
      for (i in 1:n.i) {
        # Rescale probabilities
        m.p.it_nCVD[i, ]   <- rescale_to_one(m.p.it_nCVD[i, ])
        m.p.it_MI[i, ]     <- rescale_to_one(m.p.it_MI[i, ])
        m.p.it_stroke[i, ] <- rescale_to_one(m.p.it_stroke[i, ])
        m.p.it_CVD[i, ]    <- rescale_to_one(m.p.it_CVD[i, ])
      }
      
      for (i in 1:n.i) {
        # Estimate residual
        m.p.it_nCVD[i, ]   <- replace_blanks(m.p.it_nCVD[i, ]) 
        m.p.it_MI[i, ]     <- replace_blanks(m.p.it_MI[i, ])     
        m.p.it_stroke[i, ] <- replace_blanks(m.p.it_stroke[i, ]) 
        m.p.it_CVD[i, ]    <- replace_blanks(m.p.it_CVD[i, ])
      }
      
      # remove unused variables
      rm(p_MI_CVDdeath,
         p_stroke_CVDdeath,
         p_death_death,
         p_CVDdeath_CVDdeath,
         p_CVD_MI,           
         p_CVD_stroke,       
         p_death,       
         p_nCVD_MI,          
         p_nCVD_stroke)
      
      # Create empty matrix to fill with state transition probabilities
      m.p.it <- matrix(NA, nrow = n.i, ncol = n.s) 
      
      m.p.it <- case_when(
        m.M[, t] == "No_CVD" ~ m.p.it_nCVD,
        m.M[, t] == "MI" ~ m.p.it_MI,
        m.M[, t] == "Stroke" ~ m.p.it_stroke,
        m.M[, t] == "History_CVD" ~ m.p.it_CVD,
        m.M[, t] == "CVD_death" ~ m.p.it_CVD_death,
        m.M[, t] == "Non-CVD_death" ~ m.p.it_death,
        # If none of the conditions match, keep the original m.p.it
        TRUE ~ m.p.it  
      )
      
      rm(m.p.it_nCVD,
         m.p.it_MI,
         m.p.it_stroke,
         m.p.it_CVD,
         m.p.it_CVD_death,
         m.p.it_death)      
      
      if (.track_prob == 1) {
        # print the transition probabilities or produce an error if they don't sum to 1
        ifelse((rowSums(m.p.it) <= 1.0000009) & (rowSums(m.p.it) >= 0.9999990), print("Probabilities sum to 1"), print("ERROR: Probabilities DO NOT sum to 1")) 
      }
      
      # assign names to the vector
      colnames(m.p.it) <- v.n                                                     
      
      # These ensures the same sampling from transition matrices across scenarios, else (.fix_rcat != 1) then different sampling across scenarios
      if (.fix_rcat == 1) {
        set.seed(seed)
      } else {
        set.seed(seed + seed_offset)
      }

      # sample the next health state and store that state in matrix m.M
      m.M[, t + 1] <- rcat(n.i, m.p.it)                                          
      # Mapping values to characters
      m.M[, t + 1] <- v.n[match(m.M[, t + 1], 1:6)]      

      # Update bootdata to reflect CVD history
      bootdata <- bootdata %>%
        # Update the MI variable which calculate costs & utilities to reflect whether an individual has had a diagnosis of an MI
        mutate(mi      = ifelse(m.M[, t + 1] == "MI",      
                                1, mi),    
               # Update the Stroke variable which calculate costs & utilities to reflect whether an individual has had a diagnosis of an Stroke
               stroke  = ifelse(m.M[, t + 1] == "Stroke",  
                                1, stroke))
      
      parameters[["c_absenteeism"]] <- ifelse(bootdata$age >= 65, 0, parameters[["c_absenteeism"]])
      parameters[["c_disability"]]  <- ifelse(bootdata$age >= 65, 0, parameters[["c_disability"]])
      
      # estimate costs per individual during cycle t + 1
      costs <- Costs(mydata                    = bootdata,
                     cardiac_dysrhythmia       = 0,      
                     peripheral_artery_disease = 0,
                     cci                       = 1,
                     Trt                       = Trt,
                     parameters                = parameters)
      
      # estimate effects/QALYs per individual during cycle t + 1
      effects <- Effs(mydata                    = bootdata,
                      cardiac_dysrhythmia       = 0,      
                      peripheral_artery_disease = 0,
                      cci                       = 1,
                      Trt                       = Trt,
                      parameters                = parameters) 
      
      # assign costs during cycle t + 1 to the m.C matrix
      m.C[, t + 1] <- case_when((m.M[, t + 1] == "No_CVD") ~ 
                                  # Costs are as per individual annual HC costs if No CVD
                                  costs$hc_cost, 
                                
                                (m.M[, t + 1] == "Stroke") ~ 
                                  # Costs are as per individual annual HC costs if Stroke increased by 15% in year of event
                                  (costs$hc_cost * (1 + parameters[["delta_cost_y1"]])) + 
                                  # Workplace absenteeism costs in the year of the event
                                  ifelse(.perspective == "societal", parameters[["c_absenteeism"]], 0) +
                                  # Short-term disability costs in the year of the event
                                  ifelse(.perspective == "societal", parameters[["c_disability"]], 0), 
                                
                                (m.M[, t + 1] == "MI") ~ 
                                  # Costs are as per individual annual HC costs if Stroke increased by 15% in year of event
                                  (costs$hc_cost * (1 + parameters[["delta_cost_y1"]])) + 
                                  # Workplace absenteeism costs in the year of the event
                                  ifelse(.perspective == "societal", parameters[["c_absenteeism"]], 0) +
                                  # Short-term disability costs in the year of the event
                                  ifelse(.perspective == "societal", parameters[["c_disability"]], 0), 
                                
                                (m.M[, t + 1] == "History_CVD") ~ 
                                  # Costs are as per individual annual HC costs if Stroke or MI reduced by 15% in year after event
                                  (costs$hc_cost), 
                                
                                (m.M[, t + 1] == "CVD_death" & bootdata$age <=  65) ~ 
                                  # Lost earnings from premature a CVD death (i.e. prior to retirement)
                                  ifelse(.perspective == "societal", annual_earnings, 0), 
                                
                                (m.M[, t + 1] == "CVD_death" & bootdata$age > 65) ~ 
                                  # No lost earnings assumed from a CVD death after retirement
                                  0,    
                                
                                (m.M[, t + 1] == "Non-CVD_death" & bootdata$age <=  65) ~ 
                                  # Lost earnings from premature a Non-CVD death (i.e. prior to retirement)
                                  ifelse(.perspective == "societal", annual_earnings, 0), 
                                
                                (m.M[, t + 1] == "Non-CVD_death" & bootdata$age > 65) ~ 
                                  # No lost earnings assumed from a Non-CVD death after retirement
                                  0)                           
      
      # assign effects during cycle t + 1 to the m.E matrix
      m.E[, t + 1] <- case_when((m.M[, t + 1] == "No_CVD")        ~ effects$utility,
                                (m.M[, t + 1] == "Stroke")        ~ effects$utility, 
                                (m.M[, t + 1] == "MI")            ~ effects$utility, 
                                (m.M[, t + 1] == "History_CVD")   ~ effects$utility,
                                (m.M[, t + 1] == "CVD_death")     ~ 0,
                                (m.M[, t + 1] == "Non-CVD_death") ~ 0)
      
      # Option to switch on or off the half-cycle correction
      if ( .half == 1) {
        # apply half-cycle corrections across cycles
        m.C[, t] <- (m.C[, t + 1] + m.C[, t]) * half_cycle_correction         
        m.E[, t] <- (m.E[, t + 1] + m.E[, t]) * half_cycle_correction         
      }
      
      # to run model only for non-elderly adults, i.e. stop accumulating information at age 65
      m.C[, t + 1] <- case_when((bootdata$age < .trunc_age) ~ m.C[, t + 1],
                                TRUE ~ NA)
      m.E[, t + 1] <- case_when((bootdata$age < .trunc_age) ~ m.E[, t + 1],
                                TRUE ~ NA)
      
      # display the progress of the simulation
      if (.track_cycle == 1) {
        cat('\r', paste(round(t/n.t * 100), "% done", sep = " "))                   
      }
      
    } # close the loop for the time points 
    
    # Replace m.M to have the same NA values (for those older than 65 during the model) as m.C [could also use m.E]
    m.M[is.na(m.C)] <- NA
    # Replace m.M to have "Age_65" for same NA values (for those who turned 65 during the model) as m.C [could also use m.E]
    # Assuming .trunc_age is a variable with a numeric value
    if (!is.na(.trunc_age)) {
      age_string <- paste("Age_", .trunc_age, sep = "")
      m.M[is.na(m.C)] <- age_string
    } 
    
    # Calculate the number of person-years that accumulated for non-elderly adults [this will be less then n.t * n.i because the model stops accumulating information for individuals aged 65 or older]
    # Create a subset of v.n containing only the first four values
    # Can extend to all states depending on approach to counting person_years, e.g. is a year being dead (before age 65) included as a person_year?
    if (.excl_death == 1) {
      v.n_subset <- v.n[1:4]
    } else {
      v.n_subset <- v.n[1:6]
    }
    
    # Count occurrences for each row
    person_years <- apply(m.M, MARGIN = 1, FUN = function(row) sum(row %in% v.n_subset))
    
    # Replace NA values (for those older than 65 during the model) to 0
    m.C[is.na(m.C)] <- 0
    m.E[is.na(m.E)] <- 0
    
    # total discounted cost per individual
    tc <- m.C %*% v.dwc         
    # total discounted QALYs per individual
    te <- m.E %*% v.dwe                                                            
    
    # Count number of outcomes per person
    mi_count <- apply(m.M == "MI", 1, sum)
    stroke_count <- apply(m.M == "Stroke", 1, sum)
    CVD_death_count <- apply(m.M == "CVD_death", 1, function(row) if (sum(row) >= 1) 1 else 0)
    NonCVD_death_count <- apply(m.M == "Non-CVD_death", 1, function(row) if (sum(row) >= 1) 1 else 0)
    # https://www.ncbi.nlm.nih.gov/books/NBK430746/
    
    # create a  matrix of transitions across states
    if (TS.out == TRUE) {  
      # transitions from one state to the other
      TS <- paste(m.M, cbind(m.M[, -1], NA), sep = "->")                         
      TS <- matrix(TS, nrow = n.i)     
      # name the rows 
      rownames(TS) <- paste("Ind",   1:n.i, sep = " ")  
      # name the columns
      colnames(TS) <- paste("Cycle", 0:n.t, sep = " ")                            
    } else {
      TS <- NULL
    }
    
    # create a trace from the individual trajectories
    if (TR.out == TRUE) { 
      TR <- t(apply(m.M, 2, function(x) table(factor(x, levels = v.n, ordered = TRUE))))
      # create a distribution trace
      TR <- TR / n.i  
      # name the rows 
      rownames(TR) <- paste("Cycle", 0:n.t, sep = " ")  
      # name the columns 
      colnames(TR) <- v.n                                                         
    } else {
      TR <- NULL
    }
    
    # Modify the names based on the value of Trt
    trt_suffix <- ifelse(Trt, "_TRUE", "_FALSE")
    
    results <- list(
      m.M = m.M,
      m.C = m.C,
      m.E = m.E,
      tc = tc,
      te = te,
      person_years = person_years,
      mi_count = mi_count,
      stroke_count = stroke_count,
      CVD_death_count = CVD_death_count,
      NonCVD_death_count = NonCVD_death_count,
      TS = TS,
      TR = TR,
      bootdata = bootdata,
      parameters = parameters,
      probs = probs,
      m.p.it = m.p.it
    )
    
    # Modify the names in the results list
    names(results) <- paste(names(results), trt_suffix, sep = "")

    return(results)     
    
  }  # end of the MicroSim function  
  
  #### Microsimulation bootstrapping function ----
  
  # The Bootstrapping function which runs all scenarios across same randomly sampled NHANES dataset
  Bootstrap_msm <- function(original_data, 
                            n.i,
                            seed_offset) { 
    
    # Do this before Micro-simulation to have the same sample for treatment and control scenarios
    bootdata_original <- original_data %>%
      # categorize age into specified categories each cycle that age is updated - Same categories used for cost and utility prediction as for weighting MI/stroke risk by ASCVD risk
      mutate(age_cat = cut(age,                                            
                           breaks = c(0, 24, 44, 64, Inf), 
                           labels = c('<25', '25-44', '45-64', '65+')),
             # return variables to numeric for operation (following multiple imputation)
             stroke        = as.numeric(stroke),
             mi            = as.numeric(mi),
             heart_failure = as.numeric(heart_failure),
             chd           = as.numeric(chd),
             angina        = as.numeric(angina),
             diabetes_ever = as.numeric(diabetes_ever),
             treated       = as.numeric(treated),
             smoking       = as.numeric(smoking),
             # Create variable which identifies individuals who will receive Medicaid as part of the treatment scenario (under age 65 and family income less than 138% of the federal poverty line)
             medicaid_exp   = ifelse(insurance  == "uninsured" & fpl < .fpl & age < 65, 1, 0),
             # create a family income variable to correspond to MEPS cost and utility calculations and according to the poverty line [see, https://meps.ahrq.gov/survey_comp/hc_technical_notes.shtml]
             # created within funcion because it follows imputation
             fam_income = case_when(fpl <= 1     ~ "poor",                          
                                    fpl <= 1.25  ~ "near_poor",
                                    fpl <= 2     ~ "low",
                                    fpl <= 4     ~ "medium",
                                    fpl  > 4     ~ "high",
                                    TRUE         ~ NA),
             # same for bmi categories as family income
             bmi_cat = case_when(bmi < 18.5               ~ "underweight", 
                                 (bmi >= 18.5 & bmi < 25) ~ "normal_weight",
                                 (bmi >= 25 & bmi < 30)   ~ "overweight",  
                                 bmi >= 30                ~ "obese",
                                 TRUE              ~ NA) %>%
             set_variable_labels(fam_income = "Family Income")
      )
    
    # Filter the sample to run the model for certain populations
    bootdata_original <- bootdata_original %>%
      filter(# filter to just those receiving the treatment if .ATT == 1, otherwise the whole sample (i.e. NHANES participants who do and do not receive treatment in treatment scenario - medicaid expansion) is examined (NB for distributional analysis)
        medicaid_exp == 1 | .ATT != 1)
    
    # These ensure the same dataset across bootstrapped replications, else (.fix_slice_sample != 1) then different samples across replications
    if (.fix_slice_sample == 1) {
      set.seed(seed)
    } else {
      set.seed(seed + seed_offset)
    }

    # sample NHANES data with replacement using survey weights created across waves for fasting subsample
    bootdata_original <- slice_sample(bootdata_original, n = n.i, replace = TRUE, weight_by = WTSAF)

    # These ensure the same parameters across bootstrapped replications
    if (.Deterministic == 1) {
      set.seed(seed)
    } else {
      set.seed(seed + seed_offset)
    }
    
    # Create a list of (Deterministic or Probabilistic) parameters to be used in the simulation
    parameters <- generate_parameters(Deterministic = .Deterministic, ci2se = ci2se)
    
    #### Run the simulation ----
    #running the function several times to obtain CI
    sim_no_trt  <- MicroSim(bootdata_original
                            , n.i
                            , n.t
                            , v.n
                            , d.c
                            , d.e
                            , TR.out = TRUE
                            , TS.out = TRUE
                            , Trt = FALSE
                            , parameters
                            , seed_offset = seed_offset) # run for no treatment
    
    sim_trt     <- MicroSim(bootdata_original
                            , n.i
                            , n.t
                            , v.n
                            , d.c
                            , d.e
                            , TR.out = TRUE
                            , TS.out = TRUE
                            , Trt = TRUE
                            , parameters
                            , seed_offset = seed_offset)  # run for treatment
    
    #### Collate Data ----
    
    # Define the list of variables to be divided by person_years_ntrt
    variables_to_divide_ntrt <- c("mi_ntrt", "stroke_ntrt", "CVD_death_ntrt", "nCVD_death_ntrt")
    variables_to_divide_trt <- c("mi_trt", "stroke_trt", "CVD_death_trt", "nCVD_death_trt")

    combined_data <- data.frame(
      age_baseline      = c(sim_no_trt$bootdata$age_baseline),
      female            = c(sim_no_trt$bootdata$female),
      race              = c(sim_no_trt$bootdata$race),
      insurance         = c(sim_no_trt$bootdata$insurance),
      education         = c(sim_no_trt$bootdata$education),
      fam_income        = c(sim_no_trt$bootdata$fam_income),
      mi_ntrt           = c(sim_no_trt$mi_count),
      stroke_ntrt       = c(sim_no_trt$stroke_count),
      CVD_death_ntrt    = c(sim_no_trt$CVD_death_count),
      nCVD_death_ntrt   = c(sim_no_trt$NonCVD_death_count),
      person_years_ntrt = c(sim_no_trt$person_years),
      mi_trt            = c(sim_trt$mi_count),
      stroke_trt        = c(sim_trt$stroke_count),
      CVD_death_trt     = c(sim_trt$CVD_death_count),
      nCVD_death_trt    = c(sim_trt$NonCVD_death_count),
      person_years_trt  = c(sim_trt$person_years),
      tc_ntrt           = c(sim_no_trt$tc),
      tc_trt            = c(sim_trt$tc),
      te_ntrt           = c(sim_no_trt$te),
      te_trt            = c(sim_trt$te)
    ) %>%
      mutate(
        age_cat = cut(age_baseline, breaks = c(0, 24, 44, 64, Inf), labels = c('<25', '25-44', '45-64', '65+'))
      )
    
    if (.raw_data == 1) {
      raw_data <- c(sim_no_trt, sim_trt)
    } else {
      raw_data <- list()
    }

    if(.rate == 1) {
      combined_data <- combined_data %>%
        mutate(
          across(all_of(variables_to_divide_ntrt), ~ . / person_years_ntrt),
          across(all_of(variables_to_divide_trt), ~ . / person_years_trt),
        )
    }
    
    
    # Initialize an empty list to store results for each variable
    summary_list <- list()
    
    # Define a list of variables for which you want to calculate means
    variables_to_mean <- c("tc_ntrt", "tc_trt", 
                           "te_ntrt", "te_trt", 
                           "mi_ntrt", "stroke_ntrt", "CVD_death_ntrt", "nCVD_death_ntrt", 
                           "mi_trt", "stroke_trt", "CVD_death_trt", "nCVD_death_trt")
    
    # Loop through each variable
    for (var in .grouping_vars) {
      # Get unique categories within the current variable
      categories <- unique(combined_data %>% pull({{ var }}))
      
      # Initialize an empty list to store results for each category
      category_summary_list <- list()
      
        for (cat in categories) {
          # Calculate the group summary for the current category
          group_summary <- combined_data %>%
            filter(.data[[var]] == cat) %>%
            summarise(
              variable          = var
              , category        = cat
              , num_individuals = n()
              , across(all_of(variables_to_mean), ~ mean(., na.rm = TRUE))
            ) 
          
          # Append the summary for the current category to the list
          category_summary_list[[as.character(cat)]] <- group_summary
        }
      
      # Combine category summaries for the current variable
      var_summary <- do.call(rbind, category_summary_list)
      
      # Append the summary for the current variable to the main list
      summary_list[[var]] <- var_summary
    }
    
    # Combine all variable summaries into a single dataframe
    result <- do.call(rbind, summary_list)
    
    # Calculate the total summary for all categories within each variable
    total_summary <- combined_data %>%
      summarize(
        variable            = "Total"
        , category          = "Total"
        , num_individuals   = n()
        , across(all_of(variables_to_mean), ~ mean(., na.rm = TRUE))
      )
    
    # Append the total summary to the result
    combined_output <- bind_rows(result, total_summary)
    
    # relabel category for female/male
    combined_output <- combined_output %>%
      mutate(category = case_when(category == 1 ~ "female",
                                  category == 0 ~ "male",
                                  TRUE ~ as.character(category)))   
    
    msm_results <- list(
      raw_data
      , combined_output
      , parameters
      , sim_trt
      , sim_no_trt
      )
    
    return(msm_results)
    
  }
  
  #### Deterministic Sensitivity Analysis ----
  
  # Function to perform Bootstrap_msm with perturbed values
  dsa <- function(perturbed_values) {
    # Copy the perturbed values to avoid overwriting base_values
    modified_values    <- base_values
    modified_values[i] <- perturbed_values[i]
    
    delta_CVD_history$mean <<- perturbed_values[1]
    delta_hba1c$mean       <<- perturbed_values[2]
    delta_sys_bp$mean      <<- perturbed_values[3]
    odds_hc$mi$mean        <<- perturbed_values[4]
    odds_hc$stroke$mean    <<- perturbed_values[5]
    hc_cost$mi$mean        <<- perturbed_values[6]
    hc_cost$stroke$mean    <<- perturbed_values[7]
    cost_absenteeism$mean  <<- perturbed_values[8]
    cost_disability$mean   <<- perturbed_values[9]
    delta_c_total$mean     <<- perturbed_values[10]
    delta_c_y1$mean        <<- perturbed_values[11]
    c_admin_enrollee$mean  <<- perturbed_values[12]
    annual_earnings        <<- perturbed_values[13]
    u$mi$mean              <<- perturbed_values[14]
    u$stroke$mean          <<- perturbed_values[15]
    d.e                    <<- perturbed_values[16]
    d.c                    <<- perturbed_values[17]

    #### Bootstrap the simulation ----
    
    # Initialize an empty dataframe for the combined output
    dsa_combined_output <- data.frame()
    
    for (dataset_index in 1:m) {
      # Get the dataset for the current iteration
      current_data <- dataList[[paste("Dataset_", dataset_index, sep = "")]]
      
      for (replication in 1:bootsize) {
        set.seed(seed)
        
        #create concatenated variable of dataset_index and replication so that seed is different for all simulations (not just replications)
        concat <- as.numeric(paste(dataset_index, replication, sep = ""))
        
        # Run the Bootstrap_msm function for the current replication on the current dataset
        dsa_result <- Bootstrap_msm(original_data = current_data,
                                    n.i = n.i,
                                    seed_offset = concat)
        
        # Add replication and dataset index variables to the result data frame
        dsa_result[[2]]$replication   <- replication
        dsa_result[[2]]$dataset_index <- dataset_index
        
        # Append the results to the combined data frame
        dsa_combined_output <- bind_rows(dsa_combined_output, dsa_result[[2]])
        
      }
      
      # Display the progress of the simulation
      if (.track_boot == 1) {
        cat('\r', paste(round(replication/bootsize * 100), "% Bootstrap done for Dataset_", dataset_index, sep = " "))
      }
    }
    
    #### Summarize the results  ----
    
    # Define a list of variables for which you want to calculate means
    variables_to_operate <- c("tc_ntrt", "tc_trt", 
                              "te_ntrt", "te_trt")
    
    # Initialize an empty list to store summarised outputs
    dsa_impute_outputs <- list()
    
    for (imputation_index in 1:m) {
      # Create the summarised_output for the current dataset
      dsa_summarised_output <- dsa_combined_output %>%
        filter(imputation_index == .data$dataset_index,
               variable == "Total") %>%  # Filter data for the current dataset
        summarize(
           across(all_of(variables_to_operate), ~ weighted.mean(. , num_individuals, na.rm = TRUE))
          , n = sum(num_individuals)
        ) %>%
        ungroup()
      
      # Store the summarised_output in the list
      dsa_impute_outputs[[paste("summarised_output_", imputation_index, sep = "")]] <- dsa_summarised_output
    }
    
    # Combine all the summarised outputs into one dataset
    dsa_combined_summarised_output <- bind_rows(dsa_impute_outputs) 
    
    # Calculate the mean for selected variables
    dsa_summarised_output <- dsa_combined_summarised_output %>%
      summarize(# Use Rubins rules to combine se's across imputation results
        across(all_of(variables_to_operate), ~ weighted.mean(. , n, na.rm = TRUE))
        , num_individuals = sum(n)
      ) %>%
      ungroup()
    
    # If only one imputed dataset is used then need "IF" statement to calculate SE's
    if (length(dsa_combined_summarised_output$category[dsa_combined_summarised_output$category == "Total"]) == 1) 
    {
      dsa_summarised_output <- dsa_combined_summarised_output %>%
        mutate(num_individuals = n)
    }
    
    # Calculate mean costs and standard errors for the current category
    v.C_ntrt <- dsa_summarised_output$tc_ntrt
    v.C_trt <- dsa_summarised_output$tc_trt
    v.C <- c(v.C_ntrt, v.C_trt)
    
    # Calculate mean QALYs and standard errors for the current category
    v.E_ntrt <- dsa_summarised_output$te_ntrt
    v.E_trt <- dsa_summarised_output$te_trt
    v.E <- c(v.E_ntrt, v.E_trt)
    
    # calculate incremental costs
    delta.C <- v.C[2] - v.C[1]      
    # calculate incremental QALYs
    delta.E <- v.E[2] - v.E[1]    
    
    # calculate the INMB
    inhb_values <- delta.E - (delta.C / wtp)
    
    # Store the INMB for this perturbation
    result_list <- list(inhb_values = inhb_values, 
                        perturbed_values = modified_values)
    
    return(result_list)
  }
  
  #### Cost-effectiveness plane function ----
  
  create_ce_plane <- function(data, x_var, y_var, color_var, title, xlim) {
    ce_plane <- ggplot(data, aes(x = {{x_var}}, y = {{y_var}}, color = factor({{color_var}}))) +
      geom_point(size = 0.5) +
      scale_fill_viridis(discrete = TRUE) +
      geom_abline(intercept = 0, color = "black", slope = wtp, linetype = "dashed") +
      geom_hline(yintercept = 0, color = "grey", linetype = "solid") +
      geom_vline(xintercept = 0, color = "grey", linetype = "solid") +
      labs(title = title,
           x = "Incremental Effect",
           y = "Incremental Cost") +
      theme_minimal() +
      theme(legend.position = c(0.9, 0.9)) +
      coord_cartesian(xlim = xlim) 
    
    return(ce_plane)  # Return the plot object
  }  
  
  #### INHB Bar-chart function ----     
  
  create_grouped_bar_charts <- function(data, variable_values) {
    charts <- list()
    for (variable_value in variable_values) {
      filtered_data <- data %>% filter(variable == variable_value)
      
      # Reorder the 'category' variable based on 'inhb' in descending order
      filtered_data <- filtered_data %>%
        mutate(category = fct_reorder(category, -inhb, .fun = min))
      
      chart <- ggplot(filtered_data, aes(x = category, y = inhb, fill = category)) +
        geom_bar(stat = "identity", position = "dodge", show.legend = FALSE) +
        scale_y_continuous(labels = function(x) format(x, scientific = FALSE)) +  # Turn off scientific notation for y-axis
        labs(title = "",
             x = " ",
             y = "INHB (QALYs)",
             fill = "Category") +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
        scale_fill_viridis(discrete = TRUE) 
     # + coord_cartesian(ylim = c(-0.02, 0.04))  # Set y-axis limits
      
      charts[[variable_value]] <- chart
    }
    return(charts)
  }
  
  #### Costs, QALYs, and Net Health Benefit Bar-chart function ----
  
  create_bar_chart_se <- function(data, var_prefix) {

    # Reshape the data from wide to long format for the main data
    bar_data_long <- pivot_longer(
      data,
      cols = starts_with(var_prefix) & ends_with("trt"),
      names_to = "type",
      values_to = "value"
    ) %>%
      mutate(type = ifelse(str_detect(type, "ntrt$"), "No Medicaid Expansion", "Medicaid Expansion"))
    
    # Reshape the standard error data from wide to long format
    se_data_long <- pivot_longer(
      data,
      cols = starts_with(var_prefix) & ends_with("se"),
      names_to = "type",
      values_to = "se"
    ) %>%
      mutate(type = ifelse(str_detect(type, "ntrt_se$"), "No Medicaid Expansion", "Medicaid Expansion"))
    
    # Combine the main data and standard error data
    combined_data <- left_join(bar_data_long, se_data_long, by = c("variable", "category", "type"))
    
    # Your bar chart code with bars ordered by tc_ntrt within each facet and error bars
    p <- ggplot(combined_data, aes(x = reorder(category, -value), y = value, fill = type, group = interaction(category, type))) +
      geom_bar(stat = "identity", 
               position = position_dodge(width = 0.8), 
               width = 0.6) +
      geom_errorbar(aes(ymin = value - (1.96*se), ymax = value + (1.96*se)), 
                    position = position_dodge(width = 0.8),
                    width = 0.2, size = 0.7) +  # Add error bars 
      labs(x = " ", y = " ", fill = "Treatment Type") +
      scale_color_viridis(discrete = TRUE) +
      theme_minimal() +
      theme(legend.title = element_blank(),
            legend.position = "bottom",
            legend.justification = "bottom",
            legend.box.just = "center",
            legend.key.size = unit(0.75, "lines")) +
      theme(axis.text.x = element_text(angle = 45, hjust = 1)) +  # Rotate x-axis labels
      facet_wrap(~variable, scales = "free_x")
    
    # Set the plot title based on var_prefix
    p <- p + ggtitle(var_prefix)
    
    return(p)
  }
  
  #### Equally Distributed Equivalent Health function ----
  
  # The calculate_ede function using Atkinsons formula
  calculate_ede <- function(health_scores, 
                            epsilon, 
                            kp = FALSE,
                            population_size) {
    
    # Default index is the Atkinson index (set kp == TRUE for Kolm-Pollak index)
    # https://www.ncbi.nlm.nih.gov/pmc/articles/PMC7818272/
    if (kp == FALSE) {
      if (epsilon == 1) {
        ede <- log(health_scores)
      } else {
        ede_hscore <- case_when(health_scores >= 0 ~ (health_scores ^ (1 - epsilon)),
                                health_scores <  0 ~ ((sqrt(health_scores ^ 2)) ^ (1 - epsilon)) * -1,
                                TRUE ~ (health_scores^(1 - epsilon)))
        ede_sum <- (sum(ede_hscore * (population_size/sum(population_size))))
        ede <- (ede_sum^(1 / (1 - epsilon)))
      }
    } else {
      # Kolm-Pollak index
      # https://www.ncbi.nlm.nih.gov/pmc/articles/PMC7818272/
      ede_sum <- log((sum((exp(-(health_scores * epsilon))) * (population_size/sum(population_size)))))
      ede <- -(ede_sum * (1 / epsilon)) 
    }
    
    # Check if ede is not numeric (e.g., NaN or Inf)
    if (!is.numeric(ede) || is.na(ede) || !is.finite(ede)) {
      stop("EDEH calculation resulted in a non-numeric value.")
    }
    
    return(ede)
  }
  
} # Close "FUNCTIONS" section
##### DATA ----
# Set .table1 == 1 to create and save Table 1 with and without missing values
.table1 <- 0
# Set .counts == 1 to create weighted counts across population characteristics (necessary for EDEH by population)
.counts <- 1
{
  #### Directories, folders and output names ----
  
  setwd(wd)
  getwd()
  
  # Use create functions folder below to create folders when first starting project
  #create_folders()
  
  # set-up domains for loading/saving data
  data_domain   <- paste0(wd, "/data")
  graph_domain  <- paste0(wd, "/figures")
  table_domain  <- paste0(wd, "/tables")

  # Set the date scalar
  date_scalar <- format(Sys.Date(), "%y%m%d")
  
  # Combine all table filenames to be saved into a list
  table_filenames <- c("cea_table", "ir_table", "combined_output")
  
  # name the version of the model being run
  version <- paste0(.perspective, "_trunc", .trunc_age, "_ATT", .ATT,
                    "_", n.i, "_", n.t, "_", bootsize, "_", m,
                    "_Det", .Deterministic, "_rcat", .fix_rcat, "_slice", .fix_slice_sample,
                    "_seed", seed)

  # Create filename for combining tables and figures in one excel file
  results_tables          <- paste0("results_tables", "_", date_scalar,"_", version, ".xlsx")
  results_graphs          <- paste0("results_graphs", "_", date_scalar,"_", version, ".pptx")
  results_combined_output <- paste0("combined_output", "_", date_scalar,"_", version, ".rds")
  results_dsa             <- paste0("dsa_output", "_", date_scalar,"_", version, ".rds")
  
  #### Source data ----
  
  # Use this to bring in non-CVD mortality rates from the CDC
  # Define the filename
  filename <- "nonCVDmortality.R"
  
  # Combine the base path and filename to create the full file path
  full_file_path <- file.path(filename)
  source(full_file_path)
  
  # Use this to the nhanes dataset using the nhanesA package
  # Define the filename
  filename <- "nhanesA.R"
  
  # Combine the base path and filename to create the full file path
  full_file_path <- file.path(filename)
  # slow so dont do every time
  if (.nhanesA == 1) {
    source(full_file_path)
  }
  
  #### Load the data ----
  ## Load the .rds file
  nhanes_orig <- read_rds(here(data_domain, "appended_nhanes_data.rds")) %>%
    as_tibble() %>%
    select(age_years
           , stroke_dx
           , heart_attack_dx
           , chd_dx
           , cong_heart_fail_dx
           , angina_dx
           , race_5cat
           , ratio_poverty_inc
           , edu_lev_5cat
           , hdl_chol_mgdl
           , total_chol_mgdl
           , health_ins_cover
           , health_ins_private
           , medicare
           , medicaid
           , smoke_3cat
           , BMXBMI
           , meds_hyperten
           , bp_systolic
           , diabetes_dx
           , male
           , hba1c
           , wave_year
           , WTINT2YR
           , WTMEC2YR
           , WTSAF2YR
           , SDMVPSU
           , SDMVSTRA
           , SEQN) %>% 
    filter(wave_year == "J_17" | wave_year == "I_15" | wave_year == "H_13" | 
             wave_year == "G_11") %>%
    print(n = 10, width = Inf)
  
  #### New sample weights combined across waves ----
  # https://wwwn.cdc.gov/nchs/nhanes/tutorials/weighting.aspx
  
  # Calculate the number of years based on unique values in "wave_year"
  num_years <- length(unique(nhanes_orig$wave_year)) * 2
  
  # Loop through the variables you want to divide
  variables_to_divide <- c("WTINT2YR", "WTMEC2YR", "WTSAF2YR")
  
  for (variable in variables_to_divide) {
    # Create the new variable name with the updated number
    new_variable_name <- paste0(substr(variable, 1, 5))
    
    # Perform the transformation and assign the result
    nhanes_orig[[new_variable_name]] <- nhanes_orig[[variable]] / num_years
  }
  
  #### Format the data ----
  
  nhanes_count <- nhanes_orig %>% 
    # rename variables to correspond to those used in the functions
    rename(hdl_cholesterol   = hdl_chol_mgdl,                                     
           total_cholesterol = total_chol_mgdl,
           mean_systolic_bp  = bp_systolic,
           fpl               = ratio_poverty_inc,
           bmi               = BMXBMI
           ) %>%
    mutate(age           = as.numeric(age_years),                                
           stroke        = as.numeric(stroke_dx),                                
           mi            = as.numeric(heart_attack_dx),                          
           heart_failure = as.numeric(cong_heart_fail_dx),                       
           chd           = as.numeric(chd_dx), 
           angina        = as.numeric(angina_dx),
           diabetes_ever = as.numeric(diabetes_dx),                              
           treated       = as.numeric(meds_hyperten),                            
           hba1c         = as.numeric(hba1c),                                    
           female = if_else(male == 1, 0, 1), 
           #create single insurance variable to correspond to MEPS cost and utility calculations and for cleaner distinction of treatment scenario
           insurance = case_when(health_ins_private == "1"  ~ "private",          
                                 medicare           == "1"  ~ "medicare",
                                 medicaid           == "1"  ~ "medicaid",
                                 health_ins_cover   == "1"  ~ "other_plan",
                                 health_ins_cover   == "0"  ~ "uninsured",
                                 TRUE                       ~ NA),
           # create single insurance variable to correspond to MEPS cost and utility calculations and for cleaner distinction of treatment scenario
           education = case_when(edu_lev_5cat == "1"  ~ "no_degree",              
                                 edu_lev_5cat == "2"  ~ "no_degree",
                                 edu_lev_5cat == "3"  ~ "ged_hs",
                                 edu_lev_5cat == "4"  ~ "associate_bachelor",
                                 edu_lev_5cat == "5"  ~ "master_doctorate",
                                 TRUE                 ~ NA),
           # create single race variable to correspond to MEPS cost and utility calculations and for cleaner distinction of treatment scenario
           race = case_when(race_5cat == "1"  ~ "hispanic",                       
                            race_5cat == "2"  ~ "white",
                            race_5cat == "3"  ~ "black",
                            race_5cat == "4"  ~ "asian",
                            race_5cat == "5"  ~ "other_race",
                            TRUE              ~ NA),   
           # create "current smoker" (smoking) variable to correspond to MEPS cost and utility calculations and ASCVD calculation
           smoking = case_when(smoke_3cat         == "1"  ~ "1",                  
                               smoke_3cat         == "2"  ~ "0",
                               smoke_3cat         == "0"  ~ "0",
                               TRUE                       ~ NA),
           smoking = as.numeric(smoking),
           education           = `attr<-`(education        , "label", "Education")
           , smoking           = `attr<-`(smoking          , "label", "Current Smoker")
           , bmi               = `attr<-`(bmi              , "label", "BMI (kg/m2)")
           , fpl               = `attr<-`(fpl              , "label", "Raio of Family Income to FPL")
           , age               = `attr<-`(age              , "label", "Age (yrs)")
           , stroke            = `attr<-`(stroke           , "label", "Stroke")
           , mi                = `attr<-`(mi               , "label", "Myocardial Infarction")
           , heart_failure     = `attr<-`(heart_failure    , "label", "Heart Failure")
           , chd               = `attr<-`(chd              , "label", "Coronary Heart Disease")
           , angina            = `attr<-`(angina           , "label", "Angina Pectoris")
           , diabetes_ever     = `attr<-`(diabetes_ever    , "label", "Diabetes")
           , treated           = `attr<-`(treated          , "label", "Treated for Hypertension")
           , hba1c             = `attr<-`(hba1c            , "label", "HbA1c (%)")
           , female            = `attr<-`(female           , "label", "Female (Y/N)")
           , insurance         = `attr<-`(insurance        , "label", "Insurance Status")
           , hdl_cholesterol   = `attr<-`(hdl_cholesterol  , "label", "HDL Cholesterol (mg/dL)")
           , total_cholesterol = `attr<-`(total_cholesterol, "label", "Total Cholesterol (mg/dL)")
           , mean_systolic_bp  = `attr<-`(mean_systolic_bp , "label", "Systolic Blood Pressue (mmHg)")
           ) %>%
    # drop unused variables after formatting
    select(-male
           , -age_years
           , -smoke_3cat
           , -stroke_dx
           , -angina_dx
           , -heart_attack_dx
           , -cong_heart_fail_dx
           , -chd_dx
           , -diabetes_dx
           , -meds_hyperten
           , -race_5cat
           , -edu_lev_5cat
           , -health_ins_cover
           , -health_ins_private
           , -medicare
           , -medicaid
           , -WTINT2YR
           , -WTINT
           , -WTMEC2YR
           , -WTMEC) %>%                                                              
    print(n = 10, width = Inf)       
  
  nhanes <- nhanes_count %>%
  select(-WTSAF2YR
          , -SDMVPSU
          , -SDMVSTRA)
  
  #### Table 1 ----
  
  if (.table1 == 1) {
    # Use this to bring in non-CVD mortality rates from the CDC
    # Define the filename
    filename <- "Table1.R"
    
    # Combine the base path and filename to create the full file path
    full_file_path <- file.path(filename)
    
    # Source code for conducting MI
    source(full_file_path)
  }
  
} # Close "DATA" section
##### MODEL ----
  {
  # Set timer == 1 in "Scenario Analysis" to track model run-time
  if (.timer == 1) {
    # Start timing
    start_time <- Sys.time()
  }
  #### Examine Missingness ----
    # check missing values across variables
    miss_plot_full <- nhanes %>%
      ggmiss(.)
    
    miss_map_full <-  nhanes %>%
      vis_miss(.)
  
  {
    ### FLOW DIAGRAM DATA ----
      {
        ### SECTION 1 ----
        # number of participants per wave
        length(nhanes$SEQN[nhanes$wave_year == "G_11"])
        length(nhanes$SEQN[nhanes$wave_year == "H_13"])
        length(nhanes$SEQN[nhanes$wave_year == "I_15"])
        length(nhanes$SEQN[nhanes$wave_year == "J_17"])
        
        # number of participants across waves = nhanes
        length(nhanes$SEQN)
        
        # number of participants without fasting blood sample weights
        sum(is.na(nhanes$WTSAF))
        
        # proportion of participants without fasting blood sample weights
        sum(is.na(nhanes$WTSAF))/length(nhanes$SEQN)
        
        ### SECTION 2 ----
        # remove those without fasting blood sample weights
        nhanes_filter <- nhanes %>%
          filter(!is.na(WTSAF))
        
        # number of participants remaining = nhanes_filter
        length(nhanes_filter$SEQN)
        
        # number of participants age 18 years or younger
        length(nhanes_filter$SEQN[nhanes_filter$age < .start_age])
        
        # corresponding proportion
        length(nhanes_filter$SEQN[nhanes_filter$age < .start_age]) / length(nhanes_filter$SEQN)
        
        # number of participants age 65 years or older
        length(nhanes_filter$SEQN[nhanes_filter$age > .stop_age])
        
        # corresponding proportion
        length(nhanes_filter$SEQN[nhanes_filter$age > .stop_age]) / length(nhanes_filter$SEQN)
        
        ### SECTION 3 ----
        # remove those older and younger
        nhanes_filter <- nhanes_filter %>%
          filter(age <= .stop_age, # non-elderly adults
                 age >= .start_age)
        # number of participants remaining = nhanes_filter
        length(nhanes_filter$SEQN)
        
        ### SECTION 4 ----
        # Present complete-case data
        # Function to calculate missing values and percentages
        missing_summary <- function(x) {
          missing_count <- sum(is.na(x))
          missing_percentage <- (missing_count / length(x)) * 100
          return(c(missing_count, missing_percentage))
        }
        # Calculate and print missing values and percentages
        missing_summary_df <- nhanes_filter %>%
          summarise(across(everything(), ~ missing_summary(.))) %>%
          t() %>%
          as.data.frame()
        colnames(missing_summary_df) <- c("Missing Count", "Missing Percentage")
        
        # Calculate the total number of observations with missing values
        total_missing_observations <- sum(apply(nhanes_filter, 1, anyNA))
        total_missing_percentage <- (total_missing_observations / nrow(nhanes_filter)) * 100
        
        # Add total missing observations and proportion to the data frame
        total_missing <- c(total_missing_observations, total_missing_percentage)
        missing_summary_df <- rbind(missing_summary_df, total_missing)
        rownames(missing_summary_df)[nrow(missing_summary_df)] <- "Total"
        
        # Assuming missing_summary_df is your data frame
        # To sort in descending order, you can use the 'decreasing' argument
        sorted_missing_summary_df <- missing_summary_df[order(missing_summary_df$`Missing Count`, decreasing = TRUE), ]
        print(sorted_missing_summary_df)
        
        # plot missingness for all
        miss_plot_nhanes_filter <- ggmiss(nhanes_filter)
        print(miss_plot_nhanes_filter)
        
        # plot missingness across all
        miss_map_nhanes_filter <- vis_miss(nhanes_filter)
        print(miss_map_nhanes_filter)
        
        # plot missingness for only variables with missing observations and remove double counting from derived variables, e.g. bmi and bmi_cat
        miss_plot_final <- nhanes %>%
          # filter to relevant sample
          filter(!is.na(WTSAF), # no missing fasting weights
                 age <= .stop_age, # non-elderly adults
                 age >= .start_age) %>%
          # remove irrelevant variables or those with 0 % missingness
          select(-WTSAF,
                 -race,
                 -female,
                 -age,
                 -SEQN,
                 -wave_year,
          ) %>%
          ggmiss()
        
        ### SECTION 5 ----
        # remove missing values
        nhanes_miss <- nhanes_filter %>%
          na.omit() %>%
          mutate(m_exp = case_when((insurance == "uninsured" & fpl < 1.38) ~ 1,
                                   (insurance != "uninsured" & !is.na(insurance)) | (fpl >= 1.38 & !is.na(fpl)) ~ 0,
                                   TRUE ~ NA))
        
        # count after missing values removed
        length(nhanes_miss$SEQN)
        
        # count eligble for medicaid
        table(nhanes_miss$m_exp)
        # proportion eligble for medicaid
        prop_mexp <- sum(nhanes_miss$m_exp[nhanes_miss$m_exp==1]) / length(nhanes_miss$m_exp[!is.na(nhanes_miss$m_exp)])
        print(prop_mexp)
        print(1-prop_mexp)
        
        #remove unnecessary data
        rm(nhanes_miss)
      }
    }
    
  #### Multiple Imputation ----
  
  #re-run multiple imputation or use previously saved MI datasets
  if (.mi == 1) {
    # Use this to bring in non-CVD mortality rates from the CDC
    # Define the filename
    filename <- "Imputation.R"
    
    # Combine the base path and filename to create the full file path
    full_file_path <- file.path(filename)
    
    # Source code for conducting MI
    source(full_file_path)
    
  } else {
    # Load MI datasets
    dataList <- readRDS(paste0(data_domain, "/dataList.rds"))
  }
  
  #### Weighted counts ----
    if (.counts == 1) {
      
      .count_data <- dataList$Dataset_10
      
      # Use this to bring in non-CVD mortality rates from the CDC
      # Define the filename
      filename <- "Weighted_counts.R"
      
      # Combine the base path and filename to create the full file path
      full_file_path <- file.path(filename)
      
      # Source code for conducting MI
      source(full_file_path)
    }
    
  #### Bootstrap the simulation ----
  
  # Initialize an empty dataframe for the combined output
  combined_output <- data.frame()
  
  # Initialize an empty list for raw data
  raw_data <- list()
  
  for (dataset_index in 1:m) {
    # Get the dataset for the current iteration
    current_data <- dataList[[paste("Dataset_", dataset_index, sep = "")]]
    
    for (replication in 1:bootsize) {
    set.seed(seed)
      
      #create concatenated variable of dataset_index and replication so that seed is different for all simulations (not just replications)
      concat <- as.numeric(paste(dataset_index, replication, sep = ""))
      
      # Run the Bootstrap_msm function for the current replication on the current dataset
      result <- Bootstrap_msm(original_data = current_data,
                              n.i = n.i,
                              seed_offset = concat)
      
      # Add replication and dataset index variables to the result data frame
      result[[2]]$replication   <- replication
      result[[2]]$dataset_index <- dataset_index
      
      # Append the results to the combined data frame
      combined_output <- bind_rows(combined_output, result[[2]])
      
      if (.raw_data == 1) {
      # Add replication and dataset index variables to the result data frame
      result[[1]]$replication   <- replication
      result[[1]]$dataset_index <- dataset_index
      
      # Append the results to the raw data list
      raw_data <- c(raw_data, list(result[[1]]))
      }
      
        # Display the progress of the simulation
        if (.track_boot == 1) {
          cat('\r', paste(round(replication/bootsize * 100), "% Bootstrap done for Dataset_", dataset_index, sep = " "))
        }
    }
  }
  
  #### Summarize the results  ----
  
  # Define a list of variables for which you want to calculate means
  variables_to_mean <- c("tc_ntrt", "tc_trt", 
                         "te_ntrt", "te_trt", 
                         "mi_ntrt", "stroke_ntrt", "CVD_death_ntrt", "nCVD_death_ntrt", 
                         "mi_trt", "stroke_trt", "CVD_death_trt", "nCVD_death_trt",
                         "diff_mi", "diff_stroke", "diff_CVDdeath", "diff_nCVDdeath",
                         "inmb", "inhb", "icer")

  # Initialize an empty list to store summarised outputs
  impute_outputs <- list()
  
  # Across a large number of iterations, it is possible to get "Inf" values so
  # Columns to replace 'Inf' values
  cols_to_check <- c(
    "mi_ntrt", "stroke_ntrt", "CVD_death_ntrt", "nCVD_death_ntrt",
    "mi_trt", "stroke_trt", "CVD_death_trt", "nCVD_death_trt"
  )
  
  # Loop through specified columns and replace 'Inf' values with NA
  for (col in cols_to_check) {
    combined_output[[col]][is.infinite(combined_output[[col]])] <- NA
  }

  for (imputation_index in 1:m) {
    # Create the summarised_output for the current dataset
    summarised_output <- combined_output %>%
      filter(imputation_index == .data$dataset_index) %>%  # Filter data for the current dataset
      mutate(  diff_mi        = mi_trt - mi_ntrt
             , diff_stroke    = stroke_trt - stroke_ntrt
             , diff_CVDdeath  = CVD_death_trt - CVD_death_ntrt
             , diff_nCVDdeath = nCVD_death_trt - nCVD_death_ntrt
             , icer           = (tc_trt - tc_ntrt) / (te_trt - te_ntrt)
             , inhb           = (te_trt - te_ntrt) - ((tc_trt - tc_ntrt) / wtp)
             , inmb           = ((te_trt - te_ntrt) * wtp) - (tc_trt - tc_ntrt)) %>%
      group_by(variable, category) %>%
      summarize(
          across(all_of(variables_to_mean), ~ weighted.se.mean(., num_individuals, na.rm = TRUE), .names = "{.col}_se")
        , across(all_of(variables_to_mean), ~ weighted.mean(., num_individuals, na.rm = TRUE))
        , n = sum(num_individuals)
      ) %>%
      ungroup()
    
    # Store the summarised_output in the list
    impute_outputs[[paste("summarised_output_", imputation_index, sep = "")]] <- summarised_output
  }
  
  # Combine all the summarised outputs into one dataset
  combined_summarised_output <- bind_rows(impute_outputs) 
  
  # Calculate the mean for selected variables
  summarised_output <- combined_summarised_output %>%
    group_by(variable, category) %>%
    summarize(
      # Use Rubin's rules to combine SEs across imputation results
      across(all_of(variables_to_mean),
             ~ rubin_se(.x, get(paste0(cur_column(), "_se"))), .names = "{.col}_se"),
      across(all_of(variables_to_mean), ~ weighted.mean(., n, na.rm = TRUE)),
      num_individuals = sum(n)
    ) %>%
    ungroup()
  
  # If only one imputed dataset is used then need "IF" statement to calculate SE's
  if (length(combined_summarised_output$category[combined_summarised_output$category == "Total"]) == 1) 
  {
    summarised_output <- combined_summarised_output %>%
      mutate(num_individuals = n)
  }
    
  #### Population counts ----
  
  # create weighted population counts to merge
  pop_counts <- weighted_count_data_pe %>%
    as_tibble() %>%
    rename(category = variables,
           pop_counts = weighted_count_pe) %>%
    mutate(category = case_when(  category == "total"               ~ "Total"
                                , category == "inc_poor"               ~ "poor"
                                , category == "inc_near_poor"          ~ "near_poor"
                                , category == "inc_lower"              ~ "low"
                                , category == "inc_medium"             ~ "medium"
                                , category == "inc_high"               ~ "high"
                                , category == "race_white"             ~ "white"
                                , category == "race_black"             ~ "black"
                                , category == "race_asian"             ~ "asian"
                                , category == "race_hispanic"          ~ "hispanic"
                                , category == "race_other"             ~ "other_race"
                                , category == "edu_no_degree"          ~ "no_degree"
                                , category == "edu_ged_hs"             ~ "ged_hs"
                                , category == "edu_associate_bachelor" ~ "associate_bachelor"
                                , category == "edu_master_doctorate"   ~ "master_doctorate"
                                , category == "age_24minus"            ~ "<25"
                                , category == "age_25to44"             ~ "25-44"
                                , category == "age_45to64"             ~ "45-64"),
           category = as_factor(category)) 

  #### Cost-effectiveness analysis ----
  
  # Merge pop. counts
  cea_output <- summarised_output %>%
    left_join(., pop_counts, by = "category")
  
  # Initialize an empty list to store individual table_micro data frames
  table_list <- list()
  
  # Unique category values
  unique_categories <- unique(cea_output$category)
  
  # Loop through each category
  for (category in unique_categories) {
    
    # Subset data for the current category
    category_data <- cea_output[cea_output$category == category, ]
    
    # Calculate mean costs and standard errors for the current category
    v.C_ntrt <- category_data$tc_ntrt
    v.C_trt <- category_data$tc_trt
    v.C <- c(v.C_ntrt, v.C_trt)
    
    se.C_ntrt <- category_data$tc_ntrt_se
    se.C_trt <- category_data$tc_trt_se
    se.C <- c(se.C_ntrt, se.C_trt)
    
    # Calculate mean QALYs and standard errors for the current category
    v.E_ntrt <- category_data$te_ntrt
    v.E_trt <- category_data$te_trt
    v.E <- c(v.E_ntrt, v.E_trt)
    
    se.E_ntrt <-  category_data$te_ntrt_se
    se.E_trt <- category_data$te_trt_se
    se.E <- c(se.E_ntrt, se.E_trt)
    
    # Calculate incremental costs
    delta.C <- v.C_trt - v.C_ntrt
    delta.E <- v.E_trt - v.E_ntrt
    
    # Calculate net health benefits
    nhb_ntrt <- v.E_ntrt - (v.C_ntrt / wtp)
    nhb_trt  <- v.E_trt - (v.C_trt / wtp)
    
    # calculate the ICER
    ICER    <- delta.C / delta.E  
    # calculate the INMB
    INMB    <- (delta.E * wtp) - delta.C
    inmb_alt   <- category_data$inmb
    inmb_alt_se   <- category_data$inmb_se
    
    # calculate the INHB
    INHB    <- delta.E - (delta.C / wtp)
    inhb_alt   <- category_data$inhb
    inhb_alt_se   <- category_data$inhb_se
    
    n <- category_data$num_individuals
    N <- category_data$pop_counts
    
    # Create full incremental cost-effectiveness analysis table
    cea_table_micro <- data.frame(
      c(category, "", ""),
      c(round(v.C, 0),  ""),                # costs per arm
      c(round(se.C, 0), ""),                # MCSE for costs
      c(round(v.E, 5),  ""),                # health outcomes per arm
      c(round(se.E, 5), ""),                # MCSE for health outcomes
      c("", round(delta.C, 0),   ""),       # incremental costs
      c("", round(delta.E, 5),   ""),       # incremental QALYs 
      c("", "", ""),
      c(round(nhb_ntrt, 5), round(nhb_trt, 5),   ""),      # net health benefits
      c("", "", ""),
      c("", round(ICER, 0),       ""),       # ICER
      c("", round(INMB, 0),       ""),       # INMB
      c("", round(inmb_alt, 0),       ""),       # INMB
      c("", round(inmb_alt_se, 0),       ""),   # INMB
      c("", round(INHB, 5),       ""),       # INHB
      c("", round(inhb_alt, 5),       ""),       # INMB
      c("", round(inhb_alt_se, 5),       ""),   # INMB
      c("", "", ""),
      c("", round(n, 0), ""),  # Number of individuals
      c("", round(N, 0), "")   # Pop. size
    )
    
    # Inside the loop, add the table_micro to the list
    table_list[[category]] <- cea_table_micro
    
    # After the loop, combine all the individual table_micro data frames into one
    cea_table <- do.call(rbind, table_list)
    
    # name the columns
    colnames(cea_table) <- c("Category",
                             "Costs", "se",  
                             "QALYs", "se", 
                             "Incremental Costs", 
                             "QALYs Gained", 
                             " ",
                             "Net Health Benefit",
                             " ",
                             "ICER", 
                             "INMB", "INMB_alt", "INMB_alt_se",
                             "INHB", "INHB_alt", "INHB_alt_se",
                             " ",
                             "n",
                             "N") 
    
  }
  
  #### Incidence rates  ----
  
  # Initialize an empty list to store individual table_micro data frames
  table_list <- list()
  
  # Unique category values
  unique_categories <- unique(summarised_output$category)
  
  # Loop through each category
  for (category in unique_categories) {
    
    # Subset data for the current category
    category_data <- summarised_output[summarised_output$category == category, ]
    
    # Incidence of model outcomes
    mi_ir     <- c(category_data$mi_ntrt, 
                   category_data$mi_trt, 
                   category_data$diff_mi)
    mi_ir_se  <- c(category_data$mi_ntrt_se, 
                   category_data$mi_trt_se, 
                   category_data$diff_mi_se)

    stroke_ir     <- c(category_data$stroke_ntrt, 
                       category_data$stroke_trt, 
                       category_data$diff_stroke)
    stroke_ir_se  <- c(category_data$stroke_ntrt_se, 
                       category_data$stroke_trt_se, 
                       category_data$diff_stroke_se)

    cvd_death_ir     <- c(category_data$CVD_death_ntrt, 
                          category_data$CVD_death_trt, 
                          category_data$diff_CVDdeath)
    cvd_death_ir_se  <- c(category_data$CVD_death_ntrt_se, 
                          category_data$CVD_death_trt_se, 
                          category_data$diff_CVDdeath_se)

    ncvd_death_ir     <- c(category_data$nCVD_death_ntrt, 
                           category_data$nCVD_death_trt,
                           category_data$diff_nCVDdeath)
    ncvd_death_ir_se  <- c(category_data$nCVD_death_ntrt_se, 
                           category_data$nCVD_death_trt_se,
                           category_data$diff_nCVDdeath_se)

    n <- category_data$num_individuals

    # Create full incremental cost-effectiveness analysis table
    ir_table_micro <- data.frame(
      c(category, "", "Difference", ""),
      c(round(mi_ir      * .ir_py, 5), ""),       
      c(round(mi_ir_se   * .ir_py, 5), ""),  
      c("", "", "", ""),
      c(round(stroke_ir      * .ir_py, 5), ""),   
      c(round(stroke_ir_se   * .ir_py, 5), ""),
      c("", "", "", ""),
      c(round(cvd_death_ir     * .ir_py, 5), ""),   
      c(round(cvd_death_ir_se  * .ir_py, 5), ""),
      c("", "", "", ""),
      c(round(ncvd_death_ir     * .ir_py, 5), ""),   
      c(round(ncvd_death_ir_se  * .ir_py, 5), ""),
      c("", "", "", ""),
      c(round(n, 0), "", "", "")  # Number of individuals
    )
    
    # Inside the loop, add the table_micro to the list
    table_list[[category]] <- ir_table_micro
    
    # After the loop, combine all the individual table_micro data frames into one
    ir_table <- do.call(rbind, table_list)
    
    # name the columns
    colnames(ir_table) <- c("Category",
                            "IR MI",
                            "se",
                            " ",
                            "IR Stroke",
                            "se",
                            " ",
                            "IR CVD-death",
                            "se",
                            " ",
                            "IR Non-CVD_death",
                            "se",
                            " ",
                            "N"
    ) 
    
  }
  
  # To estimate incidence rates...         
  # https://www.ncbi.nlm.nih.gov/books/NBK430746/
  
  # To compare results with MI and stroke statistics... 
  # https://doi.org/10.1161/CIR.0000000000001052
  
  
  #### Deterministic One-way Sensitivity Analysis ----
  
  # Global assignment change (<<-) is required to make this function work
  
  if (.Deterministic == 1 & 
      .one_way == 1 & 
      .fix_rcat == 1 &
      .fix_slice_sample == 1) {
    
    # List of variable names and corresponding base values
    variable_names <- c(
      "delta_CVD_history$mean",
      "delta_hba1c$mean",
      "delta_sys_bp$mean",
      "odds_hc$mi$mean",
      "odds_hc$stroke$mean",
      "hc_cost$mi$mean",
      "hc_cost$stroke$mean",
      "cost_absenteeism$mean",
      "cost_disability$mean",
      "delta_c_total$mean",
      "delta_c_y1$mean",
      "c_admin_enrollee$mean",
      "annual_earnings",
      "u$mi$mean",
      "u$stroke$mean",
      "d.e",
      "d.c"
    )  # Add your variable names
    
    # List of variable names and corresponding base values
    plot_names <- c(
      "CVD history risk mutliplier",
      "Delta non-CVD mortality",
      "Delta HbA1c",
      "Delta systolic BP",
      "Odds non-zero HC-MI",
      "Odds non-zero HC-Stroke",
      "Non-zero HC-MI",
      "Non-zero HC-Stroke",
      "Absenteeism Cost",
      "Disability Cost",
      "Delta HC-Medicaid",
      "Delta y1 HC",
      "Medcaid admin cost/enrollee",
      "Annual earnings",
      "Utility-MI",
      "Utility-Stroke",
      "Discount rate - Utility",
      "Discount rate - Costs"
    )  # Add your variable names
    
    # Define the base_values vector at the start of your script
    base_values <- c(
      delta_CVD_history$mean
      , delta_hba1c$mean
      , delta_sys_bp$mean
      , odds_hc$mi$mean
      , odds_hc$stroke$mean
      , hc_cost$mi$mean
      , hc_cost$stroke$mean
      , cost_absenteeism$mean
      , cost_disability$mean
      , delta_c_total$mean
      , delta_c_y1$mean
      , c_admin_enrollee$mean
      , annual_earnings
      , u$mi$mean
      , u$stroke$mean
      , d.e
      , d.c
    )  # Define your base values
    
    # Initialize a list to store results
    results <- list()
    
    # Loop through each variable
    for (i in seq_along(variable_names)) {
      perturbation_results <- list()  # Initialize a list to store perturbation results
      for (perturbation in c(1 - delta_one_way, 1 + delta_one_way)) {
        perturbed_values <- base_values  # Reset perturbed_values to base_values for each perturbation
        perturbed_values[i] <- base_values[i] * perturbation
        
        result <- dsa(perturbed_values)
        
        # Store the inmb_values for this perturbation
        perturbation_results[[paste(variable_names[i], "Perturbation", perturbation, sep = "_")]] <- result
      }
      
      # display the progress of the DSA
      if (.track_dsa == 1) {
        cat('\r', paste(round(i/length(variable_names) * 100), "% DSA done", sep = " "))                
      }
      
      # Store the perturbation_results for this variable
      results[[variable_names[i]]] <- perturbation_results
    }
    
    # Save the list to a file
    saveRDS(results, file = (here(data_domain, results_dsa)))

    # Calculate INHB values for each variable
    INHB_values <- sapply(variable_names, function(var_name) {
      perturbation_results <- results[[var_name]]
      perturbed_values <- sapply(perturbation_results, function(result) result$inhb_values)
      max_inhb <- max(perturbed_values)
      min_inhb <- min(perturbed_values)
      return(max_inhb - min_inhb)
    })

    # Order variables based on the largest differences
    ordered_variables <- variable_names[order(INHB_values, decreasing = TRUE)]
    
    # Create a data frame for plotting without changing factor levels initially
    tornado_data <- data.frame(
      Variable = factor(ordered_variables, levels = ordered_variables),
      Lower_INHB = sapply(ordered_variables, function(var_name) {
        perturbation_results <- results[[var_name]]
        perturbed_values <- sapply(perturbation_results, function(result) result$inhb_values)
        min_inhb <- min(perturbed_values)
        return(min_inhb)
      }),
      Upper_INHB = sapply(ordered_variables, function(var_name) {
        perturbation_results <- results[[var_name]]
        perturbed_values <- sapply(perturbation_results, function(result) result$inhb_values)
        max_inhb <- max(perturbed_values)
        return(max_inhb)
      })
    )
    
    
    # Calculate the average INMB
    average_INHB <- summarised_output$inhb[summarised_output$category == "Total"]
    
    # Filter the top 5 most impactful variables
    #top_5_variables <- head(ordered_variables, 5)
    #tornado_data <- tornado_data[tornado_data$Variable %in% top_5_variables, ]
    
    # Create a factor with plot_names labels maintaining the original order
    ordered_plot_names <- plot_names[match(ordered_variables, variable_names)]
    tornado_data$Variable <- factor(tornado_data$Variable, levels = unique(ordered_variables))
    levels(tornado_data$Variable) <- ordered_plot_names
    
    # Create the tornado diagram plot with the correct order and labels
    tornado_plot <- ggplot(tornado_data, aes(x = fct_rev(Variable), ymin = Lower_INHB, ymax = Upper_INHB)) +
      geom_linerange(position = position_dodge(width = 0.9), color = "blue") +
      geom_hline(yintercept = average_INHB, linetype = "dashed", color = "red") +
      geom_text(aes(x = 1, y = average_INHB, label = paste("Baseline INHB:", round(average_INHB))), 
                vjust = 3, hjust = 1.05, color = "black", size = 4) +  # Add INHB label
      labs(title = "", y = "Incremental Net Health Benefit", x = "Variable") +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
      coord_flip()
    
    # Filter the top 5 most impactful variables and reverse the order
    tornado_plot_data <- tornado_data %>%
      #filter(Variable %in% head(ordered_variables, 20)) %>%
      arrange(desc(Variable))
    
    print(tornado_plot + geom_linerange(data = tornado_plot_data, color = "blue"))
    
    # Create a PowerPoint presentation
    ppt <- read_pptx()
    
    # Add a slide to the presentation (Title and Content layout)
    slide <- add_slide(ppt, layout = "Title and Content")
    
    # Use ph_with() to add your plot to the slide's body
    slide <- ph_with(slide, value = tornado_plot, location = ph_location_type("body"))
    
    # Save the PowerPoint presentation
    output_pptx <- file.path(graph_domain, "tornado_plot.pptx")
    print(ppt, target = output_pptx)
    
    # Print the saved filename
    cat("PowerPoint presentation saved:", output_pptx, "\n")
    
    rm(tornado_data,
       ordered_variables,
       result)
    
  }
  
  
  #### Cost-Effectiveness Acceptability Curve ----
    
    # Create dataset for CE planes
    ceac <- combined_output %>% 
      select(dataset_index, 
             replication,
             variable,
             category,
             tc_ntrt,
             tc_trt,
             te_ntrt,
             te_trt,) %>% 
      mutate(inc_cost = (tc_trt - tc_ntrt),
             inc_effect = (te_trt - te_ntrt)) %>% 
      filter(category == "Total")

    # Define the range of wtp values and the increment
    wtp_range <- seq(0, 1000000, by = 100)
    
    # Function to calculate pr_CE for a given wtp
    calculate_pr_ce <- function(wtp) {
      pr_CE <- ifelse(((ceac$inc_effect * wtp) - ceac$inc_cost) > 0, 1, 0)
      return(mean(pr_CE))
    }
    
    # Calculate pr_CE for each value of wtp using sapply
    pr_CE <- sapply(wtp_range, calculate_pr_ce)
    pr_nCE <- 1- pr_CE
    
    # Create a data frame for plotting
    ceac_plot_data <- data.frame(wtp_range = wtp_range, pr_CE = pr_CE, pr_nCE = pr_nCE)

    # Create the plot using ggplot2
    ceac_plot <- ggplot(ceac_plot_data, aes(x = wtp_range)) +
      geom_line(aes(y = pr_CE, color = "Medicaid Expansion"), linetype = "solid", linewidth = 1) +
      geom_line(aes(y = pr_nCE, color = "No Medicaid Expansion"), linetype = "solid", linewidth = 1) +
      geom_vline(xintercept = 100000, linetype = "dashed", color = "black") +  # Adjust the xintercept to 100,000
      annotate("text", x = 100000, y = 0.05, label = "$100,000", vjust = 0.5, hjust = -0.05, color = "black") +  # Add label
      labs(title = "Cost-effectiveness acceptability curve",
           x = "Cost-effectiveness threshold",
           y = "Probability of cost-effectiveness",
           color = " ") +
      scale_color_manual(values = c("Medicaid Expansion" = "blue", "No Medicaid Expansion" = "red")) +
      theme_minimal() +
      theme(legend.position = c(0.5, 0.9)) +
      coord_cartesian(ylim = c(0, 1))  # Set y-axis limits
    
    ceac_plot
    
    # Display the plot
    print(ceac_plot)
    
    rm(ceac,
       ceac_plot_data)
    
  #### Distributional Analysis ----
  {
    ### PARETO DOMINANCE ----

        # Create dataset for barcharts
        bar_data <- summarised_output %>% 
          filter(variable != "Total"
                 , variable != "age_cat"
                 , variable != "female") %>%
          # create single insurance variable to correspond to MEPS cost and utility calculations and for cleaner distinction of treatment scenario
          mutate(
            category = fct_recode(
              category
              , "No degree"          = "no_degree"
              , "GED/HS"             = "ged_hs"
              , "Associate/Bachelor" = "associate_bachelor"
              , "Master/Doctorate"   = "master_doctorate"
              , "Poor"      = "poor"
              , "Near Poor" = "near_poor"
              , "Low"       = "low"
              , "Medium"    = "medium"
              , "High"      = "high"
              , "Asian"      = "asian"
              , "Hispanic"   = "hispanic"
              , "Other race" = "other_race"
              , "Black"      = "black"
              , "White"      = "white"),
            variable = fct_recode(
              variable
              , "Education" = "education"
              , "Family Income" = "fam_income"
              , "Race/Ethnicity" = "race")
            , inc_cost = (tc_trt - tc_ntrt)
            , inc_effect = (te_trt - te_ntrt)
            , nhb_trt = (te_trt - (tc_trt / wtp))
            , nhb_ntrt = (te_ntrt - (tc_ntrt / wtp))
            , inhb = inc_effect - (inc_cost / wtp)
          ) 

        # Use other Bar-chart function to show the INHB across categories of variables and see if there is Pareto dominance 
        
        variables_to_summarise <- c("inhb")
        
        # Define unique variable values
        unique_variable_values <- .equity_vars
          #unique(bar_data$variable)
        
        # Use INHB Bar-chart function to show the INHB across categories of variables and see if there is Pareto dominance 
        # i.e. are all groups made better-off by intervention?
        
        # Create a list of bar charts with different titles
        INHB_bar_charts_list <- create_grouped_bar_charts(bar_data, unique_variable_values)
        
        # Modify titles for each individual plot
        titles <- c("Race/Ethnicity", "Education", "Family Income")  # Replace with your desired titles for each plot
        
        for (i in seq_along(INHB_bar_charts_list)) {
          # Modify the plot title
          INHB_bar_charts_list[[i]] <- INHB_bar_charts_list[[i]] +
            labs(title = titles[i]) +  # Update the title for each plot
            theme(plot.title = element_text(size = 10)) +  # Change the title size
            if (i > 1) {  # Condition to remove y-axis title from plots 2 and 3
              theme(axis.title.y = element_blank())  # Remove y-axis title for plots 2 and 3
            }
        }
        
        # Combine plots horizontally using patchwork
        combined_INHB_plots <- wrap_plots(INHB_bar_charts_list, ncol = length(INHB_bar_charts_list))
        
        # Display the combined plots
        print(combined_INHB_plots)
        
    ### EQUALLY DISTRIBUTED EQUIVALENT HEALTH (EDEH) ----
        
        ### Using population weights ###
        
        # Calculate the EDEH for treatment and control across bootstrapped replications
        ede_data_pop <- summarised_output %>%
          filter( variable != "female"
                 , variable != "age_cat"
          ) %>%
          mutate(  nhb_ntrt = (te_ntrt) - ((tc_ntrt) / wtp)
                   , nhb_trt  = (te_trt) - ((tc_trt) / wtp)
                   , inhb     = nhb_trt - nhb_ntrt) %>%
          left_join(., pop_counts, by = "category") %>%
          group_by(variable) %>%
          summarise(  ede_ntrt    = (calculate_ede(nhb_ntrt, epsilon, kp = FALSE, population_size = pop_counts)) * sum(pop_counts)
                    , ede_trt     = (calculate_ede(nhb_trt, epsilon, kp = FALSE, population_size = pop_counts)) * sum(pop_counts)
                    , edeh        = ede_trt - ede_ntrt
                    , nhb_ntrt    = sum(nhb_ntrt*pop_counts) 
                    , nhb_trt     = sum(nhb_trt*pop_counts) 
                    , inhb        = sum(inhb*pop_counts) 
                    , Ae_ntrt     = 1 - (ede_ntrt / nhb_ntrt)
                    , Ae_trt      = 1 - (ede_trt / nhb_trt)
                    , Ae          = -(Ae_trt - Ae_ntrt)
                    ) %>%
          ungroup() %>%
          select(variable,
                 edeh,
                 Ae,
                 inhb) %>%
          mutate(sample = "population")
        
        # Creating a new column with custom labels
        ede_data_pop$custom_labels <- ifelse(ede_data_pop$variable == "education", "Education", 
                                             ifelse(ede_data_pop$variable == "fam_income", "Family Income", 
                                                    ifelse(ede_data_pop$variable == "race", "Race/Ethnicity", NA)))
        
        # Plotting using the custom labels for text
        ede_plot_pop <- ggplot(ede_data_pop, aes(x = Ae, y = inhb, color = variable, shape = variable, label = custom_labels)) +
          geom_point(size = 6) +
          labs(
            x = "Equity (Index of inequality)",
            y = "Efficiency (INHB)",
            color = "Equity-relevant variable",
            shape = "Equity-relevant variable"
          ) +
          scale_color_viridis(
            discrete = TRUE,
            breaks = c("education", "fam_income", "race"),
            labels = c("Education", "Family Income", "Race/Ethnicity")
          ) +
          scale_shape_manual(
            values = c("education" = 16, "fam_income" = 17, "race" = 18),
            guide = guide_legend(override.aes = list(size = 4))
          ) +
          geom_vline(xintercept = 0, color = "black") +
          geom_hline(yintercept = 0, color = "black") +
          theme_minimal() +
          theme(
            legend.position = "none"  # Remove the legend
          ) +
          geom_text(
            vjust = -2, hjust = 1,  # Adjust vertical and horizontal justification
            size = 3.2, color = "black"  # Set size and color of the labels
          ) + 
          coord_cartesian(xlim = c((min(ede_data_pop$Ae) + min(ede_data_pop$Ae*0.4)), 
                                   (max(ede_data_pop$Ae) + max(ede_data_pop$Ae*0))))  # Set x-axis limits

        print(ede_plot_pop)
        
    ### DCEA BY WTP AND INEQUALITY AVERSION ----

        # Define a range of values for wtp and epsilon; see sensitivity analysis parameters

        # Define the list of variable values
        variable_values <- c("fam_income", "education", "race")
        
        # Create an empty list to store the plots for each variable
        plots_list <- list()
        
        # Loop over the variable values
        for (v in variable_values) {
          # Create an empty matrix to store the edeh values
          results_matrix <- matrix(NA, nrow = length(wtp_values), ncol = length(epsilon_values))
          
          # Loop over wtp values
          for (i in seq_along(wtp_values)) {
            wtp <- wtp_values[i]
            
            # Loop over epsilon values
            for (j in seq_along(epsilon_values)) {
              epsilon <- epsilon_values[j]
              
              # cat("Evaluating combination: wtp =", wtp, "epsilon =", epsilon, "\n")
              
              # Calculate the EDEH for treatment and control
              edeh_data <- summarised_output %>%
                filter(variable == v) %>%
                mutate(  nhb_ntrt = (te_ntrt) - ((tc_ntrt) / wtp)
                         , nhb_trt  = (te_trt) - ((tc_trt) / wtp)
                         , inhb     = nhb_trt - nhb_ntrt) %>%
                left_join(., pop_counts, by = "category") %>%
                group_by(variable) %>%
                summarise(ede_ntrt    = (calculate_ede(nhb_ntrt, epsilon, kp = FALSE, population_size = pop_counts)) * sum(pop_counts)
                          , ede_trt     = (calculate_ede(nhb_trt, epsilon, kp = FALSE, population_size = pop_counts)) * sum(pop_counts)
                          , edeh        = ede_trt - ede_ntrt
                          , nhb_ntrt    = sum(nhb_ntrt*pop_counts) 
                          , nhb_trt     = sum(nhb_trt*pop_counts) 
                          , inhb        = sum(inhb*pop_counts) 
                          , Ae_ntrt     = 1 - (ede_ntrt / nhb_ntrt)
                          , Ae_trt      = 1 - (ede_trt / nhb_trt)
                          , Ae          = -(Ae_trt - Ae_ntrt)) %>%
                select(variable,
                       edeh,
                       Ae)
              
              # Check if any non-numeric values exist in the results
              if (any(!is.numeric(edeh_data$edeh) | is.na(edeh_data$edeh) | !is.finite(edeh_data$edeh))) {
                cat("Problematic combination: wtp =", wtp, "epsilon =", epsilon, "\n")
                # Add additional debugging or print statements here to inspect data for that combination
              }
              
              # Store the edeh value in the results matrix
              edeh_data$CE_threshold <- wtp
              edeh_data$Epsilon <- epsilon
              results_matrix[i, j] <- mean(edeh_data$edeh)
            }
          }
          
          # Create a data frame with correct reshaping
          dcea_data <- data.frame(
            Epsilon = rep(epsilon_values, each = length(wtp_values)),
            CE_threshold = rep(wtp_values, times = length(epsilon_values)),
            EDEH = as.vector(results_matrix)
          )
          
          ede_epsilon_plot <- ggplot(dcea_data, aes(x = Epsilon, y = EDEH, color = factor(CE_threshold))) +
            geom_line() +  # Use a line plot
            geom_hline(yintercept = 0, color = "black") +
            labs(title = "EDEH vs. Epsilon", x = "Inequality Aversion (Epsilon)", y = "EDEH") +
            theme_minimal()
          
          # Store the plots in the list
          plots_list[[v]] <- list(ede_epsilon_plot = ede_epsilon_plot)
        }
        
  } # Close "Distributional Analysis" section
 # Close "MODEL" section

}
##### Store Results ----

  # STORE TABLES
  # Create the Excel workbook
  wb <- createWorkbook()
  
  # Loop through each output and add it as a separate sheet
  for (filename in table_filenames) {
    sheet_name <- filename
    sheet_data <- get(filename)  # Assuming you have already created these outputs
    
    addWorksheet(wb, sheetName = sheet_name)
    writeData(wb, sheet = sheet_name, x = sheet_data)
  }
  
  # Remove the existing file if it exists
  if (file.exists(here(table_domain, results_tables))) {
    file.remove(here(table_domain, results_tables))
  }
  
  # Save the Excel workbook
  saveWorkbook(wb, file = (here(table_domain, results_tables)))
  
  # Print the saved filename
  cat("Excel file saved:", results_tables, "\n")
  
  rm(sheet_data)
  
  # STORE GRAPHS
  
  # Create a list of ggplot objects (replace this with your own code)
  ggplot_list <- list(
     combined_INHB_plots
    , ede_plot_pop
    , plots_list[["fam_income"]]$ede_epsilon_plot
    , plots_list[["education"]]$ede_epsilon_plot
    , plots_list[["race"]]$ede_epsilon_plot
    , ceac_plot
    , miss_plot_full
    , miss_plot_final
    , miss_map_full
    , miss_map_nhanes_filter
  )
  
  # Create a PowerPoint presentation
  ppt <- read_pptx()
  
  # Loop through each ggplot and add it as a separate slide
  for (plot_cvd in ggplot_list) {
    slide <- add_slide(ppt, layout = "Title and Content")
    
    # Insert the ggplot into the slide
    plot_area <- ph_with(slide, value = plot_cvd, location = ph_location_fullsize(center = TRUE))
    
  }
  
  # Save the PowerPoint presentation
  output_pptx <- file.path(graph_domain, results_graphs)
  print(ppt, target = output_pptx)
  
  # Print the saved filename 
  cat("PowerPoint presentation saved:", output_pptx, "\n")
  
  if (.timer == 1) {
    # End timing
    end_time <- Sys.time()
    # Calculate elapsed time
    elapsed_time <- end_time - start_time
    print(elapsed_time)
  }
  
  # save the combined_output list
  # Save the list to a file
   saveRDS(combined_output, file = (here(data_domain, results_combined_output)))

