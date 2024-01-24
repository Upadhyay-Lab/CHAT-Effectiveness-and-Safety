*********************************************
* Title: CHAT Safety and Effectiveness
* Purpose: Table 3
* Created: 10/31/22
* Last updated: 7/29/23
* Created by: Leah
* Last updated by: Leah
*********************************************

*********************************************
*Macros
*********************************************
local tables = "~/tables/"
local dta = "~/data/"
local do = "~/dofiles"
local log = "~/log"

local date : display %tdCYND date(c(current_date), "DMY")
local varlist = "async age5 csrace5 urban2 csab2 csbirth2 cpgdur6 csscreenus"
local reglist = "i.async i.age4 i.csrace4 i.urban2 i.csab2 i.csbirth2 i.cpgdur3 i.csscreenus"
*********************************************
*Set up
*********************************************

foreach var in `varlist' {
tab `var' if exclude==0, m
}

cd "`tables'"
putexcel set "CHAT_safetyefficacy_table3.xlsx", replace
putexcel A1:D30 = "", fpattern(solid ,white) font("Times New Roman")

*********************************************
*Log
*********************************************
*Log
capture log c
cd "`log'"
log using chat_se_t3_`date'.smcl, replace
*********************************************

*********************************************
*Set up
*********************************************
cd "`dta'"
use "CHATclean.dta", clear

*********************************************
*Headers
*********************************************
cd "`tables'"
putexcel A1:F1 = "Table 3. Medication abortion additional interventions and serious adverse events", bold merge vcenter
putexcel B2:B3 = "No.", hcenter border(bottom) merge
putexcel C2 = "Complete Case"
putexcel D2 = "Imputed"
putexcel C3:D3 = "Rate 95% CI", border(bottom)

putexcel A3 = "Effectiveness", bold border(top)
putexcel A4 = "Complete abortion without intervention"
putexcel A5 = "Intervention to Complete Abortion †"
putexcel A6 = "   Procedure, aspiration, or surgery"
putexcel A7 = "   Prescribed >1,600 mcg misoprostol, mifepristone, or other medications"
putexcel A8 = "   Treatment for ectopic pregnancy*"
putexcel A9 = "   Suspected or confirmed continuing pregnancy"
putexcel A10 = "Safety", bold border(top)
putexcel A11 = "No Major abortion-related adverse event*"
putexcel A12 = "Major abortion-related adverse events*"
putexcel A13 = "   Blood transfusion*"
putexcel A14 = "   Other major surgery, including treatment of ectopic pregnancy*"
putexcel A15 = "   Hospital Admission*"
putexcel A16 = "Other outcomes", bold border(top)
putexcel A17 = "   Emergency Department visits"
*********************************************
*Column B: Ns
*********************************************

*Complete abortion without intervention
count if anyint==0 & exclude==0 & seq2teffdenom==1
putexcel B4 = `r(N)' 

*Intervention to Complete Abortion
count if anyint==1 & exclude==0 & seq2teffdenom==1
putexcel B5 = `r(N)' 
   
   *Aspiration, second trimester procedure, or surgery
	count if intproc==1 & exclude==0 & seq2teffdenom==1
	putexcel B6 = `r(N)' 
	
  *Prescribed >1,600 mcg misoprostol, mifepristone, or other medications
   	count if intmed==1 & exclude==0 & seq2teffdenom==1
	putexcel B7 = `r(N)' 

   *Treatment for ectopic pregnancy
    count if intectopic==1 & exclude==0 & seq2teffdenom==1
	putexcel B8 = `r(N)' 

   *Continuing pregnancy without intervention
    count if contpg==1 & exclude==0 & seq2teffdenom==1
	putexcel B9 = `r(N)' 

*No Major abortion-related adverse event
count if saeany==0 & exclude==0 & seq2teffdenom==1
putexcel B11 = `r(N)' 
*Major abortion-related adverse events
count if saeany==1 & exclude==0 & seq2teffdenom==1
putexcel B12 = `r(N)' 
   *Blood transfusion
	count if saeblood==1 & exclude==0 & seq2teffdenom==1
	putexcel B13 = `r(N)'  
   *Other major surgery, including treatment of ectopic pregnancy
	count if saesurg==1 & exclude==0 & seq2teffdenom==1
	putexcel B14 = `r(N)'     
	*Hospital Admission
	count if saehosp==1 & exclude==0 & seq2teffdenom==1
	putexcel B15 = `r(N)'   	
*ED visit
count if seq2fedhosp==1 & exclude==0 & seq2teffdenom==1
putexcel B17 = `r(N)'
*********************************************
*Column C and D: Unimputed
*********************************************
local row = 4

foreach var in noint intany intproc intmed intectopic contpg saeanyrcd saeany saeblood saesurg saehosp seq2fedhosp {
count if `var'==1 & exclude==0 & seq2teffdenom==1
local n = `r(N)'
logistic `var'  if exclude ==0 & seq2teffdenom==1 

margins
mat a = r(table)
local b = a[1,1]*100
local ll = a[5,1]*100
local ul = a[6,1]*100

if `b'>=1 {
local bs = string(`b', "%9.1f") 
local lls = string(`ll', "%9.1f") 
local uls = string(`ul', "%9.1f") 
}
else if `b'<1 {
local bs = string(`b', "%4.2f") 
if `ll'<0 {
	local lls = "0.00"
	}
	else {
	local lls = string(`ll', "%4.2f") 
	}
local uls = string(`ul', "%4.2f") 
}
local cell = "`bs' (`lls', `uls')"
di "`cell'"

putexcel C`row' = "`cell'"

local row = `row' + 1
if `row'==10|`row'==16 {
local row = `row' + 1
}
}

*********************************************
*Column E and F: Imputed
*********************************************
local reglist = "i.async i.age4 i.csrace4 i.urban2 i.csab2 i.csbirth2 i.cpgdur3 i.csscreenus"
local dta = "/Users/koenig/Library/CloudStorage/Box-Box/CHAT Study/Data/data/Safety and Effectiveness"
cd "`dta'"
use "2_chat_imputed.dta", clear

cd "`tables'"
putexcel set "CHAT_safetyefficacy_table3.xlsx", modify
local row = 4

foreach var in noint intany intproc intmed intectopic contpg saeanyrcd saeany saeblood saesurg saehosp seq2fedhosp {  
di "`var'"
count if `var'==1 & exclude==0 & seq2teffdenom==1
local n = `r(N)'

if `n'>=14 {
capture mi estimate, post errorok esampvaryok: logistic `var' `reglist', dots difficult 	
}
if `n'<14|"`var'"=="saeanyrcd"|"`var'"=="saeany" {
mi estimate, post errorok esampvaryok: logistic `var', dots difficult 
putexcel E`row' = "UNADJUSTED"
}

est sto regimp`var'
est replay regimp`var'

mimrgns, post expression(exp(predict(xb))/(1+exp(predict(xb))))	

mat a = r(table)
local b = a[1,1]*100
local ll = a[5,1]*100
local ul = a[6,1]*100

if `b'>=1 {
local bs = string(`b', "%9.1f") 
local lls = string(`ll', "%9.1f") 
local uls = string(`ul', "%9.1f") 
}
else if `b'<1 {
local bs = string(`b', "%4.2f") 
if `ll'<0 {
	local lls = "0.00"
	}
	else {
	local lls = string(`ll', "%4.2f") 
	}
local uls = string(`ul', "%4.2f") 
}
local cell = "`bs' (`lls', `uls')"
di "`cell'"

if "`var'" == "anysae1" {
global ovsafeb = "`bs'"
global ovsafelb = "`lls'"
global ovsafeub = "`uls'"
}

putexcel D`row' = "`cell'"
local row = `row' + 1
if `row'==10|`row'==16 {
local row = `row' + 1
}
}
putexcel A`row':D`row' = "Models for estimates with N.>15 are adjusted for synchronicity of care, patient age, race/ethnicity, residence, prior abortion, pregnancy duration at intake, and pre-abortion screening ultrasonography. Estimates are calculated from multivariable logistic regression models with missing outcomes and covariates imputed using multiple imputation with chained equations. Outcomes denoted with an * are unadjusted.", merge txtwrap border(top)
local row = `row'+1
putexcel A`row':D`row' = "† – Sub-categories are not mutually exclusive", merge txtwrap border(top bottom)
