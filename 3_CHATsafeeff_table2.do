*********************************************
* Title: CHAT Safety and Effectiveness
* Purpose: Table 2
* Created: 9/14/23
* Last updated: 1/23/24
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

*********************************************
*Log
*********************************************
*Log
cd "`log'"
capture log c
log using chat_se_t2_`date'.smcl, replace

*********************************************
*Set up
*********************************************
local varlist = "async age5 csrace5 urban2 csab2 csbirth2 cpgdur6 csscreenus"

foreach var in `varlist' {
tab `var' if exclude==0, m
}

cd "`tables'"
putexcel set "CHAT_safetyefficacy_table2_`date'.xlsx", replace
putexcel A1:E43 = "", fpattern(solid ,white) font("Times New Roman")

*********************************************
*Headers
*********************************************
putexcel A1:E1 = "Table 2. Rate of complete abortion by characteristics of the sample from complete cases and multiple imputation analysis 1", bold merge vcenter
putexcel B2:C2 = "Adjusted Complete Case", hcenter merge
putexcel B3 = "Effectiveness Rate (95% CI)", hcenter
putexcel C3 = "p-value", hcenter
putexcel D2:E2 = "Adjusted Imputed Analysis", hcenter merge
putexcel D3 = "Effectiveness Rate (95% CI)", hcenter
putexcel E3 = "p-value", hcenter

count if exclude==0 & seq2teffdenom==1
putexcel B4:C4 = "N=`r(N)'", hcenter merge

count if exclude==0 
putexcel D4:E4 = "N=`r(N)'", hcenter merge

putexcel A5 = "Overall", bold
*********************************************
*Column A: Labels
*********************************************
local row = 6
foreach var in `varlist' {
local varname : var label `var'
putexcel A`row' = "`varname'", bold

tab `var' if exclude==0 & baseline==1, matcell(a)
mata: st_matrix("freq", (100*st_matrix("a")  :/ colsum(st_matrix("a"))))

levelsof `var'
	local lev = 1
	foreach level in `r(levels)' {
	local row = `row' + 1
	local vallab : label `var' `level'
	putexcel A`row' = "   `vallab'"
	local lev = `lev' + 1
	}
local row = `row' + 1	
}

*********************************************
*Column B:C: Effectiveness - Unimputed Adjusted
*********************************************
cd "`dta'"
use "CHATclean.dta", clear

cd "`tables'"
putexcel set "CHAT_safetyefficacy_table2.xlsx", modify

local row = 6
local n = 1 
foreach var in `varlist' { 
	
	*Categories	
	local storerow = `row'
	
			*Populate Cols
			local row = `storerow'
			
	logistic noint i.`var' if exclude ==0 & seq2teffdenom==1 

	est sto effregression
	mat c = r(table)

	est replay effregression
	margins 

	mat a = r(table)
	local prop = string(100*(a[1,1]), "%4.1f")			
	local l_95 = string(100*(a[5,1]), "%4.1f")
	local u_95 = string(100*(a[6,1]), "%4.1f")
	local cell = "`prop' (`l_95', `u_95')"

	putexcel C6 = "`cell'"

	mat ov = r(table)
			est restore effregression
			
			margins i.`var', post
			
			mat b = r(table)
						
			local variable = `var'
			levelsof `var'
			local loop=0
			
			foreach level in `r(levels)' {
			local loop = `loop'+1
			di "loop: `loop'"
			local loop1 = `loop'+1
			local row = `row'+1
			local lab_val : label `var' `level'
						
			local prop = string(100*(b[1,`loop']), "%4.1f")			
			local l_95 = string(100*(b[5,`loop']), "%4.1f")
			local u_95 = string(100*(b[6,`loop']), "%4.1f")
				if b[6,`loop']>1 {
				local u_95 = string(100, "%4.1f")
				}
			local pval  = c[4,`n']
			di "pvalue: `pval'"
			local p  = string((c[4,`n']), "%04.3f")
			local cell = "`prop' (`l_95', `u_95')"
			
			putexcel B`row' = "`cell'"
			if `loop'==1 {
				putexcel C`row' ="ref", hcenter italic 
				}	
			else {
				if `pval'<0.001 {
				putexcel C`row' = "<0.001", hcenter
				}
				else {
				putexcel C`row' = "`p'"
				}
				}
			
			local n = `n' + 1 
			}
			local row = `row'+1
		}
*********************************************
*Column D:E: Effectiveness - Imputed Adjusted
*********************************************
preserve
cd "`dtaimp'"
use "2_chat_imputed.dta", clear

cd "`tables'"
putexcel set "CHAT_safetyefficacy_table2.xlsx", modify

count
putexcel D4:E4 = "N=`r(N)'", hcenter merge


mi estimate, esampvaryok errorok: logistic noint i.async i.age5 i.csrace5 i.urban2 i.csab2 i.csbirth2 i.cpgdur6 i.csscreenus if exclude ==0

mat c = r(table)
est sto effregressionimp

mimrgns, post expression(exp(predict(xb))/(1+exp(predict(xb))))	esampvaryok errorok

mat a = r(table)
local prop = string(100*(a[1,1]), "%4.1f")			
local l_95 = string(100*(a[5,1]), "%4.1f")
local u_95 = string(100*(a[6,1]), "%4.1f")
local cell = "`prop' (`l_95', `u_95')"

global oveffb = "`prop'"
global ovefflb = "`l_95'"
global oveffub = "`u_95'"

putexcel D5 = "`cell'"

local row = 6
local n = 1 
foreach var in `varlist' { 
	
	*Categories	
	local storerow = `row'
	
			*Populate Cols
			local row = `storerow'
			
			est restore effregressionimp

			mimrgns i.`var', post expression(exp(predict(xb))/(1+exp(predict(xb)))) esampvaryok errorok	
			
			mat b = r(table)
						
			local variable = `var'
			levelsof `var'
			local loop=0
			
			foreach level in `r(levels)' {

			
			local loop = `loop'+1
			di "loop: `loop'"
			local loop1 = `loop'+1
			local row = `row'+1
			local lab_val : label `var' `level'
			
			local prop = string(100*(b[1,`loop']), "%4.1f")			
			local l_95 = string(100*(b[5,`loop']), "%4.1f")
			local u_95 = string(100*(b[6,`loop']), "%4.1f")
				if b[6,`loop']>1 {
				local u_95 = string(100, "%4.1f")
				}
			local pval  = c[4,`n']
			di "pvalue: `pval'"
			local p  = string((c[4,`n']), "%04.3f")
			local cell = "`prop' (`l_95', `u_95')"
			
			if "`var'"=="async" {
			if `level'==0 {
			global synceffb = "`prop'"
			global syncefflb = "`l_95'"
			global synceffub = "`u_95'"
			global bysynceffp = "`p'"
			}
			else {
			global asynceffb = "`prop'"
			global asyncefflb = "`l_95'"
			global asynceffub = "`u_95'"				
			}
			}
			putexcel D`row' = "`cell'"
			if `loop'==1 {
				putexcel E`row' ="ref", hcenter italic 
				}	
			else {
				if `pval'<0.001 {
				putexcel E`row' = "<0.001", hcenter
				}
				else {
				putexcel E`row' = "`p'"
				}
				}
			if ("`var'"=="csrace5" & `level'==5)|("`var'"=="csab2" & `level' == 1)|("`var'"=="csbirth2" & `level' == 1)|("`var'"=="cpgdur6" & `level'==6) {
			local row = `row'+1
			}
			
			local n = `n' + 1 
			}
			local row = `row'+1
		}
*********************************************
*********************************************
*Safety
*********************************************
*********************************************

*********************************************
*Column B:C: Effectiveness - Proportion and Unimputed Adjusted
*********************************************

local row = 5
local n = 1 
foreach var in `varlist' { 
	
	*Categories	
	local storerow = `row'
	
		*Populate Cols
		local row = `storerow'
	
	logistic saeanyrcd i.`var' if exclude ==0 & seq2teffdenom==1 

	est sto effregression
	mat c = r(table)

	est replay effregression
	margins 

	mat a = r(table)
	local prop = string(100*(a[1,1]), "%4.1f")			
	local l_95 = string(100*(a[5,1]), "%4.1f")
	local u_95 = string(100*(a[6,1]), "%4.1f")
	local cell = "`prop' (`l_95', `u_95')"

	putexcel F5 = "`cell'"

	mat ov = r(table)
	est restore effregression
			
	margins i.`var', post
			
	mat b = r(table)
						
	local variable = `var'
	levelsof `var' if `var'!=99
	local loop=0
			
	foreach level in `r(levels)' {
	
	*Add age subcategories 
	if "`var'"== "age5" & `level'==2 {
	tab age6 saeanyrcd if exclude ==0 & seq2teffdenom==1, matcell(p1)
	mata : st_matrix("p2", rowsum(st_matrix("p1")))

	logistic saeanyrcd i.age6 if exclude ==0 & seq2teffdenom==1 
	margins i.age6
	mat age6 = r(table)
	
	local row = `row'+1
	
	levelsof age6 if age6<3
	foreach lev in `r(levels)' {
	local vallab : label age6 `lev'
	putexcel A`row' = "   `vallab'", txtindent(1)
	local num = p1[`lev',2]
	local den = p2[`lev',1]
	local prop = `num'/`den'
	
	if `prop'==1 {
	putexcel C`row' = "100.0 (100.0, 100.0)"
	}
	else {
	local prop = string(100*(age6[1,1]), "%4.1f")			
	local l_95 = string(100*(age6[5,1]), "%4.1f")
	local u_95 = string(100*(age6[6,1]), "%4.1f")
	if age6[6,1]>1 {
	local u_95 = string(100, "%4.1f")
	}
	local cell = "`prop' (`l_95', `u_95')"
			
	putexcel C`row' = "`cell'"	
	}
	local row = `row'+1
	}
	local row = `row'-1
	}
	
	*Estimate
	local loop = `loop'+1
			di "loop: `loop'"
			local loop1 = `loop'+1
			local row = `row'+1
			local lab_val : label `var' `level'
						
			local prop = string(100*(b[1,`loop']), "%4.1f")			
			local l_95 = string(100*(b[5,`loop']), "%4.1f")
			local u_95 = string(100*(b[6,`loop']), "%4.1f")
				if b[6,`loop']>1 {
				local u_95 = string(100, "%4.1f")
				}
			local pval  = c[4,`n']
			di "pvalue: `pval'"
			local p  = string((c[4,`n']), "%04.3f")
			local cell = "`prop' (`l_95', `u_95')"
			
			putexcel C`row' = "`cell'"
			
			local n = `n' + 1 
		*Proportion 
		tab `var' saeanyrcd if exclude ==0 & seq2teffdenom==1, matcell(p1)
		mata : st_matrix("p2", rowsum(st_matrix("p1")))
		local num = p1[`loop',2]
		local den = p2[`loop',1]

		local prop = `num'/`den'
		if `prop'==1 {
		putexcel C`row' = "100.0 (100.0, 100.0)"
		}		
			}
			local row = `row'+1
		}
	local row = `row' - 1
*********************************************
*Column D:E: Effectiveness - Imputed Adjusted
*********************************************

preserve
cd "`dtaimp'"
use "2_chat_imputed.dta", clear

cd "`tables'"
putexcel set "CHAT_safetyefficacy_table3safe_`date'.xlsx", modify

count
putexcel D4:E4 = "N=`r(N)'", hcenter merge

			mi estimate, esampvaryok errorok: logistic saeanyrcd if exclude ==0

			mat c = r(table)
			est sto effregressionimp

			mimrgns, post expression(exp(predict(xb))/(1+exp(predict(xb))))	esampvaryok errorok

			mat a = r(table)
			local prop = string(100*(a[1,1]), "%4.1f")			
			local l_95 = string(100*(a[5,1]), "%4.1f")
			local u_95 = string(100*(a[6,1]), "%4.1f")
			local cell = "`prop' (`l_95', `u_95')"

			global oveffb = "`prop'"
			global ovefflb = "`l_95'"
			global oveffub = "`u_95'"

			putexcel D5 = "`cell'"

			
local row = 5
local n = 1 
foreach var in `varlist' { 
	
	*Categories	
	local storerow = `row'
	
			*Populate Cols
			local row = `storerow'
			
			capture mi estimate, esampvaryok errorok: logistic saeanyrcd i.`var' if exclude ==0

			mat c = r(table)
			est sto effregressionimp

			capture mimrgns i.`var', post expression(exp(predict(xb))/(1+exp(predict(xb)))) esampvaryok errorok	
			
			mat b = r(table)
						
			local variable = `var'
			levelsof `var' if `var'!=99
			local loop=0
			
			foreach level in `r(levels)' {

			
			local loop = `loop'+1
			di "loop: `loop'"
			local loop1 = `loop'+1
			local row = `row'+1
			local lab_val : label `var' `level'
			
			local prop = string(100*(b[1,`loop']), "%4.1f")			
			local l_95 = string(100*(b[5,`loop']), "%4.1f")
			local u_95 = string(100*(b[6,`loop']), "%4.1f")
				if b[6,`loop']>1 {
				local u_95 = string(100, "%4.1f")
				}
			local pval  = c[4,`n']
			di "pvalue: `pval'"
			local p  = string((c[4,`n']), "%04.3f")
			local cell = "`prop' (`l_95', `u_95')"
			
			if "`var'"=="async" {
			if `level'==0 {
			global synceffb = "`prop'"
			global syncefflb = "`l_95'"
			global synceffub = "`u_95'"
			global bysynceffp = "`p'"
			}
			else {
			global asynceffb = "`prop'"
			global asyncefflb = "`l_95'"
			global asynceffub = "`u_95'"				
			}
			}
			putexcel D`row' = "`cell'"
			if `loop'==1 {
				putexcel E`row' ="ref", hcenter italic 
				}	
			else {
				if `pval'<0.001 {
				putexcel E`row' = "<0.001", hcenter
				}
				else {
				putexcel E`row' = "`p'"
				}
				}
			if ("`var'"=="csrace5" & `level'==5)|("`var'"=="csab2" & `level' == 1)|("`var'"=="csbirth2" & `level' == 1)|("`var'"=="cpgdur6" & `level'==6) {
			local row = `row'+1
			}
			
			local n = `n' + 1 
			}
			local row = `row'+1
		}
		
local row = `row'+1
putexcel A`row':E`row' = "Patient age was collapsed in the multivariable models to facilitate model convergence (16-25 29 years old, 20-24 years, 25-29 years old, 30-34 years old, 35 years and older)", txtwrap border(top bottom) merge
local row = `row'+1
putexcel A`row':E`row' = "Estimates are derived from marginal estimates from logistic regressions. Estimates are not adjusted for multiple comparisons. Imputation models included patient age, urbanicity, whether the patient obtained screening ultrasonography, whether the patient obtained synchronous or asynchronous telehealth care, whether the patient participated in CHAT surveys, virtual clinic, and whether the patient used an abortion fund to pay for any portion of their abortion.", txtwrap border(top bottom) merge

restore


