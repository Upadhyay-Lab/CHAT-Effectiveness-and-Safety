*********************************************
* Title: CHAT Safety and Effectiveness
* Purpose: Extended Table 1
* Last updated: 1/23/24
* Created by: Leah
* Last updated by: Leah
*********************************************
cls
*********************************************
*Macros
*********************************************
local tables = "~/tables/"
local dta = "~/dta/"
local maindo = "~/dofiles/"
*********************************************
*Set up
*********************************************
cd "`dta'"
use "CHATclean.dta", clear

local varlist =  "age6 csrace7 urban2 csab2 csbirth2 cpgdur6 csscreenus async"
foreach var in `varlist' {
tab `var' if exclude==0, m
}

cd "`tables'"
putexcel set "CHAT_safetyefficacy_tablea1.xlsx", replace
putexcel A1:C63 = "", fpattern(solid ,white) font("Times New Roman")

*********************************************
*Headers
*********************************************
putexcel A1:C1 = "Table S1. Characteristics of the clinical chart sample and survey subsample", bold merge hcenter vcenter
putexcel B2:C2 = "% (n)", hcenter merge
putexcel B3 = "Clinical chart sample", border(bottom)
putexcel C3 = "Survey sample", border(bottom)
*********************************************
*Column B: Clinical Sample Description
*********************************************
local row = 5

count if exclude==0
putexcel B4 = "n=`r(N)'"

count if exclude==0 & surveyexclude==0
putexcel C4 = "n=`r(N)'"


foreach var in `varlist' {

local varname : var label `var'
putexcel A`row' = "`varname'", bold

tab `var' if exclude==0, matcell(ov)
mata: st_matrix("perov", (100*st_matrix("ov")  :/ colsum(st_matrix("ov"))))

tab `var' if exclude==0, matcell(freq)
mata: st_matrix("per", (100*st_matrix("freq")  :/ colsum(st_matrix("freq"))))

tab `var' if exclude==0 & surveyexclude==0, matcell(svyov)
mata: st_matrix("svyperov", (100*st_matrix("svyov")  :/ colsum(st_matrix("svyov"))))


levelsof `var'
	local lev = 1
	foreach level in `r(levels)' {

	local row = `row' + 1

	local vallab : label `var' `level'
	putexcel A`row' = "   `vallab'"
	
	local ov = ov[`lev',1]
	local perov = string(perov[`lev',1], "%4.1f")
	putexcel B`row' = "`ov' (`perov'%)"
	
	local freq = svyov[`lev',1]
	local per = string(svyperov[`lev',1], "%4.1f")
	putexcel C`row' = "`freq' (`per'%)"

	
	local lev = `lev' + 1
	}
	
local row = `row' + 1	
}
