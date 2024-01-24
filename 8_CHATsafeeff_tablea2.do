*********************************************
* Title: CHAT Safety and Effectiveness
* Purpose: Extended Table 2
* Created: 9/8/23
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
local log = "~/log"

*********************************************
*Log
*********************************************
*Log
cd "`log'"
capture log using chat_se_ta2_`date'.smcl, replace

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
putexcel set "CHAT_safetyefficacy_tablea2.xlsx", replace
putexcel A1:D60 = "", fpattern(solid ,white) font("Times New Roman")

*********************************************
*Headers
*********************************************
putexcel A1:D1 = "Table A2. Characteristics of sample with and without known abortion outcomes", bold merge hcenter vcenter
putexcel B2:C2 = "n (%)", merge hcenter
putexcel D2 = "p-value", hcenter
*********************************************
*Column B: Clinical Sample Description
*********************************************
local row = 4

count if exclude==0
putexcel B3:C3 = "n=`r(N)'", merge hcenter

foreach var in `varlist' {
local row = `row' + 1

local varname : var label `var'
putexcel A`row' = "`varname'", bold

tab `var' seq2teffdenom if exclude==0, matcell(ov) col chi
local p = string(`r(p)', "%04.3f")
if `r(p)' < 0.001 {
local p = "<0.001"
}
mata: st_matrix("per", (100*st_matrix("ov")  :/ rowsum(st_matrix("ov"))))

levelsof `var'
	local lev = 1
	foreach level in `r(levels)' {
	
	local row = `row' + 1
	if `lev'==1 {
	local toprow = `row' 
	}
	local vallab : label `var' `level'
	putexcel A`row' = "   `vallab'"
	
	local ov = ov[`lev',1]
	local per = string(per[`lev',1], "%4.1f")
	putexcel B`row' = "`ov' (`per'%)"
	
	local ov = ov[`lev',2]
	local per = string(per[`lev',2], "%4.1f")
	putexcel C`row' = "`ov' (`per'%)"
		
	local lev = `lev' + 1 
	}
	
	putexcel D`toprow':D`row' = "`p'", merge hcenter vcenter

}
capture log c
