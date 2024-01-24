********************************************************************************
* Title: CHAT Safety and Effectiveness
* Purpose: Print table 1 - Overall sample description and 
  *chi-2/fisher's comparison of sync vs. async
* Created: 10/31/22
* Last updated: 1/23/24
* Created by: Leah Koenig
* Last updated by: Leah Koenig
* Notes:
  *Please note the entire do file needs to run together for the macros to run properly
********************************************************************************
cls

********************************************************************************
*Store date
********************************************************************************
local date : display %tdCYND date(c(current_date), "DMY")

********************************************************************************
*Store for pathways - needs to be updated to local computer pathway
********************************************************************************
local tables = "~/tables/"
local dta = "~/data/"
local do = "~/dofiles"
local log = "~/log"

*********************************************
*Log
*********************************************
cd "`log'"
capture log using chat_se_t1_`date'.smcl, replace

*********************************************
*Set up
*********************************************
cd "`dta'" 
use "CHATclean.dta", clear

local varlist =  "age6 csrace7 urban2 csab2 csbirth2 cpgdur6 csscreenus seq2teffdenom async"
foreach var in `varlist' {
tab `var' if exclude==0, m
}

cd "`tables'"
putexcel set "CHAT_safetyefficacy_table1.xlsx", replace
putexcel A1:E60 = "", fpattern(solid ,white) font("Times New Roman")

*********************************************
*Headers
*********************************************
putexcel A1:E1 = "Table 1. Characteristics of samples obtaining synchronous and asynchronous telehealth medication abortion care", bold merge vcenter
putexcel B2:D2 = "% (n)", hcenter merge

*********************************************
*Column B: Clinical Sample Description
*********************************************
local row = 4

count if exclude==0
putexcel B3:D3 = "n=`r(N)'", merge hcenter

foreach var in `varlist' {

local varname : var label `var'
putexcel A`row' = "`varname'", bold

tab `var' if exclude==0, matcell(ov)
mata: st_matrix("perov", (100*st_matrix("ov")  :/ colsum(st_matrix("ov"))))

*P-value
	tab `var' sync if exclude==0, matcell(freq) chi expect
	local p = string(`r(p)', "%04.3f")
	if `r(p)' < 0.001 {
	local p = "<0.001"
	}

	*Test for fisher's exact test - this needs to be manually added if any expect cells <5
	capture tab `var' sync if exclude==0, expect
	capture tab `var' sync if exclude==0, exact
	
mata: st_matrix("per", (100*st_matrix("freq")  :/ colsum(st_matrix("freq"))))

levelsof `var'
	local lev = 1
	foreach level in `r(levels)' {

	local row = `row' + 1

	local vallab : label `var' `level'
	putexcel A`row' = "   `vallab'"
	
	local ov = ov[`lev',1]
	local perov = string(perov[`lev',1], "%4.1f")
	putexcel B`row' = "`ov' (`perov'%)"
	
	local freq = freq[`lev',1]
	local per = string(per[`lev',1], "%4.1f")
	putexcel C`row' = "`per' (`freq')"

	local freq = freq[`lev',2]
	local per = string(per[`lev',2], "%4.1f")
	putexcel D`row' = "`per' (`freq')"
	
	local lev = `lev' + 1
	}
	local toprow = `row'-(`r(r)'-1)
	putexcel E`toprow':E`row' = "`p'", merge hcenter vcenter
	
local row = `row' + 1	
}

*Footer
putexcel A`row':E`row' = "* – p-value derived from two-sided Fisher's Exact Test", merge border(top bottom)
local row = `row' + 1
putexcel A`row':E`row' = "p-values derived from two-sided Chi-squared tests, unless otherwise noted", merge
