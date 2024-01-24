*********************************************
* Title: CHAT Safety and Effectiveness
* Purpose: Sensitivity analyses - pattern mixture modeling
* Created: 1/1/24
* Last updated: 1/23/24
* Created by: Leah
* Last updated by: Leah
*********************************************
***Help file example***
/*
use "http://www.homepages.ucl.ac.uk/~rmjwiww/stata/missing/smoke.dta", clear

tab rand quit, miss

*Analysis assuming missing = smoking:

gen quit2 = quit

replace quit2 = 0 if missing(quit)

logistic quit2 rand

*Same analysis using rctmiss:

rctmiss, pmmdelta(0, expdelta): logistic quit rand
rctmiss, pmmdelta(0.1, expdelta): logistic quit rand
rctmiss, pmmdelta(0.2, expdelta): logistic quit rand

*Sensitivity analysis based around missing=smoking:

rctmiss, sens(rand) pmmdelta(0(0.1)1, expdelta base(0)): logistic quit rand

* MNAR, assuming missing value are 5 units lower than observed in both groups
xi: rctmiss, pmmdelta(0): logistic quit rand
margins 

xi: rctmiss, pmmdelta(-.2): logistic quit rand
margins 

xi: rctmiss, pmmdelta(-.3): logistic quit rand
margins 

xi: rctmiss, pmmdelta(-.4): logistic quit rand
margins 
*/

*****
*Open data
clear
use "/Users/koenig/Library/CloudStorage/Box-Box/CHAT Study/Data/data/CHATclean.dta", clear


*Complete case for comparison 
logistic noint i.async i.age5 i.csrace5 i.urban2 i.csab2 i.csbirth2 i.cpgdur6 i.csscreenus if exclude ==0 & seq2teffdenom==1 

*Effectiveness 
macro drop _all
local n = 1
foreach d in .1 .2 .5 1 2 5 10 {

	*Overall 
	xi: rctmiss, pmmdelta(`d', expdelta): logistic noint i.async i.age5 i.csrace5 i.urban2 i.csab2 i.csbirth2 i.cpgdur6 i.csscreenus if exclude ==0
	margins, expression(exp(predict(xb))/(1+exp(predict(xb))))
	mat e = r(table)
	local d`n'e_b = string(100*e[1,1], "%4.1f")
	local d`n'e_ll = string(100*e[5,1], "%4.1f")
	local d`n'e_ul = string(100*e[6,1], "%4.1f")

	*By sync/async
	xi: rctmiss, pmmdelta(`d', expdelta): logistic noint i.async i.age5 i.csrace5 i.urban2 i.csab2 i.csbirth2 i.cpgdur6 i.csscreenus if exclude ==0
	est sto reg
	
	*RENAME FACTOR VARIABLE
	local colnms: coln e(b)
	local colnms = subinstr("`colnms'","_Iasync_1", "1.async", .)
	mat b= e(b)
	mat colnames b= `colnms'

	erepost b= b, rename
	
	qui margins i.async, expression(exp(predict(xb))/(1+exp(predict(xb))))
	mat e = r(table)
	local d`n'es_b = string(100*e[1,1], "%4.1f")
	local d`n'es_ll = string(100*e[5,1], "%4.1f")
	local d`n'es_ul = string(100*e[6,1], "%4.1f")
	
	local d`n'ea_b = string(100*e[1,2], "%4.1f")
	local d`n'ea_ll = string(100*e[5,2], "%4.1f")
	local d`n'ea_ul = string(100*e[6,2], "%4.1f")	
	
	xi: rctmiss, pmmdelta(`d', expdelta): logistic noint i.async i.age5 i.csrace5 i.urban2 i.csab2 i.csbirth2 i.cpgdur6 i.csscreenus if exclude ==0
	mat p = r(table)
	local pe`n' = string(p[4,1], "%04.3f")
	
*Safety 
	*Overall
	xi: rctmiss, pmmdelta(`d', expdelta): logistic saeanyrcd i.async if exclude ==0
	margins, expression(exp(predict(xb))/(1+exp(predict(xb))))
	mat s = r(table)
	local d`n's_b = string(100*s[1,1], "%4.1f")
	local d`n's_ll = string(100*s[5,1], "%4.1f")
	local d`n's_ul = string(100*s[6,1], "%4.1f")
	
	*By sync/async
	xi: rctmiss, pmmdelta(`d', expdelta): logistic saeanyrcd i.async if exclude ==0
	est sto reg
	
	*RENAME FACTOR VARIABLE
	local colnms: coln e(b)
	local colnms = subinstr("`colnms'","_Iasync_1", "1.async", .)
	mat b= e(b)
	mat colnames b= `colnms'

	erepost b= b, rename

	qui margins i.async, expression(exp(predict(xb))/(1+exp(predict(xb))))
	mat s = r(table)
	local d`n'ss_b = string(100*s[1,1], "%4.1f")
	local d`n'ss_ll = string(100*s[5,1], "%4.1f")
	local d`n'ss_ul = string(100*s[6,1], "%4.1f")
	
	local d`n'sa_b = string(100*s[1,2], "%4.1f")
	local d`n'sa_ll = string(100*s[5,2], "%4.1f")
	local d`n'sa_ul = string(100*s[6,2], "%4.1f")	
	
	xi: rctmiss, pmmdelta(`d', expdelta): logistic saeanyrcd i.async if exclude ==0
	mat p = r(table)
	local ps`n' = string(p[4,1], "%04.3f")
	
	local n = `n'+1
}
************
*Putexcel
************
cd "/Users/koenig/Library/CloudStorage/Box-Box/CHAT Study/Papers and Manuscripts/CHAT Safety and Effectiveness Paper/Analysis/tables"
putexcel set pmm_delta.xlsx, replace
putexcel A1:I10, fpattern(solid ,white) font("Times New Roman")
putexcel A1:I1 = "Table X. Pattern-mixture modeling with delta correction to explore potential impact of informative missingness in abortion outcomes", merge bold
putexcel A2 = "Delta correction, exp", bold
putexcel A4 = "0 (assumes all abortion missing outcomes required additional intervention/had sae)"
putexcel B2:D2 = "Effectiveness", bold merge
putexcel B3 = "Overall", bold
putexcel C3 = "Synchronous", bold
putexcel D3 = "Asynchronous", bold
putexcel E3 = "p-value", bold
putexcel F2:I2 = "Safety", bold merge
putexcel F3 = "Overall", bold
putexcel G3 = "Synchronous", bold
putexcel H3 = "Asynchronous", bold
putexcel I3 = "p-value", bold

local row = 4 
foreach i in .1 .2 .5 1 2 5 10 {
if `i'<1 {
local n = string(`i', "%04.1f")
}
if `i'>=1 {
local n = string(`i', "%4.0f")
}
putexcel A`row' = "`n'"
local row = `row'+1
}
putexcel A8 = "1"

local row = 4
local n = 1 
forvalues i=1/7 {
putexcel B`row' = "`d`n'e_b' (`d`n'e_ll', `d`n'e_ul')"
putexcel C`row' = "`d`n'es_b' (`d`n'es_ll', `d`n'es_ul')"
putexcel D`row' = "`d`n'ea_b' (`d`n'ea_ll', `d`n'ea_ul')"
putexcel E`row' = "`pe`n''"
putexcel F`row' = "`d`n's_b' (`d`n's_ll', `d`n's_ul')"
putexcel G`row' = "`d`n'ss_b' (`d`n'ss_ll', `d`n'ss_ul')"
putexcel H`row' = "`d`n'sa_b' (`d`n'sa_ll', `d`n'sa_ul')"
putexcel I`row' = "`ps`n''"
local row = `row' + 1 
local n = `n'+1
}
macro dir 
