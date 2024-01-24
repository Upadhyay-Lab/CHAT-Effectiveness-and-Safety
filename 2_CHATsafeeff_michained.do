*********************************************
* Title: CHAT Safety and Effectiveness
* Purpose: MI Chained
* Do file order: 2
* Created: 10/31/22
* Last updated: 1/23/24
* Created by: Leah
* Last updated by: Leah
*********************************************
local tables = "~/tables/"
local dta = "~/data"
local log = "~/log"

local date : display %tdCYND date(c(current_date), "DMY")
*********************************************
*Log
*********************************************
*Log
cd "`log'"
capture log using chat_se_mi_`date'.smcl, replace

cd "`dta'"
use "CHATclean.dta", clear

capture mi unset

preserve

keep if exclude==0

save "1_includeonly.dta", replace

local imputevarlist = "intany noint anyintref intproc intmed intectopic contpg contpgwoint saeany saeanyrcd saeblood saesurg saehosp csrace5 csrace4 csab2 csbirth2 cpgdur6 cpgdur4 cpgdur3 casetotalmiso2 seq2acasetotalmiso age4under18 cpgdur cpgdur63d seq2fedhosp anyintref"

foreach var in `imputevarlist' {
	replace `var ' = . if `var'==99
	tab `var',m
	}

****************************
*Open Data
****************************
mi set wide
mi register imputed `imputevarlist'
mi register regular urban2 csscreenus age5 async seq2usurvey c1q6zipcode clinic c1q9abfund

mi impute chained ///
(logit) intany noint anyintref intproc intmed intectopic contpg contpgwoint saeany saeanyrcd saeblood saesurg saehosp csab2 csbirth2 seq2fedhosp anyintref ///
(mlogit) csrace5 csrace4 ///
(ologit) seq2acasetotalmiso cpgdur6 cpgdur4 ///
= urban2 csscreenus age5 async seq2usurvey c1q6zipcode clinic c1q9abfund ///
, add(100) rseed(1234) dots augment noisily noimp

save 2_chat_imputed.dta, replace

restore
