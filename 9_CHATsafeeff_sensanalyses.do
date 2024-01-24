*********************************************
* Title: CHAT Safety and Effectiveness
* Purpose: Sensitivity anallyses
* Created: 9/8/23
* Last updated: 1/23/24
* Created by: Leah
* Last updated by: Leah
*********************************************
cls

*********************************************
*Macros
*********************************************

local tables = "/Users/koenig/Library/CloudStorage/Box-Box/CHAT Study/Papers and Manuscripts/CHAT Safety and Effectiveness Paper/Analysis/tables/"
local dta = "/Users/koenig/Library/CloudStorage/Box-Box/CHAT Study/Data/data/Safety and Effectiveness/"
local log = "/Users/koenig/Library/CloudStorage/Box-Box/CHAT Study/Papers and Manuscripts/CHAT Safety and Effectiveness Paper/Analysis/log"

local reglist = "i.async i.age5 i.csrace5 i.urban2 i.csab2 i.csbirth2 i.cpgdur6 i.csscreenus"
local date : display %tdCYND date(c(current_date), "DMY")
*********************************************
cd "`dta'"
use 2_chat_imputed.dta, clear

*Log
cd "`log'"
capture log using chat_se_sensanalyses_`date'.smcl, replace

*Sensitivity analysis with survey participants only
mi estimate: logistic noint `reglist' if exclude == 0 & surveyexclude==0 & seq2usurvey == 1
est sto intsurv
mimrgns, post expression(exp(predict(xb))/(1+exp(predict(xb))))	
est restore intsurv
mimrgns i.async, post expression(exp(predict(xb))/(1+exp(predict(xb))))	

mi estimate: logistic saeanyrcd i.async if exclude == 0 & surveyexclude==0 & seq2usurvey == 1
est sto saesurv
mimrgns, post expression(exp(predict(xb))/(1+exp(predict(xb))))	
est restore saesurv
mimrgns i.async, post expression(exp(predict(xb))/(1+exp(predict(xb))))	

*Sensitivity analysis all referred to in-person care coded as complication
local reglist = "i.async i.age5 i.csrace5 i.urban2 i.csab2 i.csbirth2 i.cpgdur6 i.csscreenus"
mi estimate: logistic anyintref `reglist' if exclude == 0 
est sto sensref
mimrgns, post expression(exp(predict(xb))/(1+exp(predict(xb))))	
est restore sensref
mimrgns i.async, post expression(exp(predict(xb))/(1+exp(predict(xb))))	

*ED visits
local reglist = "i.async i.age4 i.csrace4 i.urban2 i.csab2 i.csbirth2 i.cpgdur3 i.csscreenus"
mi estimate: logistic seq2fedhosp `reglist' if exclude == 0 
est sto edvisit
mimrgns, post expression(exp(predict(xb))/(1+exp(predict(xb))))	
est restore edvisit
mimrgns i.async, post expression(exp(predict(xb))/(1+exp(predict(xb))))	

capture log c
