*********************************************
* Title: CHAT Safety and Effectiveness
* Purpose: Figure 2
* Created: 3/14/23
* Last updated: 1/23/24
* Created by: Leah
* Last updated by: Leah
*********************************************

*********************************************
*Macros
local graph = "~/graph"
*********************************************

/*
Effectiveness 

Overall: 97.7% (95% CI: 97.2%, 98.1%) 
Synchronous: 98.2% (95% CI: 97.7%, 99.0%) 
Asynchronous: 97.4% (95% CI: 96.9%, 98.0%)

Safety
Overall: 0.25% (95% CI: 0.12%, 0.37%) 
Synchronous: 0.24% (95% CI: 0.00%, 0.47%)
Asynchronous: 0.25% (95% CI: 0.10%, 0.40%) 
*/

clear
set obs 8

*Overall
gen est = .
gen lb = .
gen ub = .
gen lab = _n
	replace lab = 1 if lab==5
	replace lab = 2 if lab==6
	replace lab = 3 if lab==7
	replace lab = 4 if lab==8

gen measure = 1 in 1/4
replace measure = 2 in 5/8

*Effectiveness
	*Overall
	replace est = 97.7 in 2
	replace lb = 97.3 in 2
	replace ub = 98.1 in 2

	*Literature
	replace est = 97.4 in 1
	
	*Synchronous
	replace est = 98.2 in 3
	replace lb = 97.7 in 3
	replace ub = 98.0 in 3
	
	*Asynchronous
	replace est = 97.4 in 4
	replace lb = 96.9 in 4
	replace ub = 98.0 in 4

*Safety
	*Overall
	replace est = 99.75 in 6
	replace lb = 99.63 in 6
	replace ub = 99.88 in 6

	*Literature
	replace est = 99.7 in 5
	
	*Synchronous
	replace est = 99.76 in 7
	replace lb = 99.53 in 7
	replace ub = 100.00 in 7
	
	*Asynchronous
	replace est = 99.75 in 8
	replace lb = 99.60 in 8
	replace ub = 99.90 in 8
	

*Labels
la def lab 2 "CHAT Overall" 1 "Published estimates of in-person dispensing*" 3 "CHAT Synchronous" 4 "CHAT Asynchronous" 6 "CHAT Overall" 5 "Published estimates of in-person dispensing*" 7 "CHAT Synchronous" 8 "CHAT Asynchronous", replace
la val lab lab

gen tolabel = string(est, "%2.1f") + "%" if measure==1
replace tolabel = string(est, "%4.2f") + "%" if measure==2

*Plot
cd "`graphs'"
local gray = "235 235 233"
set scheme white_tableau

preserve
drop if measure==2
splitvallabels lab, length(12)
twoway bar est lab , barw(0.6) lcolor(none) || ///
bar est lab if lab==1, barw(0.6) bcolor(gs11) lcolor(none) || ///
rcap lb ub lab  , lcolor(black) legend(off) ///
yla(0 "0" 20 "20%" 40 "40%" 60 "60%" 80 "80%" 100 "100%", ang(h) labsize(2.5)) ///
xlab(`r(relabel)', valuelabel labsize(2.5) ) || ///
scatter est lab , msym(none) mlab(tolabel) mlabpos(12) mlabcolor(black) ytitle("") msize(2.75)  xtitle("") ///
title("Effectiveness" " ", size(med)) xtitle("") fysize(75)
graph save effectiveness.gph, replace
restore

preserve
drop if measure==1
splitvallabels lab, length(12)
twoway bar est lab , barw(0.6)  lcolor(none) || ///
bar est lab if lab==1, barw(0.6) bcolor(gs11) lcolor(none) || ///
rcap lb ub lab , lcolor(black) legend(off) ///
yla(0 "0" 20 "20%" 40 "40%" 60 "60%" 80 "80%" 100 "100%", ang(h) labsize(2.5)) ///
xlab(`r(relabel)', valuelabel labsize(2.5) ) || ///
scatter est lab , msym(none) mlab(tolabel) mlabpos(12) mlabcolor(black) ytitle("") msize(2.75)  xtitle("") ///
title("Safety" " ", size(med)) xtitle("") fysize(75)
graph save safety.gph, replace
restore

local note = `"*Published rates of in-person dispensing draw from US Food and Drug Administration Label for Mifepristone, 2016."'

graph combine effectiveness.gph safety.gph, ///
title("", pos(11) size(2.75)) note("`note'", size(2.25))

graph save "safetyefffig2.gph", replace
graph export "safetyefffig2.eps", replace
