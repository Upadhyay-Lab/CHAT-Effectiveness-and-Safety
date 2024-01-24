*********************************************
* Title: CHAT Safety and Effectiveness
* Purpose: Flow diagram
* Created: 8/8/23
* Created: Last updated
* Created by: Leah
* Last updated by: Leah
*********************************************

/*More info:

https://theesspreckelsen.wordpress.com/2014/08/23/stata-consort-flowchart/
https://theesspreckelsen.files.wordpress.com/2014/08/flowchart.pdf

*/

/* Original study:

We initially received clinical charts for 6,974 abortions provided within our study period. After excluding records for abortions where mifepristone and misoprostol were not provided (n=819) and those where the patient took neither mifepristone nor misoprostol (n=120), our analytic sample included 6,035 abortions provided to 6,020 patients. After dispensing the abortion medications, 76% (n=4,616) cases had any follow-up contact after the medications were dispensed (Figure 1). Abortion outcomes were known for 74% (4,443) of the included sample. 

*/

cls
*Definition of boxes and line styles.
local osmall = 	", box margin(small) fcolor(white) size(vsmall) justification(left)"
local omain =	", box margin(small) fcolor(white)"
local bc = ", lwidth(medthick) lcolor(black)" 
local bca = ", lwidth(medthick) lcolor(black) mlcolor(black) mlwidth(medthick) msize(medlarge)"
local topofbox = 6
local botofbox = 2.2
local boxw = 5.2
local boxl = 0

*set the scheme
set scheme white_tableau

*Define the exclusion criteria variable (you can also do this manually - I do a mix)
capture drop exclude_diagram
gen exclude_diagram = exclude if exclude!=4 & exclude!=3 & baseline==1
la val exclude_diagram exclude

*Count overall N
local N = string(6974, "%15.0fc")

*Count if excluded because no medicaitons were dispensed
local Nexcnocare = string(820, "%15.0fc")

*Count abortions provided
local Nabs = string(6154, "%15.0fc")

*Count if excluded because did not take meds
count if exclude_diagram==2 
local Nexcnomed = string(120, "%15.0fc")

*Count if included
count if exclude_diagram==0 
local Ninclude = string(`r(N)', "%15.0fc")

*Count if no follow-up
count if seq2ssafetydenom==0 & exclude_diagram==0
local Nnofu = string(`r(N)', "%15.0fc")

*Count if any follow-up
count if seq2ssafetydenom==1 & exclude_diagram==0
local Nanyfu = string(`r(N)', "%15.0fc")

*Count if any follow-up but no ab outcome
count if seq2teffdenom==0 & seq2ssafetydenom==1 
local Nnoabout = string(`r(N)', "%15.0fc")

*Count if aboutcomeknown
count if seq2teffdenom==1 & seq2ssafetydenom==1 & exclude_diagram==0
local Naboutcomeknown = string(`r(N)', "%15.0fc")

*Save labels for each value label
levelsof exclude_diagram
foreach level in `r(levels)' {
count if exclude_diagram==`level'
local n_`level' = `r(N)'
local val_`level' : label exclude_diagram `level'
di "`val_`level''"
}
macro dir

count if exclude_diagram >0 & exclude_diagram < 9
local exclude1 = `r(N)'

*Drawing the graph
cd "/Users/koenig/Library/CloudStorage/Box-Box/CHAT Study/Papers and Manuscripts/CHAT Safety and Effectiveness Paper/Analysis/graphs"
twoway  /// 1) PCI to draw a box 2) PIC horizontal lines 3) pcarrowi: connecting arrows.
   pci `boxw' `boxl' `boxw' `topofbox' `bc' || ///
   pci `boxw' `topofbox' `botofbox' `topofbox' `bc' || ////
   pci `botofbox' `topofbox' `botofbox' `boxl' `bc' || ///
   pci `botofbox' `boxl' `boxw' `boxl' `bc' ///
	|| pci 4.6 2 4.6 4.6 `bc' ///
	|| pci 3.9 2 3.9 4.5 `bc'  ///
	|| pci 3.3 2 3.3 4.5 `bc'  ///
	|| pci 2.7 2 2.7 4.5 `bc'  ///
	|| pcarrowi 5.1 2 4.35 2 `bca' ///
	|| pcarrowi 4.1 2 3.71 2 `bca'  ///
	|| pcarrowi 3.7 2 3.11 2 `bca'  ///
	|| pcarrowi 3.4 2 2.51 2 `bca'  ///	
, /// Text placed using "added text" [ACHTUNG sizes change with content]
text(5 2.1 "Clinical charts abstracted" "(n= `N')" `omain') ///
	text(4.6 3.61 "No medications dispensed (n=`Nexcnocare')" `osmall') ///
text(4.2 2.1 "Charts with abortion" "medications dispensed" "(n= `Nabs')" `omain') ///
	text(3.9 4.1 "Took neither mifepristone nor misoprostol (n=`Nexcnomed')" `osmall') ///
text(3.6 2 "Clinical charts included" "(n= `Ninclude')" `omain') ///
	text(3.3 3.45 "No follow-up contact (n=`Nnofu')"   `osmall') ///
text(3 2 "Any follow-up contact" "(n= `Nanyfu')" `omain') ///
	text(2.7 3.625 "Abortion outcome unknown (n=`Nnoabout')"   `osmall') ///
text(2.4 2 "Abortion outcome known" "(n= `Naboutcomeknown')" `omain') ///
legend(off) ///
xlabel("") ylabel("") xtitle("") ytitle("") ///
plotregion(lcolor(white)) ///
graphregion(lcolor(white)) xscale(range(0 6)) ///
xsize(2) ysize(3) /// A4 aspect ratio
note("" ///
, size(tiny)) ///
saving("safetyefffig1.gph", replace)

graph export safetyefffig1.png, replace
