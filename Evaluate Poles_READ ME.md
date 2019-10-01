Name:	Evaluate Poles

Filename:	EvaluatePoles+PrintPoleBasemaps+EqArea.py

Author:	Casey Luskin

Summary:	This text-based Python script evaluates paleomagnetic data to determine whether a paleopole is a “verified” pole, a “key” pole, or neither, according to the quality (Q) criteria of Van der Voo (1990). It then generates equal area plots, pole maps, and an Excel spreadsheet displaying those poles from user-specified time-periods that meet certain levels of quality criteria (e.g., “verified” or “key”).

The program begins by reading an Excel spreadsheet containing paleomagnetic results of prior paleomagnetic studies on the Kaapvaal / Kalahari Craton—e.g., paleopole location, mean direction, confidence ellipses, age, as well as scores for quality (Q) criteria Q1 – Q6 of Van der Voo (1990). The quality criteria of these prior studies were scored by this author, as reported in Appendix I. The program then evaluates Q scores to determine which poles are reliable or “verified.” A pole is “verified” if Q ≥ 3, and a pole is designated a “key” pole if it has a well-constrained age of magnetization (Q1), adequate demagnetization techniques (Q3), and is based upon positive field tests and/or presence of reversals (Q4 and/or Q6). The program then outputs equal area plots, Robinson projection pole maps, and an Excel spreadsheet showing which poles are have been designated “verified”, and which are “key” poles. 

Similarity to younger poles (criterion Q7) is not scored in the Excel spreadsheet, and the program thus compares all poles to younger verified poles and scores Q7 accordingly. This is done as follows: Starting with the youngest pole with Q ≥ 2 for Q1 – Q6, a list of verified poles is created (the youngest pole with Q ≥ 2 for Q1 – Q6 will have an overall score of Q ≥ 3 for Q1 – Q7, and it thus a verified pole). Older poles are then compared to the list of younger verified poles to determine, based upon user-entered specifications, if the older pole is similar enough to the younger pole to be considered an “overprint.” As poles are deemed “verified,” then they are added to the list of verified poles and progressively older poles are compared with them when evaluating their Q7. For example, if an older pole scores 2 for Q1 – Q6, but scores 1 for Q7 when compared to younger poles, then its overall score is Q = 3 and it is added to the list of “verified poles.” Progressively ollder poles are then compared to it when evaluating their Q7. 

The user can specify time periods from which poles should be outputted. Multiple time periods can be specified and results from each time is outputted independently.
Intended Scope of use: 	This program was written for use in generating Appendix 1 of this study.
Requirements:	Python 2.7 installed on Windows or Mac with standard built-in Python libraries/modules, as well as numpy, matplotlib, Basemap, Circles, and Shapely. Note that the Circles module has been specially modified by this author to properly wrap distorted circles across edges of a Robinson Projection. 
Input:	An Excel spreadsheet named “Prior Work--Quickbook.xlsx”, which contains data from prior paleomagnetic studies, including paleopole location + K and A95, mean direction + k and α95, and Q criteria scores. The user specifies the minimum angular distance (in degrees) that must separate poles (+ α95s) in order to be considered “different” for Q7, as well as the minimum time (in Ma) that must separate poles in order to be considered an overprint for Q7. The user also specifies from which time periods (in Ma) pole data should be outputted.  
Output:	Output is generated in four formats containing poles from the time period specified by the user: 

(1) an Excel spreadsheet containing three tabs: (a) all prior poles within the time period, color-coded according to whether they are verified, key, or nonverified poles, (b) only verified poles within the specified time period, and (c) only key poles within the specified time period. 

(2) Equal area plots only showing only verified poles, only key poles, and both verified and key poles with hatching distinguishing the two. 

(3) Robinson projection pole maps showing only verified poles, only key poles, and both verified and key poles with hatching distinguishing the two pole types. 

(4) Pole numbering in (2) and (3) is kept constant for each time period, and thus a textfile “legend” is outputted explaining which poles correspond to which numbers on the diagrams in (2) and (3). 
Number of lines of code:	1,465

Other Credits:	Code for drawing properly distorted 95% confidence ellipses on Robinson Projections in Basemap Circle code adapted from https://github.com/urschrei/Circles/blob/master/README.md, DOI 10.5281/zenodo.10084

Download and Support:	For support or assistance, please contact the author Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com.
