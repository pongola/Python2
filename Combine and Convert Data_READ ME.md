Name:	Combine and Convert Data

Filename:	jr6+Sam.COMBINE-CONVERT-FULL.crl.python2.py

Author:	Casey Luskin

Summary:	This text-based Python script combines demagnetization step data from a site that is stored in multiple datafile formats into a single set of datafiles for the site written in the RAPID Squid format. It automatically detects whether .jr6 / .txt datafiles created by a AGICO JR-6A dual speed “Spinner” magnetometer, RAPID Squid data files, and .dat data files created by a 2G magnetometer are present. It then combines all steps into the proper order for individual samples within a site, and saves them in a RAPID SQUID format. 

Intended Scope of use: 	This script was originally written because machinery breakdowns required multiple different magnetometers to be used to take demagnetization measurements. It was written for use by members of the UJ Paleomagnetism library and has been used by multiple investigators at the lab.  

Requirements:	Python 2.7 installed on Windows or Mac with standard built-in Python libraries/modules. Early versions of this script were also adapted for Python 3.0.

Input:	User enters site name, and script must run in a folder containing data files for a given site. Data files must be written in jr6 / .txt data format created by a AGICO JR-6A dual speed “Spinner” magnetometer, RAPID SQUID data format, or .dat data format created by a 2G magnetometer.

Output:	Combined datafiles in RAPID SQUID data format. Original files are backed up. 

Number of lines of code:	757

Other Credits:	Some segments of filereading code adapted from a script written by Michiel de Kock.

Download and Support:	The latest version of this script can be downloaded from https://github.com/pongola/Python2. For support or assistance, please contact the author Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com.
