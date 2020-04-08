# -*- coding: utf-8 -*-
###############################
#
# This script smooths IRM Acquistion Data at and above the 200 mT step
# This text-based Python script smooths thermomagnetic susceptibility data in a .cur file using a running 5-point average.
# It takes the average of the nearest 5 values to a given datapoint in the following manner: the previous two values, the
# next two values, and the value of the current data point are averaged, and the datapoint is replaced by the average.
# It also then calculates the first and second derivatives of the smoothed susceptibility data via simple slope calculations.
#
# Script written by Casey Luskin.
# For support, please contact Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com
#
###############################

import os
import math
import sys
import re
from shutil import copy

# Constants:
fileextention = ".txt"

# Program:

os.system('mode con: cols=175 lines=60')

scriptdir = os.path.dirname(os.path.realpath(__file__))
backupdirnum = 0
for root, dirs, files in os.walk(scriptdir, topdown=False):
    for name in dirs:
        if "IRMBackup" in name:
            newbackupnum_str = re.sub('[^0-9]', "", name)
            if (newbackupnum_str != ''):
                newbackupnum = int(newbackupnum_str)
                if newbackupnum > backupdirnum:
                    backupdirnum = newbackupnum
backupdirnum = backupdirnum + 1
backupdirnum_str = str(backupdirnum)
backupdirname = "IRMBackup" + backupdirnum_str

print('\n' + "Hello. This script is ready to SMOOTH your IRM Acquisition (.txt) data.")
print ("\nWARNING: This script will OVERWRITE any datafiles in this folder. It is recommended that you back up your original datafiles before proceeding...")
print (" - HOWEVER: This script will automatically back up your original datafiles into a backup folder, \"%s\"." % backupdirname)

irmfileshort = raw_input('\n' + " --> Please enter the filename name (do NOT type "".cur""): ")
irmfilename = irmfileshort.upper() + fileextention
irmfilepresent = os.path.isfile(irmfilename)

if not(irmfilepresent):
    print ('\n' + "ERROR: This folder contains no file named %s, which is necessary for this program to run." % irmfilename)
    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
    sys.exit()
else:
    print ("\nGood news! The file %s has been found!" % (irmfilename))
    print ("\n----------------------------")

# Now to read the sample names from the .txt file

irmdata = []
f = open (irmfilename,'r') #Opens the .txt file for reading

for line in f:
    linelist = line.split()
    irmval = linelist[3]
    irmval = irmval.rstrip(',')
    irmval = float(irmval)
    irmdata.append([linelist,irmval])
f.close()

totlength = len(irmdata)
for linenum, line in enumerate(irmdata): # smooths the data
    irm_present = line[1]
    irmdata[linenum].append(irm_present)
    if (19 <= linenum) and (linenum == (totlength - 2)):
        irm1 = irmdata[linenum-1][1]
        irm3 = irmdata[linenum+1][1]
        smoothedirm = (irm1 + irm_present + irm3) / 3
        irmdata[linenum][2] = smoothedirm
    elif 19 <= linenum < (totlength - 2):
        irm1 = irmdata[linenum-2][1]
        irm2 = irmdata[linenum-1][1]
        irm4 = irmdata[linenum+1][1]
        irm5 = irmdata[linenum+2][1]
        smoothedirm = (irm1 + irm2 + irm_present + irm4 + irm5) / 5
        irmdata[linenum][2] = smoothedirm

os.makedirs(backupdirname)
print("Backup folder \"%s\" created." % backupdirname)
backupdirpath = scriptdir + "\\" + backupdirname
copy(irmfilename, backupdirpath) # Save irmfile in backup folder

print ("*** IMPORTANT NOTE: Original .cur data file was automatically backed up to folder \"%s\"." % backupdirname)
print ("\n----------------------------\n")

os.remove(irmfilename)
smoothfilename = irmfileshort + "_IRMsmooth" + fileextention
smoothfile = open(smoothfilename,'w') #Opens a new irmfile for writing

for line in irmdata:
    linelist = line[0]
    smoothedirmval = line[2]

    col1 = linelist[0]
    col1_sp = "  "

    col2 = linelist[1]
    col2_sp = "  "

    col3 = linelist[2]
    col3_sp = (10 - len(col3)) * " "

    col4 = str("%.3f" % smoothedirmval) + ","
    col4_sp = (11 - len(col4)) * " "

    col5 = linelist[4]
    col5_sp = (10 - len(col5)) * " "

    col6 = linelist[5]
    col6_sp = (10 - len(col6)) * " "

    col7 = linelist[6]
      
    smoothline = col1 + col1_sp + col2 + col2_sp + col3 + col3_sp + col4 + col4_sp + col5 + col5_sp + col6 + col6_sp + col7 + '\n'
    smoothfile.write(smoothline)

smoothfile.close()

print("-----------------------------------------------\n")
print("IRM data files processed, smoothed, and saved.")
endchoice = raw_input("\n--- Program Complete! Please press enter to exit. ---")
