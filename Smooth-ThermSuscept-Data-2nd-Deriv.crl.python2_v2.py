# -*- coding: utf-8 -*-
#
##############################################
# 
# This text-based Python script smooths thermomagnetic susceptibility data in a .cur file using a running 5-point average.
# It takes the average of the nearest 5 values to a given datapoint as follows: the previous two values, the next two values,
# and the value of the current data point are averaged, and the datapoint is replaced by the average. It also then calculates
# the first and second derivatives of the smoothed susceptibility data via simple slope calculations.
#
# Script written by Casey Luskin.
# For support, please contact Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com
#
##############################################

import os
import math
import sys
import re
from shutil import copy

# Constants:
fileextention = ".cur"

# Program:

os.system('mode con: cols=175 lines=60')

scriptdir = os.path.dirname(os.path.realpath(__file__))
backupdirnum = 0
for root, dirs, files in os.walk(scriptdir, topdown=False):
    for name in dirs:
        if "Backup" in name:
            newbackupnum_str = re.sub('[^0-9]', "", name)
            if (newbackupnum_str != ''):
                newbackupnum = int(newbackupnum_str)
                if newbackupnum > backupdirnum:
                    backupdirnum = newbackupnum
backupdirnum = backupdirnum + 1
backupdirnum_str = str(backupdirnum)
backupdirname = "Backup" + backupdirnum_str

print('\n' + "Hello. This script is ready to SMOOTH your Thermomagnetic Susceptibility (.cur) data.")
print ("\nWARNING: This script will OVERWRITE any datafiles in this folder. It is recommended that you back up your original datafiles before proceeding...")
print (" - HOWEVER: This script will automatically back up your original datafiles into a backup folder, \"%s\"." % backupdirname)

curfileshort = raw_input('\n' + " --> Please enter the filename name (do NOT type "".cur""): ")
curfilename = curfileshort.upper() + fileextention
curfilepresent = os.path.isfile(curfilename)

if not(curfilepresent):
    print ('\n' + "ERROR: This folder contains no file named %s, which is necessary for this program to run." % curfilename)
    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
    sys.exit()
else:
    print ("\nGood news! The file %s has been found!" % (curfilename))
    print ("\n----------------------------")

# Now to read the sample names from the .cur file

curdata = []
f = open (curfilename,'r') #Opens the .cur file for reading

for linenum, line in enumerate(f):
    if linenum == 0:
        curdata.append(line)
    else:
        tempval = float(line[0:6])
        suscval = float(line[9:15])
        firstpart = str(line[0:9])
        lastpart = str(line[15:])
        curdata.append([tempval,suscval,firstpart,lastpart])
f.close()

totlength = len(curdata)
for linenum, line in enumerate(curdata): # smooths the data
    if 2 < linenum < (totlength - 3):
        susc1 = curdata[linenum-2][1]
        susc2 = curdata[linenum-1][1]
        susc3_present = line[1]
        susc4 = curdata[linenum+1][1]
        susc5 = curdata[linenum+2][1]
        smoothedsusc = (susc1 + susc2 + susc3_present + susc4 + susc5) / 5
        curdata[linenum].append(smoothedsusc)
    elif linenum != 0:
        susc_present = line[1]
        curdata[linenum].append(susc_present)

for linenum, line in enumerate(curdata): # calculate first derivative using smoothed susceptibility value
    if 2 < linenum < (totlength - 3):
        susc2 = curdata[linenum+1][4]
        susc1 = curdata[linenum-1][4]
        temp2 = curdata[linenum+1][0]
        temp1 = curdata[linenum-1][0]
        if temp2 - temp1 == 0:
            firstderiv = 0
        else:
            firstderiv = (susc2-susc1) / (temp2-temp1)
    else:
        firstderiv = 0
    if linenum != 0:
        curdata[linenum].append(firstderiv)

hitturnaround = False
for linenum, line in enumerate(curdata): # calculate second derivative using first derivative
    if 3 < linenum < (totlength - 4):
        firstderiv2 = curdata[linenum+1][5]
        firstderiv1 = curdata[linenum-1][5]
        temp = curdata[linenum][0]
        temp2 = curdata[linenum+1][0]
        temp1 = curdata[linenum-1][0]
        if temp > 700:
            hitturnaround = True
            secondderiv = 0
        elif (temp2 - temp1) == 0:
            secondderiv = 0
        else:
            secondderiv = (firstderiv2-firstderiv1) / (temp2-temp1)
    else:
        secondderiv = 0
    if linenum != 0:
        curdata[linenum].append(secondderiv)
        curdata[linenum].append(hitturnaround)

os.makedirs(backupdirname)
print("Backup folder \"%s\" created." % backupdirname)
backupdirpath = scriptdir + "\\" + backupdirname
copy(curfilename, backupdirpath) # Save cur file in backup folder

print ("*** IMPORTANT NOTE: Original .cur data file was automatically backed up to folder \"%s\"." % backupdirname)
print ("\n----------------------------\n")

os.remove(curfilename)
smoothfilename = curfileshort + "_smooth" + fileextention
derivfilename =  curfileshort + "_deriv" + fileextention

smoothfile = open(smoothfilename,'w') #Opens a new curfile for writing
derivfile = open(derivfilename,'w') #Opens a new curfile for writing
for linenum, line in enumerate(curdata):
    if linenum == 0:
        smoothfile.write(line)

        derivline = "  TEMP    TSUSC      1stDeriv      2ndDeriv\n"
        derivfile.write(derivline)
    else:
        smoothedsusc = line[4]
        smoothedsusc_str = str("%.1f" % smoothedsusc)
        smoothedsusc_str_spaced = ((6 - len(smoothedsusc_str)) * " ") + smoothedsusc_str

        firstderiv = line[5]
        firstderiv_str = str("%.5f" % firstderiv)
        firstderiv_str_spaced = ((14 - len(firstderiv_str)) * " ") + firstderiv_str

        secondderiv = line[6]
        secondderiv_str = str("%.5f" % secondderiv)
        secondderiv_str_spaced = ((14 - len(secondderiv_str)) * " ") + secondderiv_str
      
        smoothline = line[2] + smoothedsusc_str_spaced + line[3]
        smoothfile.write(smoothline)

        derivline = line[2] + smoothedsusc_str_spaced+ firstderiv_str_spaced + secondderiv_str_spaced + '\n'

        if not line[7]:
            derivfile.write(derivline)

smoothfile.close()
derivfile.close()

print("-----------------------------------------------\n")
print("Cur data files processed, smoothed, and saved.")
endchoice = raw_input("\n--- Program Complete! Please press enter to exit. ---")
