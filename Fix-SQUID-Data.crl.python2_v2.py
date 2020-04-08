# -*- coding: utf-8 -*-
#
##########################################################
#
# This script fixes SQUID data that was measured with a "Z" parameter with inverse sign. This led to a wrong sign on the core inclination, which led to geographic and tilt coordinates that were incorrect.  
# Script written by Casey Luskin using a kernel of code by Michiel de Kock for reading and writing SQUID data, and an Excel spreadsheet created by Michiel de Kock with proper formulas, "Core to strat coordinates_fixed.xlsx".
# For the complete mathematics on spherical coordinate transformations, see Cox and Hart's textbook, Plate Tectonics: How it Works (1986), pp. 226-228.
#
# Script written by Casey Luskin.
# For support, please contact Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com
#
##########################################################

import os
import math
import sys
import re
from shutil import copy

# Constants:
fileextention = ".sam"
yes = set(['yes','y','ye'])
no = set(['no','n'])
nospacesteplist = ["NRM","AF1000"]

# Procedures:

def rotatecoords(rotation_angle,rotation_pole_axis_lon,rotation_pole_axis_lat,dec,inc):    
    dec_rad = math.radians(dec)
    inc_rad = math.radians(inc)

    north = math.cos(dec_rad)*math.cos(inc_rad)
    east = math.sin(dec_rad)*math.cos(inc_rad)
    down = -math.sin(inc_rad)

    rotation_pole_axis_lon_rad = math.radians(rotation_pole_axis_lon) # This is the Euler Pole longitude in radians -- the point around which the coordinate are being rotated.
    rotation_pole_axis_lat_rad = math.radians(rotation_pole_axis_lat) # This is the Euler Pole latitude in radians -- the point around which the coordinate are being rotated.

    EP_x = math.cos(rotation_pole_axis_lon_rad) * math.cos(rotation_pole_axis_lat_rad) # "ep" stands for Euler Pole
    EP_y = math.sin(rotation_pole_axis_lon_rad) * math.cos(rotation_pole_axis_lat_rad)
    EP_z = -math.sin(rotation_pole_axis_lat_rad)

    rotation_angle_rad = math.radians(rotation_angle)

    R11 = (EP_x * EP_x * (1 - math.cos(rotation_angle_rad))) + (math.cos(rotation_angle_rad))
    R12 = (EP_x * EP_y * (1 - math.cos(rotation_angle_rad))) - (EP_z * math.sin(rotation_angle_rad))
    R13 = (EP_x * EP_z * (1 - math.cos(rotation_angle_rad))) + (EP_y * math.sin(rotation_angle_rad))
    
    R21 = (EP_y * EP_x * (1 - math.cos(rotation_angle_rad))) + (EP_z * math.sin(rotation_angle_rad))
    R22 = (EP_y * EP_y * (1 - math.cos(rotation_angle_rad))) + (math.cos(rotation_angle_rad))
    R23 = (EP_y * EP_z * (1 - math.cos(rotation_angle_rad))) - (EP_x * math.sin(rotation_angle_rad))

    R31 = (EP_z * EP_x * (1 - math.cos(rotation_angle_rad))) - (EP_y * math.sin(rotation_angle_rad))
    R32 = (EP_z * EP_y * (1 - math.cos(rotation_angle_rad))) + (EP_x * math.sin(rotation_angle_rad))
    R33 = (EP_z * EP_z * (1 - math.cos(rotation_angle_rad))) + (math.cos(rotation_angle_rad))

    north_rot = (R11 * north) + (R12 * east) + (R13 * down)
    east_rot  = (R21 * north) + (R22 * east) + (R23 * down)
    down_rot  = (R31 * north) + (R32 * east) + (R33 * down)

    if north_rot > 0:
        lon_rad_rot = math.atan(east_rot/north_rot)
    else:
        lon_rad_rot = math.pi + math.atan(east_rot/north_rot)

    lat_rad_rot = -math.atan(down_rot/(math.sqrt(north_rot**2+east_rot**2)))

    lon_rot = math.degrees(lon_rad_rot) % 360
    lat_rot = math.degrees(lat_rad_rot)
    
    newdirs = [lon_rot,lat_rot]
    
    return(newdirs)

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

print('\n' + "Hello. This script is ready to fix your SQUID data that was measured with a \"Z\" parameter with the wrong sign.")
print("- It will correct the sign on your core inclination measurements and recalculate geographic and tilt-corrected directions.")
print ('\n' + "PLEASE NOTE--This script assumes/requires that:")
print ("  1. You have a full set of .sam and sample files from the SQUID for the entire site in one folder.")
print ("  2. Any SQUID-derived data in sample files has a core inclination measurement with the WRONG sign.")
print ("  3. Any NON-SQUID-derived data has \"JR6\" at the every end of the line, indicating it was measured on the spinner.")
print ("  4. Tilt correction is optional--it may or may not be present in your datafiles.")
print ("\nWARNING: This script will OVERWRITE any datafiles in this folder. It is recommended that you back up your original datafiles before proceeding...")
print (" - HOWEVER: This script will automatically back up your original datafiles into a backup folder, \"%s\"." % backupdirname)

sitename = raw_input('\n' + " --> Please enter the site name: ")
sitename = sitename.upper()

samname = str(sitename + fileextention)
sampresent = os.path.isfile(samname)

if not(sampresent):
    print ('\n' + "ERROR: This folder contains no .sam file named %s, which is necessary for this program to run." % samname)
    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
    sys.exit()
else:
    print ("\nGood news! The .sam file %s has been found!" % (samname))
    print ("\n----------------------------")

# Now to read the sample names from the .sam file

samList = [ ]
f = open (samname,'r') #Opens the .sam file for reading

first_line = f.readline()
first_line_List = first_line.split()
if first_line_List[0] == 'CIT':   # some .sam files have a "CIT" at the first line and the first sample isn't until the 4th line; others have the first sample on the third line. This does a check for which kind.
    startlist = 3
else:
    startlist = 2

f.seek(0)
counter = 0
for line in f:
    lineList = line.split()
    print(lineList)
    if counter == startlist - 1:
        try:
            sitedec = float(lineList[2])
        except:
            sitedec = 0
        print("Site Declination = %s" % sitedec)
    elif counter >= startlist:
        samList.append(lineList[0])
    counter = counter + 1
f.close()

samSet = set(samList)
sampleList = list(samSet)
sampleListSorted = sorted(sampleList)

n = len(sampleListSorted)
if n > 1:
    print ("\n%s samples found in site %s. " % (n, sitename) + "Here are their names:")
elif n == 1:
    print ("\n%s sample found in site %s. " % (n, sitename) + "Here is its name:")
print (sampleListSorted)
print ("\n---------------------------\n")

missingsamplefile = False
for filename in sampleListSorted:
    tempsamplefilename = ("%s.temp" % filename)
    if os.path.isfile(tempsamplefilename): # remove tempfile if it exists.
        os.remove(tempsamplefilename)
        
    if not os.path.isfile(filename):
        print ("ERROR: Missing a sample file for sample %s" % filename)
        missingsamplefile = True

if not(missingsamplefile):
    print ("All sample files present.")
else:
    print ("\nBecause one or more sample files are missing -- see above for the name(s) -- this program cannot run properly.")
    print ("  TIP: If you are absolutely sure no sample files are missing, try deleting problematic sample file names from the .sam file.")
    endchoice = raw_input("\n----- This program will now end. Please press enter to exit. -----")
    import sys
    sys.exit()

os.makedirs(backupdirname)
print("Backup folder \"%s\" created." % backupdirname)
backupdirpath = scriptdir + "\\" + backupdirname
for filename in sampleListSorted:
    copy(filename, backupdirpath) # Save samples in backup folder
copy(samname, backupdirpath) # Save samples in backup folder
print ("*** IMPORTANT NOTE: Original sample data files have been automatically backed up to folder \"%s\"." % backupdirname)
print ("\n----------------------------\n")

showdatachoice = ''
showdatachoice = raw_input(" --> Do you want to print data to the screen? (Y/N) ").lower()
while not ((showdatachoice in yes) or (showdatachoice in no)):
    print("\nINVALID RESPONSE: Please respond with 'yes' or 'no' (or 'y' or 'n').")
    showdatachoice = raw_input("To print data to the screen, hit Y. Otherwise hit N." + '\n' + " --> Do you want to print data to the screen? (Y/N) ").lower()
printdata = showdatachoice in yes

zerodecchoice = ''
zerodecchoice = raw_input(" --> Do you want to change the site declination in the .sam file to 0 and correct core plate strike and site strike for site declination in sample files?\n(Select YES when planning to combine with Spinner data) (Y/N) ").lower()
while not ((zerodecchoice in yes) or (zerodecchoice in no)):
    print("\nINVALID RESPONSE: Please respond with 'yes' or 'no' (or 'y' or 'n').")
    zerodecchoice = raw_input("To change the site declinationi n the same file to 0, hit Y. Otherwise hit N." + '\n' + " --> Do you want to change the site declination in the .sam file to 0 and correct core plate strike and site strike for site declination in sample files?\n(Select YES when planning to combine with Spinner data) (Y/N) ").lower()
zerodec = zerodecchoice in yes

if zerodec:
    tempsamfilename = ("%s.temp" % samname)
    os.rename(samname,tempsamfilename) # convert .sam file to to tempfile.
    f = open (tempsamfilename,'r') #Opens the .sam tempfile for reading
    newsamfile = open (samname,'a') #Creates new sam file    
    counter = 0
    f.seek(0)
    for line in f:
        if counter == startlist - 1:
            lineList = line.split()
            decline = ""
            for lineindex, item in enumerate(lineList):
                if lineindex != 2:
                    decline = decline + item + " "
                else:
                    decline = decline + "0.0" + " "
            decline = decline + '\n'
            newsamfile.write("%s" % decline)
        else:
            newsamfile.write("%s" % line)
        counter = counter + 1
    f.close()
    newsamfile.close()
    os.remove(tempsamfilename) # remove temp sam file

print("\nProcessing datafiles...\n")

for filename in sampleListSorted:
    tempsamplefilename = ("%s.temp" % filename)
    os.rename(filename,tempsamplefilename) # convert sample files to to tempfiles.
    with open(tempsamplefilename) as f:
        tempcontent = f.readlines()

    lencontent = len(tempcontent)
    newsamplefile = open (filename,'a') #Creates new sample file

    if printdata:
        print ("------------------------------\nNew Sample: " + tempcontent[0])

    newsamplefile.write("%s" % tempcontent[0])

    paramlinetext = tempcontent[1].split()
    coreplatestrike = float(paramlinetext[1]) # core plate strike
    coreplatedip = float(paramlinetext[2]) # core plate dip
    sitestrike = float(paramlinetext[3]) # site strike
    sitedip = float(paramlinetext[4]) # site dip
    coreplatestrike_corrected = (coreplatestrike + sitedec) % 360
    sitestrike_corrected = (sitestrike + sitedec) % 360

    if zerodec:
        paramlinenew = "  "
        for lineindex, item in enumerate(paramlinetext):
            if lineindex == 1:
                newcoreplatestrike = ("%.1f" % coreplatestrike_corrected)
                spacer = (6 - len(newcoreplatestrike)) * " "
                paramlinenew = paramlinenew + spacer + newcoreplatestrike
            elif lineindex == 3:
                newsitestrike = ("%.1f" % sitestrike_corrected)
                spacer = (6 - len(newsitestrike)) * " "
                paramlinenew = paramlinenew + spacer + newsitestrike                
            else:
                spacer = (6 - len(item)) * " "
                paramlinenew = paramlinenew + spacer + item
        paramlinenew = paramlinenew + '\n'
        newsamplefile.write("%s" % paramlinenew)
    else:
        newsamplefile.write("%s" % tempcontent[1])

    if printdata:
        print ("- Core Plate Strike = %s; site declination = %s; corrected Core Plate Strike = %s" % (coreplatestrike, sitedec, coreplatestrike_corrected))
        print ("- Core Plate Dip = %s" % coreplatedip)
        print ("- Site Strike = %s; site declination = %s; corrected Site Strike = %s" % (sitestrike, sitedec, sitestrike_corrected))
        print ("- Site Dip = %s\n" % sitedip)

    if lencontent > 1:
        for linenum in range(2,lencontent):
            newline = tempcontent[linenum]
            newlineitems = newline.split()

            if len(newline.strip()) != 0:
                if newlineitems[-1] != "JR6":
                    if newlineitems[0] in nospacesteplist:
                        indexadd = 0
                        stepname = newlineitems[0]
                        stepspace = ""
                    else:
                        indexadd = 1
                        if len(newlineitems[1]) == 1:
                            stepspace = "   "
                        elif len(newlineitems[1]) == 2:
                            stepspace = "  "
                        elif len(newlineitems[1]) == 3:
                            stepspace = " "
                        stepname = newlineitems[0] + stepspace + newlineitems[1]

                    coredec_str = newlineitems[7 + indexadd]
                    coredec = float(coredec_str)
                    coredec_spacer = " " * (6 - len(coredec_str))

                    coreinc_str = newlineitems[8 + indexadd]
                    coreinc = float(coreinc_str)
                    if coreinc != 0:
                        coreinc_fixed = -coreinc # The minus sign here is what corrects the parameter error on SQUID measurements.
                    else:
                        coreinc_fixed = 0.0
                        
                    coreinc_fixed_str = str(coreinc_fixed)
                    coreinc_fixed_spacer = " " * (6 - len(coreinc_fixed_str))

                    if printdata:
                        print ("Linenum = %s, Core Dec = %s, Core Inc (FIXED) = %s" % (linenum, coredec, coreinc_fixed))
                    
                    strike_rotation_angle = (90 - coreplatestrike_corrected) % 360
                    newgeogset_strike = rotatecoords(strike_rotation_angle,0,90,coredec,coreinc_fixed)

                    dip_rotation_angle = -coreplatedip
                    newgeogset = rotatecoords(dip_rotation_angle,coreplatestrike_corrected,0,newgeogset_strike[0],newgeogset_strike[1])
                    
                    geogdec = newgeogset[0]
                    geoginc = newgeogset[1]

                    dip_rotation_angle = sitedip # This should NOT be the negative of the sitedip...Michiel's excel file was wrong to make it negative.
                    newtiltset = rotatecoords(dip_rotation_angle, sitestrike_corrected, 0, geogdec, geoginc)
                    
                    tiltdec = newtiltset[0]
                    tiltinc = newtiltset[1]

                    geogdec_str = "%.1f" % geogdec
                    geogdec_spacer = " " * (6 - len(geogdec_str))

                    if stepname == "NRM":
                        geogdec_spacer = geogdec_spacer + "   "

                    geoginc_str = "%.1f" % geoginc
                    geoginc_spacer = " " * (6 - len(geoginc_str))

                    tiltdec_str = "%.1f" % tiltdec
                    tiltdec_spacer = " " * (6 - len(tiltdec_str))

                    tiltinc_str = "%.1f" % tiltinc
                    tiltinc_spacer = " " * (6 - len(tiltinc_str))

                    intensity_str = " " + newlineitems[5 + indexadd]
                    
                    error_str = newlineitems[6 + indexadd]
                    error_spacer = " " * (6 - len(error_str))

                    breaker_str = newlineitems[9 + indexadd]
                    post_breaker_str = newline.split(breaker_str,1)[1]
                    post_string = " " + breaker_str + post_breaker_str

                    newline_update = stepname + geogdec_spacer + geogdec_str + geoginc_spacer + geoginc_str + tiltdec_spacer + tiltdec_str + tiltinc_spacer + tiltinc_str + intensity_str + error_spacer + error_str + coredec_spacer+ coredec_str + coreinc_fixed_spacer + coreinc_fixed_str + post_string
                    
                    newsamplefile.write("%s" % newline_update)
                    if printdata:
                        print ("New Line %s, Text: %s" % (linenum, newline_update))                    
                else:
                    newsamplefile.write("%s" % newline) # If it was JR6 data then just write the original line without doing anything.
                    if printdata:
                        print ("JR6 Data in Line %s (No modifications made), Text: %s" % (linenum, newline))
            else:
                print("Line %s: Empty Line" % linenum)
    else:
        print("No measurement data in file %s" % filename)

    newsamplefile.close()
    
    os.remove(tempsamplefilename) # remove tempfile

print("-----------------------------------------------\n")
print("All data files processed.")
print("All tempfiles removed.")
endchoice = raw_input("\n--- Program Complete! All data converted--Have fun with your new datafiles! Please press enter to exit. ---")
