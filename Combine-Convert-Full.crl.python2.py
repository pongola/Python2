# -*- coding: utf-8 -*-
#############################################################################
# This script combines automatically detects whether Spinner .jr6 and .txt and/or Squid .sam and sample files from a site are present
# and converts and/or combines any files that are present into Rapid Squid file format for use in PaleoMagX/Paleomag 3.1b3.
# This script also sorts all combined steps in the proper order.

  # Note: This program appends a "p" at the end of the name of any Spinner steps.
  # Note: This program uses site dec of 0.0 and uses data that forces all sample declinations to zero. See flowchart for details. 

# Script mostly written by Casey Luskin based upon an original core written by Michiel de Kock
# For support, please contact Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com
#
#############################################################################

import os
import re
import sys
from shutil import copy

postfix = ['th','st','nd','rd','th','th','th','th','th','th']
stepsort_dict = {'NRM':1, 'AF':2, 'TT':3}

os.system('mode con: cols=150 lines=60')
scriptdir = os.path.dirname(os.path.realpath(__file__))

backupdirnum = 0 # Calculate Backup Dir Name
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

print('\n' + "Hello. This script is ready to convert/combine your Jr6 Spinner and Rapid Squid data files for a single data.")
print ("\nWARNING: This script will OVERWRITE any Rapid Squid datafiles in this folder. It is recommended that you back up your original datafiles before proceeding...")
print (" - HOWEVER: This script will automatically back up your original datafiles into a backup folder, \"%s\"." % backupdirname)
sitename = raw_input('\n' + "Please enter the site name: ")

sitename = sitename.upper()
samname = sitename + ".sam"
txtname = sitename + ".txt"
jr6name = sitename + ".jr6"
sampresent = os.path.isfile(samname)
txtpresent = os.path.isfile(txtname)
jr6present = os.path.isfile(jr6name)
dosquid = sampresent
dospinner = (txtpresent and jr6present)
squidconvertmode = False
spinnerconvertmode = False
combinemode = False

if not(dosquid) and not(dospinner):
    if not(txtpresent) and not(jr6present):
        print('\n' + "No sam file, no txt file, and no jr6 file present for site %s. Program cannot run." % sitename)
    elif not(txtpresent):
        print('\n' + "No sam file present and no txt file present for site %s. Program cannot run." % sitename)
    else:
        print('\n' + "No sam file present and no jr6 file present for site %s. Program cannot run." % sitename)
    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
    sys.exit()
elif dosquid and not(dospinner):
    squidconvertmode = True
    if not(txtpresent):
        print ('\n' + "Sam file present but no txt file present for site %s." % sitename)
    else:
        print ('\n' + "Sam file present but no jr6 file present for site %s." % sitename)
    print('\n' + "Program will run in Convert Squid Only Mode.")
elif not(dosquid) and dospinner:
    spinnerconvertmode = True
    print('\n' + "Txt file and jr6 files present, but no sam file present for site %s." % sitename)
    print('\n' + "Program will run in Convert Spinner Only Mode.")
else:
    combinemode = True
    print ('\n' + "Sam file, txt file, and jr6 file all found for site %s." % sitename)
    print('\n' + "Program will run in Combine Squid and Spinner Mode.")

samplelist = [] # Create the list of samples
if squidconvertmode or combinemode:
    f = open(samname,'r') #Opens the .sam file for reading to get sample names.
    first_line = f.readline()
    first_line_List = first_line.split()
    if first_line_List[0] == 'CIT':   # some .sam files have a "CIT" at the first line and the first sample isn't until the 4th line; others have the first sample on the third line. This does a check for which kind.
        squidstartlist = 2
    else:
        squidstartlist = 1
    for linenum, line in enumerate(f):
        linelist = line.split()
        if linenum >= squidstartlist:
            samnamenew = linelist[0].upper()
            samplelist.append(samnamenew)
    f.close()
if spinnerconvertmode or combinemode:
    f = open (jr6name,'r') #Opens the jr6 file for reading
    for line in f:
        linelist = line.split()
        if linelist != []:
            if sitename == linelist[0]:
                samnamenewspace = linelist[0],linelist[1]
                samnamenew = ''.join(samnamenewspace)
                samnamenew = samnamenew.upper()
                samplelist.append(samnamenew)
            else:
                samnamenew = linelist[0].upper()
                samplelist.append(samnamenew)
    f.close()
    
sampleset = set(samplelist)
samplelist = list(sampleset)
samplelistsorted = sorted(samplelist)
n = len(samplelistsorted)

print ("\n-----------------------------------------------\n")

if n > 1:
    print ("%s samples found in site %s. " % (n, sitename) + "Here are their names:")
elif n == 1:
    print ("%s sample found in site %s. " % (n, sitename) + "Here is its name:")
print (samplelistsorted)

print ("\n-----------------------------------------------\n")

# Create backup folder:
os.makedirs(backupdirname)
print("Backup folder \"%s\" created." % backupdirname)
backupdirpath = scriptdir + "\\" + backupdirname

# Save data files in backup folder:
if squidconvertmode or combinemode: # copy any Squid files
    copy(samname, backupdirpath) 
    for filename in samplelistsorted:
        if os.path.isfile(filename):
            copy(filename, backupdirpath)

if spinnerconvertmode or combinemode: # copy any Spinner files
    copy(jr6name, backupdirpath)
    copy(txtname, backupdirpath)

print ("*** IMPORTANT NOTE: Original sample data files have been automatically backed up to folder \"%s\"." % backupdirname)

# Create new sam file:
tempsamname = samname + ".temp"
if os.path.isfile(tempsamname):
    os.remove(tempsamname)
if squidconvertmode or combinemode:
    os.rename(samname,tempsamname)

    oldsam = open(tempsamname,'r') #Opens the .sam file for reading
    newsam = open(samname,'w') #Creates a new .sam file for reading

    first_line = oldsam.readline()
    newsam.write(first_line)

    if squidstartlist == 2:   # some .sam files have a "CIT" at the first line and the first sample isn't until the 4th line; others have the first sample on the third line. This does a check for which kind.
        second_line = oldsam.readline()
        newsam.write(second_line)

    dataline = oldsam.readline()
    dataline_list = dataline.split()

    sitelat = dataline_list[0]
    sitelon = dataline_list[1]
    latspacer = (5 - (len(sitelat)) ) * " "
    lonspacer = (6 - (len(sitelon)) ) * " "
    
    dataline_write = latspacer + sitelat + lonspacer + sitelon + "   0.0"
    newsam.write(dataline_write)

    for samplename in samplelistsorted:
        newsam.write("\n" + samplename)
    newsam.write("\n")
    oldsam.close()
    newsam.close()
    if os.path.isfile(tempsamname):
        os.remove(tempsamname)

if spinnerconvertmode: 
    newsam = open(samname,'w') #Creates a new .sam file for reading
    newsam.write(sitename + '\n')

    entersitelat = raw_input('\n' + "Please enter the site latitude: ")
    try:   # verify that the site latitude entered is a number.
        float(entersitelat)
    except ValueError:
        enterfloat = False
        while not(enterfloat):
            entersitelat = raw_input('\n'+ "Invalid response. Site latitude must be a number." + '\n' + "Please re-enter the site latitude: ")
            try:
                float(entersitelat)
                enterfloat = True
            except ValueError:  
                enterfloat = False

    entersitelon = raw_input('\n' + "Please enter the site longitude: ")
    try:   # verify that the site longitude entered is a number.
        float(entersitelon)
    except ValueError:
        enterfloat = False
        while not(enterfloat):
            entersitelon = raw_input('\n'+ "Invalid response. Site longitude must be a number." + '\n' + "Please re-enter the site longitude: ")
            try:
                float(entersitelon)
                enterfloat = True
            except ValueError:  
                enterfloat = False

    latspacer = (5 - (len(entersitelat)) ) * " "
    lonspacer = (6 - (len(entersitelon)) ) * " "
    
    dataline_write = latspacer + entersitelat + lonspacer + entersitelon + "   0.0"
    newsam.write(dataline_write)

    for samplename in samplelistsorted:
        newsam.write("\n" + samplename)
    newsam.write("\n")
    newsam.close()

# Load data from data files:
class Createnewsample:
    def __init__(sample,newsamplename):
        sample.name = newsamplename
        sample.firstline = ""
        sample.data = []
        sample.coreplatestrike = ""
        sample.coreplatedip = ""
        sample.beddingstrike = ""
        sample.beddingdip = ""
        sample.lastdatasource = ""
        sample.sampfilepresent = False
        return

sampledatalist = [] # Create sample data list
for samplename in samplelistsorted:
    sampledatalist.append(Createnewsample(samplename))

if squidconvertmode or combinemode:
    for sample in sampledatalist:
        if os.path.isfile(sample.name):
            sample.lastdatasource = "Squid"
            sample.sampfilepresent = True
            samplefile = open(sample.name,'r')
            sample.firstline = samplefile.readline()
            secondline = samplefile.readline()
            sample.coreplatestrike = secondline[8:13].replace(" ","")
            sample.coreplatedip = secondline[14:19].replace(" ","")
            sample.beddingstrike = secondline[20:25].replace(" ","")
            sample.beddingdip = secondline[26:31].replace(" ","")
            for linenum, dataline in enumerate(samplefile):
                dataline_list = dataline.split()
                if (dataline_list[0] == "NRM") or ("p" in dataline_list[0]) or (re.sub('[0-9]', '', dataline_list[0]) in ["NRMp","Tp","T","T.","AFp","A","Ap","A."]):
                    dmagstep = dataline_list[0]
                else:
                    dmagstep = dataline_list[0] + dataline_list[1]
                    
                dmagstep_mod = re.sub('[0-9]', '', dmagstep)
                if dmagstep_mod in ["TT","Tp","T","T."]:
                    dmagstep_type = "TT"
                    dmagstep_num = int(re.sub('[^0-9]','', dmagstep))
                elif dmagstep_mod in ["AF","AFp","A","Ap","A."]:
                    dmagstep_type = "AF"
                    dmagstep_num = int(re.sub('[^0-9]','', dmagstep))
                elif dmagstep_mod in ["NRM","NRMp"]:
                    dmagstep_type = "NRM"
                    dmagstep_num = 0
                else:
                    dmagstep_type = dmagstep_mod
                sortpriority = stepsort_dict[dmagstep_type]
                newstep = (sortpriority,dmagstep_type,dmagstep_num,"Squid",dataline)
                sample.data.append(newstep)
            samplefile.close()

yesdashes = False
nodashes = False

yesspaces = False
nospaces = False

nogeog = False
yesgeog = False

notilt = False
yestilt = False

if spinnerconvertmode or combinemode:
    for sample in sampledatalist:
        jr6file = open(jr6name,'r')
        samplefound = False
        for dataline in jr6file:
            dataline_list = dataline.split()
            if dataline_list[0] == sample.name:
                samplefound = True
                sample.lastdatasource = "Spinner"
                cpsnew = dataline[41:44]
                sample.coreplatestrike = str(float(cpsnew))

                cpdnew = dataline[45:48]
                sample.coreplatedip = str(float(cpdnew))

                bsnew = dataline[49:52]
                sample.beddingstrike = str(float(bsnew))

                bdnew = dataline[53:56]
                sample.beddingdip = str(float(bdnew))
                break

        dmagstep = "NRM"
        if samplefound:
            measurementnum = 0
            try:
                with open(txtname) as txtfile:
                    for line1 in txtfile:
                        line2 = txtfile.next()
                        line3 = txtfile.next()
                        line4 = txtfile.next()
                        line5 = txtfile.next()
                        line6 = txtfile.next()
                        line7 = txtfile.next()
                        line8 = txtfile.next()
                        line9 = txtfile.next()
                        line10 = txtfile.next()
                        line11 = txtfile.next()
                        line12 = txtfile.next()
                        line13 = txtfile.next()
                        line14 = txtfile.next()
                        line15 = txtfile.next()
                        line16 = txtfile.next()
                        line17 = txtfile.next()
                        line18 = txtfile.next()
                        line19 = txtfile.next()
                        line20 = txtfile.next()
                        line21 = txtfile.next()
                                        
                        measurementnum = measurementnum + 1
                        if 11 <= measurementnum % 100 <= 19:
                            postfixindex = 0
                        else:
                            postfixindex = measurementnum % 10
                        
                        words1 = line2.split()
                                   
                        if words1: # Check if line is empty and there's an error.
                            if sitename == words1[0]:  # Automatically Determine if spaces are within samples names
                                yesspaces = True
                                # check for words1[1]
                                specnamespace = words1[0],words1[1]
                                specimen = ''.join(specnamespace)
                                dmagstepindex = 3
                            elif words1[0] in samplelistsorted:
                                nospaces = True
                                specimen = (words1[0])
                                dmagstepindex = 2
                            elif not (words1[0] in samplelistsorted): # Check first word isn't a sample name, meaning there's an error.
                                print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find a sample name but instead found: '%s'" % words1[0])
                                print ("The format error occurred in the %s%s measurement of the .txt file." % (measurementnum, postfix[postfixindex]))
                                endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                sys.exit()
                        else: # Check if there's an empty line instead of the filename.
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find a sample name but instead found an empty line.")
                            print ("The format error occurred in the %s%s measurement of the .txt file." % (measurementnum, postfix[postfixindex]))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()

                        if len(words1) <= (dmagstepindex - 1): # Check if a demag step is present. 
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find a Demag Step but found nothing there.")
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s" % (measurementnum, postfix[postfixindex], specimen))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()

                        if '-' in specimen:
                            dashes = True
                        else:
                            nodashes = True

                        dmagstep = words1[dmagstepindex]
                        dmagstepspace = (6 - len(dmagstep)) * " "
                            
                        words2 = line14.split()

                        if not(words2):   # check if words2 is empty, if so meaning there's an error in .txt file)
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find specimen Modulus measurements but instead found an empty line.")
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()
                        elif not(words2[0] == 'Modulus'):  # check if words3 does not begin with 'Modulus' again meaning there's an error in .txt file)
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find 'Modulus' but instead found: '%s'" % words2[0])
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()
                      
                        if len(words2) <= 1: # Check if a modulus measurement is present.
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find modulus measurements name but found nothing there.")
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()

                        modulus = words2[1]

                        try: # Check if the modulus measurement is actually a number.
                            float(modulus)     
                        except ValueError: 
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find modulus found a non-number: %s." % modulus)
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()          
                                  
                        modnum = float(modulus)
                        modint = int(modnum)
                        lenmodint = len(str(modint))
                        powerlenmodint = lenmodint - 1
                        modreduced = float(modnum / (10**powerlenmodint))
                        modfinal = str("%.2f" % round(modreduced,2))

                        if len(words2) <= 2: # Check if a modulus power is present.
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find modulus power but found nothing there.")
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()
                       
                        power = words2[2]
                        power10 = str(power.strip('A/m'))
                        
                        try: # Check if modulus power is a number.
                            powernum = int(power10.strip('E-'))
                        except ValueError: 
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find modulus power but found a non-number: %s." % str(power10.strip('E-')))
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()
                        
                        powernum = powernum + 2 - powerlenmodint
                       
                        intensity = "%sE-0%s" % (modfinal, powernum)

                        if len(words2) <= 4: # Check if precision measurement is present.
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find precision but found nothing there.")
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()

                        error = words2[4]
                        error = error[:-1]
                        errorspace = (6 - len(error) ) * " "
                        
                        try: # Check if precision measurement is a number.
                            float(error)     
                        except ValueError:
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find precision but found a non-number: %s." % error)
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()

                        words3 = line18.split()

                        if not(words3):   # check if words3 is empty, if so meaning there's an error in .txt file)
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Core measurements but instead found an empty line.")
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()
                        elif not(words3[0] == 'SPEC.'):  # check if words3 does not begin with 'SPEC' again meaning there's an error in .txt file)
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Core measurements name but instead found: '%s.'" % words3[0])
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()

                        if len(words3) <= 2: # Check if a Core Dec measurement is present.
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Core Dec but found nothing there.")
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()
                            
                        coredecl = (words3[2])

                        try: # Check if Core Dec measurement is a number.
                            float(coredecl)     
                        except ValueError:
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Core Dec but found a non-number: %s." % coredecl)
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()
                            
                        cD = str(coredecl)
                        cDspace = (4 - len(cD) ) * " "

                        if len(words3) <= 3: # Check if a Core Inc measurement is present.
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Core Inc but found nothing there.")
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()
                            
                        coreincl = (words3[3])

                        try: # Check if Core Inc measurement is a number.
                            float(coreincl)     
                        except ValueError:
                            print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Core Inc but found a non-number: %s." % coreincl)
                            print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                            endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                            sys.exit()
                            
                        cI = str(coreincl)
                        cIspace = (6 - len(cI) ) * " "

                        geogwords = line19.split() # Determine if there is geographic coordinates for a given sample
                        
                        if not(geogwords): # check if line is empty, meaning there's an error in .txt file)
                                print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Unexpectedly found an empty line.")
                                print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                sys.exit()

                        usegeog = False
                        if geogwords[0] == 'GEOGR.S.':
                            usegeog = True
                            yesgeog = True
                            line22 = txtfile.next()   # process to line22 if geographic coordinates present. 
                        
                            if len(geogwords) <= 1: # Check if a Geographic Dec measurement is present.
                                print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Geographic Dec but found nothing there.")
                                print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                sys.exit()
                            
                            geodecl = (geogwords[1])

                            try: # Check if Geographic Dec measurement is a number.
                                float(geodecl)     
                            except ValueError:
                                print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Geographic Dec but found a non-number: %s." % geodecl)
                                print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                sys.exit()
                            
                            if len(geogwords) <= 2: # Check if a Geographic Inc measurement is present.
                                print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Geographic Inc but found nothing there.")
                                print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                sys.exit()
                            
                            geoincl = (geogwords[2])

                            try: # Check if Geographic Inc measurement is a number.
                                float(geoincl)     
                            except ValueError:
                                print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Geographic Inc but found a non-number: %s." % geoincl)
                                print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                sys.exit()

                        else:
                            nogeog = True
                            print("Warning: Sample %s at step %s does not have geographic coordinates." % (specimen, dmagstep))
                            geodecl = coredecl
                            geoincl = coreincl

                        gD = str(geodecl)

                        gI = str(geoincl)
                        gIspace = (6 - len(gI) ) * " "

                        if geogwords[0] == 'TILT':
                            tiltwords = geogwords
                            line22 = txtfile.next()   # process to line22 if tilt coordinates present but geographic coordinates not present.
                            yestilt = True
                        else:
                            tiltwords = line20.split() # Determine if there is a tilt correction and use line 23 if there is a tilt correction
                                        
                        if usegeog and not(tiltwords): # check if line is empty, meaning there's an error in .txt file)
                                print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Unexpectedly found an empty line.")
                                print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                sys.exit()

                        if tiltwords:
                            if tiltwords[0] == 'TILT':
                                if not(geogwords[0] == 'TILT'):
                                    yestilt = True
                                    line23 = txtfile.next()   # process to line23 if tilt coordinates present after geographic coordinates
                                
                                if len(tiltwords) <= 2: # Check if a Tilt Dec measurement is present.
                                    print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Tilt Dec but found nothing there.")
                                    print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                    sys.exit()
                            
                                tiltdecl = (tiltwords[2])

                                try: # Check if Tilt Dec measurement is a number.
                                    float(tiltdecl)     
                                except ValueError:
                                    print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Tilt Inc but found a non-number: %s." % tiltdecl)
                                    print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                    sys.exit()

                                if len(tiltwords) <= 3: # Check if a Tilt Inc measurement is present.
                                    print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Tilt Inc but found nothing there.")
                                    print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                    sys.exit()
                                
                                tiltincl = (tiltwords[3])

                                try: # Check if Tilt Inc measurement is a number.
                                    float(tiltincl)     
                                except ValueError:
                                    print ('\n'+ "Advisory FORMATTING ERROR in .txt file: Expected to find Tilt Inc but found a non-number: %s." % tiltincl)
                                    print ("The format error occurred in the %s%s measurement of the .txt file: Sample %s, Step %s." % (measurementnum, postfix[postfixindex], specimen, dmagstep))
                                    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
                                    sys.exit()     
                            else:
                                notilt = True
                                tiltdecl = geodecl
                                tiltincl = geoincl
                        else:
                            notilt = True
                            tiltdecl = geodecl
                            tiltincl = geoincl    

                        tD = str(tiltdecl)
                        tDspace = (6 - len(tD) ) * " "

                        tI = str(tiltincl)
                        tIspace = (6 - len(tI) ) * " "
                   
                        dmagstep_mod = re.sub('[0-9]', '', dmagstep)    
                        if dmagstep_mod in ["TT","Tp","T","T."]:
                            dmagstep_type = "TT"
                            dmagstep_num = int(re.sub('[^0-9]','', dmagstep))
                        elif dmagstep_mod in ["AF","AFp","A","Ap","A."]:
                            dmagstep_type = "AF"
                            dmagstep_num = int(re.sub('[^0-9]','', dmagstep))
                        elif dmagstep_mod in ["NRM","NRMp"]:
                            dmagstep_type = "NRM"
                            dmagstep_num = 0
                        else:
                            dmagstep_type = dmagstep_mod
                        sortpriority = stepsort_dict[dmagstep_type]
                        dmagstep = dmagstep + "p"
                        gDspace = (10 - (len(dmagstep) + len(gD) )) * " "
                        
                        if sample.name == specimen:
                            dataline = ("%s%s%s%s%s%s%s%s%s   %s%s%s%s%s%s%s   0.00E+00 0.00E+00 0.00E+00 JR6\n" % (dmagstep, gDspace, gD, gIspace, gI, tDspace, tD, tIspace, tI, intensity, errorspace, error, cDspace, cD, cIspace, cI))

                            newstep = (sortpriority,dmagstep_type,dmagstep_num,"Spinner",dataline)
                            sample.data.append(newstep)
            except StopIteration:
                print ('\n' + "Problem encountered with sample %s, step %s." % (sample.name, dmagstep)) # do whatever you need to do with line1 alone
                continuechoice = raw_input('\n' + "----- Please press enter to continue. -----")

            txtfile.close()
        jr6file.close()
        

print("\n-----------------------------------------------\n")
print("Writing Sample Files\n")

def input_default(prompt, default):
    return raw_input("%s" % (prompt)) or default

for sample in sampledatalist:
    datalist = sample.data
    datalist = sorted(datalist, key = lambda item: item[2])
    datalist = sorted(datalist, key = lambda item: item[0])    
    sample.data = datalist
    samplefile = open(sample.name,"w")
    if squidconvertmode or combinemode:
        if combinemode and not(sample.sampfilepresent):
            samplefile.write(sample.name + "\n")
        else:
            samplefile.write(sample.firstline)
    else:
        samplefile.write(sample.name + "\n")
    
    if sample.lastdatasource == "Squid":
        originalcps = sample.coreplatestrike
        print("\nSample %s: Original core plate strike taken from Squid sample file is: %s" % (sample.name, originalcps))
        sample.coreplatestrike = input_default("\nPlease enter a new core plate strike for sample %s or hit enter to accept default value (%s): " % (sample.name, originalcps),originalcps)
        try:   # verify that the coreplatestrike entered is a number.
            float(sample.coreplatestrike)
        except ValueError:
            enterfloat = False
            while not(enterfloat):
                sample.coreplatestrike = input_default("\nInvalid response. Core plate strike must be a number.\nPlease re-enter core plate strike or or hit enter to accept default value (%s): " % (originalcps),originalcps)
                try:
                    float(sample.coreplatestrike)
                    enterfloat = True
                except ValueError:  
                    enterfloat = False

        originalbd = sample.beddingstrike
        print("\nSample %s: Original bedding strike taken from Squid sample file is: %s" % (sample.name, originalbd))
        sample.beddingstrike = input_default("\nPlease enter a new bedding strike for sample %s or hit enter to accept default value (%s): " % (sample.name, originalbd),originalbd)
        try:   # verify that the beddingstrike entered is a number.
            float(sample.beddingstrike)
        except ValueError:
            enterfloat = False
            while not(enterfloat):
                sample.beddingstrike = input_default("\nInvalid response. Bedding strike must be a number.\nPlease re-enter bedding strike or or hit enter to accept default value (%s): " % (originalbd),originalbd)
                try:
                    float(sample.coreplatestrike)
                    enterfloat = True
                except ValueError:  
                    enterfloat = False
        
    cps = float(sample.coreplatestrike)
    cps_space = (6 - len(str(cps)) ) * " "
    
    cpd = float(sample.coreplatedip)
    cpd_space = (6 - len(str(cpd)) ) * " "
    
    bs = float(sample.beddingstrike)
    bs_space = (6 - len(str(bs)) ) * " "
    
    bd = float(sample.beddingdip)
    bd_space = (6 - len(str(bd)) ) * " "

    samplefile.write("      0%s%.1f%s%.1f%s%.1f%s%.1f   1.0\n" % (cps_space,cps,cpd_space,cpd,bs_space,bs,bd_space,bd) )

    for step in sample.data:
        samplefile.write(step[4])

    samplefile.close()

print("All Sample Files Written Successfully!")    
print ("\n-----------------------------------------------\n")
print ("Other Notes:")

if yestilt and not(notilt):
    print ('\n' + "All samples in your site have tilt corrected data.")
elif not(yestilt) and notilt:
    print ('\n' + "Your site does not have any tilt corrected data.")
elif yestilt and notilt:
    print ('\n' + "Your site datafile has a mix of samples with and without a tilt correction.")

if yesgeog and not(nogeog):
    print ("All samples in your site have data in geographic coordinates.")
elif not(yesgeog) and nogeog:
    print ("Your datafile does not has geographic coordinates..")
elif yesgeog and nogeog:
    print ("Your site datafile has a mix of samples with and without geographic coordinates.")
    
if yesspaces and not(nospaces):
    print ("Your data file use spaces within sample names.")
elif not(yesspaces) and nospaces:
    print ("Your data files do NOT use spaces within sample names.")
elif yesspaces and nospaces:
    print ("Your data files use a mix of some samples names that have spaces and some that do NOT have spaces.")

if yesdashes and not(nodashes):
    print ("Your data files use dashes within sample names.")
elif not(yesdashes) and nodashes:
    print ("Your data files do NOT use dashes within sample names.")
elif yesdashes and nodashes:
    print ("Your data files use a mix of some samples names that have dashes and some that do NOT have dashes.")
    
endchoice = raw_input('\n' + "----- Program complete. Goodbye! Please press enter to exit. -----")
