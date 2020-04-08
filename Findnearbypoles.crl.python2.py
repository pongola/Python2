# -*- coding: utf-8 -*-
#
###################################################################
#
# This text-based Python script was an early iteration of what became the program "Evaluate Poles", in which the user can find prior poles on the Kaapvaal Craton that are near a user-specified pole location. 
#
# First, the user enters the latitude and longitude of a pole, as well as the pole name, and then the user enters the angular distance to search for nearby poles.
# The program then evaluates all pole locations in an Excel spreadsheet, “Prior Work--Quickbook.xlsx,” and compares the angular distance between the user-specified pole
# and each pole in the workbook. If the distance is less than the user-specified distance then a pole is deemed “nearby.”
#
# Data for each nearby pole is both outputted to the screen and encoded and saved in a GPlates Markup Language .gpml file (Qin et al., 2012).
# The resultant .gpml files can then be opened in a GPlates project (Boyden et al., 2011).
#
# Script written by Casey Luskin.
# For support, please contact Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com
#
###################################################################

import sys
import os
import math
from shutil import copy

os.system('mode con: cols=150 lines=60')

print('\n' + "Hello. This script is ready to search your Excel spreadsheet for poles near the pole you specify, within an angular distance you specify.")

excelfilename = "Prior Work--Quickbook.xlsx"
yes = set(['yes','y','ye'])
no = set(['no','n'])
badchars = "\/:*?\"<>|"

if not(os.path.isfile(excelfilename)): # check if the .excelfilename exists in this folder.
    print ('\n' + "This folder does NOT contain the file named %s, which is necessary for this program to run." % excelfilename)
    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
    sys.exit()
else:
    print ('\n' + "The Excel source file %s has been found!" % excelfilename)
    print ('\n' + "----------------------------")
    writefileschoice = ''
    writefileschoice = raw_input('\n' + "Do you want to write .gpml files of nearby poles?" + '\n' + " --> If you select YES, the script will overwrite any pre-existing nearby poles for your entered pole." + '\n' + " --> If you select NO it will only print nearby poles to the screen." + '\n' + "Do you choose to write nearby pole files? (Y/N) ").lower()
    while not ((writefileschoice in yes) or (writefileschoice in no)):
        print('\n' + "Invalid Response--please respond with 'yes' or 'no' (or 'y' or 'n').")
        writefileschoice = raw_input("To write nearby pole .gpml files, hit Y. Otherwise hit N." + '\n' + "Do you choose to overwrite prior pole files? (Y/N) ").lower()
    writefiles = writefileschoice in yes

if writefiles:
    nearbypoledir = raw_input('\n' + "Enter the name of the pole: ")
    while any((char in badchars) for char in nearbypoledir):
        nearbypoledir = raw_input('\n' + "Pole name cannot have the characters %s. Please re-enter pole name: " % badchars)
    nearbypoledir = "Poles Near " + nearbypoledir
    while os.path.exists(nearbypoledir):
        nearbypoledir = raw_input('\n' + "Directory \"%s\" already exists. Please re-enter pole name: " % nearbypoledir)
        while any((char in badchars) for char in nearbypoledir):
            nearbypoledir = raw_input('\n' + "Pole name cannot have the characters %s. Please re-enter pole name: " % badchars)
        nearbypoledir = "Poles Near " + nearbypoledir
    
enterpoleLat = raw_input('\n' + "Please enter the pole latitude: ")

try:   # verify that the pole latitude entered is a number.
    enterpoleLat = float(enterpoleLat)
except ValueError:
    enterfloat = False
    while not(enterfloat):
        enterpoleLat = raw_input('\n'+ "Invalid response. Pole latitude must be a number." + '\n' + "Please re-enter the pole latitude: ")
        try:
            enterpoleLat = float(enterpoleLat)
            enterfloat = True
        except ValueError:  
            enterfloat = False

goodlat = (abs(enterpoleLat) <= 90)

while not(goodlat):
    enterpoleLat = raw_input('\n'+ "Invalid response. Pole latitude must be between (or equal to) 90 and -90." + '\n' + "Please re-enter the pole latitude: ")
    try:   # verify that the pole latitude entered is a number.
        enterpoleLat = float(enterpoleLat)
    except ValueError:
        enterfloat = False
        while not(enterfloat):
            enterpoleLat = raw_input('\n'+ "Invalid response. Pole latitude must be a number." + '\n' + "Please re-enter the pole latitude: ")
            try:
                enterpoleLat = float(enterpoleLat)
                enterfloat = True
            except ValueError:  
                enterfloat = False
    goodlat = (abs(enterpoleLat) <= 90)

primary_poleLat = float(enterpoleLat)
rad_primary_poleLat = math.radians(float(primary_poleLat))

enterpoleLon = raw_input('\n' + "Please enter the pole longitude: ")

try:   # verify that the site longitude entered is a number.
    enterpoleLon = float(enterpoleLon)
except ValueError:
    enterfloat = False
    while not(enterfloat):
        enterpoleLon = raw_input('\n'+ "Invalid response. Pole longitude must be a number." + '\n' + "Please re-enter the pole longitude: ")
        try:
            enterpoleLon = float(enterpoleLon)
            enterfloat = True
        except ValueError:  
            enterfloat = False

goodlon = (abs(enterpoleLon) <= 360)

while not(goodlon):
    enterpoleLon = raw_input('\n'+ "Invalid response. Pole longitude must be between (or equal to) 360 and -360." + '\n' + "Please re-enter the pole longitude: ")
    try:   # verify that the pole longitude entered is a number.
        enterpoleLon = float(enterpoleLon)
    except ValueError:
        enterfloat = False
        while not(enterfloat):
            enterpoleLon = raw_input('\n'+ "Invalid response. Pole longitude must be a number." + '\n' + "Please re-enter the pole longitude: ")
            try:
                enterpoleLon = float(enterpoleLon)
                enterfloat = True
            except ValueError:  
                enterfloat = False
    goodlon = (abs(enterpoleLon) <= 360)
            
primary_poleLon = float(enterpoleLon)

if primary_poleLon < 0:
    primary_poleLon = primary_poleLon + 360

enterpoleangdist = raw_input('\n' + "Please enter the angular distance to search for nearby poles: ")

try:   # verify that the angular distance entered is a number.
    enterpoleangdist = float(enterpoleangdist)
except ValueError:
    enterfloat = False
    while not(enterfloat):
        enterpoleangdist = raw_input('\n'+ "Invalid response. Angular distance must be a number." + '\n' + "Please re-enter the angular distance to search for nearby poles: ")
        try:
            enterpoleangdist = float(enterpoleangdist)
            enterfloat = True
        except ValueError:  
            enterfloat = False

goodangdist = (0 <= enterpoleangdist <= 360)

while not(goodangdist):
    enterpoleangdist = raw_input('\n'+ "Invalid response. Angular distance must be between (or equal to) 0 and 360." + '\n' + "Please re-enter the pole longitude: ")
    try:   # verify that the pole angular distance entered is a number.
        enterpoleangdist = float(enterpoleangdist)
    except ValueError:
        enterfloat = False
        while not(enterfloat):
            enterpoleangdist = raw_input('\n'+ "Invalid response. Angular distance must be a number." + '\n' + "Please re-enter the angular distance to search for nearby poles: ")
            try:
                enterpoleangdist = float(enterpoleangdist)
                enterfloat = True
            except ValueError:  
                enterfloat = False
    goodangdist = (0 <= enterpoleangdist <= 360)
 
poleangdist = float(enterpoleangdist)

if writefiles:
    os.makedirs(nearbypoledir)        
    scriptdir = os.path.dirname(os.path.realpath(__file__))
    nearbypolepoledirpath = scriptdir + "\\" + nearbypoledir
    print('\n' + "Created directory for .gmpl files %s." % nearbypolepoledirpath)

    nicecontinentsfile = "NiceWhiteFilledContinents.gpmlz"
    nicecontinentsorigpath = ("D:\\Documents\\Dropbox\\Geo2\\Thesis\\Sampling\\Prior Work\\Gplates\\Nice Backgrounds\\%s" % nicecontinentsfile)
    nicecontinentsdestpath = nearbypolepoledirpath + "\\" + nicecontinentsfile
    copy(nicecontinentsorigpath,nicecontinentsdestpath)
    print('\n' + "Copied file %s into directory %s." % (nicecontinentsfile,nearbypolepoledirpath))

from xlrd import open_workbook
wb = open_workbook(excelfilename)
nearbypolefound = False

print ('\n' + "----------------------------")

for sheet in wb.sheets():   

    if sheet.name == "Poledata":
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        rows = []
        if poleangdist == 1:
            print ('\n' + "The following poles are within %s degree of your entered pole (Lat %s, Lon %s):" % (poleangdist, primary_poleLat, primary_poleLon))
        else:
            print ('\n' + "The following poles are within %s degrees of your entered pole (Lat %s, Lon %s):" % (poleangdist, primary_poleLat, primary_poleLon))
    #    print number_of_rows
    #    print number_of_columns

        for row in range(1, number_of_rows):
            rowitems = []
            values = []
    #        print ("row = %s" % row)
            
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                rowitems.append(value)
    #            print ("col = %s" % col)

            unit = rowitems[0]
            component = rowitems[1]
            code = rowitems[2]
            polelat = rowitems[3]
            polelon = rowitems[4]
            antipolelat = rowitems[5]
            antipolelon = rowitems[6]
            polea95 = rowitems[7]
            polek = rowitems[8]
            poledp = rowitems[9]
            poledm = rowitems[10]
            paleolat = rowitems[11]
            dec = rowitems[12]
            inc = rowitems[13]
            directiona95 = rowitems[14]
            directionk = rowitems[15]
            age = rowitems[16]
            agemod = rowitems[17]
            vandervooQ = rowitems[18]
            ref = rowitems[19]
            plateid = rowitems[27]
            platerevision = rowitems[28]

            if polea95 == '':
                if poledm != '':
                    polea95 = poledm
                elif poledp != '':
                    polea95 = poledp
                else:
                    polea95 = 0
                
     #       print ("unit = %s" % unit)
     #       print ("component = %s" % component)
     #       print ("code = %s" % code)
     #       print ("polelat = %s" % polelat)
     #       print ("polelon = %s" % polelon)
     #       print ("polea95 = %s" % polea95)
     #       print ("polek = %s" % polek)
     #       print ("poledp = %s" % poledp)
     #       print ("poledm = %s" % poledm)
     #       print ("paleolat = %s" % paleolat)
     #       print ("dec = %s" % dec)
     #       print ("inc = %s" % inc)
     #       print ("directiona95 = %s" % directiona95)
     #       print ("directionk = %s" % directionk)
     #       print ("age = %s" % age)
     #       print ("agemod = %s" % agemod)
     #       print ("ref = %s" % ref)

            age_int = int(round(age))
            age_str = str(age_int)

            polelon = float(polelon)
            polelat = float(polelat)
            
            if polelon < 0:
                polelon = polelon + 360

            if (abs(primary_poleLon - polelon) % 360) > 180:
                deltalon = 360 - (abs(primary_poleLon - polelon) % 360)
            else:
                deltalon = abs(primary_poleLon - polelon) % 360
                        
            poledist = math.degrees(math.acos((math.sin(rad_primary_poleLat) * math.sin(math.radians(polelat))) + (math.cos(rad_primary_poleLat) * math.cos(math.radians(polelat)) * math.cos(math.radians(deltalon)))))

            if poledist <= poleangdist:
                nearbypolefound = True
                if component:
                    outputgpmlfilename = ("%s.%s [%s][%s].gpml" % (age_str, unit, component, ref))
                    layername = ("%s.%s [%s][%s]" % (age_str, unit, component, ref))
                else:
                    outputgpmlfilename = ("%s.%s [%s].gpml" % (age_str, unit, ref))
                    layername = ("%s.%s [%s]" % (age_str, unit, ref))
                outputgpmlfilename = outputgpmlfilename.replace("\\", "-")
                outputgpmlfilename = outputgpmlfilename.replace("/", "-")               
                print '\n' + outputgpmlfilename
                print ("   Pole Dist = %.2f degrees (Pole Lat = %s, Pole Lon = %s, A95 = %s)" % (poledist, polelat, polelon, polea95))

                if writefiles:
                    print ("   Writing file: %s" % outputgpmlfilename)
                    if os.path.isfile(os.path.join(nearbypolepoledirpath, outputgpmlfilename)):
                        os.remove(os.path.join(nearbypolepoledirpath, outputgpmlfilename))

                    sampleFile = open(os.path.join(nearbypolepoledirpath, outputgpmlfilename), 'a')

                    sampleFile.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
                    sampleFile.write('\n' + "<gpml:FeatureCollection xmlns:gpml=\"http://www.gplates.org/gplates\" xmlns:gml=\"http://www.opengis.net/gml\" xmlns:xsi=\"http://www.w3.org/XMLSchema-instance\" gpml:version=\"1.6.0336\" xsi:schemaLocation=\"http://www.gplates.org/gplates ../xsd/gpml.xsd http://www.opengis.net/gml ../../../gml/current/base\">")
                    sampleFile.write('\n' + "    <gml:featureMember>")
                    sampleFile.write('\n' + "        <gpml:VirtualGeomagneticPole>")
                    sampleFile.write('\n' + "            <gpml:identity>%s</gpml:identity>" % plateid)
                    sampleFile.write('\n' + "            <gpml:revision>%s</gpml:revision>" % platerevision)
                    sampleFile.write('\n' + "                        <gml:name>%s</gml:name>" % layername) 
                    sampleFile.write('\n' + "            <gpml:polePosition>")
                    sampleFile.write('\n' + "                <gpml:ConstantValue>")
                    sampleFile.write('\n' + "                    <gpml:value>")
                    sampleFile.write('\n' + "                        <gml:Point>")
                    sampleFile.write('\n' + "                            <gml:pos>%s %s</gml:pos>" % (polelat, polelon))
                    sampleFile.write('\n' + "                        </gml:Point>")
                    sampleFile.write('\n' + "                    </gpml:value>")
                    sampleFile.write('\n' + "                    <gml:description></gml:description>")
                    sampleFile.write('\n' + "                    <gpml:valueType xmlns:gml=\"http://www.opengis.net/gml\">gml:Point</gpml:valueType>")
                    sampleFile.write('\n' + "                </gpml:ConstantValue>")
                    sampleFile.write('\n' + "            </gpml:polePosition>")
                    sampleFile.write('\n' + "            <gpml:reconstructionPlateId>")
                    sampleFile.write('\n' + "                <gpml:ConstantValue>")
                    sampleFile.write('\n' + "                    <gpml:value>0</gpml:value>")
                    sampleFile.write('\n' + "                    <gml:description></gml:description>")
                    sampleFile.write('\n' + "                    <gpml:valueType xmlns:gpml=\"http://www.gplates.org/gplates\">gpml:plateId</gpml:valueType>")
                    sampleFile.write('\n' + "                </gpml:ConstantValue>")
                    sampleFile.write('\n' + "            </gpml:reconstructionPlateId>")
                    sampleFile.write('\n' + "            <gpml:averageAge>0</gpml:averageAge>")
                    sampleFile.write('\n' + "            <gpml:poleA95>%s</gpml:poleA95>" % polea95)
                    sampleFile.write('\n' + "        </gpml:VirtualGeomagneticPole>")
                    sampleFile.write('\n' + "    </gml:featureMember>")
                    sampleFile.write('\n' + "</gpml:FeatureCollection>")

                    sampleFile.close()

            if antipolelon < 0:
                antipolelon  = antipolelon + 360

            if (abs(primary_poleLon - antipolelon) % 360) > 180:
                antideltalon = 360 - (abs(primary_poleLon - antipolelon) % 360)
            else:
                antideltalon = abs(primary_poleLon - antipolelon) % 360

            antipoledist = math.degrees(math.acos((math.sin(rad_primary_poleLat) * math.sin(math.radians(antipolelat))) + (math.cos(rad_primary_poleLat) * math.cos(math.radians(antipolelat)) * math.cos(math.radians(antideltalon)))))

            if antipoledist <= poleangdist:
                nearbypolefound = True
                if component:
                    antipolegpmlfilename = ("%s.%s [ANTI-POLE][%s][%s].gpml" % (age_str, unit, component, ref))
                    antipolelayername = ("%s.%s [ANTI-POLE][%s][%s]" % (age_str, unit, component, ref))
                else:
                    antipolegpmlfilename = ("%s.%s [ANTI-POLE][%s].gpml" % (age_str, unit, ref))
                    antipolelayername = ("%s.%s [ANTI-POLE][%s]" % (age_str, unit, ref))

                antipolegpmlfilename = antipolegpmlfilename.replace("\\", "-")
                antipolegpmlfilename = antipolegpmlfilename.replace("/", "-")
                print '\n' + antipolegpmlfilename
                print ("   Pole Dist = %.2f degrees (Pole Lat = %s, Pole Lon = %s, A95 = %s)" % (antipoledist, antipolelat, antipolelon, polea95))

                if writefiles:
                    print ("   Writing file: %s" % antipolegpmlfilename)
                    if os.path.isfile(os.path.join(nearbypolepoledirpath, antipolegpmlfilename)):
                        os.remove(os.path.join(nearbypolepoledirpath, antipolegpmlfilename))

                    sampleFile = open(os.path.join(nearbypolepoledirpath, antipolegpmlfilename), 'a')

                    sampleFile.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
                    sampleFile.write('\n' + "<gpml:FeatureCollection xmlns:gpml=\"http://www.gplates.org/gplates\" xmlns:gml=\"http://www.opengis.net/gml\" xmlns:xsi=\"http://www.w3.org/XMLSchema-instance\" gpml:version=\"1.6.0336\" xsi:schemaLocation=\"http://www.gplates.org/gplates ../xsd/gpml.xsd http://www.opengis.net/gml ../../../gml/current/base\">")
                    sampleFile.write('\n' + "    <gml:featureMember>")
                    sampleFile.write('\n' + "        <gpml:VirtualGeomagneticPole>")
                    sampleFile.write('\n' + "            <gpml:identity>%s</gpml:identity>" % plateid)
                    sampleFile.write('\n' + "            <gpml:revision>%s</gpml:revision>" % platerevision)
                    sampleFile.write('\n' + "                        <gml:name>%s</gml:name>" % antipolelayername) 
                    sampleFile.write('\n' + "            <gpml:polePosition>")
                    sampleFile.write('\n' + "                <gpml:ConstantValue>")
                    sampleFile.write('\n' + "                    <gpml:value>")
                    sampleFile.write('\n' + "                        <gml:Point>")
                    sampleFile.write('\n' + "                            <gml:pos>%s %s</gml:pos>" % (antipolelat, antipolelon))
                    sampleFile.write('\n' + "                        </gml:Point>")
                    sampleFile.write('\n' + "                    </gpml:value>")
                    sampleFile.write('\n' + "                    <gml:description></gml:description>")
                    sampleFile.write('\n' + "                    <gpml:valueType xmlns:gml=\"http://www.opengis.net/gml\">gml:Point</gpml:valueType>")
                    sampleFile.write('\n' + "                </gpml:ConstantValue>")
                    sampleFile.write('\n' + "            </gpml:polePosition>")
                    sampleFile.write('\n' + "            <gpml:reconstructionPlateId>")
                    sampleFile.write('\n' + "                <gpml:ConstantValue>")
                    sampleFile.write('\n' + "                    <gpml:value>0</gpml:value>")
                    sampleFile.write('\n' + "                    <gml:description></gml:description>")
                    sampleFile.write('\n' + "                    <gpml:valueType xmlns:gpml=\"http://www.gplates.org/gplates\">gpml:plateId</gpml:valueType>")
                    sampleFile.write('\n' + "                </gpml:ConstantValue>")
                    sampleFile.write('\n' + "            </gpml:reconstructionPlateId>")
                    sampleFile.write('\n' + "            <gpml:averageAge>0</gpml:averageAge>")
                    sampleFile.write('\n' + "            <gpml:poleA95>%s</gpml:poleA95>" % polea95)
                    sampleFile.write('\n' + "        </gpml:VirtualGeomagneticPole>")
                    sampleFile.write('\n' + "    </gml:featureMember>")
                    sampleFile.write('\n' + "</gpml:FeatureCollection>")

                    sampleFile.close()

if not nearbypolefound:
    if poleangdist == 1:
        print ('\n' + "   DON'T FEEL TOO BAD BUT...No other nearby poles within %s degree found." % poleangdist)
    else:
        print ('\n' + "   DON'T FEEL TOO BAD BUT...No other nearby poles within %s degrees found." % poleangdist)

endchoice = raw_input('\n' + "----- Program complete. Goodbye! Please press enter to exit. -----")
sys.exit()
