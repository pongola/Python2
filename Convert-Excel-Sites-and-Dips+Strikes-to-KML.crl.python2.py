# -*- coding: utf-8 -*-
# 
#############################################
# 
# This script converts Excel ALL SITE AND DIPS-STRIKES to KML. It does the whole excel file. 
# This text-based Python script reads an Excel file “Site-Dip-and-Dip-Directions.xlsx” containing all sampling site locations in latitude and longitude and their bedding orientation,
# plus a table of all bedding orientation measurements taken and used in this study, and converts them to a Keyhole Markup Language (.kml) file. This .kml file can be opened as a Google
# Earth overlay displaying all sites and bedding measurements used in this study. Arrow icons used for bedding measurement locations are oriented to within 15° of actual dip direction,
# allowing for easy visual identification of bedding dip-direction on map. Mouse hovering over a site or measurement location in Google Earth leads to a popup window that contains
# information from the Excel spreadsheet about the site or measurement.
#
# For support, please contact Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com
#
#############################################

linebreak = chr(10)
leftsinglequote = u'\u2018'
rightsinglequote = u'\u2019'
leftdoublequote = u'\u201c'
rightdoublequote = u'\u201d'
endash = u'\u2013'
emdash = u'\u2014'
degree_symbol = u'\u00B0'
goodarrowpath = "D:/Documents/Google Earth/Icons/Arrow"
badarrowpath = "D:/Documents/Google Earth/Icons/Black/Arrow"
goodcolor = "ff0000ff"
badcolor = "ffffff00"
weirdset = [31,50,47,19,20]

import sys
import os

os.system('mode con: cols=150 lines=60')

print('\n' + "Hello. This script is ready to conver ALL SITES and ALL DIPS+STRIKES from your Excel Site-Dip-and-Dip-Directions.xlsx to KML.")
      
excelfilename = "Site-Dip-and-Dip-Directions.xlsx"
outputkmlfile = "Sites-Dips-Strikes.kml"
sitesheet = "Sites"
measuresheet = "Measurements"
yes = set(['yes','y','ye'])
no = set(['no','n'])

import os     # check if the .excelfilename exists in this folder.

if not(os.path.isfile(excelfilename)): # check if excelfile is present
    print ('\n' + "This folder contains no file named %s, which is necessary for this program to run." % excelfilename)
    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
    sys.exit()
else:
    print ('\n' + "The Excel source file %s has been found!" % excelfilename)
    continuechoice = ''
    if os.path.isfile(outputkmlfile):
        print ('\n' + "If you continue, this program will overwrite your kml file %s." % outputkmlfile)
        continuechoice = raw_input('\n' + "To continue (and overwrite %s), hit Y. To end this program and NOT overwrtite %s, hit N." % (outputkmlfile, outputkmlfile) + '\n' + "Do you choose to overwrite and continue?  ").lower()
        while not ((continuechoice in yes) or (continuechoice in no)):
            print('\n' + "Invalid Response--please respond with 'yes' or 'no' (or 'y' or 'n').")
            continuechoice = raw_input("To continue (and overwrite %s), hit Y. To end this program and NOT overwrtite %s, hit N." % (outputkmlfile, outputkmlfile) + '\n' + "Do you choose to overwrite and continue?  ").lower()

    if continuechoice in no:
        print ('\n' + "Program terminated by user." + '\n')
        sys.exit()
    elif os.path.isfile(outputkmlfile):
        os.remove(outputkmlfile)
        print ('\n' + "----------------------------" + '\n')

print ("Opening workbook %s" % excelfilename)
from xlrd import open_workbook
wb = open_workbook(excelfilename)

degree_sign = u'\xb0'
arrownumlist = [0,15,30,45,60,75,90,105,120,135,150,165,180,195,210,225,240,255,270,285,300,315,330,345]

sampleFile = open (outputkmlfile, 'a')

print ("Writing file %s" % outputkmlfile)
sampleFile.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>") # Write Intro KML Code
sampleFile.write('\n' + "<kml xmlns=\"http://www.opengis.net/kml/2.2\" xmlns:gx=\"http://www.google.com/kml/ext/2.2\" xmlns:kml=\"http://www.opengis.net/kml/2.2\" xmlns:atom=\"http://www.w3.org/2005/Atom\">")

sampleFile.write('\n' + "<Folder>") # Write Folder for Sites
sampleFile.write('\n' + "   <name>Sites and Dip/Strike Measurements from Casey Luskin's PhD Research</name>")
sampleFile.write('\n' + "	<open>1</open>")
sampleFile.write('\n' + "<Folder>")
sampleFile.write('\n' + "   <name>Sites</name>")
sampleFile.write('\n' + "	<open>1</open>")

for sheet in wb.sheets(): # Write Site KML Code
    if sheet.name == sitesheet:
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        
        for row in range(1, number_of_rows):
            rowitems = []
            values = []
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                rowitems.append(value)

            sitenum = int(round(rowitems[0]))
            sitename = rowitems[1]
            latitude = rowitems[2]
            longitude = rowitems[3]
            sitedec = rowitems[4]
            sitedec_stdev = rowitems[5]
            dipdir = rowitems[6]
            measnum = rowitems[7]
            dip = rowitems[10]
            
            structuralnotes = rowitems[11]
            structuralnotes_decoded = structuralnotes.replace(linebreak,"  ").replace(leftsinglequote, "'").replace(rightsinglequote, "'").replace(leftdoublequote, "\"").replace(rightdoublequote, "\"").replace(endash,"--").replace(emdash,"--").replace(degree_symbol," degrees").encode('latin1')

            sitenotes = rowitems[12]
            sitenotes_decoded = sitenotes.replace(linebreak,"  ").replace(leftsinglequote, "'").replace(rightsinglequote, "'").replace(leftdoublequote, "\"").replace(rightdoublequote, "\"").replace(endash,"--").replace(emdash,"--").replace(degree_symbol," degrees").encode('latin1')

            sampleFile.write('\n' + "	<Document>")
            sampleFile.write('\n' + "		<name>%s. %s (Site)</name>" % (sitenum, sitename))
            sampleFile.write('\n' + "		<StyleMap id=\"m_ylw-pushpin\">")
            sampleFile.write('\n' + "			<Pair>")
            sampleFile.write('\n' + "				<key>normal</key>")
            sampleFile.write('\n' + "				<styleUrl>#s_ylw-pushpin</styleUrl>")
            sampleFile.write('\n' + "			</Pair>")
            sampleFile.write('\n' + "			<Pair>")
            sampleFile.write('\n' + "				<key>highlight</key>")
            sampleFile.write('\n' + "				<styleUrl>#s_ylw-pushpin_hl</styleUrl>")
            sampleFile.write('\n' + "			</Pair>")
            sampleFile.write('\n' + "		</StyleMap>")
            sampleFile.write('\n' + "		<Style id=\"s_ylw-pushpin\">")
            sampleFile.write('\n' + "			<IconStyle>")
            sampleFile.write('\n' + "				<scale>1.1</scale>")
            sampleFile.write('\n' + "				<Icon>")
            sampleFile.write('\n' + "					<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>")
            sampleFile.write('\n' + "				</Icon>")
            sampleFile.write('\n' + "				<hotSpot x=\"20\" y=\"2\" xunits=\"pixels\" yunits=\"pixels\"/>")
            sampleFile.write('\n' + "			</IconStyle>")
            sampleFile.write('\n' + "		</Style>")
            sampleFile.write('\n' + "		<Style id=\"s_ylw-pushpin_hl\">")
            sampleFile.write('\n' + "			<IconStyle>")
            sampleFile.write('\n' + "				<scale>1.3</scale>")
            sampleFile.write('\n' + "				<Icon>")
            sampleFile.write('\n' + "					<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>")
            sampleFile.write('\n' + "				</Icon>")
            sampleFile.write('\n' + "				<hotSpot x=\"20\" y=\"2\" xunits=\"pixels\" yunits=\"pixels\"/>")
            sampleFile.write('\n' + "			</IconStyle>")
            sampleFile.write('\n' + "		</Style>")
            sampleFile.write('\n' + "		<Placemark>")
            sampleFile.write('\n' + "			<name>%s. %s</name>" % (sitenum, sitename))
            sampleFile.write('\n' + "			<LookAt>")
            sampleFile.write('\n' + "				<longitude>%s</longitude>" % longitude)
            sampleFile.write('\n' + "				<latitude>%s</latitude>" % latitude)
            sampleFile.write('\n' + "            		<altitude>0</altitude>")
            sampleFile.write('\n' + "				<heading>-2.153898092502043e-007</heading>")
            sampleFile.write('\n' + "				<tilt>0</tilt>")
            sampleFile.write('\n' + "				<range>8412.23478461989</range>")
            sampleFile.write('\n' + "				<gx:altitudeMode>relativeToSeaFloor</gx:altitudeMode>")
            sampleFile.write('\n' + "			</LookAt>")
            sampleFile.write('\n' + "			<styleUrl>#m_ylw-pushpin</styleUrl>")
            sampleFile.write('\n' + "			<Point>")
            sampleFile.write('\n' + "				<gx:drawOrder>1</gx:drawOrder>")
            sampleFile.write('\n' + "				<coordinates>%s,%s</coordinates>" % (longitude, latitude))
            sampleFile.write('\n' + "			</Point>")
            sampleFile.write('\n' + "			<name>%s. Site %s</name>" % (sitenum, sitename))
            sampleFile.write('\n' + "			<description><p><b>Site %s Notes:</b> %s</p>" % (sitename, sitenotes_decoded))
            sampleFile.write('\n' + "			             <p><b>Site %s Structural Data:</b><br>Dip: %s" % (sitename, dip))
            sampleFile.write(degree_sign.encode('utf8'))
            sampleFile.write("</br><br>Dip Dir: %s" % dipdir)
            sampleFile.write (degree_sign.encode('utf8'))
            sampleFile.write(".</br></p>")
            sampleFile.write('\n' + "			             <p><b>Site %s Structural Notes:</b> " % sitename)
            sampleFile.write("%s</p></description>" % (structuralnotes_decoded))
            sampleFile.write('\n' + "		</Placemark>")
            sampleFile.write('\n' + "	</Document>")
            sampleFile.write('\n')
sampleFile.write("</Folder>" + '\n') # Write End of Site Folder

sampleFile.write('\n' + "<Folder>") # Write Folder for Measurements
sampleFile.write('\n' + "<name>Dip/Dip Dir Measurements</name>")
sampleFile.write('\n' + "	<open>1</open>")

for sheet in wb.sheets(): # Write measurement kml code
    if sheet.name == measuresheet:
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        for row in range(1, number_of_rows):
            rowitems = []
            values = []
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                rowitems.append(value)

            measnum = int(round(rowitems[0]))
            latitude = rowitems[1]
            longitude = rowitems[2]

            dip = int(round(rowitems[4]))
            dipdir = int(round(rowitems[5]))
            if dipdir > 352:
                dipdir = 0            
            arrownum = min(arrownumlist, key=lambda x:abs(x-dipdir))

            uncorrecteddipdir_text = rowitems[6]
            try:
                uncorrecteddipdir = int(uncorrecteddipdir_text)
            except:
                uncorrecteddipdir = uncorrecteddipdir_text.replace(linebreak,"  ").replace(leftsinglequote, "'").replace(rightsinglequote, "'").replace(leftdoublequote, "\"").replace(rightdoublequote, "\"").replace(endash,"--").replace(emdash,"--").replace(degree_symbol," degrees").encode('latin1')
            
            deccorrection_text = rowitems[7]
            try:
                deccorrection = int(deccorrection_text)
            except:
                deccorrection = deccorrection_text.replace(linebreak,"  ").replace(leftsinglequote, "'").replace(rightsinglequote, "'").replace(leftdoublequote, "\"").replace(rightdoublequote, "\"").replace(endash,"--").replace(emdash,"--").replace(degree_symbol," degrees").encode('latin1')

            measurementOK_text = rowitems[8]
            measurementOK_text_decoded = measurementOK_text.replace(linebreak,"  ").replace(leftsinglequote, "'").replace(rightsinglequote, "'").replace(leftdoublequote, "\"").replace(rightdoublequote, "\"").replace(endash,"--").replace(emdash,"--").replace(degree_symbol," degrees").encode('latin1')
            if measurementOK_text_decoded == "":
                measurementOK_text_decoded = "??"
            if measurementOK_text_decoded == "Yes":
                arrowcolor = goodcolor
                arrowpath = goodarrowpath
            else:
                arrowcolor = badcolor
                arrowpath = badarrowpath

            dec_explanation = rowitems[9]
            dec_explanation_decoded = dec_explanation.replace(linebreak,"  ").replace(leftsinglequote, "'").replace(rightsinglequote, "'").replace(leftdoublequote, "\"").replace(rightdoublequote, "\"").replace(endash,"--").replace(emdash,"--").replace(degree_symbol," degrees").encode('latin1')
            
            tripnotes = rowitems[10]
            tripnotes_decoded = tripnotes.replace(linebreak,"  ").replace(leftsinglequote, "'").replace(rightsinglequote, "'").replace(leftdoublequote, "\"").replace(rightdoublequote, "\"").replace(endash,"--").replace(emdash,"--").replace(degree_symbol," degrees").encode('latin1')
            
            lithnotes = rowitems[11]
            lithnotes_decoded = lithnotes.replace(linebreak,"  ").replace(leftsinglequote, "'").replace(rightsinglequote, "'").replace(leftdoublequote, "\"").replace(rightdoublequote, "\"").replace(endash,"--").replace(emdash,"--").replace(degree_symbol," degrees").encode('latin1')

            sampleFile.write('\n' + "<Document>")
            sampleFile.write('\n' + "	<name>%s. Dip/Dip Dir: %s" % (measnum, dip) + degree_sign.encode('utf8') + "/%s" % dipdir + degree_sign.encode('utf8') + "</name>")
            sampleFile.write('\n' + "	<Style id=\"sh_Arrow%s\">" % arrownum + '\n')
            sampleFile.write('\n' + "	<color>%s</color>" % arrowcolor)
            sampleFile.write('\n' + "		<IconStyle>")
            sampleFile.write('\n' + "			<Icon>")
            sampleFile.write('\n' + "				<href>%s%s.png</href>" % (arrowpath, arrownum) + '\n')
            sampleFile.write('\n' + "			</Icon>")
            sampleFile.write('\n' + "			<hotSpot x=\"1.0\" y=\"1.0\" xunits=\"fraction\" yunits=\"fraction\"/>")
            sampleFile.write('\n' + "		</IconStyle>")            
            sampleFile.write('\n' + "		<LabelStyle>")
            sampleFile.write('\n' + "			<color>ff00ffff</color>")
            sampleFile.write('\n' + "		</LabelStyle>")
            sampleFile.write('\n' + "		<ListStyle>")
            sampleFile.write('\n' + "		</ListStyle>")
            sampleFile.write('\n' + "	</Style>")
            sampleFile.write('\n' + "	<StyleMap id=\"msn_Arrow%s\">" % arrownum + '\n')
            sampleFile.write('\n' + "		<Pair>")
            sampleFile.write('\n' + "			<key>normal</key>")
            sampleFile.write('\n' + "			<styleUrl>#sn_Arrow%s</styleUrl>" % arrownum + '\n')
            sampleFile.write('\n' + "		</Pair>")
            sampleFile.write('\n' + "		<Pair>")
            sampleFile.write('\n' + "			<key>highlight</key>")
            sampleFile.write('\n' + "			<styleUrl>#sh_Arrow%s</styleUrl>" % arrownum + '\n')
            sampleFile.write('\n' + "		</Pair>")
            sampleFile.write('\n' + "  </StyleMap>")
            sampleFile.write('\n' + "	<Style id=\"sn_Arrow%s\">" % arrownum + '\n')
            sampleFile.write('\n' + "		<IconStyle>")
            sampleFile.write('\n' + "			<color>%s</color>" % arrowcolor)
            sampleFile.write('\n' + "			<Icon>")
            sampleFile.write('\n' + "				<href>%s%s.png</href>" % (arrowpath, arrownum) + '\n')
            sampleFile.write('\n' + "			</Icon>")
            sampleFile.write('\n' + "			<hotSpot x=\"1.0\" y=\"1.0\" xunits=\"fraction\" yunits=\"fraction\"/>")
            sampleFile.write('\n' + "		</IconStyle>")
            sampleFile.write('\n' + "		<LabelStyle>")
            sampleFile.write('\n' + "			<color>ff00ffff</color>" )
            sampleFile.write('\n' + "		</LabelStyle>")
            sampleFile.write('\n' + "		<ListStyle>")
            sampleFile.write('\n' + "		</ListStyle>")
            sampleFile.write('\n' + "	</Style>")
            sampleFile.write('\n' + "	<Placemark>")           
            sampleFile.write('\n' + "		<name>%s. Dip/Dip Dir: %s" % (measnum, dip) + degree_sign.encode('utf8') +"/%s" % dipdir + degree_sign.encode('utf8') + "</name>")
            sampleFile.write('\n' + "		<description><![CDATA[<TABLE><TR><TD><img src=\"file:///%s%s.png\"/></TD><TD><b>Dip:</b> %s" % (arrowpath, arrownum, dip) + degree_sign.encode('utf8') +
                                    "<BR><b>Dip Dir (Corrected):</b> %s" % (dipdir) + degree_sign.encode('utf8') +
                                    "<BR></TD></TR></TABLE><BR><b>Uncorrected Dip Dir:</b> %s" % (uncorrecteddipdir) + degree_sign.encode('utf8') +
                                    "<BR><BR><b>Dec Correction Applied in Calculations To Restore to True Dip Dir:</b> %s" % (deccorrection) + degree_sign.encode('utf8') + 
                                    "<BR><BR><b>Am I Absolutely 100%% Sure About Dec Correction Based Upon Field Notes?</b> %s<BR><i>...Explanation:</i> %s<BR><BR><b>Trip Notes:</b> %s<BR><BR><b>Site Notes:</b> %s]]></description>" % (measurementOK_text_decoded, dec_explanation_decoded, tripnotes_decoded, lithnotes_decoded))
            sampleFile.write('\n' + "		<LookAt>")
            sampleFile.write('\n' + "			<longitude>%s</longitude>" % longitude)
            sampleFile.write('\n' + "			<latitude>%s</latitude>" % latitude + '\n')
            sampleFile.write('\n' + "			<altitude>0</altitude>")
            sampleFile.write('\n' + "			<heading>0.005035767148373351</heading>")
            sampleFile.write('\n' + "			<tilt>0</tilt>")
            sampleFile.write('\n' + "			<range>2322.042274646679</range>")
            sampleFile.write('\n' + "			<gx:altitudeMode>relativeToSeaFloor</gx:altitudeMode>")
            sampleFile.write('\n' + "		</LookAt>")
            if not (measnum in weirdset) and 1==2:
                sampleFile.write('\n' + "		<styleUrl>#msn_Arrow%s</styleUrl>" % arrownum + '\n')
            else:
                sampleFile.write('\n' + "  <StyleMap id=\"msn_Arrow%s\">" % arrownum + '\n')
                sampleFile.write('\n' + "		<Pair>")
                sampleFile.write('\n' + "			<key>normal</key>")
                sampleFile.write('\n' + "			<styleUrl>#sn_Arrow%s</styleUrl>" % arrownum + '\n')
                sampleFile.write('\n' + "		</Pair>")
                sampleFile.write('\n' + "		<Pair>")
                sampleFile.write('\n' + "			<key>highlight</key>")
                sampleFile.write('\n' + "			<styleUrl>#sh_Arrow%s</styleUrl>" % arrownum + '\n')
                sampleFile.write('\n' + "		</Pair>")
                sampleFile.write('\n' + "  </StyleMap>")
            sampleFile.write('\n' + "		<Point>")
            sampleFile.write('\n' + "			<gx:drawOrder>1</gx:drawOrder>")
            sampleFile.write('\n' + "			<coordinates>%s,%s,0</coordinates>"  % (longitude, latitude) + '\n')
            sampleFile.write('\n' + "		</Point>")
            sampleFile.write('\n' + "	</Placemark>")
            sampleFile.write('\n' + "</Document>")
            sampleFile.write('\n')
sampleFile.write('\n' + "</Folder>") # Write End of Measurement Folder
sampleFile.write('\n' + "</Folder>")
sampleFile.write('\n' + "</kml>") # Write outro text

sampleFile.close()
print ("File %s is written!" % outputkmlfile)

endchoice = raw_input('\n' + "----- Program complete. Goodbye! Please press enter to exit. -----")
sys.exit()
