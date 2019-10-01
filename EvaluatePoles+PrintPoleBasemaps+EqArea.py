# -*- coding: utf-8 -*-

####################### INTRODUCTION #######################
# This script evaluates poles for Q criteria (including finding similar younger poles) and prints basemaps with pole diagrams and equal area plots.
# The input Excel file must be named: "Prior Work--Quickbook.xlsx".
# Circle code adapted from https://github.com/urschrei/Circles/blob/master/README.md

####################### IMPORTS ############################
import sys
import os
import math
import re

from shutil import copy
from xlrd import open_workbook
import xlwt

from os.path import dirname
import matplotlib.pyplot as plt
from mpl_toolkits.basemap import Basemap
import numpy as np

from Circles.circles import circle
from Circles.circles import circle_wrap
# Note that circles is found in C:\Python27\Lib\site-packages\Circles
# Note that I modified circle to create circle_wrap
from shapely.geometry import Polygon, Point
from descartes import PolygonPatch

verifiedqmin = 0 #Change this to make verified poles require higher vqq

####################### CONSTANTS #######################
excelfilename = "MesoarcheanQuickbook.xlsx"
finalpolesexcelfilename = "Prior Work Verified.xls"
geologicaltimescale = [ [1,"Quarternary",0,2.58],[2,"Neogene",2.58,23.03],
                        [3,"Paleogene",23.03,66],[4,"Cretaceous",66,145],
                        [5,"Jurassic",145,201.3],[6,"Triassic",201.3,251.9],
                        [7,"Permian",251.9,298.9],[8,"Carboniferous",298.9,358.9],
                        [9,"Devonian",358.9,419.2],[10,"Silurian",419.2,485.4],
                        [11,"Devonian",358.9,419.2],[12,"Silurian",419.2,485.4],
                        [13,"Cambrian",485.4,541],[14,"Ediacaran",541,635],
                        [15,"Cryogenian",635,720],[16,"Tonian",720,1000],
                        [17,"Stenian",1000,1200],[18,"Ectasian",1200,1400],
                        [19,"Calymmian",1400,1600],[20,"Statherian",1600,1800],
                        [21,"Orosirian",1800,2050],[22,"Rhyacian",2050,2300],
                        [23,"Siderian",2300,2500],[24,"Neoarchean",2500,2800],
                        [25,"Mesoarchean",2800,3200],[26,"Paleoarchean",3200,3600],
                        [27,"Eoarchean",3600,4000] ]

alpha_symbol = u'\N{GREEK SMALL LETTER ALPHA}'
subscript9 = u'\N{SUBSCRIPT NINE}'
subscript5 = u'\N{SUBSCRIPT FIVE}'
a95_word_lower = "%s%s%s" % (alpha_symbol,subscript9,subscript5)
A95_word_upper = "A%s%s" % (subscript9,subscript5)

cols = [["Unit (Component)",40],["Age (Ma)",15],["Dec",6],["Inc",6],["k",5],[a95_word_lower,5],
        ["Plat",6],["Plon",6],["dp",5],["dm",5],["K",5],[A95_word_upper,5],
        ["1",2],["2",2],["3",2],["4",2],["5",2],["6",2],["7",2],["Q",2],["Paleomag Reference\n(Age Reference)",40],["Younger/Similar Poles",40],["Pole Type",3]]

eqcircletick_length_dict = {10:8, 15:8, 20:8, 30:8, 40:8, 45:10, 50:8, 60:8, 70:8, 75:8, 80:8, 90:12,
                            100:8, 105:8, 110:8, 120:8, 130:8, 135:10, 140:8, 150:8, 160:8, 165:8, 170:8, 180:12,
                            190:8, 195:8, 200:8, 210:8, 220:8, 225:10, 230:8, 240:8, 250:8, 255:8, 260:8, 270:12,
                            280:8, 285:8, 290:8, 300:8, 310:8, 315:10, 320:8, 330:8, 340:8, 345:8, 350:8, 0:12}

eqaxestick_length_dict = {10:5, 15:5, 20:5, 30:10, 40:5, 45:10, 50:5, 60:10, 70:5, 75:5, 80:5}

specialcolorset = ["ALICEBLUE","ANTIQUEWHITE","ANTIQUEWHITE1","ANTIQUEWHITE2","ANTIQUEWHITE3","ANTIQUEWHITE4","AQUA","AQUAMARINE1","AQUAMARINE2","AQUAMARINE3","AQUAMARINE4","AZURE1","AZURE2","AZURE3","AZURE4",
                   "BANANA","BEIGE","BISQUE1","BISQUE2","BISQUE3","BISQUE4","BLACK","BLANCHEDALMOND","BLUE","BLUE2","BLUE3","BLUE4","BLUEVIOLET","BRICK","BROWN","BROWN1","BROWN2","BROWN3","BROWN4","BURLYWOOD","BURLYWOOD1",
                   "BURLYWOOD2","BURLYWOOD3","BURLYWOOD4","BURNTSIENNA","BURNTUMBER","CADETBLUE","CADETBLUE1","CADETBLUE2","CADETBLUE3","CADETBLUE4","CADMIUMORANGE","CADMIUMYELLOW","CARROT","CHARTREUSE1","CHARTREUSE2",
                   "CHARTREUSE3","CHARTREUSE4","CHOCOLATE","CHOCOLATE1","CHOCOLATE2","CHOCOLATE3","CHOCOLATE4","COBALT","COBALTGREEN","COLDGREY","CORAL","CORAL1","CORAL2","CORAL3","CORAL4","CORNFLOWERBLUE","CORNSILK1",
                   "CORNSILK2","CORNSILK3","CORNSILK4","CRIMSON","CYAN2","CYAN3","CYAN4","DARKGOLDENROD","DARKGOLDENROD1","DARKGOLDENROD2","DARKGOLDENROD3","DARKGOLDENROD4","DARKGRAY","DARKGREEN","DARKKHAKI","DARKOLIVEGREEN",
                   "DARKOLIVEGREEN1","DARKOLIVEGREEN2","DARKOLIVEGREEN3","DARKOLIVEGREEN4","DARKORANGE","DARKORANGE1","DARKORANGE2","DARKORANGE3","DARKORANGE4","DARKORCHID","DARKORCHID1","DARKORCHID2","DARKORCHID3","DARKORCHID4",
                   "DARKSALMON","DARKSEAGREEN","DARKSEAGREEN1","DARKSEAGREEN2","DARKSEAGREEN3","DARKSEAGREEN4","DARKSLATEBLUE","DARKSLATEGRAY","DARKSLATEGRAY1","DARKSLATEGRAY2","DARKSLATEGRAY3","DARKSLATEGRAY4","DARKTURQUOISE",
                   "DARKVIOLET","DEEPPINK1","DEEPPINK2","DEEPPINK3","DEEPPINK4","DEEPSKYBLUE1","DEEPSKYBLUE2","DEEPSKYBLUE3","DEEPSKYBLUE4","DIMGRAY","DODGERBLUE1","DODGERBLUE2","DODGERBLUE3","DODGERBLUE4","EGGSHELL","EMERALDGREEN",
                   "FIREBRICK","FIREBRICK1","FIREBRICK2","FIREBRICK3","FIREBRICK4","FLESH","FLORALWHITE","FORESTGREEN","GAINSBORO","GHOSTWHITE","GOLD1","GOLD2","GOLD3","GOLD4","GOLDENROD","GOLDENROD1","GOLDENROD2","GOLDENROD3",
                   "GOLDENROD4","GRAY","GRAY1","GRAY10","GRAY11","GRAY12","GRAY13","GRAY14","GRAY15","GRAY16","GRAY17","GRAY18","GRAY19","GRAY2","GRAY20","GRAY21","GRAY22","GRAY23","GRAY24","GRAY25","GRAY26","GRAY27","GRAY28",
                   "GRAY29","GRAY3","GRAY30","GRAY31","GRAY32","GRAY33","GRAY34","GRAY35","GRAY36","GRAY37","GRAY38","GRAY39","GRAY4","GRAY40","GRAY42","GRAY43","GRAY44","GRAY45","GRAY46","GRAY47","GRAY48","GRAY49","GRAY5",
                   "GRAY50","GRAY51","GRAY52","GRAY53","GRAY54","GRAY55","GRAY56","GRAY57","GRAY58","GRAY59","GRAY6","GRAY60","GRAY61","GRAY62","GRAY63","GRAY64","GRAY65","GRAY66","GRAY67","GRAY68","GRAY69","GRAY7","GRAY70",
                   "GRAY71","GRAY72","GRAY73","GRAY74","GRAY75","GRAY76","GRAY77","GRAY78","GRAY79","GRAY8","GRAY80","GRAY81","GRAY82","GRAY83","GRAY84","GRAY85","GRAY86","GRAY87","GRAY88","GRAY89","GRAY9","GRAY90","GRAY91",
                   "GRAY92","GRAY93","GRAY94","GRAY95","GRAY97","GRAY98","GRAY99","GREEN","GREEN1","GREEN2","GREEN3","GREEN4","GREENYELLOW","HONEYDEW1","HONEYDEW2","HONEYDEW3","HONEYDEW4","HOTPINK","HOTPINK1","HOTPINK2","HOTPINK3",
                   "HOTPINK4","INDIANRED","INDIANRED1","INDIANRED2","INDIANRED3","INDIANRED4","INDIGO","IVORY1","IVORY2","IVORY3","IVORY4","IVORYBLACK","KHAKI","KHAKI1","KHAKI2","KHAKI3","KHAKI4","LAVENDER","LAVENDERBLUSH1",
                   "LAVENDERBLUSH2","LAVENDERBLUSH3","LAVENDERBLUSH4","LAWNGREEN","LEMONCHIFFON1","LEMONCHIFFON2","LEMONCHIFFON3","LEMONCHIFFON4","LIGHTBLUE","LIGHTBLUE1","LIGHTBLUE2","LIGHTBLUE3","LIGHTBLUE4","LIGHTCORAL",
                   "LIGHTCYAN1","LIGHTCYAN2","LIGHTCYAN3","LIGHTCYAN4","LIGHTGOLDENROD1","LIGHTGOLDENROD2","LIGHTGOLDENROD3","LIGHTGOLDENROD4","LIGHTGOLDENRODYELLOW","LIGHTGREY","LIGHTPINK","LIGHTPINK1","LIGHTPINK2","LIGHTPINK3",
                   "LIGHTPINK4","LIGHTSALMON1","LIGHTSALMON2","LIGHTSALMON3","LIGHTSALMON4","LIGHTSEAGREEN","LIGHTSKYBLUE","LIGHTSKYBLUE1","LIGHTSKYBLUE2","LIGHTSKYBLUE3","LIGHTSKYBLUE4","LIGHTSLATEBLUE","LIGHTSLATEGRAY",
                   "LIGHTSTEELBLUE","LIGHTSTEELBLUE1","LIGHTSTEELBLUE2","LIGHTSTEELBLUE3","LIGHTSTEELBLUE4","LIGHTYELLOW1","LIGHTYELLOW2","LIGHTYELLOW3","LIGHTYELLOW4","LIMEGREEN","LINEN","MAGENTA","MAGENTA2","MAGENTA3",
                   "MAGENTA4","MANGANESEBLUE","MAROON","MAROON1","MAROON2","MAROON3","MAROON4","MEDIUMORCHID","MEDIUMORCHID1","MEDIUMORCHID2","MEDIUMORCHID3","MEDIUMORCHID4","MEDIUMPURPLE","MEDIUMPURPLE1","MEDIUMPURPLE2",
                   "MEDIUMPURPLE3","MEDIUMPURPLE4","MEDIUMSEAGREEN","MEDIUMSLATEBLUE","MEDIUMSPRINGGREEN","MEDIUMTURQUOISE","MEDIUMVIOLETRED","MELON","MIDNIGHTBLUE","MINT","MINTCREAM","MISTYROSE1","MISTYROSE2","MISTYROSE3",
                   "MISTYROSE4","MOCCASIN","NAVAJOWHITE1","NAVAJOWHITE2","NAVAJOWHITE3","NAVAJOWHITE4","NAVY","OLDLACE","OLIVE","OLIVEDRAB","OLIVEDRAB1","OLIVEDRAB2","OLIVEDRAB3","OLIVEDRAB4","ORANGE","ORANGE1","ORANGE2",
                   "ORANGE3","ORANGE4","ORANGERED1","ORANGERED2","ORANGERED3","ORANGERED4","ORCHID","ORCHID1","ORCHID2","ORCHID3","ORCHID4","PALEGOLDENROD","PALEGREEN","PALEGREEN1","PALEGREEN2","PALEGREEN3","PALEGREEN4",
                   "PALETURQUOISE1","PALETURQUOISE2","PALETURQUOISE3","PALETURQUOISE4","PALEVIOLETRED","PALEVIOLETRED1","PALEVIOLETRED2","PALEVIOLETRED3","PALEVIOLETRED4","PAPAYAWHIP","PEACHPUFF1","PEACHPUFF2","PEACHPUFF3",
                   "PEACHPUFF4","PEACOCK","PINK","PINK1","PINK2","PINK3","PINK4","PLUM","PLUM1","PLUM2","PLUM3","PLUM4","POWDERBLUE","PURPLE","PURPLE1","PURPLE2","PURPLE3","PURPLE4","RASPBERRY","RAWSIENNA","RED1","RED2","RED3",
                   "RED4","ROSYBROWN","ROSYBROWN1","ROSYBROWN2","ROSYBROWN3","ROSYBROWN4","ROYALBLUE","ROYALBLUE1","ROYALBLUE2","ROYALBLUE3","ROYALBLUE4","SALMON","SALMON1","SALMON2","SALMON3","SALMON4","SANDYBROWN","SAPGREEN",
                   "SEAGREEN1","SEAGREEN2","SEAGREEN3","SEAGREEN4","SEASHELL1","SEASHELL2","SEASHELL3","SEASHELL4","SEPIA","SGIBEET","SGIBRIGHTGRAY","SGICHARTREUSE","SGIDARKGRAY","SGIGRAY12","SGIGRAY16","SGIGRAY32","SGIGRAY36",
                   "SGIGRAY52","SGIGRAY56","SGIGRAY72","SGIGRAY76","SGIGRAY92","SGIGRAY96","SGILIGHTBLUE","SGILIGHTGRAY","SGIOLIVEDRAB","SGISALMON","SGISLATEBLUE","SGITEAL","SIENNA","SIENNA1","SIENNA2","SIENNA3","SIENNA4","SILVER",
                   "SKYBLUE","SKYBLUE1","SKYBLUE2","SKYBLUE3","SKYBLUE4","SLATEBLUE","SLATEBLUE1","SLATEBLUE2","SLATEBLUE3","SLATEBLUE4","SLATEGRAY","SLATEGRAY1","SLATEGRAY2","SLATEGRAY3","SLATEGRAY4","SNOW1","SNOW2","SNOW3",
                   "SNOW4","SPRINGGREEN","SPRINGGREEN1","SPRINGGREEN2","SPRINGGREEN3","STEELBLUE","STEELBLUE1","STEELBLUE2","STEELBLUE3","STEELBLUE4","TAN","TAN1","TAN2","TAN3","TAN4","TEAL","THISTLE","THISTLE1","THISTLE2",
                   "THISTLE3","THISTLE4","TOMATO1","TOMATO2","TOMATO3","TOMATO4","TURQUOISE","TURQUOISE1","TURQUOISE2","TURQUOISE3","TURQUOISE4","TURQUOISEBLUE","VIOLET","VIOLETRED","VIOLETRED1","VIOLETRED2","VIOLETRED3",
                   "VIOLETRED4","WARMGREY","WHEAT","WHEAT1","WHEAT2","WHEAT3","WHEAT4","WHITE","WHITESMOKE","YELLOW1","YELLOW2","YELLOW3","YELLOW4","RED","LIME","YELLOW","CYAN"]

specialcolordict = {"ALICEBLUE":"#F0F8FF","ANTIQUEWHITE":"#FAEBD7","ANTIQUEWHITE1":"#FFEFDB","ANTIQUEWHITE2":"#EEDFCC","ANTIQUEWHITE3":"#CDC0B0","ANTIQUEWHITE4":"#8B8378","AQUA":"#00FFFF","AQUAMARINE1":"#7FFFD4",
                    "AQUAMARINE2":"#76EEC6","AQUAMARINE3":"#66CDAA","AQUAMARINE4":"#458B74","AZURE1":"#F0FFFF","AZURE2":"#E0EEEE","AZURE3":"#C1CDCD","AZURE4":"#838B8B","BANANA":"#E3CF57","BEIGE":"#F5F5DC",
                    "BISQUE1":"#FFE4C4","BISQUE2":"#EED5B7","BISQUE3":"#CDB79E","BISQUE4":"#8B7D6B","BLACK":"#000000","BLANCHEDALMOND":"#FFEBCD","BLUE":"#0000FF","BLUE2":"#0000EE","BLUE3":"#0000CD","BLUE4":"#00008B",
                    "BLUEVIOLET":"#8A2BE2","BRICK":"#9C661F","BROWN":"#A52A2A","BROWN1":"#FF4040","BROWN2":"#EE3B3B","BROWN3":"#CD3333","BROWN4":"#8B2323","BURLYWOOD":"#DEB887","BURLYWOOD1":"#FFD39B","BURLYWOOD2":"#EEC591",
                    "BURLYWOOD3":"#CDAA7D","BURLYWOOD4":"#8B7355","BURNTSIENNA":"#8A360F","BURNTUMBER":"#8A3324","CADETBLUE":"#5F9EA0","CADETBLUE1":"#98F5FF","CADETBLUE2":"#8EE5EE","CADETBLUE3":"#7AC5CD",
                    "CADETBLUE4":"#53868B","CADMIUMORANGE":"#FF6103","CADMIUMYELLOW":"#FF9912","CARROT":"#ED9121","CHARTREUSE1":"#7FFF00","CHARTREUSE2":"#76EE00","CHARTREUSE3":"#66CD00","CHARTREUSE4":"#458B00",
                    "CHOCOLATE":"#D2691E","CHOCOLATE1":"#FF7F24","CHOCOLATE2":"#EE7621","CHOCOLATE3":"#CD661D","CHOCOLATE4":"#8B4513","COBALT":"#3D59AB","COBALTGREEN":"#3D9140","COLDGREY":"#808A87","CORAL":"#FF7F50",
                    "CORAL1":"#FF7256","CORAL2":"#EE6A50","CORAL3":"#CD5B45","CORAL4":"#8B3E2F","CORNFLOWERBLUE":"#6495ED","CORNSILK1":"#FFF8DC","CORNSILK2":"#EEE8CD","CORNSILK3":"#CDC8B1","CORNSILK4":"#8B8878",
                    "CRIMSON":"#DC143C","CYAN2":"#00EEEE","CYAN3":"#00CDCD","CYAN4":"#008B8B","DARKGOLDENROD":"#B8860B","DARKGOLDENROD1":"#FFB90F","DARKGOLDENROD2":"#EEAD0E","DARKGOLDENROD3":"#CD950C",
                    "DARKGOLDENROD4":"#8B6508","DARKGRAY":"#A9A9A9","DARKGREEN":"#006400","DARKKHAKI":"#BDB76B","DARKOLIVEGREEN":"#556B2F","DARKOLIVEGREEN1":"#CAFF70","DARKOLIVEGREEN2":"#BCEE68","DARKOLIVEGREEN3":"#A2CD5A",
                    "DARKOLIVEGREEN4":"#6E8B3D","DARKORANGE":"#FF8C00","DARKORANGE1":"#FF7F00","DARKORANGE2":"#EE7600","DARKORANGE3":"#CD6600","DARKORANGE4":"#8B4500","DARKORCHID":"#9932CC","DARKORCHID1":"#BF3EFF",
                    "DARKORCHID2":"#B23AEE","DARKORCHID3":"#9A32CD","DARKORCHID4":"#68228B","DARKSALMON":"#E9967A","DARKSEAGREEN":"#8FBC8F","DARKSEAGREEN1":"#C1FFC1","DARKSEAGREEN2":"#B4EEB4","DARKSEAGREEN3":"#9BCD9B",
                    "DARKSEAGREEN4":"#698B69","DARKSLATEBLUE":"#483D8B","DARKSLATEGRAY":"#2F4F4F","DARKSLATEGRAY1":"#97FFFF","DARKSLATEGRAY2":"#8DEEEE","DARKSLATEGRAY3":"#79CDCD","DARKSLATEGRAY4":"#528B8B",
                    "DARKTURQUOISE":"#00CED1","DARKVIOLET":"#9400D3","DEEPPINK1":"#FF1493","DEEPPINK2":"#EE1289","DEEPPINK3":"#CD1076","DEEPPINK4":"#8B0A50","DEEPSKYBLUE1":"#00BFFF","DEEPSKYBLUE2":"#00B2EE",
                    "DEEPSKYBLUE3":"#009ACD","DEEPSKYBLUE4":"#00688B","DIMGRAY":"#696969","DODGERBLUE1":"#1E90FF","DODGERBLUE2":"#1C86EE","DODGERBLUE3":"#1874CD","DODGERBLUE4":"#104E8B","EGGSHELL":"#FCE6C9",
                    "EMERALDGREEN":"#00C957","FIREBRICK":"#B22222","FIREBRICK1":"#FF3030","FIREBRICK2":"#EE2C2C","FIREBRICK3":"#CD2626","FIREBRICK4":"#8B1A1A","FLESH":"#FF7D40","FLORALWHITE":"#FFFAF0",
                    "FORESTGREEN":"#228B22","GAINSBORO":"#DCDCDC","GHOSTWHITE":"#F8F8FF","GOLD1":"#FFD700","GOLD2":"#EEC900","GOLD3":"#CDAD00","GOLD4":"#8B7500","GOLDENROD":"#DAA520","GOLDENROD1":"#FFC125",
                    "GOLDENROD2":"#EEB422","GOLDENROD3":"#CD9B1D","GOLDENROD4":"#8B6914","GRAY":"#808080","GRAY1":"#030303","GRAY10":"#1A1A1A","GRAY11":"#1C1C1C","GRAY12":"#1F1F1F","GRAY13":"#212121",
                    "GRAY14":"#242424","GRAY15":"#262626","GRAY16":"#292929","GRAY17":"#2B2B2B","GRAY18":"#2E2E2E","GRAY19":"#303030","GRAY2":"#050505","GRAY20":"#333333","GRAY21":"#363636","GRAY22":"#383838",
                    "GRAY23":"#3B3B3B","GRAY24":"#3D3D3D","GRAY25":"#404040","GRAY26":"#424242","GRAY27":"#454545","GRAY28":"#474747","GRAY29":"#4A4A4A","GRAY3":"#080808","GRAY30":"#4D4D4D","GRAY31":"#4F4F4F",
                    "GRAY32":"#525252","GRAY33":"#545454","GRAY34":"#575757","GRAY35":"#595959","GRAY36":"#5C5C5C","GRAY37":"#5E5E5E","GRAY38":"#616161","GRAY39":"#636363","GRAY4":"#0A0A0A","GRAY40":"#666666",
                    "GRAY42":"#6B6B6B","GRAY43":"#6E6E6E","GRAY44":"#707070","GRAY45":"#737373","GRAY46":"#757575","GRAY47":"#787878","GRAY48":"#7A7A7A","GRAY49":"#7D7D7D","GRAY5":"#0D0D0D","GRAY50":"#7F7F7F",
                    "GRAY51":"#828282","GRAY52":"#858585","GRAY53":"#878787","GRAY54":"#8A8A8A","GRAY55":"#8C8C8C","GRAY56":"#8F8F8F","GRAY57":"#919191","GRAY58":"#949494","GRAY59":"#969696","GRAY6":"#0F0F0F",
                    "GRAY60":"#999999","GRAY61":"#9C9C9C","GRAY62":"#9E9E9E","GRAY63":"#A1A1A1","GRAY64":"#A3A3A3","GRAY65":"#A6A6A6","GRAY66":"#A8A8A8","GRAY67":"#ABABAB","GRAY68":"#ADADAD","GRAY69":"#B0B0B0",
                    "GRAY7":"#121212","GRAY70":"#B3B3B3","GRAY71":"#B5B5B5","GRAY72":"#B8B8B8","GRAY73":"#BABABA","GRAY74":"#BDBDBD","GRAY75":"#BFBFBF","GRAY76":"#C2C2C2","GRAY77":"#C4C4C4","GRAY78":"#C7C7C7",
                    "GRAY79":"#C9C9C9","GRAY8":"#141414","GRAY80":"#CCCCCC","GRAY81":"#CFCFCF","GRAY82":"#D1D1D1","GRAY83":"#D4D4D4","GRAY84":"#D6D6D6","GRAY85":"#D9D9D9","GRAY86":"#DBDBDB","GRAY87":"#DEDEDE",
                    "GRAY88":"#E0E0E0","GRAY89":"#E3E3E3","GRAY9":"#171717","GRAY90":"#E5E5E5","GRAY91":"#E8E8E8","GRAY92":"#EBEBEB","GRAY93":"#EDEDED","GRAY94":"#F0F0F0","GRAY95":"#F2F2F2","GRAY97":"#F7F7F7",
                    "GRAY98":"#FAFAFA","GRAY99":"#FCFCFC","GREEN":"#008000","GREEN1":"#00FF00","GREEN2":"#00EE00","GREEN3":"#00CD00","GREEN4":"#008B00","GREENYELLOW":"#ADFF2F","HONEYDEW1":"#F0FFF0","HONEYDEW2":"#E0EEE0",
                    "HONEYDEW3":"#C1CDC1","HONEYDEW4":"#838B83","HOTPINK":"#FF69B4","HOTPINK1":"#FF6EB4","HOTPINK2":"#EE6AA7","HOTPINK3":"#CD6090","HOTPINK4":"#8B3A62","INDIANRED":"#CD5C5C","INDIANRED1":"#FF6A6A",
                    "INDIANRED2":"#EE6363","INDIANRED3":"#CD5555","INDIANRED4":"#8B3A3A","INDIGO":"#4B0082","IVORY1":"#FFFFF0","IVORY2":"#EEEEE0","IVORY3":"#CDCDC1","IVORY4":"#8B8B83","IVORYBLACK":"#292421",
                    "KHAKI":"#F0E68C","KHAKI1":"#FFF68F","KHAKI2":"#EEE685","KHAKI3":"#CDC673","KHAKI4":"#8B864E","LAVENDER":"#E6E6FA","LAVENDERBLUSH1":"#FFF0F5","LAVENDERBLUSH2":"#EEE0E5","LAVENDERBLUSH3":"#CDC1C5",
                    "LAVENDERBLUSH4":"#8B8386","LAWNGREEN":"#7CFC00","LEMONCHIFFON1":"#FFFACD","LEMONCHIFFON2":"#EEE9BF","LEMONCHIFFON3":"#CDC9A5","LEMONCHIFFON4":"#8B8970","LIGHTBLUE":"#ADD8E6","LIGHTBLUE1":"#BFEFFF",
                    "LIGHTBLUE2":"#B2DFEE","LIGHTBLUE3":"#9AC0CD","LIGHTBLUE4":"#68838B","LIGHTCORAL":"#F08080","LIGHTCYAN1":"#E0FFFF","LIGHTCYAN2":"#D1EEEE","LIGHTCYAN3":"#B4CDCD","LIGHTCYAN4":"#7A8B8B",
                    "LIGHTGOLDENROD1":"#FFEC8B","LIGHTGOLDENROD2":"#EEDC82","LIGHTGOLDENROD3":"#CDBE70","LIGHTGOLDENROD4":"#8B814C","LIGHTGOLDENRODYELLOW":"#FAFAD2","LIGHTGREY":"#D3D3D3","LIGHTPINK":"#FFB6C1",
                    "LIGHTPINK1":"#FFAEB9","LIGHTPINK2":"#EEA2AD","LIGHTPINK3":"#CD8C95","LIGHTPINK4":"#8B5F65","LIGHTSALMON1":"#FFA07A","LIGHTSALMON2":"#EE9572","LIGHTSALMON3":"#CD8162","LIGHTSALMON4":"#8B5742",
                    "LIGHTSEAGREEN":"#20B2AA","LIGHTSKYBLUE":"#87CEFA","LIGHTSKYBLUE1":"#B0E2FF","LIGHTSKYBLUE2":"#A4D3EE","LIGHTSKYBLUE3":"#8DB6CD","LIGHTSKYBLUE4":"#607B8B","LIGHTSLATEBLUE":"#8470FF",
                    "LIGHTSLATEGRAY":"#778899","LIGHTSTEELBLUE":"#B0C4DE","LIGHTSTEELBLUE1":"#CAE1FF","LIGHTSTEELBLUE2":"#BCD2EE","LIGHTSTEELBLUE3":"#A2B5CD","LIGHTSTEELBLUE4":"#6E7B8B","LIGHTYELLOW1":"#FFFFE0",
                    "LIGHTYELLOW2":"#EEEED1","LIGHTYELLOW3":"#CDCDB4","LIGHTYELLOW4":"#8B8B7A","LIMEGREEN":"#32CD32","LINEN":"#FAF0E6","MAGENTA":"#FF00FF","MAGENTA2":"#EE00EE","MAGENTA3":"#CD00CD","MAGENTA4":"#8B008B",
                    "MANGANESEBLUE":"#03A89E","MAROON":"#800000","MAROON1":"#FF34B3","MAROON2":"#EE30A7","MAROON3":"#CD2990","MAROON4":"#8B1C62","MEDIUMORCHID":"#BA55D3","MEDIUMORCHID1":"#E066FF","MEDIUMORCHID2":"#D15FEE",
                    "MEDIUMORCHID3":"#B452CD","MEDIUMORCHID4":"#7A378B","MEDIUMPURPLE":"#9370DB","MEDIUMPURPLE1":"#AB82FF","MEDIUMPURPLE2":"#9F79EE","MEDIUMPURPLE3":"#8968CD","MEDIUMPURPLE4":"#5D478B",
                    "MEDIUMSEAGREEN":"#3CB371","MEDIUMSLATEBLUE":"#7B68EE","MEDIUMSPRINGGREEN":"#00FA9A","MEDIUMTURQUOISE":"#48D1CC","MEDIUMVIOLETRED":"#C71585","MELON":"#E3A869","MIDNIGHTBLUE":"#191970",
                    "MINT":"#BDFCC9","MINTCREAM":"#F5FFFA","MISTYROSE1":"#FFE4E1","MISTYROSE2":"#EED5D2","MISTYROSE3":"#CDB7B5","MISTYROSE4":"#8B7D7B","MOCCASIN":"#FFE4B5","NAVAJOWHITE1":"#FFDEAD","NAVAJOWHITE2":"#EECFA1",
                    "NAVAJOWHITE3":"#CDB38B","NAVAJOWHITE4":"#8B795E","NAVY":"#000080","OLDLACE":"#FDF5E6","OLIVE":"#808000","OLIVEDRAB":"#6B8E23","OLIVEDRAB1":"#C0FF3E","OLIVEDRAB2":"#B3EE3A","OLIVEDRAB3":"#9ACD32",
                    "OLIVEDRAB4":"#698B22","ORANGE":"#FF8000","ORANGE1":"#FFA500","ORANGE2":"#EE9A00","ORANGE3":"#CD8500","ORANGE4":"#8B5A00","ORANGERED1":"#FF4500","ORANGERED2":"#EE4000","ORANGERED3":"#CD3700",
                    "ORANGERED4":"#8B2500","ORCHID":"#DA70D6","ORCHID1":"#FF83FA","ORCHID2":"#EE7AE9","ORCHID3":"#CD69C9","ORCHID4":"#8B4789","PALEGOLDENROD":"#EEE8AA","PALEGREEN":"#98FB98","PALEGREEN1":"#9AFF9A",
                    "PALEGREEN2":"#90EE90","PALEGREEN3":"#7CCD7C","PALEGREEN4":"#548B54","PALETURQUOISE1":"#BBFFFF","PALETURQUOISE2":"#AEEEEE","PALETURQUOISE3":"#96CDCD","PALETURQUOISE4":"#668B8B",
                    "PALEVIOLETRED":"#DB7093","PALEVIOLETRED1":"#FF82AB","PALEVIOLETRED2":"#EE799F","PALEVIOLETRED3":"#CD6889","PALEVIOLETRED4":"#8B475D","PAPAYAWHIP":"#FFEFD5","PEACHPUFF1":"#FFDAB9","PEACHPUFF2":"#EECBAD",
                    "PEACHPUFF3":"#CDAF95","PEACHPUFF4":"#8B7765","PEACOCK":"#33A1C9","PINK":"#FFC0CB","PINK1":"#FFB5C5","PINK2":"#EEA9B8","PINK3":"#CD919E","PINK4":"#8B636C","PLUM":"#DDA0DD","PLUM1":"#FFBBFF",
                    "PLUM2":"#EEAEEE","PLUM3":"#CD96CD","PLUM4":"#8B668B","POWDERBLUE":"#B0E0E6","PURPLE":"#800080","PURPLE1":"#9B30FF","PURPLE2":"#912CEE","PURPLE3":"#7D26CD","PURPLE4":"#551A8B","RASPBERRY":"#872657",
                    "RAWSIENNA":"#C76114","RED1":"#FF0000","RED2":"#EE0000","RED3":"#CD0000","RED4":"#8B0000","ROSYBROWN":"#BC8F8F","ROSYBROWN1":"#FFC1C1","ROSYBROWN2":"#EEB4B4","ROSYBROWN3":"#CD9B9B","ROSYBROWN4":"#8B6969",
                    "ROYALBLUE":"#4169E1","ROYALBLUE1":"#4876FF","ROYALBLUE2":"#436EEE","ROYALBLUE3":"#3A5FCD","ROYALBLUE4":"#27408B","SALMON":"#FA8072","SALMON1":"#FF8C69","SALMON2":"#EE8262","SALMON3":"#CD7054",
                    "SALMON4":"#8B4C39","SANDYBROWN":"#F4A460","SAPGREEN":"#308014","SEAGREEN1":"#54FF9F","SEAGREEN2":"#4EEE94","SEAGREEN3":"#43CD80","SEAGREEN4":"#2E8B57","SEASHELL1":"#FFF5EE","SEASHELL2":"#EEE5DE",
                    "SEASHELL3":"#CDC5BF","SEASHELL4":"#8B8682","SEPIA":"#5E2612","SGIBEET":"#8E388E","SGIBRIGHTGRAY":"#C5C1AA","SGICHARTREUSE":"#71C671","SGIDARKGRAY":"#555555","SGIGRAY12":"#1E1E1E","SGIGRAY16":"#282828",
                    "SGIGRAY32":"#515151","SGIGRAY36":"#5B5B5B","SGIGRAY52":"#848484","SGIGRAY56":"#8E8E8E","SGIGRAY72":"#B7B7B7","SGIGRAY76":"#C1C1C1","SGIGRAY92":"#EAEAEA","SGIGRAY96":"#F4F4F4","SGILIGHTBLUE":"#7D9EC0",
                    "SGILIGHTGRAY":"#AAAAAA","SGIOLIVEDRAB":"#8E8E38","SGISALMON":"#C67171","SGISLATEBLUE":"#7171C6","SGITEAL":"#388E8E","SIENNA":"#A0522D","SIENNA1":"#FF8247","SIENNA2":"#EE7942","SIENNA3":"#CD6839",
                    "SIENNA4":"#8B4726","SILVER":"#C0C0C0","SKYBLUE":"#87CEEB","SKYBLUE1":"#87CEFF","SKYBLUE2":"#7EC0EE","SKYBLUE3":"#6CA6CD","SKYBLUE4":"#4A708B","SLATEBLUE":"#6A5ACD","SLATEBLUE1":"#836FFF",
                    "SLATEBLUE2":"#7A67EE","SLATEBLUE3":"#6959CD","SLATEBLUE4":"#473C8B","SLATEGRAY":"#708090","SLATEGRAY1":"#C6E2FF","SLATEGRAY2":"#B9D3EE","SLATEGRAY3":"#9FB6CD","SLATEGRAY4":"#6C7B8B","SNOW1":"#FFFAFA",
                    "SNOW2":"#EEE9E9","SNOW3":"#CDC9C9","SNOW4":"#8B8989","SPRINGGREEN":"#00FF7F","SPRINGGREEN1":"#00EE76","SPRINGGREEN2":"#00CD66","SPRINGGREEN3":"#008B45","STEELBLUE":"#4682B4","STEELBLUE1":"#63B8FF",
                    "STEELBLUE2":"#5CACEE","STEELBLUE3":"#4F94CD","STEELBLUE4":"#36648B","TAN":"#D2B48C","TAN1":"#FFA54F","TAN2":"#EE9A49","TAN3":"#CD853F","TAN4":"#8B5A2B","TEAL":"#008080","THISTLE":"#D8BFD8",
                    "THISTLE1":"#FFE1FF","THISTLE2":"#EED2EE","THISTLE3":"#CDB5CD","THISTLE4":"#8B7B8B","TOMATO1":"#FF6347","TOMATO2":"#EE5C42","TOMATO3":"#CD4F39","TOMATO4":"#8B3626","TURQUOISE":"#40E0D0",
                    "TURQUOISE1":"#00F5FF","TURQUOISE2":"#00E5EE","TURQUOISE3":"#00C5CD","TURQUOISE4":"#00868B","TURQUOISEBLUE":"#00C78C","VIOLET":"#EE82EE","VIOLETRED":"#D02090","VIOLETRED1":"#FF3E96",
                    "VIOLETRED2":"#EE3A8C","VIOLETRED3":"#CD3278","VIOLETRED4":"#8B2252","WARMGREY":"#808069","WHEAT":"#F5DEB3","WHEAT1":"#FFE7BA","WHEAT2":"#EED8AE","WHEAT3":"#CDBA96","WHEAT4":"#8B7E66","WHITE":"#FFFFFF",
                    "WHITESMOKE":"#F5F5F5","YELLOW1":"#FFFF00","YELLOW2":"#EEEE00","YELLOW3":"#CDCD00","YELLOW4":"#8B8B00","RED":"#FF0000","LIME":"#00FF00","YELLOW":"#FFFF00","CYAN":"#00FFFF"}

cellcolordict = {"A":"white","V":"rose","K":"pale_blue"}

####################### PROCEDURES #######################

def getperiod(age):
    for period in geologicaltimescale:
        starttime = period[2]
        endtime = period[3]
        if starttime < age <= endtime:
            return(period)

def enterpoleangdist():
    enterpoleangdist = raw_input('\n' + "What is the minimum angular distance (in degrees) that must separate poles (+ %s95s)\nin order to be considered different for criterion 7? " % alpha_symbol)
    try:   # verify that the minimum distance entered is a number.
        enterpoleangdist = float(enterpoleangdist)
    except ValueError:
        enterfloat = False
        while not(enterfloat):
            enterpoleangdist = raw_input('\n'+ "Invalid response. Minimum angular distance must be a number." + '\n' + "Please re-enter the minimum angular distance: ")
            try:
                enterpoleangdist = float(enterpoleangdist)
                enterfloat = True
            except ValueError:  
                enterfloat = False

    goodangdist = (0 <= enterpoleangdist <= 360)
    while not(goodangdist):
        enterpoleangdist = raw_input('\n'+ "Invalid response. Minimum angular distance must be >= 0 and <= 360." + '\n' + "Please re-enter minimum angular distance: ")
        try:   # verify that the minimum distance entered is a number.
            enterpoleangdist = float(enterpoleangdist)
        except ValueError:
            enterfloat = False
            while not(enterfloat):
                enterpoleangdist = raw_input('\n'+ "Invalid response. Minimum angular distance must be a number." + '\n' + "Please re-enter the minimum angular distance: ")
                try:
                    enterpoleangdist = float(enterpoleangdist)
                    enterfloat = True
                except ValueError:  
                    enterfloat = False
        goodangdist = (0 <= enterpoleangdist <= 360)

    minpoleangdist = float(enterpoleangdist)
    return(minpoleangdist)

def enterpoletempdist():
    enterpoletempdist = raw_input('\n' + "What is the minimum time (in Ma) that must separate poles (+ %s95s)\nin order to be considered an overprint for criterion 7? " % alpha_symbol)
    try:   # verify that the minimum time entered is a number.
        enterpoletempdist = float(enterpoletempdist)
    except ValueError:
        enterfloat = False
        while not(enterfloat):
            enterpoletempdist = raw_input('\n'+ "Invalid response. Minimum time must be a number." + '\n' + "Please re-enter the minimum time: ")
            try:
                enterpoletempdist = float(enterpoletempdist)
                enterfloat = True
            except ValueError:  
                enterfloat = False

    goodtimedist = (0 <= enterpoletempdist)
    while not(goodtimedist):
        enterpoletempdist = ('\n'+ "Invalid response. Minimum time must be >= 0." + '\n' + "Please re-enter minimum time: ")
        try:   # verify that the minimum time entered is a number.
            enterpoletempdist = float(enterpoletempdist)
        except ValueError:
            enterfloat = False
            while not(enterfloat):
                enterpoletempdist = raw_input('\n'+ "Invalid response. Minimum time must be a number." + '\n' + "Please re-enter the minimum time: ")
                try:
                    enterpoletempdist = float(enterpoletempdist)
                    enterfloat = True
                except ValueError:  
                    enterfloat = False
        goodtimedist = (0 <= enterpoletempdist)

    minpoletempdist = float(enterpoletempdist)
    return(minpoletempdist)

def getdatachartfoldername(filepath):
    data_chartfoldernum = 0
    for root, dirs, files in os.walk(filepath, topdown=False):
        for name in dirs:
            if "Database_Charts" in name:
                newdata_chartfoldernum_str = re.sub('[^0-9]', "", name)
                if (newdata_chartfoldernum_str != ''):
                    newdata_chartfoldernum = int(newdata_chartfoldernum_str)
                    if newdata_chartfoldernum > data_chartfoldernum:
                        data_chartfoldernum = newdata_chartfoldernum
    data_chartfoldernum = data_chartfoldernum + 1
    data_chartfoldernum_str = str(data_chartfoldernum)
    data_chartfoldername = "Database_Charts" + data_chartfoldernum_str
    return(data_chartfoldername)

def loadallpoles(wb):
    allpoles_list = []
    for sheet in wb.sheets(): #Load all poles into allpoles_list
        if sheet.name == "Poledata":
            number_of_rows = sheet.nrows
            number_of_columns = sheet.ncols
            rows = []
            for row in range(1, number_of_rows):
                rowitems = []
                values = []
        #        print("row = %s" % row)
                
                for col in range(number_of_columns):
                    value  = (sheet.cell(row,col).value)
                    rowitems.append(value)
        #            print("col = %s" % col)

                inpaleomagia = rowitems[0]

                unit = str(rowitems[1])
                unit_decoded = unit.encode('latin1')
                
                component = str(rowitems[2])
                component_decoded = component.encode('latin1')

                if component_decoded != "":
                    unittext = "%s (%s)" % (unit_decoded,component_decoded)
                else:
                    unittext = unit_decoded

                code = rowitems[3]
                code_decoded = code.encode('latin1')
                
                polelat = rowitems[4]
                polelat_num = float(polelat)
                polelat_str = ("%.1f" % polelat_num)

                polelon = rowitems[5]
                polelon_num = float(polelon)
                polelon_str = ("%.1f" % polelon_num)
                
                antipolelat = rowitems[6]
                antipolelat_num = float(antipolelat)
                
                antipolelon = rowitems[7]
                antipolelon_num = float(antipolelon)
                
                polea95 = rowitems[8]
                if polea95:
                    polea95_num = float(polea95)
                    polea95_str = ("%.1f" % polea95_num)
                else:
                    polea95_num = ""
                    polea95_str = ""
               
                polek = rowitems[9]
                if polek:
                    polek_num = float(polek)
                    polek_str = ("%.1f" % polek_num)
                else:
                    polek_num = ""
                    polek_str = ""

                poledp = rowitems[10]
                if poledp:
                    poledp_num = float(poledp)
                    poledp_str = ("%.1f" % poledp_num)
                else:
                    poledp_num = ""
                    poledp_str = ""

                poledm = rowitems[11]
                if poledm:
                    poledm_num = float(poledm)
                    poledm_str = ("%.1f" % poledm_num)
                else:
                    poledm_num = ""
                    poledm_str = ""
                
                if polea95_num != "":
                    pole_radius = polea95_num
                else:
                    pole_radius = max(poledp_num,poledm_num)

                paleolat = rowitems[12]
                paleolat_num = float(paleolat)

                dec = rowitems[13]
                dec_num = float(dec)
                dec_str = ("%.1f" % dec_num)

                inc = rowitems[14]
                inc_num = float(inc)
                inc_str = ("%.1f" % inc_num)

                directiona95 = rowitems[15]
                if directiona95 == "":
                    directiona95_num = 0
                else:
                    if directiona95 == "90":
                        directiona95_num = 90.0001
                    else:
                        directiona95_num = float(directiona95)
                directiona95_str = ("%.1f" % directiona95_num)

                directionk = rowitems[16]
                if directionk:
                    directionk_num = float(directionk)
                    directionk_str = ("%.1f" % directionk_num)
                else:
                    directionk_num = ""
                    directionk_str = ""

                ageprefix = rowitems[17]
                try:
                    ageprefix_num = float(ageprefix)
                    ageprefix_decoded = str(ageprefix_num)
                except:
                    ageprefix_decoded = ageprefix.encode('latin1')
                
                age = rowitems[18]
                age_int = int(round(age))
                age_str = str(age_int)

                agepostfix = rowitems[19]
                try:
                    agepostfix_num = int(float(agepostfix))
                    agepostfix_decoded = str(agepostfix_num)
                except:
                    agepostfix_decoded = agepostfix.encode('latin1')

                upperage = re.sub("[^0-9]", "",agepostfix_decoded)
                if upperage == "":
                    upperage = "99999"
                upperage_int = int(upperage)

                agetext = ageprefix_decoded + age_str
                if agepostfix_decoded != "":
                    agetext = agetext + agepostfix_decoded

                pmagref = rowitems[20]
                pmagref_decoded = pmagref.encode('latin1')

                ageref = rowitems[21]
                ageref_decoded = ageref.encode('latin1')
                
                pmagref_text = "%s\n(%s)" % (pmagref_decoded,ageref_decoded)

                vvq1 = int(rowitems[22])
                vvq2 = int(rowitems[23])
                vvq3 = int(rowitems[24])
                vvq4 = int(rowitems[25])
                vvq5 = int(rowitems[26])
                vvq6 = int(rowitems[27])
                vvq7 = 0
                vvqtotal = int(rowitems[29])

                vvq1_str = str(vvq1)
                vvq2_str = str(vvq2)
                vvq3_str = str(vvq3)
                vvq4_str = str(vvq4)
                vvq5_str = str(vvq5)
                vvq6_str = str(vvq6)
                vvqtotal_str = str(vvqtotal)
                
                plateid = "GPlates-0c9adcac-2142-486c-ac9f-95868320515f"
                platerevision = "GPlates-441bb92b-b8b4-44c6-8c5f-e283f2312cc7"
                
         #       print("unit = %s" % unit)
         #       print("component = %s" % component)
         #       print("code = %s" % code)
         #       print("polelat = %s" % polelat)
         #       print("polelon = %s" % polelon)
         #       print("polea95 = %s" % polea95)
         #       print("polek = %s" % polek)
         #       print("poledp = %s" % poledp)
         #       print("poledm = %s" % poledm)
         #       print("paleolat = %s" % paleolat)
         #       print("dec = %s" % dec)
         #       print("inc = %s" % inc)
         #       print("directiona95 = %s" % directiona95)
         #       print("directionk = %s" % directionk)
         #       print("age = %s" % age)
         #       print("agemod = %s" % agemod)
         #       print("ref = %s" % ref)
                 
                writetext_list = [unittext,agetext,dec_str,inc_str,directionk_str,directiona95_str,polelat_str,polelon_str,poledp_str,poledm_str,polek_str,polea95_str,
                                  vvq1_str,vvq2_str,vvq3_str,vvq4_str,vvq5_str,vvq6_str,"0",vvqtotal_str,pmagref_text,"","A"]
                
                newpole = [ageprefix_decoded,age_int,agepostfix_decoded,unit_decoded,component_decoded,code_decoded,
                           polelat_num,polelon_num,antipolelat_num,antipolelon_num,polea95_num,polek_num,poledp_num,poledm_num,pole_radius,
                           paleolat_num,dec_num,inc_num,directiona95_num,directionk_num,pmagref_decoded,ageref_decoded,
                           vvq1,vvq2,vvq3,vvq4,vvq5,vvq6,vvq7,vvqtotal,plateid,platerevision,writetext_list,"A",upperage_int]

                allpoles_list.append(newpole)
                allpoles_list_sorted_age = sorted(allpoles_list, key = lambda item: item[1])
    return(allpoles_list_sorted_age)

def processverifiedpoles():
    global allpoles_list_sorted_age
    
    verifiedpoles_list = []
    for aindex, pole in enumerate(allpoles_list_sorted_age):
        vvqtotal = pole[29]
        if vvqtotal >= (verifiedqmin-1):
            allpoles_list_sorted_age[aindex][32][22] = "V"
            allpoles_list_sorted_age[aindex][33] = "V"
            verifiedpoles_list.append(allpoles_list_sorted_age[aindex])
            break

    if verifiedpoles_list == []:
        return(verifiedpoles_list)
    
    for aindex, pole in enumerate(allpoles_list_sorted_age):
        age = pole[1]
        code = pole[5]
        #period = getperiod(age)
        pole_radius = pole[14]
        vvq1 = pole[22]
        vvq3 = pole[24]
        vvq4 = pole[25]
        vvq6 = pole[27]
        vvqtotal = pole[29]

        vvq7 = 1
        inlist = False
        youngerpoles_text = ""
        for vpole in verifiedpoles_list:
            v_code = vpole[5]
            if v_code != code:
                v_age = vpole[1]
                #v_period = getperiod(v_age)
                #if period[0] - v_period[0] > 1:
                agediff = age - v_age
                if agediff > minpoletempdist:
                    v_pole_radius = vpole[14]
                    poledist = getpoledist(pole,vpole)
                    a95dist = poledist - pole_radius - v_pole_radius
                    if a95dist <= minpoleangdist:
                        vvq7 = 0
                        if vpole[4] != "":
                            spacer = " "
                        else:
                            spacer= ""
                        youngerpoles_text = youngerpoles_text + ("%s%s%s, Age: %s Ma; " % (vpole[3], spacer, vpole[4], v_age) )
            else:
                inlist = True

        newvvqtotal = vvqtotal + vvq7
        allpoles_list_sorted_age[aindex][28] = vvq7
        allpoles_list_sorted_age[aindex][29] = newvvqtotal
        allpoles_list_sorted_age[aindex][32][18] = str(vvq7)
        allpoles_list_sorted_age[aindex][32][19] = str(newvvqtotal)
        allpoles_list_sorted_age[aindex][32][21] = youngerpoles_text
        if (newvvqtotal >= verifiedqmin) and not inlist:
            allpoles_list_sorted_age[aindex][32][22] = "V"
            allpoles_list_sorted_age[aindex][33] = "V"
            verifiedpoles_list.append(allpoles_list_sorted_age[aindex])
    return(verifiedpoles_list)

def processkeypoles():
    global allpoles_list_sorted_age
    
    keypoles_list = []
    for aindex, pole in enumerate(allpoles_list_sorted_age):
        vvq1 = pole[22]
        vvq3 = pole[24]
        vvq4 = pole[25]
        vvq6 = pole[27]
        if (vvq1 == 1) and (vvq3 == 1) and (vvq4 == 1 or vvq6 == 1):
            allpoles_list_sorted_age[aindex][32][22] = "K"
            allpoles_list_sorted_age[aindex][33] = "K"
            keypoles_list.append(allpoles_list_sorted_age[aindex])
            break

    if keypoles_list == []:
        return(keypoles_list)

    for aindex, pole in enumerate(allpoles_list_sorted_age):
        age = pole[1]
        code = pole[5]
        #period = getperiod(age)
        pole_radius = pole[14]
        vvq1 = pole[22]
        vvq3 = pole[24]
        vvq4 = pole[25]
        vvq6 = pole[27]
        vvqtotal = pole[29]

        vvq7 = 1
        inlist = False
        youngerpoles_text = ""
        for index, kpole in enumerate(keypoles_list):
            polenum = index + 1
            k_code = kpole[5]
            if k_code != code:
                k_age = kpole[1]
                #k_period = getperiod(k_age)
                #if period[0] - k_period[0] > 1:
                agediff = age - k_age
                if agediff > minpoletempdist:
                    k_pole_radius = kpole[14]
                    poledist = getpoledist(pole,kpole)
                    a95dist = poledist - pole_radius - k_pole_radius
                    if a95dist <= minpoleangdist:
                        vvq7 = 0
                        if kpole[4] != "":
                            spacer = " "
                        else:
                            spacer= ""
                        youngerpoles_text = youngerpoles_text + ("%s%s%s, Age: %s Ma; " % (kpole[3], spacer, kpole[4], k_age) )
            else:
                inlist = True

        newvvqtotal = vvqtotal + vvq7
        if (vvq1 == 1) and (vvq3 == 1) and (vvq4 == 1 or vvq6 == 1) and not inlist:
            allpoles_list_sorted_age[aindex][28] = vvq7
            allpoles_list_sorted_age[aindex][29] = newvvqtotal
            allpoles_list_sorted_age[aindex][32][18] = str(vvq7)
            allpoles_list_sorted_age[aindex][32][19] = str(newvvqtotal)
            allpoles_list_sorted_age[aindex][32][21] = youngerpoles_text
            allpoles_list_sorted_age[aindex][32][22] = "K"
            allpoles_list_sorted_age[aindex][33] = "K"
            keypoles_list.append(allpoles_list_sorted_age[aindex])
    return(keypoles_list)

def getpoledist(pole1,pole2):
    pole1_lat = pole1[6]
    pole1_lon = pole1[7]
    antipole1_lat = pole1[8]
    antipole1_lon = pole1[9]

    pole2_lat = pole2[6]
    pole2_lon = pole2[7]
    
    if (abs(pole1_lon - pole2_lon) % 360) > 180:
        deltalon = 360 - (abs(pole1_lon - pole2_lon) % 360)
    else:
        deltalon = abs(pole1_lon - pole2_lon) % 360

    normalpoledist = math.degrees(math.acos((math.sin(math.radians(pole1_lat)) * math.sin(math.radians(pole2_lat))) +
                                      (math.cos(math.radians(pole1_lat)) * math.cos(math.radians(pole2_lat)) * math.cos(math.radians(deltalon)))))

    if (abs(antipole1_lon - pole2_lon) % 360) > 180:
        antideltalon = 360 - (abs(antipole1_lon - pole2_lon) % 360)
    else:
        antideltalon = abs(antipole1_lon - pole2_lon) % 360

    antipoledist = math.degrees(math.acos((math.sin(math.radians(antipole1_lat)) * math.sin(math.radians(pole2_lat))) +
                                      (math.cos(math.radians(antipole1_lat)) * math.cos(math.radians(pole2_lat)) * math.cos(math.radians(antideltalon)))))
    truepoledist = min(normalpoledist,antipoledist)
    return(truepoledist)

def getoppositecolor(color):
    color = color.upper()
    if color in specialcolorset:
        hexval = specialcolordict[color]
    elif color != None:
        hexval = color
    else:
        return(None)
    redhexval = hexval[1]+hexval[2]
    greenhexval = hexval[3]+hexval[4]
    bluehexval = hexval[5]+hexval[6]
    reddecval = int(redhexval, 16)
    greendecval = int(greenhexval, 16)
    bluedecval = int(bluehexval, 16)
    red_decinvert = -reddecval + 255
    green_decinvert = -greendecval + 255
    blue_decinvert = -bluedecval + 255

    if 90 < red_decinvert < 150:
        red_hexinvert = "FF"
    elif red_decinvert <= 15:
        red_hexinvert = hex(red_decinvert).replace("x", "")
    else:
        red_hexinvert = hex(red_decinvert).replace("0x", "")

    if 90 < green_decinvert < 150:
        green_hexinvert = "FF"
    elif green_decinvert <= 15:
        green_hexinvert = hex(green_decinvert).replace("x", "")
    else:
        green_hexinvert = hex(green_decinvert).replace("0x", "")

    if 90 < blue_decinvert < 150:
        blue_hexinvert = "FF"
    if blue_decinvert <= 15:
        blue_hexinvert = hex(blue_decinvert).replace("x", "")
    else:
        blue_hexinvert = hex(blue_decinvert).replace("0x", "")
        
    color_invert_hex = "#" + red_hexinvert + green_hexinvert + blue_hexinvert
    return(color_invert_hex)

def makesavereferences(directionlist,plotname,folderpath,periodname,shownumbers,combinevk):
    if shownumbers:
        if directionlist:
            caption = "References: "
            for dirnum, direction in enumerate(directionlist):
                age = direction[0]
                refinfo = direction[7]
                refnum = dirnum + 1
                if combinevk:
                    poletype = direction[9]
                    caption = caption + "%s: (%s) %s; " % (refnum, poletype, refinfo)
                else:
                    caption = caption + "%s: %s; " % (refnum, refinfo)
            caption = caption[:-2] + "."
        else:
            caption = ""
        
        captionfilename = periodname + "_" + plotname + ".txt"
        captionfile = open(os.path.join(folderpath, captionfilename), 'a')
        captionfile.write(caption)
        captionfile.close()
        print("Saving caption file %s" % captionfilename)
    return

def makensaveeqareaplot(directionlist,plotname,folderpath,periodname,shownumbers,combinevk):
    fig = plt.figure(figsize=(12,12),facecolor='white')
    fig = plt.gcf()
    ax = fig.gca()
    plt.subplots_adjust(left=0.05, right=1.15, top=1.15, bottom=0.05)
    fig.tight_layout()

    axeslength = 500
    eqarearadius = 498

    ax.clear()
    ax.axis('off')
    ax.axis('equal')
    ax.axis([-axeslength, axeslength, -axeslength, axeslength])

    maincircle = plt.Circle((0, 0), eqarearadius, color="Black", fill=False, lw=1)
    ax.add_artist(maincircle)
    
    ax.plot([0,0],[-6,6],"Black",lw=1,zorder=0) # Draw Cross
    ax.plot([-6,6],[0,0],"Black",lw=1,zorder=0)

    tickmarker = 0 # plot circle ticks
    while tickmarker < 360: 
        tick_length = eqcircletick_length_dict[tickmarker]
        tickmarker_rad = math.radians(tickmarker)
        x1 = math.sin(tickmarker_rad)*(eqarearadius-tick_length)
        x2 = math.sin(tickmarker_rad)*eqarearadius
        y1 = math.cos(tickmarker_rad)*(eqarearadius-tick_length)
        y2 = math.cos(tickmarker_rad)*eqarearadius
        ax.plot([x1,x2],[y1,y2],"black",lw=1,zorder=0)
        tickmarker = tickmarker + 30

    tickmarker = 30 # plot axes ticks
    while tickmarker < 90:
        tick_length = eqaxestick_length_dict[tickmarker]
        tickdist = math.sqrt(1-(math.sin(math.radians(math.fabs(tickmarker)))))*eqarearadius
        ax.plot([tickdist,tickdist],[-tick_length,tick_length],"Black",lw=1,zorder=0)
        ax.plot([-tick_length,tick_length],[tickdist,tickdist],"Black",lw=1,zorder=0)
        ax.plot([-tickdist,-tickdist],[-tick_length,tick_length],"Black",lw=1,zorder=0)
        ax.plot([-tick_length,tick_length],[-tickdist,-tickdist],"Black",lw=1,zorder=0)
        tickmarker = tickmarker + 30

    chart_label_fontsize = 20
    label_fontsize = 18
    symbol = "o"
    drawcolor = "black"
    symbolsize = 500

    ax.text(0, eqarearadius + 7, "N", color="Black", ha="center", fontsize=chart_label_fontsize,zorder=0)
    ax.text(eqarearadius + 7, 0, "E", color="Black", va="center", fontsize=chart_label_fontsize,zorder=0)
    ax.text(0, -(eqarearadius + 7), "S", color="Black", ha="center", va="top", fontsize=chart_label_fontsize,zorder=0)
    ax.text(-(eqarearadius + 7), 0, "W", color="Black", va="center", ha="right", fontsize=chart_label_fontsize,zorder=0)

    if directionlist:
        for dirnum, direction in enumerate(directionlist):
            age = direction[0]
            dec = direction[1]
            inc = direction[2]
            refnum = dirnum + 1

            dec_rad = math.radians(dec)
            inc_rad = math.radians(inc)
            forward = math.sqrt(1-(math.sin(math.radians(math.fabs(inc)))))*eqarearadius
            point_x = math.sin(dec_rad) * forward
            point_y = math.cos(dec_rad) * forward

            if combinevk:
                poletype = direction[9]
                if poletype == "K":
                    drawcolor = "#0000FF" # blue
                elif poletype == "V":
                    drawcolor = "#FF0000" # red
                
            if inc >= 0:
                newpoint = ax.scatter(point_x, point_y, color=drawcolor, s=symbolsize, marker=symbol, edgecolor=None,linewidth=1,zorder=age)
                if shownumbers:
                    ax.text(point_x+0.5, point_y-1, refnum,color=getoppositecolor(drawcolor),va='center',ha='center',zorder=age,fontsize=label_fontsize,fontweight='normal')
            else:
                newpoint = ax.scatter(point_x, point_y, color="white", s=symbolsize, marker=symbol, edgecolor=drawcolor,linewidth=1,zorder=age+1)
                if shownumbers:
                    ax.text(point_x+0.5, point_y-1, refnum,color=drawcolor,va='center',ha='center',zorder=age+1,fontsize=label_fontsize,fontweight='normal')
    
    saveimagename = periodname + "_" + plotname + "_directions.png"
    saveimagepath = folderpath + "\\" + saveimagename
    print("Saving image %s" % saveimagename)
    plt.savefig(saveimagepath)
    plt.close()
    return

def goesthroughzero(angle1,angle2):
    # Code adapted from https://stackoverflow.com/questions/11406189/determine-if-angle-lies-between-2-other-angles
    angulardist = ((angle2 - angle1) % 360 + 360) % 360
    if angulardist >= 180:
        angle1,angle2 = angle2,angle1
    if angle1 <= angle2:
        anglegoesthroughzero = (angle1 <= 0) and (angle2 >= 0)
    else:
        anglegoesthroughzero = (angle1 <=0) or (angle2 >= 0)
    return(anglegoesthroughzero)

def makensaveeqareaa95plot(directionlist,plotname,folderpath,periodname,shownumbers,combinevk,doantipoles):
    fig = plt.figure(figsize=(12,12),facecolor='white')
    fig = plt.gcf()
    ax = fig.gca()
    plt.subplots_adjust(left=0.05, right=1.15, top=1.15, bottom=0.05)
    fig.tight_layout()

    axeslength = 500
    eqarearadius = 498
    cosmologicalconstant = 1.57 # Don't touch this setting or bad things will happen and the program won't work (and the universe may implode).

    ax.clear()
    ax.axis('off')
    ax.axis('equal')
    ax.axis([-axeslength, axeslength, -axeslength, axeslength])

    maincircle = plt.Circle((0, 0), eqarearadius, color="Black", fill=False, lw=1)
    ax.add_artist(maincircle)
    
    ax.plot([0,0],[-6,6],"Black",lw=1,zorder=0) # Draw Cross
    ax.plot([-6,6],[0,0],"Black",lw=1,zorder=0)

    tickmarker = 0 # plot circle ticks
    while tickmarker < 360: 
        tick_length = eqcircletick_length_dict[tickmarker]
        tickmarker_rad = math.radians(tickmarker)
        x1 = math.sin(tickmarker_rad)*(eqarearadius-tick_length)
        x2 = math.sin(tickmarker_rad)*eqarearadius
        y1 = math.cos(tickmarker_rad)*(eqarearadius-tick_length)
        y2 = math.cos(tickmarker_rad)*eqarearadius
        ax.plot([x1,x2],[y1,y2],"black",lw=1,zorder=0)
        tickmarker = tickmarker + 30

    tickmarker = 30 # plot axes ticks
    while tickmarker < 90:
        tick_length = eqaxestick_length_dict[tickmarker]
        tickdist = math.sqrt(1-(math.sin(math.radians(math.fabs(tickmarker)))))*eqarearadius
        ax.plot([tickdist,tickdist],[-tick_length,tick_length],"Black",lw=1,zorder=0)
        ax.plot([-tick_length,tick_length],[tickdist,tickdist],"Black",lw=1,zorder=0)
        ax.plot([-tickdist,-tickdist],[-tick_length,tick_length],"Black",lw=1,zorder=0)
        ax.plot([-tick_length,tick_length],[-tickdist,-tickdist],"Black",lw=1,zorder=0)
        tickmarker = tickmarker + 30

    chart_label_fontsize = 20
    label_fontsize = 24
    symbol = "o"
    symbolsize = 500
    ellipse_symbol = "o"

    ax.text(0, eqarearadius + 7, "N", color="Black", ha="center", fontsize=chart_label_fontsize,zorder=0)
    ax.text(eqarearadius + 7, 0, "E", color="Black", va="center", fontsize=chart_label_fontsize,zorder=0)
    ax.text(0, -(eqarearadius + 7), "S", color="Black", ha="center", va="top", fontsize=chart_label_fontsize,zorder=0)
    ax.text(-(eqarearadius + 7), 0, "W", color="Black", va="center", ha="right", fontsize=chart_label_fontsize,zorder=0)

    if directionlist:
        for dirnum, direction in enumerate(directionlist):
            age = direction[0]
            dec = direction[1]
            inc = direction[2]
            ell_a95 = direction[3]
          
            reference = direction[7]
            zplotorder = -int(round(ell_a95,0)) + 500
            refnum = dirnum + 1

            dec_rad = math.radians(dec)
            inc_rad = math.radians(inc)
            forward = math.sqrt(1-(math.sin(math.radians(math.fabs(inc)))))*eqarearadius
            point_x = math.sin(dec_rad) * forward
            point_y = math.cos(dec_rad) * forward

            if inc >= 0:
                up = False
                drawcolor = '#0000FF' # blue
            else:
                up = True
                drawcolor = '#FF0000' # red
            textcolor = getoppositecolor(drawcolor)

            hatch = False
            if combinevk:
                poletype = direction[9]
                if poletype == "V":
                    hatch = "\\"
                    lw = 1
                    ls = "-"
                elif poletype == "K":
                    hatch = "//"
                    lw = 4
                    ls = "--"
            else:
                hatch = ""
                lw = 1
                ls = "-"

            ls = "-" #This is to remove the hatching, etc.
            lw=1
            hatch = False
                    
            if shownumbers:
                ax.text(point_x+0.5, point_y-1, refnum,color=textcolor,va='center',ha='center',zorder=zplotorder+1,fontsize=label_fontsize,fontweight='normal')

            if ell_a95 == 0:
                continue
            
            ell_x_list = []
            ell_y_list = []
            
            if ell_a95 > 90: # Note: Much of the math in this section is borrowed from pmagpy. Thank you!
                bigellipse = True
                ell_a95 = 180 - ell_a95
                dec = dec - 180
                inc = -inc
            else:
                bigellipse = False
            
            rad_ell_a95 = math.radians(ell_a95)
           
            trans_matrix = [[0,0,0],[0,0,0],[0,0,0]]
            
            north = math.cos(dec_rad) * math.cos(inc_rad)
            east = math.sin(dec_rad) * math.cos(inc_rad)
            down = math.sin(inc_rad)

            if down < 0:
                north = -north
                east = -east
                down = -down

            trans_matrix[0][2] = north
            trans_matrix[1][2] = east
            trans_matrix[2][2] = down

            binc = inc - ( (abs(inc)/inc) * 90 )
            rad_binc = math.radians(binc)
            
            north = math.cos(dec_rad) * math.cos(rad_binc)
            east = math.sin(dec_rad) * math.cos(rad_binc)
            down = math.sin(rad_binc)

            if down < 0:
                north = -north
                east = -east
                down = -down

            trans_matrix[0][0] = north
            trans_matrix[1][0] = east
            trans_matrix[2][0] = down

            gdec = dec + 90
            rad_gdec = math.radians(gdec)
            ginc = 0
            
            north = math.cos(rad_gdec) * math.cos(ginc)
            east = math.sin(rad_gdec) * math.cos(ginc)
            down = math.sin(ginc)

            if down < 0:
                north = -north
                east = -east
                down = -down

            trans_matrix[0][1] = north
            trans_matrix[1][1] = east
            trans_matrix[2][1] = down

            #ellipse_densityfactor = 1
            ellipse_densityfactor = 50
            totdrawpoints = ell_a95 * ellipse_densityfactor
            totdrawpoints_int = int(totdrawpoints)

##            pointsperdegree = float(float(totdrawpoints) / 360)
##            startdegrees = 0
##            enddegrees = int(round(360*pointsperdegree,0))

            #print("%s. Reference %s; drawcolor = %s; a95 = %s" % (refnum, reference, drawcolor, ell_a95) )

            upsidedown = False
            upsidedown_x_list = []
            upsidedown_y_list = []
            usdnum = 0
            for pointnum in range(0,totdrawpoints_int):
                degrees = pointnum / (totdrawpoints-1) * 360
                rad_pointnum = math.radians(degrees)
                vector_matrix = [0,0,0]
                vector_matrix[0] = math.sin(rad_ell_a95) * math.cos(rad_pointnum)
                vector_matrix[1] = math.sin(rad_ell_a95) * math.sin(rad_pointnum)
                vector_matrix[2] = math.sqrt(1.0 - vector_matrix[0]**2 - vector_matrix[1]**2)
                ellipse = [0,0,0]       
                for j in range(0,3):
                    for k in range(0,3):
                        ellipse[j] = ellipse[j] + (trans_matrix[j][k]*vector_matrix[k])                    
                R = math.sqrt(1.0 - abs(ellipse[2])) / math.sqrt (ellipse[0]**2 + ellipse[1]**2)
                R_deg = math.degrees(R) / 90 * eqarearadius * cosmologicalconstant
                ell_x = (ellipse[1] * R_deg)
                ell_y = (ellipse[0] * R_deg)
                if up or bigellipse:
                    ell_x = -ell_x
                    ell_y = -ell_y
                #newpoint = ax.scatter(ell_x, ell_y, color=drawcolor, s=5, marker=ellipse_symbol,zorder=zplotorder)
                #ax.text(ell_x, ell_y, str(pointnum),color="black",va='center',ha='center',zorder=zplotorder+1,fontsize=label_fontsize,fontweight='normal')
                ell_x_list.append(ell_x)
                ell_y_list.append(ell_y)
                #print("%s: R=%.3f; R_deg=%.3f; ell[2]=%.3f; ell[1]=%.3f; ell[0]=%.3f; ell_x=%.3f; ell_y=%.3f" % (pointnum,R,R_deg,ellipse[2],ellipse[1],ellipse[0],ell_x,ell_y) )

                if ellipse[2] < 0:
                    upsidedown = True
                    upsidedown_x_list.append(ell_x)
                    upsidedown_y_list.append(ell_y)
                    #print("%s (%s): ell_x: %.3f, ell_y: %.3f" % (usdnum, pointnum,ell_x,ell_y) )
                    usdnum = usdnum + 1

            zippy = zip(ell_x_list, ell_y_list)
            pol = Polygon(zippy)
            fullellipse = PolygonPatch(pol, fc=drawcolor, ec='#000000', alpha=.5, lw=lw, ls=ls, hatch=hatch, zorder=zplotorder)
            ax.add_patch(fullellipse)

            if upsidedown:
                start_x = upsidedown_x_list[-1]
                start_y = upsidedown_y_list[-1]
                start_phi = math.degrees(math.atan2(start_y,start_x)) % 360
                #print("start values: start_x = %s; start_y = %s, start_phi = %s" % (start_x,start_y, start_phi))
                
                end_x = upsidedown_x_list[0]
                end_y = upsidedown_y_list[0]
                end_phi = math.degrees(math.atan2(end_y,end_x)) % 360
                #print("end values: end_x = %s; end_y = %s, end_phi = %s" % (end_x,end_y, end_phi))

                doreverse = False
                if not(goesthroughzero(start_phi,end_phi)):
                    #print("doesn'tgothroughzero")
                    if start_phi > end_phi:
                        start_phi,end_phi = end_phi,start_phi
                        doreverse = True
                else:
                    #print("goesthrouthzero")
                    bigger = max(start_phi,end_phi)
                    if bigger == end_phi:
                        doreverse = True
                    smaller = min(start_phi,end_phi)
                    start_phi = bigger - 360
                    end_phi = smaller
    
                #print(start_phi, end_phi)
                
                pointsperdegree = 5
                startpoint = int(round(pointsperdegree * start_phi,0))
                endpoint = int(round(pointsperdegree * end_phi,0))

                #print(startpoint,endpoint,doreverse)

                if not doreverse:
                    for pointnum in range(startpoint,endpoint):
                        degrees_pointnum = (float(pointnum)/pointsperdegree)
                        rad_pointnum = math.radians(degrees_pointnum)
                        upsidedown_x = math.cos(rad_pointnum) * eqarearadius
                        upsidedown_y = math.sin(rad_pointnum) * eqarearadius
                        upsidedown_x_list.append(upsidedown_x)
                        upsidedown_y_list.append(upsidedown_y)
                        #print("%s (%s) (degrees: %.5f): usd_x: %.3f, usd_y: %.3f" % (usdnum, pointnum, degrees_pointnum, upsidedown_x,upsidedown_y) )
                        usdnum = usdnum + 1
                else:
                    for pointnum in reversed(range(startpoint,endpoint)):
                        degrees_pointnum = (float(pointnum)/pointsperdegree)
                        rad_pointnum = math.radians(degrees_pointnum)
                        upsidedown_x = math.cos(rad_pointnum) * eqarearadius
                        upsidedown_y = math.sin(rad_pointnum) * eqarearadius
                        upsidedown_x_list.append(upsidedown_x)
                        upsidedown_y_list.append(upsidedown_y)
                        #print("%s (%s) (degrees: %.5f): usd_x: %.3f, usd_y: %.3f" % (usdnum, pointnum, degrees_pointnum, upsidedown_x,upsidedown_y) )
                        usdnum = usdnum + 1                    
                 
                zippy_usd = zip(upsidedown_x_list, upsidedown_y_list)
                pol_usd = Polygon(zippy_usd)
                usd_ellipse = PolygonPatch(pol_usd, fc='#800080', ec='#000000', alpha=.5, lw=lw, ls=ls, hatch=hatch, zorder=zplotorder)
                ax.add_patch(usd_ellipse)
                
    saveimagename = periodname + "_" + plotname + "_directionsa95.png"
    saveimagepath = folderpath + "\\" + saveimagename
    print("Saving image %s" % saveimagename)
    plt.savefig(saveimagepath)
    plt.close()
    return

def makensavepoleplot(directionlist,plotname,folderpath,periodname,shownumbers,combinevk):
    plt.clf()
    fig = plt.figure(figsize=(13,7),facecolor='white')
    fig = plt.gcf()
    ax = fig.gca()
    plt.subplots_adjust(left=0.05, right=0.06, top=0.06, bottom=0.05)
    fig.tight_layout()

    m = Basemap(projection='robin',lon_0=0,resolution='c')
    m.drawcoastlines()
    m.fillcontinents(color='silver',lake_color='silver')
    # draw parallels and meridians.
    m.drawparallels(np.arange(-90.,120.,30.))
    m.drawmeridians(np.arange(0.,360.,60.))
    m.drawmapboundary(fill_color='grey')
    #plt.title(periodname)

    if directionlist:
        for dirnum, direction in enumerate(directionlist):
            lat = direction[4]
            lon = direction[5]
            refnum = dirnum + 1
            x, y = m(lon, lat)

            if combinevk:
                poletype = direction[9]
                if poletype == "K":
                    drawcolor = "#0000FF" # blue
                elif poletype == "V":
                    drawcolor = "#FF0000" # red
            else:
                drawcolor = "#0000FF" # blue
            textcolor = getoppositecolor(drawcolor)

            plt.plot(x, y, 'o', markersize=18, color=drawcolor)
            if shownumbers:
                plt.text(x, y, refnum, ha="center",va="center",fontsize=12, color=textcolor)
    
    saveimagename = periodname + "_" + plotname + "_poles.png"
    saveimagepath = folderpath + "\\" + saveimagename
    print("Saving image %s" % saveimagename)
    plt.savefig(saveimagepath)
    plt.close()
    return

def makensavepolea95plot(directionlist,plotname,folderpath,periodname,shownumbers,combinevk,doantipoles):
    plt.clf()
    fig = plt.figure(figsize=(13,7),facecolor='white')
    fig = plt.gcf()
    ax = fig.gca()
    plt.subplots_adjust(left=0.05, right=0.06, top=0.06, bottom=0.05)
    fig.tight_layout()

    midlon = 0
    label_fontsize = 16
    boundary = (midlon - 180) % 360
    m = Basemap(projection='robin',lon_0=midlon,resolution='c')
    m.drawcoastlines()
    m.fillcontinents(color='silver',lake_color='silver')
    # draw parallels and meridians.
    m.drawparallels(np.arange(-90.,120.,30.))
    m.drawmeridians(np.arange(0.,360.,60.))
    m.drawmapboundary(fill_color='grey')
    plt.title(periodname)
    
    if directionlist:
        for dirnum, direction in enumerate(directionlist):
            lat = direction[4]
            lon = direction[5]
            a95 = direction[6]
            ref = direction[7]
            normalpole = direction[8]
            if not doantipoles:
                if not normalpole:
                    continue
    
            zplotorder = -int(round(a95,0)) + 500
            refnum = dirnum + 1

            if combinevk:
                poletype = direction[9]
                if poletype == "V":
                    hatch = "\\"
                    lw = 1
                    ls = "-"
                elif poletype == "K":
                    hatch = "//"
                    lw = 4
                    ls = "--"
            else:
                hatch = ""
                lw = 1
                ls = "-"

            ls = "-" #This is to remove the hatching, etc.
            lw=1
            hatch = False

            p1_lon_rad = math.radians(lon)
            p2_lon_rad = math.radians(boundary)
            if math.fabs((math.degrees(p2_lon_rad) - math.degrees(p1_lon_rad))) > 180:
                lon_diff = math.radians(360-math.fabs((math.degrees(p2_lon_rad) - math.degrees(p1_lon_rad))))
            else:
                lon_diff = math.radians(math.fabs((math.degrees(p2_lon_rad) - math.degrees(p1_lon_rad))))

            lon_diff_degrees = math.degrees(lon_diff)
            lat_rad = math.radians(lat)
            lon_diff_angdist = math.cos(lat_rad) * lon_diff_degrees
            lon_diff_realdist = math.cos(lat_rad) * lon_diff_degrees * 110.5
            radius = a95 * 110.5

            if normalpole:
                drawcolor = "blue"
                textcolor = "yellow"
            else:
                drawcolor = "red"
                textcolor = "cyan"

            if radius > lon_diff_realdist:
                #print("%s. IIIIIIIIIIINNNNNNNNNNNNNNN WRAP" % refnum)
                p = Point(m(lon, lat))
                buffered = p.buffer(radius * 1000)

                toohighlat = (math.fabs(lat) + a95) > 90
                circleset = circle_wrap(m, lon, lat, radius, midlon, boundary, lon_diff_angdist, toohighlat)
                
                circle_main = circleset[0]
                poly_main = Polygon(circle_main)
                polypatch_main = PolygonPatch(poly_main, fc=drawcolor, ec='#000000', alpha=.2, lw=lw, ls=ls, hatch=hatch, zorder=zplotorder)
                ax.add_patch(polypatch_main)

                if shownumbers:
                    x, y = m(lon, lat)
                    plt.text(x, y, refnum, ha="center",va="center",fontsize=label_fontsize, color=textcolor, zorder=zplotorder)

                circle_wrapped = circleset[1]
                poly_wrap = Polygon(circle_wrapped)
                polypatch_wrap = PolygonPatch(poly_wrap, fc=drawcolor, ec='#000000', alpha=.5, lw=lw, ls=ls, hatch=hatch, zorder=zplotorder)
                ax.add_patch(polypatch_wrap)

                if toohighlat:
                    circle_highlat = circleset[2]
                    poly_highlat = Polygon(circle_highlat)
                    polypatch_highlat = PolygonPatch(poly_highlat, fc=drawcolor, ec=None, alpha=.2, lw=0, ls=ls, hatch=hatch, zorder=zplotorder)
                    ax.add_patch(polypatch_highlat)
                    
                continue
            
            p = Point(m(lon, lat))
            buffered = p.buffer(radius * 1000)
            circle_main = circle(m, lon, lat, radius)
            poly_main = Polygon(circle_main)
            polypatch_main = PolygonPatch(poly_main, fc=drawcolor, ec='#000000', alpha=.5, lw=lw, ls=ls, hatch=hatch, zorder=zplotorder)
            ax.add_patch(polypatch_main)

            if shownumbers:
                x, y = m(lon, lat)
                plt.text(x, y, refnum, ha="center",va="center",fontsize=label_fontsize, color=textcolor, zorder=zplotorder)
    
    saveimagename = periodname + "_" + plotname + "_polesa95.png"
    saveimagepath = folderpath + "\\" + saveimagename
    print("Saving image %s" % saveimagename)
    plt.savefig(saveimagepath)
    plt.close()
    return

def dotimeperiod(verifiedpoles_list,keypoles_list,periodname,endtime,starttime,shownumbers,doantipoles):
    print("\n----------------------------\nDoing Time Period: %s. Start time = %s Ma, End time = %s Ma." % (periodname, starttime, endtime) )

    verifieddirectionlist = []
    for pole in verifiedpoles_list:
        age_int = pole[1]
        upperage_int = pole[34]
        if starttime < age_int <= endtime:
            dec_num = pole[16]
            inc_num = pole[17]
            dira95 = pole[18]
            polelat = pole[6]
            polelon = pole[7]
            polea95 = pole[10]

            unit_decoded = pole[3]
            component_decoded = pole[4]
            pmagref_decoded = pole[20]
            agetext = pole[32][1]
            poletype = pole[33]
            if component_decoded != "":
                spacer = " "
            else:
                spacer= ""                
            refinfo = ("%s: %s%s%s, Age: %s Ma" % (pmagref_decoded, unit_decoded, spacer, component_decoded, agetext) )

            verifieddirectionlist.append([age_int,dec_num,inc_num,dira95,polelat,polelon,polea95,refinfo,True,poletype,upperage_int])
            if doantipoles:
                antipolelat = pole[8]
                antipolelon = pole[9]
                antiinc = -inc_num
                antidec = (dec_num + 180) % 360
                if antidec < 0:
                    antidec = antidec + 360
                verifieddirectionlist.append([age_int,antidec,antiinc,dira95,antipolelat,antipolelon,polea95,refinfo,False,poletype,upperage_int])

    keydirectionlist = []
    for pole in keypoles_list:
        age_int = pole[1]
        upperage_int = pole[34]
        if starttime < age_int <= endtime:
            dec_num = pole[16]
            inc_num = pole[17]
            dira95 = pole[18]
            polelat = pole[6]
            polelon = pole[7]
            polea95 = pole[10]

            unit_decoded = pole[3]
            component_decoded = pole[4]
            pmagref_decoded = pole[20]
            agetext = pole[32][1]
            poletype = pole[33]
            if component_decoded != "":
                spacer = " "
            else:
                spacer= ""                
            refinfo = ("%s: %s%s%s, Age: %s Ma" % (pmagref_decoded, unit_decoded, spacer, component_decoded, agetext) )

            keydirectionlist.append([age_int,dec_num,inc_num,dira95,polelat,polelon,polea95,refinfo,True,poletype,upperage_int])
            if doantipoles:
                antipolelat = pole[8]
                antipolelon = pole[9]
                antiinc = -inc_num
                antidec = (dec_num + 180) % 360
                if antidec < 0:
                    antidec = antidec + 360
                keydirectionlist.append([age_int,antidec,antiinc,dira95,antipolelat,antipolelon,polea95,refinfo,False,poletype,upperage_int])

    verifieddirectionlist_sorted_upperage = sorted(verifieddirectionlist, key = lambda item: item[10])
    verifieddirectionlist_sorted_age = sorted(verifieddirectionlist_sorted_upperage, key = lambda item: item[0])
    keydirectionlist_sorted_upperage = sorted(keydirectionlist, key = lambda item: item[10])
    keydirectionlist_sorted_age = sorted(keydirectionlist_sorted_upperage, key = lambda item: item[0])

##    makesavereferences(verifieddirectionlist_sorted_age,"verifiedpoles",verifiedpath,periodname,shownumbers,False)
##    makensaveeqareaplot(verifieddirectionlist_sorted_age,"verifiedpoles",verifiedpath,periodname,shownumbers,False)
##    makensaveeqareaa95plot(verifieddirectionlist_sorted_age,"verifiedpoles",verifiedpath,periodname,shownumbers,False,doantipoles)
##    makensavepoleplot(verifieddirectionlist_sorted_age,"verifiedpoles",verifiedpath,periodname,shownumbers,False)
##    makensavepolea95plot(verifieddirectionlist_sorted_age,"verifiedpoles",verifiedpath,periodname,shownumbers,False,doantipoles)
##
##    makesavereferences(keydirectionlist_sorted_age,"keypoles",keypath,periodname,shownumbers,False)
##    makensaveeqareaplot(keydirectionlist_sorted_age,"keypoles",keypath,periodname,shownumbers,False)
##    makensaveeqareaa95plot(keydirectionlist_sorted_age,"keypoles",keypath,periodname,shownumbers,False,doantipoles)
##    makensavepoleplot(keydirectionlist_sorted_age,"keypoles",keypath,periodname,shownumbers,False)
##    makensavepolea95plot(keydirectionlist_sorted_age,"keypoles",keypath,periodname,shownumbers,False,doantipoles)

    makesavereferences(verifieddirectionlist_sorted_age,"combinedpoles",combinedpath,periodname,shownumbers,True)
    makensaveeqareaplot(verifieddirectionlist_sorted_age,"combinedpoles",combinedpath,periodname,shownumbers,True)
    makensaveeqareaa95plot(verifieddirectionlist_sorted_age,"combinedpoles",combinedpath,periodname,shownumbers,True,doantipoles)
    makensavepoleplot(verifieddirectionlist_sorted_age,"combinedpoles",combinedpath,periodname,shownumbers,True)
    makensavepolea95plot(verifieddirectionlist_sorted_age,"combinedpoles",combinedpath,periodname,shownumbers,True,doantipoles)
    return

####################### MAIN PROGRAM #######################

os.system('mode con: cols=150 lines=60')
print('\n' + "Hello. This script is ready to evaluate for poles near the pole you specify,\nwithin an angular distance you specify, and within temporal distance you specify.")

if not(os.path.isfile(excelfilename)): # check if the excelfilename exists in this folder.
    print('\n' + "This folder does NOT contain the file %s, which is necessary for this program to run." % excelfilename)
    endchoice = raw_input('\n' + "----- This program will now end. Press enter to exit. -----")
    sys.exit()

print('\n' + "The Excel source file %s has been found!" % excelfilename)
print('\n' + "----------------------------")

filepath = os.path.dirname(os.path.realpath(__file__))
data_chartfoldername = getdatachartfoldername(filepath)
data_chartfolderpath = filepath + "\\" + data_chartfoldername
os.makedirs(data_chartfolderpath)
newexcelfilename = data_chartfolderpath + "\\" + excelfilename
copy(filepath + "\\" + excelfilename, newexcelfilename) # Copy Excel File Into Chart + Data Folder
backupexcelfilename = excelfilename + ".bak"
copy(filepath + "\\" + excelfilename, data_chartfolderpath + "\\" + backupexcelfilename) # Backup Excel File in New Chart + Data Folder

verifiedpath = data_chartfolderpath + "\\" + "verifiedpoles"
os.makedirs(verifiedpath)
keypath = data_chartfolderpath + "\\" + "keypoles"
os.makedirs(keypath)
combinedpath = data_chartfolderpath + "\\" + "combinedpoles"
os.makedirs(combinedpath)

print('\n' + "Created new file directory %s and Excel file %s and backup Excel file %s into it." % (data_chartfoldername,newexcelfilename,backupexcelfilename))
print('\n' + "----------------------------")

minpoleangdist = enterpoleangdist()
minpoletempdist = enterpoletempdist()

wb = open_workbook(excelfilename)
allpoles_list_sorted_age = loadallpoles(wb)

verifiedpoles_list = processverifiedpoles()
keypoles_list = processkeypoles()

wb = xlwt.Workbook(encoding='latin-1')

borders4 = xlwt.Borders()
borders4.left = 1
borders4.right = 1
borders4.top = 1
borders4.bottom = 1

boldstyle = xlwt.XFStyle()
boldfont = xlwt.Font()
boldfont.name = "Times New Roman"
boldfont.height = 8 * 20
boldfont.bold = True
boldstyle.font = boldfont
boldstyle.borders = borders4

normalstyle = xlwt.XFStyle()
normalfont = xlwt.Font()
normalfont.name = "Times New Roman"
normalfont.height = 8 * 20
normalstyle.font = normalfont
normalstyle.borders = borders4

vpolesheet = wb.add_sheet("Verified Poles")
row = vpolesheet.row(0)
for index, coltext in enumerate(cols):
    vpolesheet.col(index).width = (coltext[1] * 256)
    row.write(index,coltext[0],style=boldstyle)
polenum = 1

for pole in verifiedpoles_list:
    row = vpolesheet.row(polenum)
    writetext_list = pole[32]
    for index, textitem in enumerate(writetext_list):
        row.write(index, textitem,style=normalstyle)
    polenum = polenum + 1

vpolesheet.set_panes_frozen(True)
vpolesheet.set_horz_split_pos(1) 

kpolesheet = wb.add_sheet("Key Poles")
row = kpolesheet.row(0)
for index, coltext in enumerate(cols):
    kpolesheet.col(index).width = (coltext[1] * 256)
    row.write(index,coltext[0],style=boldstyle)
polenum = 1
for pole in keypoles_list:
    row = kpolesheet.row(polenum)
    writetext_list = pole[32]
    for index, textitem in enumerate(writetext_list):
        row.write(index, textitem,style=normalstyle)     
    polenum = polenum + 1

kpolesheet.set_panes_frozen(True)
kpolesheet.set_horz_split_pos(1)

apolesheet = wb.add_sheet("All Poles")
row = apolesheet.row(0)
for index, coltext in enumerate(cols):
    apolesheet.col(index).width = (coltext[1] * 256)
    row.write(index,coltext[0],style=boldstyle)

polenum = 1
for pole in allpoles_list_sorted_age:
    row = apolesheet.row(polenum)
    writetext_list = pole[32]
    cellcolor = cellcolordict[pole[33]]
    colorstyle = xlwt.XFStyle()
    colorpattern = xlwt.Pattern()
    colorpattern.pattern = xlwt.Pattern.SOLID_PATTERN
    colorpattern.pattern_fore_colour = xlwt.Style.colour_map[cellcolor]
    colorstyle.pattern = colorpattern
    colorstyle.borders = borders4
    #colorstyle.font = normalfont
    for index, textitem in enumerate(writetext_list):
        row.write(index, textitem,style=colorstyle)     
    polenum = polenum + 1

apolesheet.set_panes_frozen(True)
apolesheet.set_horz_split_pos(1)

saveexcelpath_name = data_chartfolderpath + "\\" + finalpolesexcelfilename
wb.save(saveexcelpath_name)
print("Saving Excel File with Verified/Key Poles %s." % finalpolesexcelfilename)

##dotimeperiod(verifiedpoles_list,keypoles_list,"Eoarchean",4000,3600,True,False)
##dotimeperiod(verifiedpoles_list,keypoles_list,"Paleoarchean",3600,3200,True,False)
##dotimeperiod(verifiedpoles_list,keypoles_list,"Mesoarchean",3200,2800,True,False)
##dotimeperiod(verifiedpoles_list,keypoles_list,"Neoarchean",2800,2500,True,False)

##dotimeperiod(verifiedpoles_list,keypoles_list,"Early Paleoproterozoic",2500,2050,True,False)
##dotimeperiod(verifiedpoles_list,keypoles_list,"Late Paleoproterozoic",2050,1600,True,False)
##dotimeperiod(verifiedpoles_list,keypoles_list,"Mesoproterozoic",1600,1000,True,False)
##dotimeperiod(verifiedpoles_list,keypoles_list,"Neoproterozoic",1000,541,True,False)

##dotimeperiod(verifiedpoles_list,keypoles_list,"Paleozoic",541,252,True,False)
##dotimeperiod(verifiedpoles_list,keypoles_list,"Mesozoic",252,66,True,False)
##dotimeperiod(verifiedpoles_list,keypoles_list,"Cenozoic",66,0,True,False)

##dotimeperiod(verifiedpoles_list,keypoles_list,"All",4000,0,False,True)

##dotimeperiod(verifiedpoles_list,keypoles_list,"Phanerozoic",541,0,True,False)

dotimeperiod(verifiedpoles_list,keypoles_list,"Mesoarchean",3200,2800,True,False)

endchoice = raw_input('\n' + "----- Program complete. Goodbye! Please press enter to exit. -----")
sys.exit()
