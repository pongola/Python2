###########################################
#
# PROGRAM FOR CALCULATING CUTOFF ANGLE FOR VGP DISPERSION STATISTICS. Written by Casey Luskin.
# This text-based Python script reads a list of VGP latitude and longitude values and recursively calculates the cutoff angle
# (for excluding certain VGPs from VGP dispersion calculations) following the method of Vandamme (1994).
# For support, please contact Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com
#
###########################################

import scipy.stats
import os

####### CONSTANTS ############
inputfile = "4Fish_VGPs_WD_NUBMean.txt"

########### PROCEDURES #############

def get_meancoords(pointlist):
    north_vector_sum = 0
    east_vector_sum = 0
    down_vector_sum = 0
    for point in pointlist:
        point_lat_rad = math.radians(point[0])
        point_lon_rad = math.radians(point[1])

        point_north_vector = math.cos(point_lat_rad) * math.cos(point_lon_rad)
        north_vector_sum = north_vector_sum + point_north_vector

        point_east_vector = math.cos(point_lat_rad) * math.sin(point_lon_rad)
        east_vector_sum = east_vector_sum + point_east_vector

        point_down_vector = math.sin(point_lat_rad)
        down_vector_sum = down_vector_sum + point_down_vector

    numpoints = len(pointlist)
    R = math.sqrt(north_vector_sum**2 + east_vector_sum**2 + down_vector_sum**2)
    if numpoints == int(round(R,10)):
        avglat = point[0]
        avglon = point[1]
        a95 = 0
        k = "--"
    else:
        north_dircos = north_vector_sum / R
        east_dircos = east_vector_sum / R
        down_dircos = down_vector_sum / R
        hypotenuse = math.sqrt(north_dircos**2 + east_dircos**2)
        avglat = math.degrees(math.atan2(down_dircos,hypotenuse))
        avglon = math.degrees(math.atan2(east_dircos,north_dircos)) % 360
        k = (numpoints - 1) / (numpoints - R)
        if ( 1 - ((numpoints - R) / R) * ( (20**(1/(numpoints-1))) - 1 ) ) < -1:
            a95 = "max -- 180"
        else:
            a95 = math.degrees( math.acos( 1 - ((numpoints - R) / R) * ( (20**(1/(numpoints-1))) - 1 ) ) )        
    if numpoints == 2:
        a95 = "ERR"

    return([avglat,avglon,a95,k,numpoints])

def get_coorddist(p1_lat,p1_lon,p2_lat,p2_lon):
    earth_radius = 6371.0
    p1_lat_rad = math.radians(p1_lat)
    p1_lon_rad = math.radians(p1_lon%360)
    p2_lat_rad = math.radians(p2_lat)
    p2_lon_rad = math.radians(p2_lon%360)
    if p1_lat_rad == p2_lat_rad and p1_lon_rad == p2_lon_rad:
        angdist = 0
        realdist = 0
    else:
        if math.fabs((math.degrees(p2_lon_rad) - math.degrees(p1_lon_rad))) > 180:
            lon_diff = math.radians(360-math.fabs((math.degrees(p2_lon_rad) - math.degrees(p1_lon_rad))))
        else:
            lon_diff = math.radians(math.fabs((math.degrees(p2_lon_rad) - math.degrees(p1_lon_rad))))
        angdist_rad = math.acos( (math.sin(p1_lat_rad)*math.sin(p2_lat_rad)) + (math.cos(p1_lat_rad)*math.cos(p2_lat_rad)*math.cos(lon_diff)) )
        angdist = math.degrees(angdist_rad)
        realdist = angdist_rad * earth_radius
    return([angdist,realdist])

def thetamaxok(A,vgplist):
    thetamaxisok = True
    for vgp in vgplist:
        theta = vgp[2]
        if theta > A:
            thetamaxisok = False
            break
    return(thetamaxisok)


def removebiggesttheta(vgplist):
    biggesttheta = 0
    for vgp in vgplist:
        theta = vgp[2]
        if theta > biggesttheta:
            biggesttheta = theta

    newvgplist = []
    for vgp in vgplist:
        theta = vgp[2]
        if theta != biggesttheta:
            newvgplist.append(vgp)
#        else:
#            print("Removing %s" % theta)
    return(newvgplist)

########### PROGRAM ################

filepath = os.path.dirname(os.path.realpath(__file__))
filepathandname = filepath + "\\" + inputfile

f = open(filepathandname,'r')

vgplist = []

for line in f:
    line_split = line.split()
    lon_str = (line_split[0])
    lat_str = (line_split[1])
    lon = float(lon_str)
    lat = float(lat_str)
    newvgp = [lat,lon,0]
    vgplist.append(newvgp)

poleloc = get_meancoords(vgplist)

for vgpnum, vgp in enumerate(vgplist):
    vgplat = vgp[0]
    vgplon = vgp[1]
    polelat = poleloc[0]
    polelon = poleloc[1]
    angdist = get_coorddist(vgplat,vgplon,polelat,polelon)[0]
    vgplist[vgpnum][2] = angdist

iteration = 1
ASDsum = 0
numvgps = len(vgplist)
for vgp in vgplist:
    theta = vgp[2]
    newterm = (theta**2) / (numvgps - 1)
    ASDsum = ASDsum + newterm
ASD = math.sqrt(ASDsum)
A = (1.8 * ASD) + 5
print("Iteration %s: A = %s" % (iteration, A))

while not(thetamaxok(A,vgplist)):
    iteration = iteration + 1
    vgplist = removebiggesttheta(vgplist)
    poleloc = get_meancoords(vgplist)
    for vgpnum, vgp in enumerate(vgplist):
        vgplat = vgp[0]
        vgplon = vgp[1]
        polelat = poleloc[0]
        polelon = poleloc[1]
        angdist = get_coorddist(vgplat,vgplon,polelat,polelon)[0]
        vgplist[vgpnum][2] = angdist
    ASDsum = 0
    numvgps = len(vgplist)
    for vgp in vgplist:
        theta = vgp[2]
        newterm = (theta**2) / (numvgps - 1)
        ASDsum = ASDsum + newterm
    ASD = math.sqrt(ASDsum)
    A = (1.8 * ASD) + 5
    print("Iteration %s: A = %s" % (iteration,A))
    
print("FINAL CUTOFF ANGLE (A) = %s\n\n\n" % (A) )
#for vgp in vgplist:
    #print(vgp)
