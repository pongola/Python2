# -*- coding: utf-8 -*-
"""
circles.py
Created by Stephan Hügel on 2014-05-22
Copyright Stephan Hügel

Convenience functions for calculating and drawing circles on a projected map,
based on great-circle distances

Almost all code from:
http://www.geophysique.be/2011/02/20/matplotlib-basemap-tutorial-09-drawing-circles/

License: MIT

"""
__author__ = 'urschrei@gmail.com'
import numpy as np
import math


class CourseException(Exception):
    """ Simple error class """
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr(self.value)



def _gccalc(lon, lat, azimuth, maxdist=None):
    """
    Original javascript on http://williams.best.vwh.net/gccalc.htm
    Translated into python by Thomas Lecocq
    This function is a black box, because trigonometry is difficult
    
    """
    glat1 = lat * np.pi / 180.
    glon1 = lon * np.pi / 180.
    s = maxdist / 1.852
    faz = azimuth * np.pi / 180.
 
    EPS = 0.00000000005
    if ((np.abs(np.cos(glat1)) < EPS) and not (np.abs(np.sin(faz)) < EPS)):
        raise CourseException("Only North-South courses are meaningful")

    a = 6378.13 / 1.852
    f= 1 / 298.257223563
    r = 1 - f
    tu = r * np.tan(glat1)
    sf = np.sin(faz)
    cf = np.cos(faz)
    if (cf == 0):
        b = 0.
    else:
        b = 2. * np.arctan2 (tu, cf)

    cu = 1. / np.sqrt(1 + tu * tu)
    su = tu * cu
    sa = cu * sf
    c2a = 1 - sa * sa
    x = 1. + np.sqrt(1. + c2a * (1. / (r * r) - 1.))
    x = (x - 2.) / x
    c = 1. - x
    c = (x * x / 4. + 1.) / c
    d = (0.375 * x * x - 1.) * x
    tu = s / (r * a * c)
    y = tu
    c = y + 1
    while (np.abs (y - c) > EPS):
        sy = np.sin(y)
        cy = np.cos(y)
        cz = np.cos(b + y)
        e = 2. * cz * cz - 1.
        c = y
        x = e * cy
        y = e + e - 1.
        y = (((sy * sy * 4. - 3.) * y * cz * d / 6. + x) *
            d / 4. - cz) * sy * d + tu

    b = cu * cy * cf - su * sy
    c = r * np.sqrt(sa * sa + b * b)
    d = su * cy + cu * sy * cf
    glat2 = (np.arctan2(d, c) + np.pi) % (2*np.pi) - np.pi
    c = cu * cy - su * sy * cf
    x = np.arctan2(sy * sf, c)
    c = ((-3. * c2a + 4.) * f + 4.) * c2a * f / 16.
    d = ((e * cy * c + cz) * sy * c + y) * sa
    glon2 = ((glon1 + x - (1. - c) * d * f + np.pi) % (2*np.pi)) - np.pi

    baz = (np.arctan2(sa, b) + np.pi) % (2 * np.pi)

    glon2 *= 180./np.pi
    glat2 *= 180./np.pi
    baz *= 180./np.pi
    return (glon2, glat2, baz)

def circle(m, centerlon, centerlat, radius, *args, **kwargs):
    """
    Return lon, lat tuples of a "circle" which matches the chosen Basemap projection
    Takes the following arguments:
    m = basemap instance
    centerlon = originating lon
    centrelat = originating lat
    radius = radius

    """

    glon1 = centerlon
    glat1 = centerlat
    X = []
    Y = []
    for azimuth in range(0, 360):
        glon2, glat2, baz = _gccalc(glon1, glat1, azimuth, radius)
        X.append(glon2)
        Y.append(glat2)
    X.append(X[0])
    Y.append(Y[0])

    proj_x, proj_y = m(X,Y)
    return zip(proj_x, proj_y)

def circle_wrap(m, centerlon, centerlat, radius, midlon, boundary, lon_diff_angdist, toohighlat, *args, **kwargs):
    # This modified function revised by Casey Luskin. The rest of this file written by the original author, Stephan Hügel, urschrei@gmail.com, as noted above.
    """
    Return lon, lat tuples of a "circle" for a Robinson Projection map
    AND also returns lon, lat tuples of a "wrapped" circle on the other side of the Robinson Projection
    Takes the following arguments:
    m = basemap instance
    centerlon = originating lon
    centrelat = originating lat
    radius = radius
    midlon = middle longitude of the Robinson projection
    lon_diff_angdist which is the longitudinal distance between the center of the circle and the nearest boundary of the map

    """
    glon1 = centerlon
    glat1 = centerlat
    X = []
    Y = []
    X_wrap = []
    Y_wrap = []

    centerlonpos = centerlon % 360
    if centerlonpos == midlon:
        pos = "middle"
    elif (midlon < centerlonpos < (midlon+180)) or (0 < centerlonpos < ((midlon+180)%360)):
        pos = "east"
    else:
        pos = "west"

    for azimuth in range(0, 360):
        glon2, glat2, baz = _gccalc(glon1, glat1, azimuth, radius)

        p1_lon_rad = math.radians(centerlon)
        p2_lon_rad = math.radians(glon2)
        if math.fabs((math.degrees(p2_lon_rad) - math.degrees(p1_lon_rad))) > 180:
            lon_diff = math.radians(360-math.fabs((math.degrees(p2_lon_rad) - math.degrees(p1_lon_rad))))
        else:
            lon_diff = math.radians(math.fabs((math.degrees(p2_lon_rad) - math.degrees(p1_lon_rad))))

        center_glon2_lon_diff_degrees = math.degrees(lon_diff)
        centerlat_rad = math.radians(centerlat)
        glon2_lon_diff = math.cos(centerlat_rad) * center_glon2_lon_diff_degrees

        if pos == "west" and (180 < azimuth <= 360):
            if glon2_lon_diff > lon_diff_angdist:
                X_wrap.append(glon2)
                glon2 = -boundary
                #print('WESTover %s' % glon2)
                Y_wrap.append(glat2)
        elif pos == "east" and (0 < azimuth <= 180):
            if glon2_lon_diff > lon_diff_angdist:
                X_wrap.append(glon2)
                glon2 = boundary
                #print('EASTover %s' % glon2)
                Y_wrap.append(glat2)
            
        X.append(glon2)
        Y.append(glat2)
        #print(glon2, azimuth)
    X.append(X[0])
    Y.append(Y[0])

    A = X_wrap
    B = Y_wrap

    if pos == "west":
        sidelon = boundary
    elif pos == "east":
        sidelon = -boundary
    for item in reversed(Y_wrap):
        A.append(sidelon)
        B.append(item)
    
    A.append(A[0])
    B.append(B[0])

    proj_x, proj_y = m(X,Y)
    proj_a, proj_b = m(A,B)

    if not toohighlat:
        return [zip(proj_x, proj_y),zip(proj_a,proj_b)]
    else:
        northhem = Y[0] > 0
        C = []
        D = []
        if northhem:
            Y_max = max(Y)
            B_max = max(B)
            toplat = min([Y_max,B_max]) - 1
            toplat_big = int(round(toplat * 10,0))
            for latindex in range(toplat_big, 900):
                C.append(boundary)
                newlat = latindex / 10
                D.append(newlat)
            for latindex in reversed(range(toplat_big, 900)):
                C.append(-boundary)
                newlat = latindex / 10
                D.append(newlat)
        else:
            Y_max = min(Y)
            B_max = min(B)
            toplat = max([Y_max,B_max]) + 1
            toplat_big = int(round(toplat * 10,0))
            for latindex in reversed(range(-900, toplat_big)):
                C.append(-boundary)
                newlat = latindex / 10
                #print(newlat)
                D.append(newlat)
            for latindex in range(-900,toplat_big):
                C.append(boundary)
                #print(newlat)
                newlat = latindex / 10
                D.append(newlat)

        proj_c, proj_d = m(C,D)
        return [zip(proj_x, proj_y),zip(proj_a,proj_b),zip(proj_c,proj_d)]



        
            
