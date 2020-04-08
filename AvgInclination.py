##################################
#
# This text-based Python script follows the method outlined by Tauxe et al. (2010) after McFadden and Reid (1982) for finding a numerical solution
# to the mean inclination (and associated a95 and k values) of a set of inclination-only data. It numerically calculates the average inclination by
# trying values from the minimum to maximum inclination values in the dataset at 0.1 degree intervals. The inclination value that yields a solution to the
# equation that is closest to zero is the average inclination. The program outputs the solutions for each 0.1 degree interval, and at the end gives the best
# solution (closest to zero) as well as the associated a95 and k.
# 
# Script written by Casey Luskin.
# For support, please contact Casey Luskin at caseyl@uj.ac.za or casey.luskin@gmail.com
#
##################################

import math
import scipy.stats

incset1 = [-61.6, -61.1, -59.4, -57.8, -56.0, -50.6, -40.6, -40.6]
incset2 = [49.4, 52.6, 69.7]

incset3 = [-61.6, -61.1, -59.4, -57.8, -56.0, -50.6, -40.6, -40.6, -49.4, -52.6, -69.7]

incset4 = incset1 + incset2

incsetuse = incset3

closestmcfaddenreid = 100000
bestvaluetext = ""
coincset = []
for inc in incsetuse:
    if inc < 0:
        coinc = -90 - inc
    elif inc > 0:
        coinc = 90 - inc
    elif inc == 0:
        coinc = 0
    coincset.append(coinc)

N = len(coincset)
maxinc = max(coincset)
mininc = min(coincset)

maxincX10 = int(maxinc * 10)
minincX10 = int(mininc * 10)

for avgcoincX10 in range(minincX10, (maxincX10+10)):
    avgcoincX10 = float(avgcoincX10)
    ten = float(10)
    avgcoinc = avgcoincX10 / ten
    avgcoinc_rad = math.radians(avgcoinc)

    term1 = N * math.cos(avgcoinc_rad)

    factor2 = (math.sin(avgcoinc_rad)**2) - (math.cos(avgcoinc_rad)**2)

    factor3 = 0
    for coinc in coincset:
        coinc_rad = math.radians(coinc)
        factor3 = factor3 + math.cos(coinc_rad)

    factor4 = 2 * math.sin(avgcoinc_rad) * math.cos(avgcoinc_rad)

    factor5 = 0
    for coinc in coincset:
        coinc_rad = math.radians(coinc)
        factor5 = factor5 + coinc_rad

    term2 = factor2 * factor3

    term3 = factor4 * factor5

    mcfaddenreid = term1 + term2 - term3

    S = 0
    C = 0
    for coinc in coincset:
        coinc_rad = math.radians(coinc)
        S = S + math.sin(avgcoinc_rad - coinc_rad)
        C = C + math.cos(avgcoinc_rad - coinc_rad)

    k = (N - 1) / (2 * (N - C))

    if avgcoinc < 0:
        avginc = -90 - avgcoinc + (S / C)
    elif avgcoinc > 0:
        avginc = 90 - avgcoinc + (S / C)
    elif avgcoinc == 0:
        avginc = 0

    Nx2minus1 = 2 * (N - 1)
    critval = round(scipy.stats.f.isf(0.05, Nx2minus1, Nx2minus1),2)

    a95 = math.degrees(math.acos(1 - (0.5 * (S/C)**2) - (critval / (2 * C * k)) ))

    mcfaddenreid_text = "avginc = %s, mcfaddenreid = %s, a95 = %.1f, k = %.1f" % (avginc, mcfaddenreid, a95, k)

    if math.fabs(mcfaddenreid) < closestmcfaddenreid:
        closestmcfaddenreid = math.fabs(mcfaddenreid)
        bestvaluetext = mcfaddenreid_text

    print(mcfaddenreid_text)

print("\n============================================\n")
print("Best value: %s" % bestvaluetext)
