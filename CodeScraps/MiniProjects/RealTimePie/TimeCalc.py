# @Title:   TimeCalc
# @Version: 1.0
# @Author:  Paul F.
# @Date:    11-Dec-2018
# @Type:    Module
# Summary
#   Module that interacts with times based on established format
#   Format:  "{hh}:{mm}:{ss} {A/P}M"
#   Intended to help with handling the time outputs from Results.txt files
#   Assumes valid input
#
# Changelog:
#   Date            Author          Description
#   11-Dec-2018     Paul F.         First version, no bugs to speak of, but no error handling either
#


import time

def process(formattedTime):     #takes formatted input time and returns as int list [H, M, S]
    meridiem = 0
    timeAndMeridiem = formattedTime.split(' ')
    hms = timeAndMeridiem[0].split(':')
    if len(timeAndMeridiem) > 1:
        if hms[0] == "12":
            hms[0] = "00"
        if timeAndMeridiem[1][0] == 'P':
            meridiem = 12
    hmsInt = [int(item) for item in hms]  #list comprehensions ftw
    hmsInt[0] += meridiem
    return hmsInt

def format(processedTime):      #takes processed time and returns formatted time
    meridiem = 'A'
    if processedTime[0] >= 12:
        if processedTime[0] > 12:
            processedTime[0] -= 12
        meridiem = 'P'
    if processedTime[0] == 0:
        processedTime[0] = 12
    return "{:0>2d}:{:0>2d}:{:0>2d} {}M".format(processedTime[0], processedTime[1], processedTime[2], meridiem)

def formatVal(processedDuration):   #takes processed duration and returns formatted duration
    return "{:0>2d}:{:0>2d}:{:0>2d}".format(processedDuration[0], processedDuration[1], processedDuration[2])

def now():          #returns formatted current time
    return time.strftime("%I:%M:%S %p")

def elapsed(timeStart, timeEnd):    #returns processed duration between two processed times
    if timeEnd[2] < timeStart[2]:
        timeEnd[2] += 60
        timeEnd[1] -= 1
    if timeEnd[1] < timeStart[1]:
        timeEnd[1] += 60
        timeEnd[0] -= 1
    if timeEnd[0] < timeStart[0]:
        timeEnd[0] += 24
    return [timeEnd[0]-timeStart[0], timeEnd[1]-timeStart[1], timeEnd[2]-timeStart[2]]


    
