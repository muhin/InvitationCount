"""
Created on 03 Mar 2018 7:37 AM

@author: muhin
"""


def processIntervalName(dbIntervalName):
    monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
                  'November', 'December']
    intervalNameAsList = dbIntervalName.split(' ')
    monthName = intervalNameAsList[0]
    if monthName in monthNames:
        finalIntervalName = intervalNameAsList[0] + ', ' + intervalNameAsList[1]
        return finalIntervalName
    else:
        intervalNameLength = len(dbIntervalName)
        if intervalNameLength == 11:
            finalIntervalName = dbIntervalName[0:6] + ', ' + dbIntervalName[7:]
            return finalIntervalName
        elif intervalNameLength == 12:
            finalIntervalName = dbIntervalName[0:7] + ', ' + dbIntervalName[8:]
            return finalIntervalName
        elif intervalNameLength >= 5 or intervalNameLength <= 7:
            return dbIntervalName


def adhocIntervalName(inputDate):
    monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
                  'November', 'December']
    listDate = inputDate.split("-")
    monthNameIndex = int(listDate[1]) - 1
    return monthNames[monthNameIndex] + ", " + str(listDate[0])


def fileReName(inputDate):
    monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
                  'November', 'December']
    listDate = inputDate.split("-")
    monthNameIndex = int(listDate[1]) - 1
    return monthNames[monthNameIndex] + ' ' + str(listDate[0])