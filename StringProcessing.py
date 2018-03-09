"""
Created on 03 Mar 2018 7:37 AM

@author: muhin
"""

from cryptography.fernet import Fernet


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
    listDate = inputDate.split('-')
    monthNameIndex = int(listDate[1]) - 1
    return monthNames[monthNameIndex] + ", " + str(listDate[0])


def fileReName(inputDate):
    monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
                  'November', 'December']
    listDate = inputDate.split('-')
    monthNameIndex = int(listDate[1]) - 1


def processEncryptedPassword():
    f = open('credential.txt', 'r')
    msg = f.read()
    msgSplit = msg.split(',')
    bytePassword = bytes(msgSplit[2], 'utf-8')
    fernetObj = Fernet("hPh9eN7EXljKDTJ4u4KDURBUQ4IqE42f3IBjwDdzb08=")
    decryptedPassword = fernetObj.decrypt(bytePassword)
    stringDecryptedPassword = str(decryptedPassword, 'utf-8')
    f.close()
    return stringDecryptedPassword


def returnEndDate(startDate):
    startDateAsList = startDate.split('-')
    intNextMonth = int(startDateAsList[1]) + 1
    if intNextMonth == 13:
        intNextYear = int(startDateAsList[0]) + 1
        stringNextYear = str(intNextYear)
        stringNextMonth = str(stringNextYear + '-' + '01-01')
        return stringNextMonth
    else:
        nextMonthFormat = '{:02d}'.format(intNextMonth)
        stringNextMonth = str(startDateAsList[0] + '-' + nextMonthFormat + '-' + startDateAsList[2])
        return stringNextMonth
