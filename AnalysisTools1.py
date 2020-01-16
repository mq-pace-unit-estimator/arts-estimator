# -*- coding: utf-8 -*-
"""
Created on Fri Sep 13 16:08:02 2019

@author: Thomas Demkowski
"""

import xlrd

file_location = "Insert path for: Copy of 2019 AH Enrolment Numbers.xlsx"

workbook = xlrd.open_workbook(file_location)


typeDictionary = { # for selecting column of interest
        'Faculty': 1,
        'Department': 3,
        'Unit': 4,
        'Unit Title': 5,
        'Attribute': 7,
        'Study Period': 9,
        'Avail Desc': 10
}


def studentsEnrolledType(title,name,date):
    sheetObj = workbook.sheet_by_name(date)
    maxRow = 1
    while sheetObj.cell_value(maxRow,4) != '':
        maxRow += 1
    row = 1
    column = typeDictionary[title]
    studentCount = 0
    while row <= maxRow:
        if sheetObj.cell_value(row,column) == name:
            studentCount += sheetObj.cell_value(row,11)
        row += 1
    return studentCount
    
def studentsEnrolledUnit(unit,date):
    return studentsEnrolledType('Unit',unit,date)


def studentsEnrolledFaculty(faculty,date):
    return studentsEnrolledType('Faculty',faculty,date)


def studentsEnrolledDepartment(department,date):
    return studentsEnrolledType('Department',department,date)


def collectNames(title,collection,date): # example: ('Unit',units)
    sheetObj = workbook.sheet_by_name(date)
    maxRow = 1
    while sheetObj.cell_value(maxRow,4) != '':
        maxRow += 1
    maxRow -= 1
    column = typeDictionary[title]
    row = 1
    while row <= maxRow:
        collection.add(sheetObj.cell_value(row,column))
        row+= 1
            
def collectAllNames(title,collection):
    sheet_names = workbook.sheet_names()
    for x in sheet_names:
        collectNames(title,collection,x) 

units = set([])
collectNames('Unit',units,'15 Jan 2019')
print(units)


def getUnitsByLevel(level,date): # example: ('300','15 Jan 2019')
    setResult = set([])
    setObj = set([])
    collectNames('Unit',setObj,date)
    for x in setObj: # 3rd last character always represents unit level
        if(x[len(x) - 3] == level[0]):
            setResult.add(x)
    return setResult


def getUnitsByCodename(codename,date):
    setResult = set([])
    setObj = set([])
    collectNames('Unit',setObj,date)
    for x in setObj:
        if(codename in x):
            setResult.add(x)
    return setResult


def getUnitsByLevelAndCodename(level,codename,date):
    A = getUnitsByLevel(level,date)
    B = getUnitsByCodename(codename,date)
    return A.intersection(B)


def getUnitsByFaculty(faculty,date):
    setResult = set([])
    sheetObj = workbook.sheet_by_name(date)
    maxRow = 1
    while sheetObj.cell_value(maxRow,4) != '':
        maxRow += 1
    maxRow -= 1
    row = 1
    while row <= maxRow:
        if sheetObj.cell_value(row,1) == faculty:
            setResult.add(sheetObj.cell_value(row,4))
        row+= 1
    return setResult


def totalUnitsByLevel(level,date):
    return len(getUnitsByLevel(level,date))


def totalUnitsByCodename(codename,date):
    return len(getUnitsByCodename(codename,date))


def totalUnitsByLevelAndCodename(level,codename,date):
    return len(getUnitsByLevelAndCodename(level,codename,date))


def totalUnitsByFaculty(faculty,date):
    return len(getUnitsByFaculty(faculty,date))
    
# The following derives from AnalysisTools2.py
AH1 = {'AHIS108', 'AHIS140', 'AHIS110', 'EXAH130', 'AHIS118', 'EXAH100', 'AHIS191', 'AHIS120', 'AHIX150', 'AHIS170', 'AHIS100', 'AHIS168', 'AHMG101', 'AHIS158', 'AHIX118', 'AHIX108', 'AHIS150', 'AHIS178', 'AHIX110', 'AHIS190'}
AH2 = {'EXAH230', 'AHIS204', 'EXAH200', 'AHIX202', 'EXAH231', 'AHIS230', 'AHIS280', 'AHIS253', 'AHIS205', 'AHIS261', 'AHIS291', 'AHIS209', 'AHIS272', 'EXAH203', 'AHIS219', 'AHIX220', 'EXAH210', 'AHIS255', 'AHIS220', 'AHIS202', 'AHIS200', 'EXAH232', 'EXAH216', 'AHIS290', 'AHIX250', 'AHIS279', 'AHIS250'}
AH3 = {'AHIS389', 'AHIS377', 'AHIS371', 'AHIS319', 'AHIS354', 'AHIS312', 'EXAH316', 'AHIS380', 'AHIS331', 'AHIS357', 'AHIS372', 'AHIS391', 'EXAH332', 'AHIS313', 'EXAH310', 'AHIS370', 'AHIS335', 'AHIS393', 'AHIS392', 'EXAH330', 'AHIS344', 'AHIS394', 'AHIS341', 'AHIS349', 'AHIS343', 'AHIS358', 'AHIS301', 'AHIS339', 'AHIS342', 'AHIS368', 'AHIS318', 'EXAH311', 'AHIS345', 'AHIX335', 'AHIS399', 'AHIS309', 'AHIS356', 'AHIS350', 'EXAH331', 'AHIS308', 'AHIX342', 'AHIS305', 'AHIS346'}


def sortBySemester(collection,date,semesterCode):
    tmp = set({})
    for x in collection:
        sheetObj = workbook.sheet_by_name(date)
        row = 1
        while sheetObj.cell_value(row,4) != x:
            row += 1
            # if a unit isnt there then there will be an exception
        if sheetObj.cell_value(row,9) == semesterCode:
            tmp.add(x)
    return tmp