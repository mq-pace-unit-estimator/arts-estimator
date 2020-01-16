# -*- coding: utf-8 -*-
"""
Created on Thu Sep 19 17:04:25 2019

@author: Thomas Demkowski
"""

import xlrd

file_location_all_units         =   "Insert path for: all_ug_units.xlsx"
file_location_enrolment_numbers =   "Insert path for: AH Enrolment numbers 2011 to 2019 Pivot Table.xlsx"
file_location_unit_analysis     =   "Insert path for: FOA AHIS unit analysis V2.xlsx"

workbook_all_units              =   xlrd.open_workbook(file_location_all_units)
workbook_enrolment_numbers      =   xlrd.open_workbook(file_location_enrolment_numbers)
workbook_unit_analysis          =   xlrd.open_workbook(file_location_unit_analysis)

sheet_all_units                 =   workbook_all_units.sheet_by_index(0)
sheet_enrolment_numbers         =   workbook_enrolment_numbers.sheet_by_index(0)
sheet_unit_analysis             =   workbook_unit_analysis.sheet_by_index(1)


def collectAllUnits():
    A = []
    s = ""
    for m in range(1,sheet_all_units.nrows):
        s = sheet_all_units.cell_value(m,0) + sheet_all_units.cell_value(m,1)
        A.append(s)
        s = ""
    return A

units = collectAllUnits()

def getPrerequisites(unitCode):
    try:
        row = units.index(unitCode) + 1
    except:
        return "'" + unitCode + "'" + " was not found in the units set"
    return sheet_all_units.cell_value(row,4)


def getPrerequisiteUnits(unitCode): # return a list of units associated
    prerequisites = set()
    for x in units: # O(n) & units list is large
        if (x in getPrerequisites(unitCode)):
            prerequisites.add(x)
    return prerequisites


years = {
    "Grand Total": 11,"2019":10, "2018":9, "2017":8, "2016":7, "2015":6, "2014":5, "2013":4, "2012":3, "2011":2
}

def getTotalStudents(unitCode,year): #example ("MATH336","2019")
    for m in range(2,sheet_enrolment_numbers.nrows):
        if (unitCode in sheet_enrolment_numbers.cell_value(m,0)):
            return sheet_enrolment_numbers.cell_value(m,years[year])
            break
    return 0


AHunits = {'EXAH230', 'AHIS158', 'AHIS308', 'EXAH316', 'AHIS190', 'AHIX342', 'EXAH232', 'AHIX110', 'AHIS602', 'AHIS178', 'AHIS357', 'AHPG830', 'AHIS205', 'ECJS853', 'AHIS710', 'AHIS250', 'AHIS230', 'AHIS399', 'AHIS342', 'AHIS372', 'AHMG101', 'AHIS220', 'AHIS356', 'AHIS371', 'AHIS705', 'AHIX250', 'EXAH231', 'AHIX202', 'EXAH310', 'AHIS331', 'AHIS350', 'AHIS709', 'AHIS219', 'AHIS280', 'AHPG855', 'AHIS704', 'AHIS305', 'AHPG820', 'AHIS389', 'AHPG815', 'AHIS346', 'AHPG884', 'AHIS319', 'AHIS394', 'AHIS209', 'AHIS110', 'AHPG880', 'AHIS335', 'EXAH331', 'AHIS345', 'AHIS706', 'AHIS204', 'AHIS339', 'AHPG811', 'EXAH100', 'AHIS202', 'AHIS601', 'EXAH330', 'EXAH332', 'AHPG821', 'AHIX220', 'AHIS341', 'AHIX108', 'AHIS100', 'AHPG872', 'AHIS701', 'AHIS301', 'AHIS391', 'AHIS349', 'AHIS393', 'EXAH311', 'AHIS309', 'AHIS150', 'AHIS600', 'AHIS358', 'AHIS368', 'EXAH203', 'AHIS291', 'AHIS279', 'AHIS370', 'AHIX118', 'AHIS377', 'EXAH130', 'AHIS313', 'AHIS200', 'AHIS108', 'AHIS120', 'AHPG816', 'AHIS170', 'AHPG858', 'AHIX150', 'AHIS354', 'AHIS380', 'AHIS708', 'AHIS253', 'AHIS703', 'INTS600', 'AHIX335', 'AHIS191', 'AHIS255', 'AHIS168', 'AHIS118', 'AHIS344', 'AHIS290', 'AHIS261', 'AHIS312', 'EXAH216', 'EXAH210', 'AHPG824', 'AHIS702', 'AHIS343', 'EXAH200', 'AHPG883', 'AHIS700', 'AHIS392', 'AHIS272', 'AHIS707', 'AHIS318', 'AHIS140'}


def getLevelUnits(n,unitSet):
    tempSet = set([])
    for x in unitSet:
        if x[-3] == str(n):
            tempSet.add(x)
    return tempSet


def getCompleted(unitCode):
    for m in range(1,sheet_unit_analysis.nrows):
        if (unitCode == sheet_unit_analysis.cell_value(m,6)):
            if sheet_unit_analysis.cell_value(m,8) == 'Completed':
                return int(sheet_unit_analysis.cell_value(m,10))
    return -1


def getFailed(unitCode):
    for m in range(1,sheet_unit_analysis.nrows):
        if (unitCode == sheet_unit_analysis.cell_value(m,6)):
            if sheet_unit_analysis.cell_value(m,8) == 'Failed':
                return int(sheet_unit_analysis.cell_value(m,10))
    return -1


def getWWP(unitCode):
    for m in range(1,sheet_unit_analysis.nrows):
        if (unitCode == sheet_unit_analysis.cell_value(m,6)):
            if sheet_unit_analysis.cell_value(m,8) == 'Withdraw without penalty':
                return int(sheet_unit_analysis.cell_value(m,10))
    return -1


def getPercentagePassed(unitCode):
    numerator = getCompleted(unitCode)
    denominator = getCompleted(unitCode) + getFailed(unitCode) + getWWP(unitCode)
    ans = (int)(100*numerator/denominator)
    return ans
    
# collecting every Ancient History unit by year
AH1 = getLevelUnits(1,AHunits)
AH2 = getLevelUnits(2,AHunits)
AH3 = getLevelUnits(3,AHunits)
AHRelevantUnits = AH1.union(AH2).union(AH3)


def getCorrequisites(unitCode):
    try:
        row = units.index(unitCode) + 1
    except:
        return "'" + unitCode + "'" + " was not found in the units set"
    return sheet_all_units.cell_value(row,5)

    
def prerequisiteCount(collection):
    count = 0
    for x in collection:
        if getPrerequisites(x) != "":
            count += 1
    return count

    
def getPercentage(collection):
    count = 0
    for x in collection:
        if getCompleted(x) == -1 or getFailed(x) == -1 or getWWP(x) == -1:
            count += 1
    return 100*(count / len(collection))