'''
This program is an EcoCar PDF Parser
It will read a PDF and choose keywords mostly test cases and output the data to a an excell sheet or document
It uses a script created by another user. These sets of script are stored in the folder
    pdfminer-20140328
'''

'''
Using PDFMiner
pdf2txt.py [options] filename.pdf
Options"
    -o output file name
    -p comma-separated list of page numbers to extract
    -t output format (text/html/xml/tag[for Tagged PDFs])
    -O dirname (trigers extraction of images from PDF into directory)
    -P password
'''
#!/usr/bin/env python

import subprocess
import os
import array
import string
import xlsxwriter
import time

#Important Global Variables
#===============================================================
global pdfName, txtName, fileName       
global mainKeyWords 
global testCaseNames
global testCaseNamesIndex
global testCaseKeywords 
global testResInfoKeys
global testCaseInfoKeys 
global testCaseReqKeys 
global simulationKeys
global xclCount
global workbook

#Function Definitions
#===============================================================

def lineCleaner(i):
    #print "line = "+i
    i = i.strip('\n')
    i = i.strip(' ')
    end = len(i)-1
    start = 0
    try:
        character = i[start]
    except IndexError:
        return " "
    
    if(start >= end):
            return " "
    while ((character >= 'A' and character <= 'Z') or (character >= 'a' and character <= 'z') or (character>='0' and character<= '9'))==False:
        start+=1
        #print(start)
        character = i[start]
        if(start >= end):
            return " "

    try:
        character = i[end]
    except IndexError:
        return" "
    #print "lineCleaner = "+i[end]
    #print ord(i[end])
    while ((character >= 'A' and character <= 'Z') or (character >= 'a' and character <= 'z') or (character>='0' and character<= '9'))==False:
        end-=1
        #print(end)
        character = i[end]
    strLine = i[start:end+1]
    #print strLine
    return strLine

def pdfNameRet():
    end = False
    while(end == False):
        global pdfName
        pdfName = raw_input("What is the file path of the PDF you would like to convert?")
        end = os.path.exists("EcoCarReports/"+pdfName+".pdf")
        if end == False:
            print("File does not exist in the EcoCarReports folder. Please try again")
    return

def pdf2Txt():
    with open("EcoCarReportExcelSheets/"+pdfName+".txt","w+")as output:
        subprocess.call(["python","pdfminer-20140328/tools/pdf2txt.py","EcoCarReports/"+pdfName+".pdf"], stdout=output);
    global txtName
    txtName = pdfName
    output.close()
    return

def runThroughTxtFile():
    global txtName
    global testCaseNames
    lineCount = 0
    c = 0
    with open("EcoCarReportExcelSheets/"+txtName+".txt","r") as infile:
        for i in infile:
            i = lineCleaner(i)
            print i
    
    infile.close()
    return

def runThroughTxtFile2Line(end):
    j = 1
    k=0
    with open("EcoCarReportExcelSheets/"+txtName+".txt","r") as infile:
        for i in infile:
            i = lineCleaner(i)
            print i
            if j == end:
                break
            j+=1
    #while k < len(i):
    #    print i[k]
    #    k+=1
    value = ord(i[0])
    infile.close()
    return

def findTestCaseNames():
    global txtName
    global testCaseNames
    infile = open("EcoCarReportExcelSheets/"+txtName+".txt","r")
    lines = infile.readlines()   #error
    lineCount = 0
    while lineCount < len(lines):#error
        lines[lineCount] = lineCleaner(lines[lineCount])#error
        print lines[lineCount]       
        lineCount = lineCount + 1
    infile.close()
    return

def getTxtArray():
    global txtName
    global testCaseNames
    infile = open("EcoCarReportExcelSheets/"+txtName+".txt","r")
    lines = infile.readlines()
    lineCount = 0
    while lineCount < len(lines):
        lines[lineCount] = lines[lineCount].strip('\n')
        lines[lineCount] = lines[lineCount].strip(' ')
        lines[lineCount] = lines[lineCount].strip('')
        if lines[lineCount] == "":
            lines[lineCount] = " "
        #lines[lineCount] = lineCleaner(lines[lineCount])
        #print lines[lineCount]
        lineCount = lineCount+1
    infile.close()
    return lines

def runThroTxtArray(lines):
    i=0;
    while i < len(lines):
        print lines[i]
        i+=1
    return

def getTestCaseNames(lines):
    global mainKeyWords
    global testCaseNames
    global testCaseNamesIndex
    i = 0
    j = 0
    
    while i < len(lines):
        if (lines[i] == mainKeyWords[1]):
            i+=1
            break
        i+=1
    
    while i<len(lines):
        if(lines[i]==mainKeyWords[2]):
            i+=1
            break
        elif(lines[i]!=" "):
            testCaseNames.append(lines[i])
            testCaseNamesIndex.append(i)
            i+=1
        else:
            i+=1
    print "Printing Test CaseNames"
    while j<len(testCaseNames):
        print testCaseNames[j]+" at line "+str(testCaseNamesIndex[j])
        j+=1
    return i

def testCaseNameLocations(start, lines):
    global testCaseNamesIndex
    i = start
    j = 0

    while i<len(lines):
        #print j
        if(lines[i]==testCaseNames[j]):
            testCaseNamesIndex[j] = i
            j+=1
            i+=1
            if(j>=len(testCaseNames)):
                break
        else:
            i+=1
    print "Printing Test Case Names and Locations"
    k=0
    while k<len(testCaseNames):
        print testCaseNames[k]+" at line "+str(testCaseNamesIndex[k])
        k+=1
    return


#=====These functions are used for obtaining test case relevant information======
#this function identifies all the lines which contain keywords in a test case
#it starts at a line where 1 test case begins and continues until it reaches
#the starting point of the next test case
#testIndexes is an array that holds the line numbers of the different test cases
#index references the test case that is currently being used
def obtainTestKWLines(testIndexes, lines, index, key):
    keyWordIndex = []
    start = testIndexes[index]+1
    if(index+1>=len(testIndexes)):
        end = len(lines)
    else:
        end = testIndexes[index+1]
        
    while start<end:
        i = 0
        while i<len(key):
            if text[start] == key[i]:
                keyWordIndex.append(start)
                break;
            i+=1
        start+=1
    return keyWordIndex

#This function goes through a test case that has no additional and specific keywords
#For this to occur it just goes through the general test case
def printBasicTestCase(start, text, end):
    
    testKeys = ["Result Type:","Parent:","Start Time:","End Time:","Outcome:"]
    global xclCount
    keys = []
    results = []
    printAr = []
    lineCount = start+1
    emptyLineCount = 0
    blnRes = False;    #This boolean tells me the function reading keys or results
    i=0
    j=0
    while text[lineCount]== " ": #This loops is incase lineCount indexes an empty line
        lineCount+=1
    while len(results)<len(testKeys):
        if emptyLineCount == 0 and text[lineCount]!=" ":
            keys.append(text[lineCount])
        elif (emptyLineCount == 1 or emptyLineCount == 2) and text[lineCount]!=" ":
            results.append(text[lineCount])
        else:
            emptyLineCount+=1
        lineCount+=1
    lineCount = 0
    while lineCount<len(results):
        #print testKeys[lineCount]+" "+results[lineCount]
        printAr.append(testKeys[lineCount]+" "+results[lineCount])
        lineCount+=1
    lineCount = 0
    while lineCount<len(printAr):
        print "A"+str(xclCount)+" "+printAr[lineCount]
        worksheet.write('A'+str(xclCount),printAr[lineCount])
        lineCount+=1
        xclCount+=1
    return printAr

def obtainTestResultInfo(start, lines, end):
    global xclCount
    testKeys = ["Result Type:","Parent:","Start Time:","End Time:", "Outcome:", "Description:"]
    keys = []
    results = []
    printAr = []
    #key
    #res
    lineCount = start+1
    emptyLineCount = 0
    strTemp = ""
    while lines[lineCount]== " ": #This loops is incase lineCount indexes an empty line
        lineCount+=1
    while lineCount<end:
        if emptyLineCount == 0 and text[lineCount]!=" ":
            keys.append(text[lineCount])
        elif (emptyLineCount == 1 or emptyLineCount == 2) and text[lineCount]!=" ":
            results.append(text[lineCount])
        else:
            emptyLineCount+=1
        lineCount+=1
    lineCount = 0
    while lineCount<len(results):
        #print keys[lineCount]+" "+results[lineCount],
        key = keys[lineCount]
        res = results[lineCount]
        printAr.append(testKeys[lineCount]+" "+results[lineCount])
        strTemp += keys[lineCount]+" "+results[lineCount]
        lineCount+=1
        if lineCount>=len(keys) and lineCount<len(results):
            #print results[lineCount]
            res += res+results[lineCount]
            strTemp+= results[lineCount]
            printAr[lineCount-1] = printAr[lineCount-1]+results[lineCount]
            lineCount+=1
        #else:
            #print ""
    lineCount = 0
    while lineCount<len(printAr):
        print "A"+str(xclCount)+" "+printAr[lineCount]
        worksheet.write('A'+str(xclCount),printAr[lineCount])
        xclCount+=1
        lineCount+=1
    return printAr


#The data for 'Test Case Information' can appear before or after where
#Test Case Information appears in the test case line or after
#atm this function will be unused and the results for testCase Info will be hard coded until
#a more consistent pdf reader will be used
def obtainTestCaseInfo(start, lines, end):
    testKeys = ["Name:","Type:"]
    return

#The data for 'Test Case Information' can appear before or after where
#Test Case Information appears in the test case line or after
#atm this function will be unused and the results for testCase Info will be hard coded until
#a more consistent pdf reader will be used
def obtainTestSuiteInfo(start, lines, end):
    testKeys = ["Name:"]
    return


def obtainTestCaseRes(start, lines, end):
    global xclCount
    testKeys = ["Description:","Document:"]
    keys = []
    results = []
    lineCount = start+1
    emptyLineCount = 0
    while lines[lineCount]== " ": #This loops is incase lineCount indexes an empty line
        lineCount+=1
    while lineCount<end:
        if emptyLineCount == 0 and text[lineCount]!=" ":
            keys.append(text[lineCount])
        elif (emptyLineCount == 1) and text[lineCount]!=" ":
            results.append(text[lineCount])
        else:
            emptyLineCount+=1
        lineCount+=1
    lineCount = 0
    while lineCount<len(results):
        #print keys[lineCount]+" "+results[lineCount]
        printAr.append(keys[lineCount]+" "+results[lineCount])
        lineCount+=1
    lineCount = 0
    while lineCount<len(printAr):
        print "A"+str(xclCount)+" "+printAr[lineCount]
        worksheet.write('A'+str(xclCount),printAr[lineCount])
        xclCount+=1
        lineCount+=1
    return printAr

#=====These functions are used for obtaining Simulation relevant information=====
def obtainSimKWLines(testIndexes, lines, index, key):
    keyWordIndex = []
    start = testIndexes[index]+1
    if(index+1>=len(testIndexes)):
        end = len(lines)
    else:
        end = testIndexes[index+1]
    while start<end:
        i = 0
        while i<len(key):
            if text[start]==key[i]:
                keyWordIndex.append(start)
                break;
            i+1
        start+=1
    return keyWordIndex
def obtainSysUndTestInfo(start, lines, end):
    global xclCount
    testKeys = ["Model:","Harness:","Harness Owner:", "Simulation Mode","Configuration Set:","Start Time:","Stop Time:", "CheckSum:"]
    keys = []
    results = []
    printAr = []
    lineCount = start+1
    emptyLineCount = 0
    while lines[lineCount]== " ": #This loops is incase lineCount indexes an empty line
        lineCount+=1
    while lineCount<end:
        if emptyLineCount == 0 and text[lineCount]!=" ":
            keys.append(text[lineCount])
        elif (emptyLineCount == 1) and text[lineCount]!=" ":
            results.append(text[lineCount])
        else:
            emptyLineCount+=1
        lineCount+=1
    lineCount = 0
    while lineCount<len(results):
        #print keys[lineCount]+" "+results[lineCount]
        printAr.append(keys[lineCount]+" "+results[lineCount])
        lineCount+=1
        
    lineCount = 0
    while lineCount<len(printAr):
        print "A"+str(xclCount)+" "+printAr[lineCount]
        worksheet.write('A'+str(xclCount),printAr[lineCount])
        xclCount+=1
        lineCount+=1
    return printAr

def print2Excel(toExcelAr):
    return
#=======================================Main program====================================

#with open("EcoCarReportExcelSheets/Test1.txt","w+") as output:
#    subprocess.call(["python","pdfminer-20140328/tools/pdf2txt.py", "EcoCarReports/newReport.pdf"], stdout=output);
global mainKeyWords, testCaseNames, testCaseKeywords, testResInfoKEys, testCaseInfoKeys,testCaseReqKeys, simulationKeys
xclCount = 0
mainKeyWords = ['Summary','Name','Outcome','Duration','(Seconds)']
testCaseNames = []
testCaseNamesIndex = []
testKeywords = ["Test Result Information","Test Suite Information","Test Case Information","Test Case Requirements"]
testKWIndex = []
simulationKeywords = ["Simulation","System Under Test Information","Simulation Logs"]
simKWIndex = []
text = []# an array that holds all the lines of the pdf
printAr = []
i = 1

print "==============================pdfNameRet() begining=============================="
pdfNameRet()
print "==============================pdfNameRet() begining=============================="

print "==============================pdf2Txt() begining================================="
pdf2Txt()
print "==============================pdf2Txt() complete================================="

print "==============================getTxtArray() begining============================="
text = getTxtArray()
print "==============================getTxtArray() complete============================="

print "==============================runThroTxtArray() begining========================="
runThroTxtArray(text)
print "==============================runThroTxtArray() complete========================="

print "================================getTestCaseNames() begining======================"
end = getTestCaseNames(text)
print "================================getTestCaseNames() complete======================"

print "==============================testCaseNameLocations() begining==================="
testCaseNameLocations(end, text)
print "==============================testCaseNameLocations() complete==================="

print "=================Printing All Test Cases to Shell Beginning======================"
testCount = 0
workbook = xlsxwriter.Workbook("EcoCarReportExcelSheets/"+txtName+'.xlsx')
worksheet = workbook.add_worksheet()
while testCount < len(testCaseNames):
    testKWIndex = obtainTestKWLines(testCaseNamesIndex, text, testCount, testKeywords)
    i = 0
    xclCount+=1        
    print "\nA"+str(xclCount)+" "+testCaseNames[testCount]
    worksheet.write('A'+str(xclCount),testCaseNames[testCount])
    xclCount+=1
    printAr.append(testCaseNames[testCount])    
    #This prints out the keys and indexs at them
    if (len(testKWIndex) == 0):
            printAr.append(printBasicTestCase(testCaseNamesIndex[testCount],text,end))           
    while i<len(testKWIndex):
        printAr = []
        try:
            end = testKWIndex[i+1]
        except IndexError:
            try:
                end = testCaseNamesIndex[testCount+1]
            except IndexError:
                end = len(text)
                   
                
        #print text[testKWIndex[i]]+" at index "+str(testKWIndex[i])
        if (text[testKWIndex[i]] == "Test Result Information"):
            print "A"+str(xclCount)+" "+"Test Result Information"
            worksheet.write('A'+str(xclCount),'Test Result Information')
            xclCount+=1
            printAr.append("Test Results Information")
            printAr.append(obtainTestResultInfo(testKWIndex[i],text,end))
            
        elif (text[testKWIndex[i]] == "Test Suite Information"):
            print "A"+str(xclCount)+" "+"Test Suite Information"
            worksheet.write('A'+str(xclCount),'Test Suite Information')
            xclCount+=1
            print "A"+str(xclCount)+" "+"Name: "+testCaseNames[testCount]
            worksheet.write('A'+str(xclCount),'Name: '+testCaseNames[testCount])
            xclCount+=1
            printAr.append("Test Suite Information")
            printAr.append("Name: "+testCaseNames[testCount])
            
        elif (text[testKWIndex[i]] == "Test Case Information"):
            print "A"+str(xclCount)+" "+"Test Case Information"
            worksheet.write('A'+str(xclCount),'Test Case Information')
            xclCount+=1
            print "A"+str(xclCount)+" "+"Name: "+testCaseNames[testCount]
            worksheet.write('A'+str(xclCount),'Name: '+testCaseNames[testCount])
            xclCount+=1
            printAr.append("Test Case Information")
            printAr.append("Name: "+testCaseNames[testCount])
            
        else:
            #print "A"+str(xclCount)+" "+"Test Case Requirements"
            #xclCount+=1
            printAr.append("Test Case Requirements")
            printAr.append(obtainTestCaseRes(testKWIndex[i],text,end))
            
        i+=1
    testSWIndex = obtainTestKWLines(testCaseNamesIndex, text, testCount, simulationKeywords)
    i=0
    while i<len(testSWIndex):#This for loop executes Simulation Test Cases
        if text[testSWIndex[i]] == "Simulation":
            print "A"+str(xclCount)+" "+"Simulation"
            worksheet.write('A'+str(xclCount),'Simulation')
            xclCount+=1
            printAr.append("Simulation")
        elif text[testSWIndex[i]] == "System Under Test Information":
            print "A"+str(xclCount)+" "+"System Under Test Information"
            worksheet.write('A'+str(xclCount),'System Under Test Information')
            xclCount+=1
            printAr.append("Simulation")
            obtainSysUndTestInfo(testSWIndex[i],text,end)
        elif text[testSWIndex[i]] == "Simulation Logs":
            print "A"+str(xclCount)+" "+"Simulation Logs"
            worksheet.write('A'+str(xclCount),'Simulation Logs')
            xclCount+=1
            printAr.append("Simulation Logs")
        else:
            print "\nSimulation data not found"
        i+=1
    testCount+=1
    #print "=================Printing All Test Cases to Shell Complete======================="
    #k = 0
    #l = 0
    #while k<len(printAr):
    #    l=0
    #    #print "TYPE: "+str(type(printAr[k]))
    #    if(type(printAr[k])==type('str')):
    #        print printAr[k]
    #    else:
    #        while l<len(printAr[k]):
    #            print printAr[k][l]
    #            l+=1
    #    k+=1
    #print "=================Printing printAr begi==========================================="
workbook.close()


#print "runThroughTxtFile2Line() begining"
#runThroughTxtFile2Line(i)
#print "runThroughTxtFile2Line() complete"

#print "findFirstTest() begining"
#findFirstTest()
#print "findFirstTest() complete"

#lineCleaner("   asdfasfa               ")
#==========================================

#(character >= 'A' and character <= 'Z') or (character >= 'a' and character <= 'z') or (character>='1' and character<= '9')

    
'''
pdfFilePath = raw_input("What is the file path of the PDF you would like to read?")
exclFilePath = raw_input("What is the file path of the Excel document you would like to write to?")
while (end == 'N'):
    'end = raw_input("Would you like to exit 'Y/N'")
'''

'''
Unused but Helpful code
-----------------------------------
Standard File reading
infile = open("EcoCarReportExcelSheets/Test1.txt","r")

contents = infile.read()
print contents
infile.close()
----------------------------
Correct PDF reader code
import subprocess
with open("pdfminer-20140328/samples/whooo3.txt", "w+") as output:
    subprocess.call(["python","pdfminer-20140328/tools/pdf2txt.py", "pdfminer-20140328/samples/simple1.pdf"], stdout=output);
'''
