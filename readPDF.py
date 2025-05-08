#!/usr/bin/python
import pypdf
from openpyxl import load_workbook

def readManifest(in_file):
    print(f"[o] Reading file: \t {in_file}")
    reader = pypdf.PdfReader(in_file)
    n_pages = len(reader.pages)
    print(f"[o] Number of pages: ",n_pages)

    for i in range(n_pages):
        text = reader.pages[i].extract_text().split("\n")
        parseText(text)

def parseText(text_dump):
  # Parse text_dump
    for i in range(len(text_dump)):
        c1 = len(text_dump[i]) == 13
        c2 = not("To /" in text_dump[i])
        c3 = not("cartons" in text_dump[i])
        if (c1 & c2 & c3):
            container = text_dump[i]
            size = text_dump[i+1]
            weight = text_dump[i+5]
            unit = (container,size,weight)
            print(unit)


def populateExcel(in_file):
    print(f"[o] Loading File: \t {in_file}")
    wb = load_workbook(in_file)

    worksheet3 = wb["Blad3"]
#   worksheet3['I10'] = 20.000
    print(worksheet3['I10'].value)


    wb.save(in_file)



"""
===============================================================================
+ FUNCTION TESTS
===============================================================================
"""
def testPopulateExcel():
    input_file = "./stabiliteit test.xlsx"
    populateExcel(input_file)


def testReadManifest():
    input_file = "./Voyage manifest incoming 9164DEL-I.pdf"
    readManifest(input_file)



if __name__ == '__main__':
    testReadManifest()
    testPopulateExcel()
