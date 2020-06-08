# University of Michigan CAEN EACS Laptop Inventory #
# Developed by Arthur Su #
# Last edited: 6/7/2020 #

from __future__ import print_function
import pyzbar.pyzbar as pyzbar
import numpy as np
import cv2
import xlsxwriter
import re
import argparse
from pathlib import Path
from tkinter import filedialog
from tkinter import *
import os
import os.path

### EDITABLE DEFAULT VALUES ###
defOutputFilename = "inventory"
defNamePrefix = "EA"

defDepartment = "EACS"
defUnit = "EACS - Engin Admin Computer Support"
defStatus = "Onsite"
defMachineType = "Laptop"
defMachineDescription = "HP ELITEBOOK 840 G6"
defOSType = "WINDOWS"
defOSVersion = "10"
defOwner = "EACS"
defPrimaryUser = "EACS"
defRoom = "G255"
defBuilding = "LEC"
defPurchaseDate = "*FILL ME OUT*"

# Creates GUI window
window = Tk()
window.title("Inventory GUI")
window.geometry('400x440')
window.resizable(False, False)

# Creates some vars
allLines = []
fields = []
dropdowns = []

serialNumberList = []
nameList = []
macAddressList = []

machineDescription = ""
osType = ""
osVersion = ""
owner = ""
primaryUser = ""
room = ""
building = ""
department = ""
purchased = ""
unit = ""
machineType = ""
status = ""
increment = 0

# Clears some vars
folderPath = ""
outputName = ""
namePrefix = ""

# Group 1 - Select folder location
group1 = LabelFrame(window, text = "Main", padx = 1, pady = 1)
group1.grid(padx = 1, pady = 1)

labelLocalFolder = Label(group1, text = "Image Folder Location")
labelLocalFolder.grid(column = 0, row = 1)

folderLocation = Entry(group1, width = 60)
folderLocation.grid(column = 0, row = 2)

# Open file browser upon button press
def selectFolderPath():
    # Open file selecter and making path variable
    folderPath = filedialog.askdirectory()
    folderLocation.insert(END, folderPath)

# Button to open file browser
selectFolder = Button(group1, text = ". . .", command = selectFolderPath)
selectFolder.grid(column = 2, row = 2)

labelLocalFolder.grid()
folderLocation.grid()
selectFolder.grid()

# Group 2 - Adjust parameters
group2 = LabelFrame(window, text = "General", padx = 1, pady = 1)
group2.grid(padx = 0, pady = 0)

# Output filename entry field
label3 = Label(group2, text = "Output Filename (w/o .xlsx)")
label3.grid(column = 0,row = 5)
outputNameGUI = Entry(group2, width = 20)
outputNameGUI.insert(0, defOutputFilename)
outputNameGUI.grid(column = 4, row = 5)

# Laptop name prefix entry field
label4 = Label(group2, text = "Laptop Name Prefix (w/o dash, serial# suffix)")
label4.grid(column = 0, row = 7)
namePrefixGUI = Entry(group2, width = 20)
namePrefixGUI.insert(0, defNamePrefix)
namePrefixGUI.grid(column = 4, row = 7)

# Group 3 - Required information
group3 = LabelFrame(window, text = "Required info - check website for formatting", padx = 1, pady = 1)
group3.grid(padx = 0, pady = 0)

# Info to be filled out
label5 = Label(group3, text = "Department")
label5.grid(column = 0, row = 10)
departmentGUI = Entry(group3, width = 20)
departmentGUI.insert(0, defDepartment)
departmentGUI.grid(column = 4, row = 10)

label6 = Label(group3, text = "Unit")
label6.grid(column = 0, row = 12)
unitGUI = Entry(group3, width = 20)
unitGUI.insert(0, defUnit)
unitGUI.grid(column = 4, row = 12)

label7 = Label(group3, text = "Status")
label7.grid(column = 0, row = 14)
statusGUI = Entry(group3, width = 20)
statusGUI.insert(0, defStatus)
statusGUI.grid(column = 4, row = 14)

label8 = Label(group3, text = "Machine Type")
label8.grid(column = 0, row = 16)
machineTypeGUI = Entry(group3, width = 20)
machineTypeGUI.insert(0, defMachineType)
machineTypeGUI.grid(column = 4, row = 16)

label9 = Label(group3, text = "Make/Model")
label9.grid(column = 0, row = 18)
machineDescriptionGUI = Entry(group3, width = 20)
machineDescriptionGUI.insert(0, defMachineDescription)
machineDescriptionGUI.grid(column = 4, row = 18)

label10 = Label(group3, text = "OS Type")
label10.grid(column = 0, row = 20)
osTypeGUI = Entry(group3, width = 20)
osTypeGUI.insert(0, defOSType)
osTypeGUI.grid(column = 4, row = 20)

label11 = Label(group3, text = "OS Version")
label11.grid(column = 0, row = 22)
osVersionGUI = Entry(group3, width = 20)
osVersionGUI.insert(0, defOSVersion)
osVersionGUI.grid(column = 4, row = 22)

label12 = Label(group3, text = "Owner")
label12.grid(column = 0, row = 24)
ownerGUI = Entry(group3, width = 20)
ownerGUI.insert(0, defOwner)
ownerGUI.grid(column = 4, row = 24)

label13 = Label(group3, text = "Primary User")
label13.grid(column = 0, row = 26)
primaryUserGUI = Entry(group3, width = 20)
primaryUserGUI.insert(0, defPrimaryUser)
primaryUserGUI.grid(column = 4, row = 26)

label14 = Label(group3, text = "Room")
label14.grid(column = 0, row = 28)
roomGUI = Entry(group3, width = 20)
roomGUI.insert(0, defRoom)
roomGUI.grid(column = 4, row = 28)

label15 = Label(group3, text = "Building")
label15.grid(column = 0, row = 30)
buildingGUI = Entry(group3, width = 20)
buildingGUI.insert(0, defBuilding)
buildingGUI.grid(column = 4, row = 30)

label16 = Label(group3, text = "Date Purchased (**/**/****)")
label16.grid(column = 0, row = 32)
purchasedGUI = Entry(group3, width = 20)
purchasedGUI.insert(0, defPurchaseDate)
purchasedGUI.grid(column = 4, row = 32)

# Scan for barcodes and values
def decode(scan) : 
    decodedObjects = pyzbar.decode(scan)
    return decodedObjects

# Gets text input fields for inventory website
def getFields():
    fieldIn = open("fields.txt")
    fields = fieldIn.readlines()
    fields = [x.strip() for x in fields]
    return(fields)

# Gets dropdown menus for inventory website
def getDropdowns():
    dropdownIn = open("dropdowns.txt")
    dropdowns = dropdownIn.readlines()
    dropdowns = [y.strip() for y in dropdowns]
    return(dropdowns)
    
# Creates JS script for text input fields
def scriptTextFields(allLines, fieldArray, outputName, namePrefix, increment):
    # Creates array of info values
    textFieldValues = [serialNumberList[increment], namePrefix + "-" +
        serialNumberList[increment], machineDescriptionGUI.get(), osTypeGUI.get(),
        osVersionGUI.get(), ownerGUI.get(), primaryUserGUI.get(), roomGUI.get(),
        buildingGUI.get(), purchasedGUI.get(), macAddressList[increment]]

    # Increments through text fields, adds respective commands to JS script
    for obj in range(0, len(fieldArray)):
        allLines.append("document.getElementById('" + fieldArray[obj] + "').value = '" + textFieldValues[obj] + "';\n")
                    
# Creates JS script for dropdown menus
def scriptDropdowns(allLines, dropdownArray):
    # Creates array of info values
    dropdownMenuValues = [departmentGUI.get(), unitGUI.get(), 
        machineTypeGUI.get(), statusGUI.get()]
    
    # Increments through dropdown menus, adds respective commands to JS script
    for obj in range(0, len(dropdownArray)):
        if(dropdownArray[obj] != "unit"):
            allLines.append('function setSelectedIndex(s, ' + dropdownMenuValues[obj] + ') {')
            allLines.append('for (var i = 0; i < s.options.length; i++) {')
            allLines.append('if (s.options[i].text == ' + dropdownMenuValues[obj] + ') {')
            allLines.append('s.options[i].selected = true;return;}}}')
            allLines.append('setSelectedIndex(document.getElementById("' + dropdownArray[obj] + '"), "' + dropdownMenuValues[obj] + '");\n')
        else:
            # Units are tricky because of the "-" in the name
            allLines.append('document.getElementById("' + dropdownArray[obj] + '").value = "' + unitGUI.get() + '";\n')
            
# Calls scriptTextFields and scriptDropdowns to write the JS to outputName_script.js
def createScript(folderPath, allLines, getDropdowns, outputName, namePrefix, department, unit, machineType, status):
    # Creates Excel spreadsheet for script commands
    workbook = xlsxwriter.Workbook(outputName + "_script.xlsx")
    worksheet2 = workbook.add_worksheet()
    worksheet2.write_string(0, 0, "Number")
    worksheet2.write_string(0, 1, "Script Command")
    worksheet2.set_column('A:A', 7)
    
    of = outputName + "_script.txt"
    
    increment = 0
    fileCount = 0
    imgs = Path(folderPath)
    for img in imgs.iterdir():
        fileCount += 1
    print(str(fileCount) + " images processed")
    
    outfile = open(of, 'w')
    # Increments through the number of images in folder
    while(increment < fileCount):
        scriptTextFields(allLines, getFields(), outputName, namePrefix, increment)
        scriptDropdowns(allLines, getDropdowns())
        
        # Saves inventory form
        allLines.append("document.getElementById('inventoryForm-save').click();\n")
        
        # Waits two seconds, then clones the inventory form
        allLines.append("setTimeout(() => {document.getElementsByClassName('btn-warning')[0].click();}, 2000);")
        allLines.append('\n\n')
        outfile.writelines(allLines)
        
        # Writes JS script to Excel spreadsheet
        worksheet2.write_string(increment + 1, 1, "".join(allLines))
        worksheet2.write_number(increment + 1, 0, increment)
        
        increment += 1
        allLines = []
    workbook.close()
    outfile.close()        

# Action after clicking start button
def execute():
    inventory(folderLocation.get(), outputNameGUI.get(), namePrefixGUI.get())

# Inventories laptops 
def inventory(folderPath, outputName, namePrefix):
    
    # Clears vars
    barcodeNum = 0
    imageNum = 0
    nameTrigger = 0
    
    # Creates Excel spreadsheet for laptop information
    workbook = xlsxwriter.Workbook(outputName + ".xlsx")
    worksheet1 = workbook.add_worksheet()
    worksheet1.write_string(0, 0, "Serial#")
    worksheet1.write_string(0, 1, "LAN MAC")
    if namePrefix != "":
        worksheet1.write_string(0, 2, "Name")
        nameTrigger = 1
        worksheet1.set_column('C:C', 18)
    worksheet1.set_column('A:B', 18)
        
    # Iterates through images in folder
    images = Path(folderPath)
    for image in images.iterdir():
        print(image.name)
        imageNum += 1
        
        # Scans and decodes barcodes in image
        scan = cv2.imread(folderPath + "/" + image.name)
        decodedObjects = decode(scan)
        
        # Iterates through barcodes in image
        for barcode in decodedObjects:
            barcodeNum += 1
            
            # Filters for MAC address format
            macFormat = re.compile("..:..:..:..:..:..")
            
            # Filters out the WLAN MAC
            if barcodeNum == 2:
                print ("Skip WLAN MAC, result {}".format(barcodeNum))
            # Writes serial number to spreadsheet
            elif len(barcode.data.decode("utf-8")) == 10:
                print ("Serial# found, result {}".format(barcodeNum))
                worksheet1.write_string(imageNum, 0, barcode.data.decode("utf-8"))
                serialNumberList.append(barcode.data.decode("utf-8"))
                if nameTrigger == 1:
                    worksheet1.write_string(imageNum, 2, namePrefix + "-" + barcode.data.decode("utf-8"))
                    nameList.append(barcode.data.decode("utf-8"))
            # Writes LAN MAC address to spreadsheet
            elif macFormat.match(barcode.data.decode("utf-8")):
                print ("LAN MAC found, result {}".format(barcodeNum))
                worksheet1.write_string(imageNum, 1, barcode.data.decode("utf-8"))
                macAddressList.append(barcode.data.decode("utf-8"))
            # Filters out everything else
            else:
                print ("Nothing found, result {}".format(barcodeNum))
        print("")
        barcodeNum = 0  
        
    # Creates JS script
    createScript(folderPath, allLines, getDropdowns, outputName, namePrefix, departmentGUI.get(), unitGUI.get(), machineTypeGUI.get(), statusGUI.get())
    
    workbook.close()  
    
# Start button
executeButton = Button(window, text="Start", command=execute)
executeButton.grid(column=0,row=8)
window.mainloop()
