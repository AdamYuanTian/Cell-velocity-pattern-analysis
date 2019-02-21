#This code is used to identify the cells that go through tethering from all the tracked cells.

#Version 3 README (Old)
#Currently, the defination of tethering here is:
#	1. For a cell, its instantaneous velocities are lower than a specific cutoff velocity for a certain length of frames.
#		Eg. The instantaneous velocity of a cell is lower than 10 um/s for 3 continuous frames.
#Update needed: 1. For a tethering cell, if the starting velocity is smaller than a specific value, it should not be included in the final list.
#                  Since the cell is not moving at all.
#               2. Instantaneous velocity is calculated using ((x2-x1)^2+(y2-y1)^2)^0.5, so the cell area fluctuation caused by cell overlap
#                  will influence cell location (center of area), thus influence instantaneous velocity. The fluctuation should be excluded.

#Version 4 README
#Currently, the defination of tethering here is:
#	1. (Max&Min) For the instantaneous velocities of a cell, if the difference between Min and Max is equal or smaller than a cutoff value (negative), 
#	   and the Min is after Max, the cell is considered tethering. If the Min is not after Max, 2nd round assessment is needed.
#	2. (Dynamic Baseline) A dynamic velocity baseline is used, and is always replaced by a higher value when going through all the instantaneous velocities.
#	   When the difference between Current Velocity and the Baseline reaches the cutoff (a negative value), the cell is considered tethering.
#	Note: The code will go through all information, but currently we are only looking at Rolled Cells at Shear Rate 20 s-1


########### Place the py file in the folder with excel files need to be analyzed.

import os
import openpyxl
import xlwt
import pandas as pd 
from pandas import ExcelWriter
from pandas import ExcelFile

########### Custom parameters for analysis
#continuousVelocityLength = int(input('Please type in the cutoff velocity difference for tethering: '))
#velocityDiffCutoff = int(input('Please type in the cutoff velocity difference for cell tethering: '))
velocityDiffCutoff = -100 #-100 um/s is used currently for shear rate 20 s-1

########### Get the list of directory of files needs to be analyzed
sourceDir = os.getcwd() #get current working directory
items = os.listdir(".") #get all the file names under this folder
folderList = []
pathList = []
for names in items:
	if names.endswith(".xlsx"): #get all the xlsx files under this folder
		folderList.append(names) #add all the xlsx files to the folderList
		pathList.append(sourceDir + '\\' + names) #add all the directory of xlsx files to the pathList

########## Check each file in the folder
for fileName in folderList:
	print ('File: ' + fileName + ' is being processed...')
	wb = openpyxl.load_workbook(fileName) #load the file

########## Check each sheet in the file
	tetheringCellName_allSheets=[]
	tetheringCellVelocity_allSheets=[]
	
	sheetNames = wb.sheetnames #load all the sheet names and save to sheetNames
	for sheet in sheetNames: #load each sheet in the file
		if sheet=='Sheet1': #'Sheet1' is blank
			continue
		print ('Sheet: ' + sheet + ' is being processed...')
		df = pd.read_excel(fileName, sheet_name=sheet) #store all data in current sheet into dataframe

########## Store each cell velocities from the sheet into a 2D list of list
		listVelocity = df['Velocity (um/s)'] #store the whole velocity column into a list
		listLabel = df['Label'] #store the whole label column into a list
		listCellStatus = df['Cell Status'] #store the cell status into a list
		cellVelocity_temp = [] #for storing current cell velocity temporarily
		allCellVelocity = [] #for storing all cell velocity in 2D list
		allCellName = [] #for storing all cell names
		
		for i in range(len(listVelocity)): #go through each rows
			if listVelocity[i] == 0 and pd.notnull(listCellStatus[i]): #sometines the velocity can be zero in the middle, so make sure cell status is not NaN
				allCellVelocity.append(cellVelocity_temp) #add one cell velocity to 2D list for all cells
				cellVelocity_temp=[] #reset
				allCellName.append(listLabel[i]) #add current cell name to list
				continue
			cellVelocity_temp.append(listVelocity[i])
		
		allCellVelocity.append(cellVelocity_temp) #add the last cell's velocities to the 2D list
		allCellVelocity.pop(0) #clear the first empty element
		

########### Analyze each cell velocities for finding tethering
		tetheringCellVelocity = [] #store the tethering cell's velocities in current sheet
		tetheringCellName = []
		for j in range(len(allCellVelocity)): #check each cell in the 2D list
			############# Changes needed for tethering criteria
			#for n in range(len(allCellVelocity[j])-continuousVelocityLength): #check each velocity for the cell
			#	count = 0
			#	if allCellVelocity[j][n] < velocityCutoff:
			#		for m in range(continuousVelocityLength): #check the value of continuous number with the length of continuousVelocityLength
			#			if allCellVelocity[j][n+m] < velocityCutoff:
			#				count = count + 1
			#	if count == continuousVelocityLength: #compare the length 
			#		tetheringCellVelocity.append(allCellVelocity[j]) #store the velocity of tethering cells
			#		tetheringCellName.append(allCellName[j]) #store the name of tethering cells
			#		break
		
			######Criteria 1, compare the Max and Min instantaneous velocity of a cell, and compare their order
			maxInsVelocity = max(allCellVelocity[j]) #Get the max velocity for a cell
			maxInsVelocityLoc = allCellVelocity[j].index(max(allCellVelocity[j])) #Get the location of the 1st max velocity, assuming there is only one max
			
			minInsVelocity = min(allCellVelocity[j]) #Get the min velocity for a cell              
			minInsVelocityLoc = [] #Store all the locations of min velocity
			for counter, value in enumerate(allCellVelocity[j]): #Get all the locations of min velocity, there maybe multiple zero velocities for a cell
				if value == minInsVelocity:
					minInsVelocityLoc.append(counter)
			
			if minInsVelocity - maxInsVelocity <= velocityDiffCutoff and maxInsVelocityLoc < minInsVelocityLoc[-1]: #Apply Criteria 1
				tetheringCellVelocity.append(allCellVelocity[j]) #Store the velocity of tethering cells
				tetheringCellName.append(allCellName[j]) #Store the name of tethering cells
				continue #Skip Criteria 2 if Critera 1 is fulfilled
				
			######Criteria 2, use a Dynamic Baseline to assess the drop in instantaneous velocity
			baseline = allCellVelocity[j][0] #Initialize the baseline with the 1st instantaneous velocity
			
			for n in range(len(allCellVelocity[j])): #Go through each instantaneous velocity
				velocityDiff = allCellVelocity[j][n]-baseline
				if velocityDiff > 0: #Replace the baseline with a higher value
					baseline = allCellVelocity[j][n] #Dynamic Baseline
					continue #Proceed to next loop
				if velocityDiff <= velocityDiffCutoff: #Meet the 2nd criteria
					tetheringCellVelocity.append(allCellVelocity[j]) #Store the velocity of tethering cells
					tetheringCellName.append(allCellName[j]) #Store the name of tethering cells
					break #Terminate the loop
					

		tetheringCellName_allSheets.append(tetheringCellName)
		tetheringCellVelocity_allSheets.append(tetheringCellVelocity)

	#####write the data into excel and move onto next file
	wkbk = xlwt.Workbook()
	for a in range(len(sheetNames)-1):  #'Sheet1' is not included because it is blank
		ws = wkbk.add_sheet(sheetNames[a+1]) #add sheet to file
		for b in range(len(tetheringCellName_allSheets[a])):
			ws.write(0,b,tetheringCellName_allSheets[a][b]) #add cell name to the first row
			for c in range(len(tetheringCellVelocity_allSheets[a][b])):
				ws.write(c+1,b,tetheringCellVelocity_allSheets[a][b][c]) #add value into each cell, a is the sheet, b is the cell, c is the velocity
	wkbk.save(fileName[:-5]+'_Tethering Analyzed'+'.xls')

#Add a summary file at the end, including 

	print ('\n') #space between files


print ('Analysis is done!')
