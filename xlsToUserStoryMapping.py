#Licensed to the Apache Software Foundation (ASF) under one
#or more contributor license agreements.  See the NOTICE file
#distributed with this work for additional information
#regarding copyright ownership.  The ASF licenses this file
#to you under the Apache License, Version 2.0 (the
#"License"); you may not use this file except in compliance
#with the License.  You may obtain a copy of the License at
#
#  http://www.apache.org/licenses/LICENSE-2.0
#
#Unless required by applicable law or agreed to in writing,
#software distributed under the License is distributed on an
#"AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
#KIND, either express or implied.  See the License for the
#specific language governing permissions and limitations
#under the License.

#Twitter: @ManuCervello

import xlrd
import xlwt
import sys
import ConfigParser

#RELEASE_BASE: Excel row where the first release start, after Theme, feature, space rows and separator
RELEASE_BASE = 5

THEME_COLUMN_INDEX = 0
FEATURE_COLUMN_INDEX = 1
RELEASE_COLUMN_INDEX = 2
ID_COLUMN_INDEX = 3
USER_STORY_NAME_COLUMN_INDEX = 4
STATUS_COLUMN_INDEX = 5

#Separator patterns: feature separator and release separator
bottomBorder = xlwt.easyxf('borders: bottom dashed;')
bottonMediumBorder = xlwt.easyxf('borders: bottom medium;')

class commonTreeNode(object):
	def __init__(self, name):
		self.name = name
		self.status = ""
		self.children = []

	def addChild(self, obj):
		for c in self.children:
			if (c.name == obj.name):
				return obj
				
		self.children.append(obj)
		return obj

	def getChildByName(self, name):
		for c in self.children:
			if c.name == name:
				return c
		return None

class userStoryMapping(object):
	def __init__(self, rootName, releasesNames):
		self.__tree = commonTreeNode(rootName)
		self.__totalFeatures = 0
		self.releaseOrder = []
		self.__preReleaseOrder = releasesNames.split(',')
		for element in self.__preReleaseOrder:
			self.releaseOrder.append([element.translate(None, "'"), 0])

	def getThemes(self):
		return self.__tree.children

	def getFeatures(self, themeNode):
		return themeNode.children
	
	def getRelease(self, featureNode, releaseName):
		return featureNode.getChildByName(releaseName)
		
	def getUserStories(self, releaseNode):
		return releaseNode.children

	def importFromExcelFile(self, inputFile, sheetName, columns):
		backlogExcel = xlrd.open_workbook(inputFile)
		worksheet = backlogExcel.sheet_by_name(sheetName)
		
		for currentRow in range(1, (worksheet.nrows - 1), 1):
			theme = worksheet.cell_value(currentRow, columns[THEME_COLUMN_INDEX])
			feature = worksheet.cell_value(currentRow, columns[FEATURE_COLUMN_INDEX])
			release = worksheet.cell_value(currentRow, columns[RELEASE_COLUMN_INDEX]).strip()
			id = int(worksheet.cell_value(currentRow, columns[ID_COLUMN_INDEX]))
			name = worksheet.cell_value(currentRow, columns[USER_STORY_NAME_COLUMN_INDEX])
			status = worksheet.cell_value(currentRow, columns[STATUS_COLUMN_INDEX])

			themeNode = self.__addTheme(theme)
			featureNode = self.__addFeature(themeNode, feature)
			releaseNode = self.__addReleaseToFeature(featureNode, release)
			userStoryNode = self.__addUserStory(releaseNode, str(id)+'-'+name)
			self.__addStatus(userStoryNode, status)
			print "creating node for ", theme, " : ", name

	def __createNode(self, parentNode, name):
		__node = parentNode.getChildByName(name)
		if (__node == None):
			__node = parentNode.addChild(commonTreeNode(name))
		return __node

	def __existRelease(self, name):
		for x,y in self.releaseOrder:
			if x == name:
				return True
		return False

	def __setCapacity(self, releaseName, amount):
		for obj in self.releaseOrder:
			if obj[0] == releaseName:
				if obj[1] < amount:
					obj[1] = amount
				return
		
	def __addTheme(self, themeName):
		return self.__createNode(self.__tree, themeName)
		
	def __addFeature(self, themeNode, featureName):
		__node = themeNode.getChildByName(featureName)
		if (__node == None):
			self.__totalFeatures += 1
			__node = themeNode.addChild(commonTreeNode(featureName))
		return __node

	def __addReleaseToFeature(self, featureNode, releaseName):
		if self.__existRelease(releaseName) == False:
			releaseName = 'Uncategorized'
		return self.__createNode(featureNode, releaseName)

	def __addUserStory(self, releaseNode, userStoryName):
		userStoryNode = self.__createNode(releaseNode, userStoryName)
		self.__setCapacity(releaseNode.name, len(releaseNode.children))
		
		return userStoryNode

	def __addStatus(self, userStoryNode, statusName):
		userStoryNode.status = statusName
		
	def getTotalFeatures(self):
		return self.__totalFeatures
		

Config = ConfigParser.RawConfigParser()
Config.read(str(sys.argv[1]))

def fortmatColumns(sheet, userStoryMapping, featuresPattern):
	for i in range(0, (userStoryMapping.getTotalFeatures()*2), 2):
		sheet.col(i).width = 500
		sheet.col(i+1).width = 5000
		sheet.write(3, i, "", featuresPattern)
		sheet.write(3, i+1, "", featuresPattern)

def formatReleases(sheet, userStoryMapping, releaseBaseRow, releasePattern):
	currentRelease = 0
	#Write Relases with Borders
	for element in userStoryMapping.releaseOrder:
		currentRelease += element[1]*2
		sheet.write(releaseBaseRow+currentRelease, 1, element[0], releasePattern)
		for i in range(0, (userStoryMapping.getTotalFeatures()*2)-2, 1):
			sheet.write(releaseBaseRow+currentRelease, i+2, "", releasePattern)
		currentRelease += 2
		
def setHeight(sheet, row, height):
	sheet.row(row).height_mismatch = True
	sheet.row(row).height = height

def formatSheet(sheet, userStoryMapping, releaseBaseRow, height):
	setHeight(sheet, 0, height)
	setHeight(sheet, 2, height)

	currentRow = 0
	for element in userStoryMapping.releaseOrder:
		for i in range(0, (element[1]*2), 2):
			setHeight(sheet, releaseBaseRow+currentRow+i, height)
		currentRow += (element[1]*2) + 2
		

#Patterns to format cells according to their content
def getPattern(status):
	if status == 'Done':
		return xlwt.easyxf('align: wrap 1; pattern: pattern solid, fore_colour Lime;')
	elif status == 'New':
		return xlwt.easyxf('align: wrap 1; pattern: pattern solid, fore_colour 22;')
	else:
		return xlwt.easyxf('align: wrap 1; pattern: pattern solid, fore_colour Orange;')

# write information
def createUserStoryCards(sheet, myUserStoryMapping):
	themeStyle = xlwt.easyxf('font: bold 1; align: wrap 1; pattern: pattern solid, fore_colour Aqua;')
	featureStyle = xlwt.easyxf('font: bold 1; align: wrap 1; pattern: pattern solid, fore_colour Gold;')
	featuresWritten = 0
	for theme in myUserStoryMapping.getThemes():
		sheet.write(0, 1+(featuresWritten*2), theme.name, themeStyle)
		for feature in myUserStoryMapping.getFeatures(theme):
			sheet.write(2, 1+(featuresWritten*2), feature.name, featureStyle)
			currentRelease = 0
			for element in myUserStoryMapping.releaseOrder:
				currentRow = 0
				releaseNode = myUserStoryMapping.getRelease(feature, element[0])
				if (releaseNode != None):
					for userStory in myUserStoryMapping.getUserStories(releaseNode):
						sheet.write(RELEASE_BASE+currentRelease+currentRow, 1+(featuresWritten*2), userStory.name, getPattern(userStory.status))
						currentRow += 2
					
				currentRelease += (element[1]*2)+2
			featuresWritten += 1


columnIndexes = [int(Config.get('Input','ThemeColumnIndex')), int(Config.get('Input','FeatureColumnIndex')), int(Config.get('Input','ReleaseColumnIndex')), int(Config.get('Input','IdColumnIndex')), int(Config.get('Input','NameColumnIndex')), int(Config.get('Input','StatusColumnIndex'))]

myUserStoryMapping = userStoryMapping(Config.get('Output','RootNodeName'), Config.get('Input','ReleaseOrder'))
myUserStoryMapping.importFromExcelFile(Config.get('Input','InputFile'), Config.get('Input','InputSheetName'), columnIndexes)

userStoryMappingBook = xlwt.Workbook(encoding='utf-8')
sheet = userStoryMappingBook.add_sheet(Config.get('Output','OutputSheetName'))
		
fortmatColumns(sheet, myUserStoryMapping, bottomBorder)
formatReleases(sheet, myUserStoryMapping, RELEASE_BASE, bottonMediumBorder)
formatSheet(sheet, myUserStoryMapping, RELEASE_BASE, 256*5)
createUserStoryCards(sheet, myUserStoryMapping)

userStoryMappingBook.save(Config.get('Output','OutputFile'))

