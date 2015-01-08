import xlwt

#RELEASE_BASE: Excel row where the first release start, after Theme, feature, space rows and separator
RELEASE_BASE = 5
	
#Separator patterns: feature separator and release separator
bottomBorder = xlwt.easyxf('borders: bottom dashed;')
bottonMediumBorder = xlwt.easyxf('borders: bottom medium;')

class UserStoryMappingToExcel(object):

	def __init__(self, fileName, sheetName):
		self.__fileName = fileName
		self.__sheetName = sheetName
		self.__userStoryMappingBook = xlwt.Workbook(encoding='utf-8')
		self.__sheet = self.__userStoryMappingBook.add_sheet(self.__sheetName)
	
	def convert(self, userStoryMapping):

		self.__fortmatColumns(userStoryMapping, bottomBorder)
		self.__formatReleases(userStoryMapping, RELEASE_BASE, bottonMediumBorder)
		self.__formatSheet(userStoryMapping, RELEASE_BASE, 256*5)
		self.__createUserStoryCards(userStoryMapping)

		self.__userStoryMappingBook.save(self.__fileName)

	def __fortmatColumns(self, userStoryMapping, featuresPattern):
		for i in range(0, (userStoryMapping.getTotalFeatures()*2), 2):
			self.__sheet.col(i).width = 500
			self.__sheet.col(i+1).width = 5000
			self.__sheet.write(3, i, "", featuresPattern)
			self.__sheet.write(3, i+1, "", featuresPattern)

	def __formatReleases(self, userStoryMapping, releaseBaseRow, releasePattern):
		currentRelease = 0
		#Write Relases with Borders
		for element in userStoryMapping.releaseOrder:
			currentRelease += element[1]*2
			self.__sheet.write(releaseBaseRow+currentRelease, 1, element[0], releasePattern)
			for i in range(0, (userStoryMapping.getTotalFeatures()*2)-2, 1):
				self.__sheet.write(releaseBaseRow+currentRelease, i+2, "", releasePattern)
			currentRelease += 2

	def __setHeight(self, sheet, row, height):
		sheet.row(row).height_mismatch = True
		sheet.row(row).height = height

	def __formatSheet(self, userStoryMapping, releaseBaseRow, height):
		self.__setHeight(self.__sheet, 0, height)
		self.__setHeight(self.__sheet, 2, height)

		currentRow = 0
		for element in userStoryMapping.releaseOrder:
			for i in range(0, (element[1]*2), 2):
				self.__setHeight(self.__sheet, releaseBaseRow+currentRow+i, height)
			currentRow += (element[1]*2) + 2
		
	#Patterns to format cells according to their content
	def __getPattern(self, status):
		if status == 'Done':
			return xlwt.easyxf('align: wrap 1; pattern: pattern solid, fore_colour Lime;')
		elif status == 'New':
			return xlwt.easyxf('align: wrap 1; pattern: pattern solid, fore_colour 22;')
		else:
			return xlwt.easyxf('align: wrap 1; pattern: pattern solid, fore_colour Orange;')

	# write information
	def __createUserStoryCards(self, userStoryMapping):
		themeStyle = xlwt.easyxf('font: bold 1; align: wrap 1; pattern: pattern solid, fore_colour Aqua;')
		featureStyle = xlwt.easyxf('font: bold 1; align: wrap 1; pattern: pattern solid, fore_colour Gold;')
		featuresWritten = 0
		for theme in userStoryMapping.getThemes():
			self.__sheet.write(0, 1+(featuresWritten*2), theme.name, themeStyle)
			for feature in userStoryMapping.getFeatures(theme):
				self.__sheet.write(2, 1+(featuresWritten*2), feature.name, featureStyle)
				currentRelease = 0
				for element in userStoryMapping.releaseOrder:
					currentRow = 0
					releaseNode = userStoryMapping.getRelease(feature, element[0])
					if (releaseNode != None):
						for userStory in userStoryMapping.getUserStories(releaseNode):
							self.__sheet.write(RELEASE_BASE+currentRelease+currentRow, 1+(featuresWritten*2), userStory.name, self.__getPattern(userStory.status))
							currentRow += 2
					
					currentRelease += (element[1]*2)+2
				featuresWritten += 1

