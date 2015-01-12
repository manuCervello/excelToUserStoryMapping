import xlrd
from backlogTable import BacklogTable

THEME_COLUMN_INDEX = 0
FEATURE_COLUMN_INDEX = 1
RELEASE_COLUMN_INDEX = 2
ID_COLUMN_INDEX = 3
USER_STORY_NAME_COLUMN_INDEX = 4
STATUS_COLUMN_INDEX = 5


class ExcelToTable(object):

	def __init__(self, fileName, sheetName):
		self.__fileName = fileName
		self.__sheetName = sheetName
		self.__excelThemeColumn = 0
		self.__excelFeatureColumn = 0
		self.__excelReleaseColumn = 0
		self.__excelIdColumn = 0
		self.__excelUserStoryColumn = 0
		self.__excelStatusColumn = 0
		self.__new = "New"
		self.__done = "Done"
		self.__doing = "Doing"
		
	def setLiterals(self, new, done, doing):
		self.__new = new
		self.__done = done
		self.__doing = doing
		

	def setColumns(self, theme, feature, release, idNumber, userStory, status):
		self.__excelThemeColumn = theme
		self.__excelFeatureColumn = feature
		self.__excelReleaseColumn = release
		self.__excelIdColumn = idNumber
		self.__excelUserStoryColumn = userStory
		self.__excelStatusColumn = status
	
	def convert(self):
		table = BacklogTable()

		backlogExcel = xlrd.open_workbook(self.__fileName)
		worksheet = backlogExcel.sheet_by_name(self.__sheetName)

		for currentRow in range(1, (worksheet.nrows - 1), 1):
			theme = worksheet.cell_value(currentRow, self.__excelThemeColumn)
			feature = worksheet.cell_value(currentRow, self.__excelFeatureColumn)
			release = worksheet.cell_value(currentRow, self.__excelReleaseColumn).strip()
			id = int(worksheet.cell_value(currentRow, self.__excelIdColumn))
			name = worksheet.cell_value(currentRow, self.__excelUserStoryColumn)
			status = worksheet.cell_value(currentRow, self.__excelStatusColumn)
			if status == self.__done:
				status = "Done"
			elif status == self.__doing:
				status = "Doing"
			else:
				status = "New"
				
			table.append([theme, feature, release, str(id) + " - " + name, status])

		return table
