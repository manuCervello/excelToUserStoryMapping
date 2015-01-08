class BacklogTable(object):
	THEME = 0
	FEATURE = 1
	RELEASE = 2
	USER_STORY = 3
	STATUS = 4

	def __init__(self):
		self.__table = []

	def append(self, row):
		self.__table.append(row)

	def getBacklog(self):
		return self.__table