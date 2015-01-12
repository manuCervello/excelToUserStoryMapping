from treeNode import TreeNode
from backlogTable import BacklogTable


class UserStoryMapping(object):
	def __init__(self, releasesNames):
		self.__tree = TreeNode('rootNode')
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

	def getTotalFeatures(self):
		return self.__totalFeatures

	def createFromTable(self, table):
		for row in table.getBacklog():
			themeNode = self.__addTheme(row[BacklogTable.THEME])
			featureNode = self.__addFeature(themeNode, row[BacklogTable.FEATURE])
			releaseNode = self.__addReleaseToFeature(featureNode, row[BacklogTable.RELEASE])
			userStoryNode = self.__addUserStory(releaseNode, row[BacklogTable.USER_STORY])
			self.__addStatus(userStoryNode, row[BacklogTable.STATUS])

	def __createNode(self, parentNode, name):
		__node = parentNode.getChildByName(name)
		if (__node == None):
			__node = parentNode.addChild(TreeNode(name))
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
			__node = themeNode.addChild(TreeNode(featureName))
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
		