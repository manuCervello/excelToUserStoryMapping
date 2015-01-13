class UserStoryMappingToHTML(object):

	def __init__(self, fileName):
		self.__fileName = fileName
		self.__totalTableColumns = 0
	
	def __setStyle(self):
		style = "<style type='text/css'>"
		style += "*{margin: 0px; padding: 0px;}"
		style += "#fixedCards{position: fixed; z-index:999;}"
		style += "#movingCards{position: relative; top: 280px; z-index:99;}"
		style += "table{border-collapse: collapse;}"
		style += ".themeClass{background-color: #A9BCF5}"
		style += ".featureClass{background-color: #F4FA58}"
		style += ".releaseNameClass{width: 150px; font-size: large; font-weight: bold;}"
		style += ".NewClass{background-color: #BDBDBD}"
		style += ".DoneClass{background-color: #58FA58}"
		style += ".DoingClass{background-color: #FE9A2E}"
		style += ".cardClass{height: 100px; width: 200px; margin:8px; padding:8px;}"
		style += "tr.releaseRow td.releaseRow{border-bottom: 1px solid black}"
		style += "</style>"
		return style

	def __setScripts(self):
		f = open("../js/jquery-1.11.2.min.js", "r")
		string = "<script>"
		string += f.read()
		string += "$(window).scroll(function(event) {$(\"#fixedCards\").css(\"margin-left\", 0-$(document).scrollLeft());});"
		string += "</script>"
		f.close()
		return string
		
	def __insertEmptyCell(self):
		return "<td></td>"

	def __insertReleaseCell(self):
		return "<td class='releaseRow'></td>"
		
	def __insertReleaseNameCard(self, text):
		return "<div class='releaseNameClass'>" + text + "</div>"

	def __insertCard(self, text, cardPattern):
		return "<div class='" + cardPattern + " cardClass'>" + text + "</div>"
		
	def __insertThemesRow(self, userStoryMapping):
		string = "<tr>"
		string += "<td>" + self.__insertReleaseNameCard("") + "</td>"
		for theme in userStoryMapping.getThemes():
			string += "<td>" + self.__insertCard(theme.name, "themeClass") + "</td>"
			for feature in range(1, len(userStoryMapping.getFeatures(theme)), 1):
				string += self.__insertEmptyCell()
		string += "</tr>"
		return string

	def __insertFeaturesRow(self, userStoryMapping):
		string = "<tr>"
		string += "<td>" + self.__insertReleaseNameCard("") + "</td>"
		for theme in userStoryMapping.getThemes():
			for feature in userStoryMapping.getFeatures(theme):
				string += "<td>" + self.__insertCard(feature.name, "featureClass") + "</td>"
				self.__totalTableColumns += 2
		string += "</tr>"
		return string

	def __insertFeaturesSeparator(self):
		string = "<tr>"
		for cell in range(0, self.__totalTableColumns, 1):
			string += self.__insertEmptyCell()
		string += "</tr>"
		return string
	
	def __createEmptyArray(self, width, height):
		array = []
		for i in range(0, width):
			array.append([])
			for j in range(0, height):
				array[i].append(None)
		return array
		
	def __getArrayRelease(self, userStoryMapping, releaseName, depth):
		print "::::"+releaseName
		userStoriesPerRelease = self.__createEmptyArray(userStoryMapping.getTotalFeatures(), depth)
		
		currentFeature = 0
		for theme in userStoryMapping.getThemes():
			for feature in userStoryMapping.getFeatures(theme):
				releaseNode = userStoryMapping.getRelease(feature, releaseName)
				if (releaseNode != None):
					currentUserStory = 0
					for userStory in userStoryMapping.getUserStories(releaseNode):
						userStoriesPerRelease[currentFeature][currentUserStory] = [userStory.name, userStory.status]
						currentUserStory += 1

				currentFeature += 1
				
		return userStoriesPerRelease
			
	def __insertUserStoriesInRelease(self, userStoryMapping, releaseName, depth):
		userStories = self.__getArrayRelease(userStoryMapping, releaseName, depth)
		string = ""
		for row in range(0, depth):
			string += "<tr>"
			string += "<td>" + self.__insertReleaseNameCard("") + "</td>"
			for column in range(0, len(userStories)):
				if userStories[column][row] != None:
					string += "<td>" + self.__insertCard(userStories[column][row][0], userStories[column][row][1]+"Class") + "</td>"
				else:
					string += self.__insertEmptyCell()
			string += "</tr>"
		return string
			
	def __insertReleaseName(self, releaseName, totalFeatures):
		string = "<tr class='releaseRow'>"
		string += "<td class='releaseRow'>" + self.__insertReleaseNameCard(releaseName) + "</td>"
		for cell in range(0, totalFeatures):
			string += self.__insertReleaseCell()
			string += self.__insertReleaseCell()
		string += "</tr>"
		
		return string
			
		
	
	def convert(self, userStoryMapping):
		html = "<html>"
		html += "<body>"
		html += "<head>"
		html += self.__setScripts()
		html += self.__setStyle()
		html += "</head>"
		featuresWritten = 0
		html += "<div id='fixedCards'>"
		html += "<table style=\"background-color:white\">"
		html += self.__insertThemesRow(userStoryMapping)
		html += self.__insertFeaturesRow(userStoryMapping)
		html += self.__insertFeaturesSeparator()
		html += "</table>"
		html += "</div>"
		html += "<div id='movingCards'>"
		html += "<table style=\"background-color:white\">"
		for release in range(0, len(userStoryMapping.releaseOrder)):
			print userStoryMapping.releaseOrder[release][0]
			html += self.__insertUserStoriesInRelease(userStoryMapping, userStoryMapping.releaseOrder[release][0], userStoryMapping.releaseOrder[release][1])
			html += self.__insertReleaseName(userStoryMapping.releaseOrder[release][0], userStoryMapping.getTotalFeatures())

		html += "</table>"
		html += "</div>"
		html += "</body>"
		html += "</html>"

		f = open(self.__fileName, "w")
		f.write(html)
		f.close()
		
