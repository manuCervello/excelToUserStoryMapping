class TreeNode(object):
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