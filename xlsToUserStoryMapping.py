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

import sys
import time
import ConfigParser

from userStoryMapping import UserStoryMapping
from excelToTable import ExcelToTable
from userStoryMappingToExcel import UserStoryMappingToExcel
from userStoryMappingToHTML import UserStoryMappingToHTML


Config = ConfigParser.RawConfigParser()
Config.read(str(sys.argv[1]))

myExcelToTable = ExcelToTable(Config.get('Input','InputFile'), Config.get('Input','InputSheetName'))
myExcelToTable.setColumns(
	int(Config.get('Input','ThemeColumnIndex')), 
	int(Config.get('Input','FeatureColumnIndex')), 
	int(Config.get('Input','ReleaseColumnIndex')), 
	int(Config.get('Input','IdColumnIndex')), 
	int(Config.get('Input','NameColumnIndex')), 
	int(Config.get('Input','StatusColumnIndex')))

myExcelToTable.setLiterals(Config.get('Input','NewLiteral'),
	Config.get('Input','DoneLiteral'), 
	Config.get('Input','DoingLiteral'))
	
myBacklogTable = myExcelToTable.convert()

myUserStoryMapping = UserStoryMapping(Config.get('Input','ReleaseOrder'))
myUserStoryMapping.createFromTable(myBacklogTable)

today = time.strftime("%Y-%m-%d")

myUserStoryMappingToExcel = UserStoryMappingToExcel(today+"_"+Config.get('Output','OutputFile')+".xls", Config.get('Output','OutputExcelSheetName'))
myUserStoryMappingToExcel.convert(myUserStoryMapping)

myUserStoryMappingToHTML = UserStoryMappingToHTML(today+"_"+Config.get('Output','OutputFile')+".html")
myUserStoryMappingToHTML.convert(myUserStoryMapping)

