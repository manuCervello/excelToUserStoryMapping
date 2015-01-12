Requirements:

xlrd: https://pypi.python.org/pypi/xlrd
xlwt: https://pypi.python.org/pypi/xlwt
html: https://pypi.python.org/pypi/html/1.16


The basic call is python xlsToUserStoryMapping.py config.ini > output.txt

The config.ini file should contain those fields:

[Input]
InputFile: Backlog.xls
InputSheetName: Backlog Items
IdColumnIndex: 1
NameColumnIndex: 2
StatusColumnIndex: 5
ReleaseColumnIndex: 6
ThemeColumnIndex: 7
FeatureColumnIndex: 8
ReleaseOrder: 'Release 1','Release 2','Release 3','Release 4','Uncategorized'

[Output]
RootNodeName: Project
OutputExcelFile: Project_userStoryMapping.xls
OutputExcelSheetName: Project User Story Mapping
OutputHTMLFile: Project_userStoryMapping.html


The release Order is important and all those names should be present in the Backlog.xls. The last one 'Uncategorized' is used to group all the
undefined releases that we don't want to show.

