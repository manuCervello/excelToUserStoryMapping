Requirements:

Python 2.7.x installed with the following modules:
- xlrd
- xlwt
- sys (by default)
- ConfigParser (by default)

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
OutputFile: Project_userStoryMapping.xls
OutputSheetName: Project User Story Mapping


The release Order is important and all those names should be present in the Backlog.xls. The last one 'Uncategorized' is used to group all the
undefined releases that we don't want to show.

