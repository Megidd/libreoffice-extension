Sub HelloWorldMacro()
	doc = ThisComponent
	sheets = doc.Sheets

	' Get import file.
	CsvURL = FileName("Select KD output")

	'csv file read options
	Filter = "44,34,65535,1,1/1"

	' Make sure at least one sheet exists.
	exists = sheets.hasByName("Sheet1")
	if not exists then sheets.insertNewByName("Sheet1", 0)

	' Create sheet for import.
	exists = sheets.hasByName("Imported")
	if exists then sheets.removeByName("Imported")

	sheetI = doc.createInstance("com.sun.star.sheet.Spreadsheet")
	sheets.insertByName("Imported", sheetI)

	'import creating a link between the sheet and the .csv source
	sheetI.link(CsvURL, "", "Text - txt - csv (StarCalc)", _
	Filter, com.sun.star.sheet.SheetLinkMode.VALUE)

	'release link so that the document is independent
	sheetI.setLinkMode(com.sun.star.sheet.SheetLinkMode.NONE)

	' Create sheet for map.
	exists = sheets.hasByName("Map")
	if exists then sheets.removeByName("Map")

	sheetM = doc.createInstance("com.sun.star.sheet.Spreadsheet")
	sheets.insertByName("Map", sheetM)

End Sub

Function FileName(msg) As String
	' User home directory is needed to get files by file picker.
	Dim oSubst As Object, Home As String
	oSubst = CreateUnoService("com.sun.star.util.PathSubstitution")
	Home = oSubst.getSubstituteVariableValue("$(home)")

	Dim oFilePicker As Object, FilePath As String
	FilePath = ""
	'FilePicker initialization
	oFilePicker = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
	oFilePicker.DisplayDirectory = Home
	oFilePicker.appendFilter("CSV Documents", "*.csv")
	oFilePicker.CurrentFilter = "CSV Documents"
	oFilePicker.Title = msg
	'execution and return check (OK?)
	If oFilePicker.execute = _
		com.sun.star.ui.dialogs.ExecutableDialogResults.OK Then
		FilePath = oFilePicker.Files(0)
	End If

	FileName = FilePath
End Function
