Sub HelloWorldMacro()
	doc = ThisComponent
	sheets = doc.Sheets

	Cleanup()

	' Get import file.
	CsvURL = File("Select KD output")

	sheetI = doc.createInstance("com.sun.star.sheet.Spreadsheet")
	sheets.insertByName("Imported", sheetI)

	'csv file read options
	Filter = "44,34,65535,1,1/1"
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

Function Cleanup()
	doc = ThisComponent
	sheets = doc.Sheets

	' Make sure at least one sheet exists.
	exists = sheets.hasByName("Sheet1")
	if not exists then sheets.insertNewByName("Sheet1", 0)

	exists = sheets.hasByName("Imported")
	if exists then sheets.removeByName("Imported")

	exists = sheets.hasByName("Map")
	if exists then sheets.removeByName("Map")
End Function

Function File(msg) As String
	Dim oFilePicker As Object
	oFilePicker = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
	oFilePicker.appendFilter("CSV Documents", "*.csv")
	oFilePicker.CurrentFilter = "CSV Documents"
	oFilePicker.Title = msg
	'execution and return check (OK?)
	If oFilePicker.execute = _
		com.sun.star.ui.dialogs.ExecutableDialogResults.OK Then
		File = oFilePicker.Files(0)
	End If
End Function
