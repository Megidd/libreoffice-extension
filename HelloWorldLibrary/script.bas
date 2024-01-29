Sub HelloWorldMacro()
	doc = ThisComponent
	sheets = doc.Sheets

	Cleanup()

	CsvURL = File("Select KD output")

	sheetI = doc.createInstance("com.sun.star.sheet.Spreadsheet")
	sheets.insertByName("Import", sheetI)

	'csv file read options
	Filter = "44,34,65535,1,1/1"
	'Creating a link between the sheet and the .csv source
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

	exists = sheets.hasByName("Import")
	if exists then sheets.removeByName("Import")

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
