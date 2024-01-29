Sub HelloWorldMacro()
	doc = ThisComponent
	sheets = doc.Sheets

	Cleanup()

	CsvURL = File("Select KD result")
	sheetI = SheetBy(CsvURL, "Import")

	CsvURL = File("Select map")
	sheetM = SheetBy(CsvURL, "Map")
End Sub

Function Cleanup()
	doc = ThisComponent
	sheets = doc.Sheets

	' Delete previous data

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

Function SheetBy(File, Name) As Object
	doc = ThisComponent
	sheets = doc.Sheets

	sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
	sheets.insertByName(Name, sheet)

	' CSV encoding is assumed to be UTF-16 LE

	'csv file read options
	Filter = "44,34,65535,1,1/1"
	'Creating a link between the sheet and the .csv source
	sheet.link(File, "", "Text - txt - csv (StarCalc)", _
	Filter, com.sun.star.sheet.SheetLinkMode.VALUE)

	'release link so that the document is independent
	sheet.setLinkMode(com.sun.star.sheet.SheetLinkMode.NONE)
	SheetBy = sheet
End Function
