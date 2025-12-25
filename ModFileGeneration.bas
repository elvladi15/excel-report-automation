Attribute VB_Name = "ModFileGeneration"
Sub CreateMailFiles()
	Dim fileGenerated As Boolean
	Dim colGENERAR_REPORTE As String
	Dim colNOMBRE As String

	For Each mailName In ThisWorkbook.ActiveSheet.Evaluate("FILTER(CORREOS[NOMBRE], CORREOS[GENERAR CORREO?] = ""SI"")")
		Call CreateMail(CStr(mailName))
	Next mailName

	If executionMode = "MANUAL" And allFilesCreated Then MsgBox "Archivos creados correctamente."
End Sub

Sub CreateMail(mailName As String)
	Dim mailFiles As Variant
	Dim mailFileCount As Long
	Dim isOneFilePerRange As Boolean

	On Error Goto ErrorHandler

	isOneFilePerRange = ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""" & mailName & """, CORREOS[NOMBRE], CORREOS[UN ARCHIVO POR RANGO?])") = "SI"

	If Dir(baseReportFolder & "\" & mailName, vbDirectory) = "" Then MkDir baseReportFolder & "\" & mailName

	mailFileCount = Application.WorksheetFunction.CountIf(tbl_ARCHIVOS.ListColumns("CORREO").DataBodyRange, mailName)

	For Each mailFileName In ThisWorkbook.ActiveSheet.Evaluate("FILTER(ARCHIVOS[NOMBRE], ARCHIVOS[CORREO] = """ & mailName & """)")
		If isOneFilePerRange Then
			currentProcessDate = Null

			Call CreateMailFile(CStr(mailFileName))
		Else
			Dim dateValue As Date

			For dateValue = startProcessDate To endProcessDate
				currentProcessDate = dateValue

				Call CreateMailFile(CStr(mailFileName))
			Next dateValue
		End If
	Next mailFileName

	Exit Sub
	ErrorHandler:
		AppendToLogsFile ("Ha ocurrido un error al crear los archivos del correo: '" & mailName & "'.")

		continueExecution = False
End Sub

Sub CreateMailFile(mailFileName As String)
	Call AppendToLogsFile("Generando archivo: '" & mailFileName & "'...")

	Dim Workbook As Workbook
	Dim fileReports As Variant
	Dim reportDate As String
	Dim mailName As String
	Dim folder As String
	Dim fileQuantityPerMail As Long

	mailName = CStr(ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""" & mailFileName & """, ARCHIVOS[NOMBRE], ARCHIVOS[CORREO])"))
	folder = baseReportFolder & "\" & mailName
	fileQuantityPerMail = Application.WorksheetFunction.CountIf(tbl_ARCHIVOS.ListColumns("CORREO").DataBodyRange, mailName)

	Set Workbook = Workbooks.Add

	fileReports = ThisWorkbook.ActiveSheet.Evaluate("FILTER(REPORTES[NOMBRE], REPORTES[ARCHIVO] = """ & mailFileName & """)")

	For Each item In fileReports
		Call CreateFileReport(Workbook, CStr(item))
	Next item

	If Workbook.Worksheets(Workbook.Worksheets.Count).Name <> "Sheet1" Then
		For Each Query In Workbook.Queries
			Query.delete
		Next Query

		If fileQuantityPerMail > 1 Then
			If IsNull(currentProcessDate) Then
				 folder = folder & "\" & Format(startProcessDate, "dd") & "-" & Format(endProcessDate, "dd")
			Else
				folder = folder & "\" & Format(currentProcessDate, dateFormat)
			End If

			If Dir(folder, vbDirectory) = "" Then MkDir folder
		End If

		folder = folder & "\" & mailFileName & " "

		If IsNull(currentProcessDate) Then
			If startProcessDate = endProcessDate Then
				folder = folder & Format(endProcessDate, dateFormat)
			Else
				folder = folder & Format(startProcessDate, "dd") & "-" & Format(endProcessDate, "dd")
			End If
		Else
			folder = folder & Format(currentProcessDate, dateFormat)
		End If

		folder = folder & ".xlsx"

		Application.DisplayAlerts = False
		Workbook.SaveAs fileName:=folder, FileFormat:=xlOpenXMLWorkbook
		Application.DisplayAlerts = True

	   Call AppendToLogsFile("Archivo: '" & mailFileName & "' creado exitosamente.")
	Else
		MsgBox "El archivo: '" & mailFileName & "' no pudo ser creado porque no se generó ningún reporte."

		Call AppendToLogsFile("El archivo: '" & mailFileName & "' no pudo ser creado porque no se generó ningún reporte.")

		allFilesCreated = False
	End If

	Workbook.Close False
End Sub

Sub CreateFileReport(Workbook As Workbook, fileReportName As String)
	Dim Worksheet As Worksheet
	Dim reportTable As ListObject
	Dim newTbl As ListObject
	Dim mailFile As String
	Dim mailName As String

	ThisWorkbook.Activate

	Set reportTable = ThisWorkbook.Sheets(fileReportName).ListObjects(fileReportName)

	reportTable.DataBodyRange.Borders.LineStyle = xlContinuous

	If Not IsNull(currentProcessDate) Then reportTable.Range.AutoFilter Field:=reportTable.ListColumns("PROCESS_DATE_FOR_RANGE").Index, Criteria1:=Format(currentProcessDate, "dd-MM-yyyy")

	If Application.WorksheetFunction.CountA(reportTable.DataBodyRange) = 0 Then
		Call AppendToLogsFile("El reporte " & fileReportName & " no trajo registros.")

		errorReport = fileReportName

		GoTo removeFilter
	End If

	If Workbook.Worksheets(1).Name = "Sheet1" Then
		Set Worksheet = Workbook.Worksheets(1)
	Else
		Set Worksheet = Workbook.Worksheets.Add
	End If

	Worksheet.Name = fileReportName

	reportTable.Range.Resize(reportTable.ListRows.Count + 2, reportTable.ListColumns.Count - 1).Copy
	Worksheet.Range("A1").PasteSpecial Paste:=xlPasteFormats

	reportTable.Range.Resize(reportTable.ListRows.Count + 2, reportTable.ListColumns.Count - 1).Copy
	Worksheet.Range("A1").PasteSpecial Paste:=xlPasteValues

	reportTable.DataBodyRange.ClearContents

	Worksheet.Columns.AutoFit

removeFilter:
	reportTable.AutoFilter.ShowAllData
End Sub