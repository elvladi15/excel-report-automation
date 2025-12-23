Attribute VB_Name = "ModFileGeneration"
Sub CreateMailFiles()
	Dim fileGenerated As Boolean
	Dim colGENERAR_REPORTE As String
	Dim colNOMBRE As String

	For Each row In tbl_CORREOS.DataBodyRange.Rows
		colGENERAR_REPORTE = row.Cells(1, tbl_CORREOS.ListColumns("GENERAR CORREO?").Index).Value
		colNOMBRE = row.Cells(1, tbl_CORREOS.ListColumns("NOMBRE").Index).Value

		If colGENERAR_REPORTE = "SI" Then
			Call CreateMail(colNOMBRE)
		End If
	Next row

	If executionMode = "MANUAL" Then MsgBox "Archivos creados correctamente."
End Sub

Sub CreateMail(mailName As String)
	Dim mailFiles As Variant
	Dim mailFileCount As Long
	Dim isOneFilePerRange As Boolean

	Dim colNOMBRE As String
	Dim colCORREO As String

	Application.DisplayAlerts = False

	isOneFilePerRange = ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""" & mailName & """, CORREOS[NOMBRE], CORREOS[UN ARCHIVO POR RANGO?])") = "SI"

	If Dir(baseReportFolder & "\" & mailName, vbDirectory) = "" Then MkDir baseReportFolder & "\" & mailName

	mailFileCount = Application.WorksheetFunction.CountIf(tbl_ARCHIVOS.ListColumns("CORREO").DataBodyRange, mailName)

	For Each row In tbl_ARCHIVOS.DataBodyRange.Rows
		colNOMBRE = row.Cells(1, tbl_ARCHIVOS.ListColumns("NOMBRE").Index).Value
		colCORREO = row.Cells(1, tbl_ARCHIVOS.ListColumns("CORREO").Index).Value

		If colCORREO <> mailName Then GoTo continueLoop
		
		If isOneFilePerRange Then
			currentProcessDate = Null

			Call CreateMailFile(colNOMBRE)
		Else
			Dim dateValue As Date

			For dateValue = startProcessDate To endProcessDate
				currentProcessDate = dateValue

				Call CreateMailFile(colNOMBRE)
			Next dateValue
		End If
		'If Not canMailBeSent Then Call AppendToLogsFile("El correo " & mailName & " no puede ser generado porque el reporte " & errorReport & " no trajo registros.")
		continueLoop:
	Next row
End Sub

Sub CreateMailFile(mailFileName As String)
	Call AppendToLogsFile("Generando archivo " & mailFileName & "...")

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

		'If Not canMailBeSent Then
		'	Workbook.Close False

		'	Exit Sub
		'End If
	Next item

	If Workbook.Worksheets.Count > 1 Then
		Workbook.Sheets("Sheet1").delete

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

	   Workbook.SaveAs fileName:=folder, FileFormat:=xlOpenXMLWorkbook

	   Call AppendToLogsFile("Archivo " & mailFileName & " creado exitosamente.")
	End If

	Workbook.Close False
End Sub

Sub CreateFileReport(Workbook As Workbook, fileReportName As String)
	Dim Worksheet As Worksheet
	Dim reportTable As ListObject
	Dim newTbl As ListObject
	Dim rowCount As Long
	Dim mailFile As String
	Dim mailName As String

	mailFile = CStr(ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""" & fileReportName & """, REPORTES[NOMBRE], REPORTES[ARCHIVO])"))
	mailName = CStr(ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""" & mailFile & """, ARCHIVOS[NOMBRE], ARCHIVOS[CORREO])"))

	ThisWorkbook.Activate

	Set reportTable = ThisWorkbook.Sheets(fileReportName).ListObjects(fileReportName)

	If Not IsNull(currentProcessDate) Then reportTable.Range.AutoFilter Field:=reportTable.ListColumns("PROCESS_DATE_FOR_RANGE").Index, Criteria1:=Format(currentProcessDate, "dd-MM-yyyy")

	rowCount = reportTable.ListRows.Count

	If rowCount = 0 Then
		'canMailBeSent = False

		Call AppendToLogsFile("El reporte " & fileReportName & " no trajo registros.")

		errorReport = fileReportName

		GoTo removeFilter
	End If

	Set Worksheet = Workbook.Worksheets.Add

	Worksheet.Name = fileReportName

	reportTable.Range.Resize(reportTable.ListRows.Count + 2, reportTable.ListColumns.Count - 1).Copy
	Worksheet.Range("A1").PasteSpecial Paste:=xlPasteFormats

	reportTable.Range.Resize(reportTable.ListRows.Count + 2, reportTable.ListColumns.Count - 1).Copy
	Worksheet.Range("A1").PasteSpecial Paste:=xlPasteValues

	reportTable.DataBodyRange.Delete

	Worksheet.Columns.AutoFit

removeFilter:
	reportTable.AutoFilter.ShowAllData
End Sub