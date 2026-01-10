Attribute VB_Name = "ModFileGeneration"
Sub CreateMailFiles()
	Dim fileGenerated As Boolean
	Dim generateReportColumn As String
	Dim nameColumn As String
	Dim outputMesssage As String

	outputMesssage = ""

	For Each mailName In PARAMETERS.Evaluate("FILTER(MAILS[" & GetMailNameColumnName() & "], MAILS[" & GetMailGenerateMailColumnName() & "] = """ & Split(GetYesNoInCurrentLanguage(), ",")(0) & """)")
		Call CreateMail(CStr(mailName))
	Next mailName

	If executionMode = "MANUAL" Then
		If reportsNotGenerated.Count = 0 Then
			outputMesssage = outputMesssage & "Reportes generados exitosamente. "
		Else
			outputMesssage = "Los reportes:" & vbCrLf & vbCrLf

			For Each report In reportsNotGenerated
				outputMesssage = outputMesssage & report & vbCrLf
			Next report
			
			outputMesssage = outputMesssage & vbCrLf

			outputMesssage = outputMesssage & " no se pudieron generar." & vbCrLf & vbCrLf
		End If

		If mailFilesNotGenerated.Count = 0 Then
			outputMesssage = outputMesssage & "Archivos creados exitosamente."
		Else
			outputMesssage = outputMesssage & "Los archivos:" & vbCrLf & vbCrLf

			For Each mailFile In mailFilesNotGenerated
				outputMesssage = outputMesssage & mailFile & vbCrLf
			Next mailFile

			outputMesssage = outputMesssage & vbCrLf

			outputMesssage = outputMesssage & " no se pudieron crear porque no tenían ningún reporte."
		End If

		MsgBox outputMesssage
	End If

End Sub

Sub CreateMail(mailName As String)
	Dim mailFiles As Variant
	Dim mailFileCount As Long
	Dim isOneFilePerRange As Boolean

	isOneFilePerRange = PARAMETERS.Evaluate("XLOOKUP(""" & mailName & """, MAILS[" & GetMailNameColumnName() & "], MAILS[" & GetMailIsOneFilePerRangeColumnName() & "])") = Split(GetYesNoInCurrentLanguage(), ",")(0)

	If Dir(baseReportFolder & "\" & mailName, vbDirectory) = "" Then MkDir baseReportFolder & "\" & mailName

	mailFileCount = Application.WorksheetFunction.CountIf(tbl_MAIL_FILES.ListColumns(GetMailFilesMailColumnName()).DataBodyRange, mailName)

	For Each mailFileName In PARAMETERS.Evaluate("FILTER(MAIL_FILES[" & GetMailFilesNameColumnName() & "], MAIL_FILES[" & GetMailFilesMailColumnName() & "] = """ & mailName & """)")
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
End Sub

Sub CreateMailFile(mailFileName As String)
	On Error Goto ErrorHandler
	Call AppendToLogsFile("Generando archivo: '" & mailFileName & "'...")

	Dim Workbook As Workbook
	Dim fileReports As Variant
	Dim reportDate As String
	Dim mailName As String
	Dim folder As String
	Dim fileQuantityPerMail As Long

	mailName = CStr(PARAMETERS.Evaluate("XLOOKUP(""" & mailFileName & """, MAIL_FILES[" & GetMailFilesNameColumnName() & "], MAIL_FILES[" & GetMailFilesMailColumnName() & "])"))
	folder = baseReportFolder & "\" & mailName
	fileQuantityPerMail = Application.WorksheetFunction.CountIf(tbl_MAIL_FILES.ListColumns(GetMailFilesMailColumnName()).DataBodyRange, mailName)

	Set Workbook = Workbooks.Add

	fileReports = PARAMETERS.Evaluate("FILTER(FILE_REPORTS[" & GetFileReportsNameColumnName() & "], FILE_REPORTS[" & GetFileReportsFileColumnName() & "] = """ & mailFileName & """)")

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
		mailFilesNotGenerated.Add mailFileName

		Call AppendToLogsFile("El archivo: '" & mailFileName & "' no pudo ser creado porque no se generó ningún reporte.")
	End If

	Workbook.Close False

	Exit Sub

	ErrorHandler:
	Call AppendToLogsFile("Ha ocurrido un error al generar el archivo " & mailFileName & ".")
End Sub

Sub CreateFileReport(Workbook As Workbook, fileReportName As String)
	Dim Worksheet As Worksheet
	Dim reportTable As ListObject
	Dim newTbl As ListObject
	Dim mailFile As String
	Dim mailName As String

	On Error Goto ErrorHandler

	ThisWorkbook.Activate

	Set reportTable = ThisWorkbook.Sheets(fileReportName).ListObjects(fileReportName)

	If reportTable.ListRows.Count = 1 And reportTable.ListColumns.Count = 1 Then
		Call AppendToLogsFile("Hubo un error al consultar el reporte " & fileReportName & " desde la base de datos.")

		reportsNotGenerated.Add fileReportName

		Exit Sub
	End If

	If reportTable.ListRows.Count = 0 Then
		Call AppendToLogsFile("El reporte " & fileReportName & " no trajo registros.")

		reportsNotGenerated.Add fileReportName

		Exit Sub
	End If

	reportTable.DataBodyRange.Borders.LineStyle = xlContinuous

	On Error Goto no_PROCESS_DATE_FOR_RANGE_column
	reportTable.ListColumns("PROCESS_DATE_FOR_RANGE").DataBodyRange.NumberFormat = tbl_PARAMETERS.ListRows(2).Range.Cells(1, 2).NumberFormat
	
	If Not IsNull(currentProcessDate) Then reportTable.Range.AutoFilter Field:=reportTable.ListColumns("PROCESS_DATE_FOR_RANGE").Index, Criteria1:=Format(currentProcessDate, dateFormat)


	If Application.WorksheetFunction.CountA(reportTable.DataBodyRange) = 0 Then
		Call AppendToLogsFile("El reporte " & fileReportName & " no se actualizó.")

		reportsNotGenerated.Add fileReportName

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

	Worksheet.Columns.AutoFit

	removeFilter:
		reportTable.AutoFilter.ShowAllData
		Exit Sub
	no_PROCESS_DATE_FOR_RANGE_column:
		Call AppendToLogsFile("No se encontró la columna PROCESS_DATE_FOR_RANGE en el reporte " & fileReportName & ".")
		Exit Sub
	ErrorHandler:
		Call AppendToLogsFile("Ha ocurrido un error al generar el reporte " & fileReportName & ".")
		Exit Sub
End Sub