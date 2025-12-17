Attribute VB_Name = "ModFileGeneration"
Sub Manual_CreateMailFiles()
    Call CreateMailFiles("MANUAL")
End Sub
Sub Automatic_CreateMailFiles()
    Call CreateMailFiles("AUTOMATICO")
End Sub
Sub CreateMailFiles(executionMode As String)
    If isInputValidationCorrect = False Then Exit Sub
    
	InitializeGlobals
	
	Dim mails As Variant
	
	mails = Range("CORREOS[NOMBRE]")
	
	If selectedReport = "Todos" Then
		For Each item In mails
			Call CreateMail(CStr(item))
		Next item
		
		Exit Sub
	End If
	
	Call CreateMail(selectedReport)

    If executionMode = "MANUAL" Then MsgBox "Archivos creados correctamente."
End Sub

Sub CreateMail(mailName As String)
    If isInputValidationCorrect = False Then Exit Sub

	InitializeGlobals
	
	Dim mailFiles As Variant
	Dim mailFileCount As Long
	Dim isOneFilePerRange As Boolean

	Application.DisplayAlerts = False
	
	mailFiles = ThisWorkbook.ActiveSheet.Evaluate("FILTER(ARCHIVOS[NOMBRE], ARCHIVOS[CORREO] = """ & mailName & """)")
	
	isOneFilePerRange = ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""" & mailName & """, CORREOS[NOMBRE], CORREOS[UN ARCHIVO POR RANGO?])") = "SI"
	
	If Dir(baseReportFolder & "\" & mailName, vbDirectory) = "" Then MkDir baseReportFolder & "\" & mailName
	
	mailFileCount = UBound(mailFiles) - LBound(mailFiles) + 1
	
	For Each item In mailFiles
		If isOneFilePerRange Then
			currentProcessDate = Null
			
			Call CreateMailFile(CStr(item), baseReportFolder & "\" & mailName, mailFileCount)
		Else
			Dim dateValue As Date
			
			For dateValue = startProcessDate To endProcessDate
				currentProcessDate = dateValue
				
				Call CreateMailFile(CStr(item), baseReportFolder & "\" & mailName, mailFileCount)
			Next dateValue
		End If
		If canMailBeSent = False Then
			Call AppendToLogsFile("El correo " & mailName & " no puede ser generado porque el reporte " & errorReport & " no trajo registros.")
			
			Exit Sub
		End If
	Next item
End Sub

Sub CreateMailFile(mailFileName As String, folder As String, mailFileCount As Long)
	Call AppendToLogsFile("Generando archivo " & mailFileName & "...")
	
	Dim Workbook As Workbook
	Dim fileReports As Variant
	Dim reportDate As String
	
	Set Workbook = Workbooks.Add
	
	fileReports = ThisWorkbook.ActiveSheet.Evaluate("FILTER(REPORTES[NOMBRE], REPORTES[ARCHIVO] = """ & mailFileName & """)")
	
	For Each item In fileReports
		Call CreateFileReport(Workbook, CStr(item))
		
		If canMailBeSent = False Then
			Workbook.Close False
			
			Exit Sub
		End If
	Next item
	
	If Workbook.Worksheets.Count > 1 Then
        Workbook.Sheets("Sheet1").delete

		For Each Query In Workbook.Queries
			Query.delete
		Next Query
		
		If mailFileCount > 1 Then
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
	
	ThisWorkbook.Activate
	
	Set reportTable = ThisWorkbook.Sheets(fileReportName).ListObjects(fileReportName)
	
	If IsNull(currentProcessDate) Then
		rowCount = Application.WorksheetFunction.Subtotal(103, reportTable.ListColumns(1).DataBodyRange)
	Else
		reportTable.Range.AutoFilter Field:=reportTable.ListColumns("PROCESS_DATE_FOR_RANGE").Index, Criteria1:=Format(currentProcessDate, "dd-MM-yyyy")
		rowCount = reportTable.ListRows.Count
	End If
   
	If rowCount = 0 Then
		canMailBeSent = False
		
		errorReport = fileReportName
		
		GoTo removeFilter
	End If
	
	Set Worksheet = Workbook.Worksheets.Add
	
	Worksheet.Name = fileReportName
	
	reportTable.Range.Resize(reportTable.ListRows.Count + 2, reportTable.ListColumns.Count - 1).Copy
	Worksheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
	
	reportTable.Range.Resize(reportTable.ListRows.Count + 2, reportTable.ListColumns.Count - 1).Copy
	Worksheet.Range("A1").PasteSpecial Paste:=xlPasteValues

	Worksheet.Columns.AutoFit
	
removeFilter:
	reportTable.AutoFilter.ShowAllData
End Sub