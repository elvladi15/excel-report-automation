Attribute VB_Name = "ModUtils"
Sub CloseAllOtherWorkbooks()
	Dim wb As Workbook, keep As Workbook
	Set keep = ThisWorkbook

	Application.DisplayAlerts = False
	On Error GoTo CleanUp

	For Each wb In Application.Workbooks
		If Not wb Is keep Then
			wb.Close SaveChanges:=False
		End If
	Next wb

CleanUp:
	   Application.DisplayAlerts = True
End Sub

Sub Manual_RefreshAll()
    Call RefreshAll("MANUAL")
End Sub
Sub Automatic_RefreshAll()
    Call RefreshAll("AUTOMATICO")
End Sub
Sub RefreshAll(executionMode As String)
	If isInputValidationCorrect = False Then Exit Sub
	
	Call AppendToLogsFile("Cerrando los demás libros de Excel...")
	CloseAllOtherWorkbooks
	
	Call AppendToLogsFile("Refrescando hoja de cálculo...")
	ThisWorkbook.Sheets("PARAMETROS").Calculate
	
	Call AppendToLogsFile("Actualizando reportes...")
	ThisWorkbook.RefreshAll

    If(executionMode = "MANUAL") Then MsgBox("Hojas de Excel actualizadas.")
End Sub

Sub AppendToLogsFile(message As String)
    If canGenerateLogs = False Then Exit Sub

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    With fso.OpenTextFile(logsFileFolder & "\" & "Logs " & Format(Date, dateFormat) & ".txt", 8, True)
        .WriteLine Format(Now, "yyyy-MM-dd hh:mm:ss - ") & message
        .Close
    End With
End Sub

Sub AddValidationFromTableColumn()
	Dim cell As Range
	Dim listText As String
	
	Set ws = ThisWorkbook.Sheets("PARAMETROS")
	Set tbl = ws.ListObjects("PARAMETROS")
	Set cell = Range("PARAMETROS[VALOR]").Cells(Evaluate("MATCH(""Reporte a generar"",PARAMETROS[NOMBRE],0)"))
	
	listText = "Todos,"
	
	For Each item In Range("CORREOS[NOMBRE]")
		listText = listText & item.Value & ","
	Next item
	
	listText = Left(listText, Len(listText) - 1)
	
	With cell.Validation
		.delete
		.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listText
		.IgnoreBlank = True
		.InCellDropdown = True
	End With
End Sub

Sub OpenOutlookIfNotRunning()
	Dim outlook As Object

	On Error GoTo OpenOutlook
		Set outlook = GetObject(, "Outlook.Application")

		Exit Sub
	OpenOutlook:
		Shell """outlook.exe""", vbNormalFocus
End Sub