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

Sub RefreshAll()
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

Function GetBasicTableStructure() As Object
	Dim basicTableStructure As Object
	Set basicTableStructure = CreateObject("Scripting.Dictionary")
	
	Set basicTableStructure("tables") = New Collection

		basicTableStructure("tables").Add CreateObject("Scripting.Dictionary")
		With basicTableStructure("tables")(basicTableStructure("tables").Count)
			.Add "key", "PARAMETROS"
			.Add "columns", New Collection

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "key", "NOMBRE"
				.Add "rows", New Collection

					.Item("rows").Add CreateObject("Scripting.Dictionary")
						.Item("rows")(.Item("rows").Count).Add "key", "START_PROCESS_DATE"
						.Item("rows")(.Item("rows").Count).Add "value", Null

					.Item("rows").Add CreateObject("Scripting.Dictionary")
						.Item("rows")(.Item("rows").Count).Add "key", "END_PROCESS_DATE"
						.Item("rows")(.Item("rows").Count).Add "value", Null

					.Item("rows").Add CreateObject("Scripting.Dictionary")
						.Item("rows")(.Item("rows").Count).Add "key", "Directorio base reportes"
						.Item("rows")(.Item("rows").Count).Add "value", Null

					.Item("rows").Add CreateObject("Scripting.Dictionary")
						.Item("rows")(.Item("rows").Count).Add "key", "Generar logs?"
						.Item("rows")(.Item("rows").Count).Add "value", Null

					.Item("rows").Add CreateObject("Scripting.Dictionary")
						.Item("rows")(.Item("rows").Count).Add "key", "Directorio archivos de logs"
						.Item("rows")(.Item("rows").Count).Add "value", Null

					.Item("rows").Add CreateObject("Scripting.Dictionary")
						.Item("rows")(.Item("rows").Count).Add "key", "Carpeta de Outlook"
						.Item("rows")(.Item("rows").Count).Add "value", Null

					.Item("rows").Add CreateObject("Scripting.Dictionary")
						.Item("rows")(.Item("rows").Count).Add "key", "Formato de fechas"
						.Item("rows")(.Item("rows").Count).Add "value", Null
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "key", "VALOR"
				.Add "rows", New Collection
			End With
		End With

	Set GetBasicTableStructure = basicTableStructure
End Function

Function GetBasicTableStructureTableIndex(key As String) As Long
	Dim i As Long

	For i = 1 To arr.Count
		If(basicTableStructure("tables")(i)("key") = key) Then
			GetBasicTableStructureTableIndex = i
			Exit Function
		End If
	Next
End Function

Function GetBasicTableStructureColumnIndex(tableIndex As Long, key As String) As Long
	Dim i As Long

	For i = 1 To arr.Count
		If(basicTableStructure("tables")(tableIndex)("columns")(i)("key") = key) Then
			GetBasicTableStructureColumnIndex = i
			Exit Function
		End If
	Next
End Function

Function GetBasicTableStructureRowIndex(tableIndex As Long, columnIndex As Long, key As String) As Long
	Dim i As Long

	For i = 1 To arr.Count
		If(basicTableStructure("tables")(tableIndex)("columns")(columnIndex)("rows")(i)("key") = key) Then
			GetBasicTableStructureRowIndex = i
			Exit Function
		End If
	Next
End Function