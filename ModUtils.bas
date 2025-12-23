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
	On Error Goto ErrorHandler

	Call AppendToLogsFile("Cerrando los demás libros de Excel...")
	CloseAllOtherWorkbooks

	Call AppendToLogsFile("Refrescando hoja de cálculo...")
	ThisWorkbook.Sheets("PARAMETROS").Calculate

	Call AppendToLogsFile("Actualizando reportes...")
	ThisWorkbook.RefreshAll

	If executionMode = "MANUAL" Then
		MsgBox("Hojas de Excel actualizadas.")
	ElseIf executionMode = "AUTOMÁTICO"Then
		Set wsPARAMETROS = ThisWorkbook.Sheets("PARAMETROS")
		startProcessDate = CDate(CStr(ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""START_PROCESS_DATE"", PARAMETROS[NOMBRE], PARAMETROS[VALOR])")))
		endProcessDate = CDate(CStr(ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""END_PROCESS_DATE"", PARAMETROS[NOMBRE], PARAMETROS[VALOR])")))
	End If

	Exit Sub
	ErrorHandler:
		continueExecution = False
End Sub

Sub AppendToLogsFile(message As String)
	If Not canGenerateLogs Then Exit Sub

	Dim fso As Object

	Set fso = CreateObject("Scripting.FileSystemObject")

	With fso.OpenTextFile(logsFileFolder & "\" & "Logs " & Format(Date, dateFormat) & ".txt", 8, True)
		.WriteLine Format(Now, "yyyy-MM-dd hh:mm:ss - ") & message
		.Close
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
			.Add "name", "PARAMETROS"
			.Add "columns", New Collection

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "NOMBRE"
				.Add "rows", New Collection
					.Item("rows").Add "START_PROCESS_DATE"
					.Item("rows").Add "END_PROCESS_DATE"
					.Item("rows").Add "Directorio base reportes"
					.Item("rows").Add "Generar logs?"
					.Item("rows").Add "Directorio archivos de logs"
					.Item("rows").Add "Carpeta de Outlook"
					.Item("rows").Add "Formato de fechas"
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "VALOR"
				.Add "rows", New Collection
			End With
		End With

		basicTableStructure("tables").Add CreateObject("Scripting.Dictionary")
		With basicTableStructure("tables")(basicTableStructure("tables").Count)
			.Add "name", "CORREOS"
			.Add "columns", New Collection

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "NOMBRE"
				.Add "rows", New Collection
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "CONVERSACION"
				.Add "rows", New Collection
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "UN ARCHIVO POR RANGO?"
				.Add "rows", New Collection
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "GENERAR CORREO?"
				.Add "rows", New Collection
			End With
		End With

		basicTableStructure("tables").Add CreateObject("Scripting.Dictionary")
		With basicTableStructure("tables")(basicTableStructure("tables").Count)
			.Add "name", "ARCHIVOS"
			.Add "columns", New Collection

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "NOMBRE"
				.Add "rows", New Collection
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "CORREO"
				.Add "rows", New Collection
			End With
		End With

		basicTableStructure("tables").Add CreateObject("Scripting.Dictionary")
		With basicTableStructure("tables")(basicTableStructure("tables").Count)
			.Add "name", "REPORTES"
			.Add "columns", New Collection

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "NOMBRE"
				.Add "rows", New Collection
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", "ARCHIVO"
				.Add "rows", New Collection
			End With
		End With

	Set GetBasicTableStructure = basicTableStructure
End Function