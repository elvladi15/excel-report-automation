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

	Call AppendToLogsFile("Actualizando reportes...")
	ThisWorkbook.RefreshAll

	If executionMode = "MANUAL" Then
		MsgBox("Hojas de Excel actualizadas.")
	ElseIf executionMode = "AUTOMATIC"Then
		startProcessDate = CDate(CStr(PARAMETERS.Evaluate("XLOOKUP(""START_PROCESS_DATE"", PARAMETERS[NOMBRE], PARAMETERS[VALOR])")))
		endProcessDate = CDate(CStr(PARAMETERS.Evaluate("XLOOKUP(""END_PROCESS_DATE"", PARAMETERS[NOMBRE], PARAMETERS[VALOR])")))
	End If

	Exit Sub

	ErrorHandler:
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
			.Add "name", "PARAMETERS"
			.Add "columns", New Collection

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetParameterNameColumnName()
				.Add "rows", New Collection
					.Item("rows").Add GetParameterStartProcessDateName()
					.Item("rows").Add GetParameterEndProcessDateName()
					.Item("rows").Add GetParameterMaxTimeoutInSecondsName()
					.Item("rows").Add GetParameterFilesBaseFolderName()
					.Item("rows").Add GetParameterGenerateLogsName()
					.Item("rows").Add GetParameterLogFilesFolderName()
					.Item("rows").Add GetParameterOutlookFolderName()
					.Item("rows").Add GetParameterDateFormatName()
					.Item("rows").Add GetParameterScheduleTimeName()
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetParameterValueColumnName()
				.Add "rows", Null
			End With
		End With

		basicTableStructure("tables").Add CreateObject("Scripting.Dictionary")
		With basicTableStructure("tables")(basicTableStructure("tables").Count)
			.Add "name", "MAILS"
			.Add "columns", New Collection

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetMailNameColumnName()
				.Add "rows", Null
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetMailConversationColumnName()
				.Add "rows", Null
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetMailIsOneFilePerRangeColumnName()
				.Add "rows", Null
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetMailGenerateMailColumnName()
				.Add "rows", Null
			End With
		End With

		basicTableStructure("tables").Add CreateObject("Scripting.Dictionary")
		With basicTableStructure("tables")(basicTableStructure("tables").Count)
			.Add "name", "MAIL_FILES"
			.Add "columns", New Collection

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetMailFilesNameColumnName()
				.Add "rows", Null
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetMailFilesMailColumnName()
				.Add "rows", Null
			End With
		End With

		basicTableStructure("tables").Add CreateObject("Scripting.Dictionary")
		With basicTableStructure("tables")(basicTableStructure("tables").Count)
			.Add "name", "FILE_REPORTS"
			.Add "columns", New Collection

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetFileReportsNameColumnName()
				.Add "rows", Null
			End With

			.Item("columns").Add CreateObject("Scripting.Dictionary")
			With .Item("columns")(.Item("columns").Count)
				.Add "name", GetFileReportsFileColumnName()
				.Add "rows", Null
			End With
		End With

	Set GetBasicTableStructure = basicTableStructure
End Function

Function GetLanguageStructure() As Object
	Dim languageStructure As Object
	Set languageStructure = CreateObject("Scripting.Dictionary")
	
	Set languageStructure("languages") = New Collection
		languageStructure("languages").Add CreateObject("Scripting.Dictionary")

		With languageStructure("languages")(languageStructure("languages").Count)
			.Add "name", "SPANISH"
			.Add "languageNames", New Collection
				.Item("languageNames").Add CreateObject("Scripting.Dictionary")
				With .Item("languageNames")(.Item("languageNames").Count)
					.Add "language", "SPANISH"
					.Add "name", "Español"
				End With

				.Item("languageNames").Add CreateObject("Scripting.Dictionary")
				With .Item("languageNames")(.Item("languageNames").Count)
					.Add "language", "ENGLISH"
					.Add "name", "Inglés"
				End With
		End With

		languageStructure("languages").Add CreateObject("Scripting.Dictionary")
		With languageStructure("languages")(languageStructure("languages").Count)
			.Add "name", "ENGLISH"
			.Add "languageNames", New Collection
				.Item("languageNames").Add CreateObject("Scripting.Dictionary")
				With .Item("languageNames")(.Item("languageNames").Count)
					.Add "language", "SPANISH"
					.Add "name", "Spanish"
				End With

				.Item("languageNames").Add CreateObject("Scripting.Dictionary")
				With .Item("languageNames")(.Item("languageNames").Count)
					.Add "language", "ENGLISH"
					.Add "name", "English"
				End With
		End With

	Set GetlanguageStructure = languageStructure
End Function