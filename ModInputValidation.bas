Attribute VB_Name = "ModInputValidation"
Function IsInputValidationCorrect() As Boolean
	Set dictParameters = CreateObject("Scripting.Dictionary")

	If Not IsBasicTableStructureCorrent() Then Exit Function
	If Not IsParameterValidationCorrect() Then Exit Function
	If Not ValidateAllBasicTableContents() Then Exit Function
	If Not IsPowerQueryWorksheetAndTableValidationCorrect() Then Exit Function

	IsInputValidationCorrect = True
End Function

Function IsBasicTableStructureCorrent() As Boolean
	Dim tableObject As ListObject
	Dim columnObject As ListColumn
	
	Set basicTableStructure = GetBasicTableStructure()

	For Each table in basicTableStructure("tables")
		On Error Resume Next
		Set tableObject = PARAMETERS.ListObjects(table("name"))
		On Error GoTo 0

		If Err.Number <> 0 Then
			MsgBox "La tabla: '" & table("name") & "' no existe. Favor revisar nombres internos de las tablas."
			Exit Function
		End If

		For Each column in table("columns")
			On Error Resume Next
			Set columnObject = tableObject.ListColumns(column("name"))
			On Error GoTo 0

			If Err.Number <> 0 Then
				MsgBox "La columna: '" & column("name") & "' de la tabla: '" & table("name") & "' no existe. Favor revisar nombres."
				Exit Function
			End If

			If IsNull(column("rows")) Then Goto continueLoop

			For Each row in column("rows")
				If IsError(Application.Match(row, Range(table("name") & "[" & column("name") & "]"), 0)) Then
					MsgBox "El valor: '" & row & "', columna: '" & column("name") & "', tabla: '" & table("name") & "' no existe. Favor revisar nombres."
					Exit Function
				End If
			Next row

		continueLoop:
		Next column
	Next table

	IsBasicTableStructureCorrent = True
End Function

Function IsParameterValidationCorrect() As Boolean
	Dim nameColumn As String
	Dim valueColumn As String

	Dim nameParameterColumnName As String
	Dim valueParameterColumnName As String

	Dim startProcessDateParameterName As String
	Dim endProcessDateParameterName As String
	Dim maxTimeoutInSecondsParameterName As String
	Dim filesBaseFolderParameterName As String
	Dim generateLogsParameterName As String
	Dim logsFileFolderParameterName As String
	Dim outlookFolderParameterName As String
	Dim dateFormatParameterName As String
	Dim scheduleTimeParameterName As String


	nameParameterColumnName = GetNameParameterColumnName()
	valueParameterColumnName = GetValueParameterColumnName()

	startProcessDateParameterName = GetStartProcessDateParameterName()
	endProcessDateParameterName = GetEndProcessDateParameterName()
	maxTimeoutInSecondsParameterName = GetMaxTimeoutInSecondsParameterName()
	filesBaseFolderParameterName = GetFilesBaseFolderParameterName()
	generateLogsParameterName = GetGenerateLogsParameterName()
	logsFileFolderParameterName = GetLogFilesFolderParameterName()
	outlookFolderParameterName = GetOutlookFolderParameterName()
	dateFormatParameterName = GetDateFormatParameterName()
	scheduleTimeParameterName = GetScheduleTimeParameterName()

	For Each row In tbl_PARAMETERS.DataBodyRange.Rows
		nameColumn = row.Cells(1, tbl_PARAMETERS.ListColumns(nameParameterColumnName).Index).Value
		valueColumn = row.Cells(1, tbl_PARAMETERS.ListColumns(valueParameterColumnName).Index).Value

		dictParameters.Add nameColumn, valueColumn
	Next row

	For Each row In tbl_PARAMETERS.DataBodyRange.Rows
		nameColumn = row.Cells(1, tbl_PARAMETERS.ListColumns(nameParameterColumnName).Index).Value
		valueColumn = row.Cells(1, tbl_PARAMETERS.ListColumns(valueParameterColumnName).Index).Value

		If (nameColumn = startProcessDateParameterName Or nameColumn = endProcessDateParameterName) And Not IsDate(valueColumn) Then
			MsgBox "El valor del parámetro: '" & nameColumn & "' debe ser una fecha válida."
			Exit Function
		End If

		If nameColumn = maxTimeoutInSecondsParameterName And Not IsNumeric(valueColumn) Then
			MsgBox "El valor del parámetro: '" & nameColumn & "' debe ser un número."
			Exit Function
		End If

		If nameColumn = logsFileFolderParameterName And dictParameters(generateLogsParameterName) = Split(GetYesNoInCurrentLanguage(), ",")(1) Then GoTo continueLoop

		If valueColumn = "" Then
			MsgBox "El valor del parámetro: '" & nameColumn & "' no puede quedar vacío."
			Exit Function
		End If

		If nameColumn Like "Directorio*" Then
			If Dir(valueColumn, vbDirectory) = "" Then
				MsgBox "El directorio del parámetro: '" & nameColumn & "' no existe. Favor de validar ruta."
				Exit Function
			End If

			If Right(valueColumn, 1) = "\" Then
				MsgBox "El directorio del parámetro: '" & valueColumn & "' contiene el caracter \ al final. Favor de remover."
				Exit Function
			End If
		End If

		If nameColumn = scheduleTimeParameterName Then
			On Error Goto NotValidTime
			scheduleTime = TimeValue(valueColumn)
			GoTo continueLoop

			NotValidTime:
			MsgBox "La hora de ejecución: '" & valueColumn & "' no es una fecha válida."
			Exit Function
		End If

		continueLoop:
	Next row

	startProcessDate = CDate(dictParameters(startProcessDateParameterName))
	endProcessDate = CDate(dictParameters(endProcessDateParameterName))
	baseReportFolder = dictParameters(filesBaseFolderParameterName)
	logsFileFolder = dictParameters(logsFileFolderParameterName)
	outlookFolderName = dictParameters(outlookFolderParameterName)
	dateFormat = dictParameters(dateFormatParameterName)
	canGenerateLogs = dictParameters(generateLogsParameterName) = Split(GetYesNoInCurrentLanguage(), ",")(0)
	scheduleTime = TimeValue(dictParameters(scheduleTimeParameterName))

	IsParameterValidationCorrect = True
End Function

Function ValidateAllBasicTableContents() As Boolean
	If Not ValidateBasicTableContent(tbl_MAILS) Then Exit Function
	If Not ValidateBasicTableContent(tbl_MAIL_FILES) Then Exit Function
	If Not ValidateBasicTableContent(tbl_FILE_REPORTS) Then Exit Function

	ValidateAllBasicTableContents = True
End Function

Function ValidateBasicTableContent(table As ListObject)
	Dim atLeast1MailToGenerate As Boolean

	atLeast1MailToGenerate = False

	If table.ListRows.Count = 0 Then
		MsgBox "La tabla: '" & table.Name & "' está vacía."
		Exit Function
	End If

	For Each column in table.ListColumns
		For Each cell in column.DataBodyRange
			If cell.Value = "" Then
				MsgBox "Hay valores vacíos en la tabla: '" & table.Name & "'."
				Exit Function
			End If

			If table.Name = "CORREOS" Then
				If column.Name = "UN ARCHIVO POR RANGO?" Then
					GoTo continueLoop
				End If

				If column.Name = "GENERAR CORREO?" Then
					If cell.Value = Split(GetYesNoInCurrentLanguage(), ",")(0) Then
						atLeast1MailToGenerate = True
					End If

					GoTo continueLoop
				End If

			End If

			If (table.Name = "ARCHIVOS" And column.Name = "CORREO") Or table.Name = "REPORTES" Then GoTo continueLoop

			If Application.CountIf(column.DataBodyRange, cell.Value) > 1 Then
				MsgBox "Hay valores duplicados en la columna: '" & column.Name & "' de la tabla: '" & table.Name & "'."
				Exit Function
			End If

			If table.Name = "CORREOS" And column.Name = "NOMBRE" Then
				For Each mailName In tbl_MAIL_FILES.ListColumns("CORREO").DataBodyRange
					If mailName.Value = cell.Value Then Goto continueLoop
				Next mailName

				MsgBox "El correo: '" & cell.Value & "' no tiene ningún archivo asociado."

				Exit Function
			End If
			
			If table.Name = "ARCHIVOS" And column.Name = "NOMBRE" Then
				For Each mailFileName in tbl_FILE_REPORTS.ListColumns("ARCHIVO").DataBodyRange
					If mailFileName.Value = cell.Value Then Goto continueLoop
				Next mailFileName

				MsgBox "El archivo: '" & cell.Value & "' no tiene ningún reporte asociado."

				Exit Function
			End If
			continueLoop:
		Next cell
	Next column
	
	If table.Name = "CORREOS" And Not atLeast1MailToGenerate Then
		MsgBox "Debe haber al menos 1 correo a generar."
		Exit Function
	End If

	ValidateBasicTableContent = True
End Function

Function IsPowerQueryWorksheetAndTableValidationCorrect() As Boolean
	Dim Worksheet As Worksheet
	Dim table As ListObject
	Dim columnExists As Boolean
	Dim nameColumn As String

	For Each row In tbl_FILE_REPORTS.DataBodyRange.Rows
		nameColumn = row.Cells(1, tbl_FILE_REPORTS.ListColumns("NOMBRE").Index).Value

		On Error Resume Next
		Set Worksheet = ThisWorkbook.Worksheets(nameColumn)
		On Error GoTo 0
		If Err.Number <> 0 Then
			MsgBox "La hoja de cálculo: '" & nameColumn & "' no existe. Favor crearla junto a su tabla de Power Query."
			Exit Function
		End If

		On Error Resume Next
		Set table = Worksheet.ListObjects(nameColumn)
		On Error GoTo 0
		If Err.Number <> 0 Then
			MsgBox "La tabla: '" & nameColumn & "' no fue encontrada en su respectiva hoja de cálculo. Favor crear."
			Exit Function
		End If
	Next row
	IsPowerQueryWorksheetAndTableValidationCorrect = True
End Function

Function IsConversationColumnCorrect() As Boolean
	Dim conversationColumn As String

	Set outlookAppRef = CreateObject("Outlook.Application").GetNamespace("MAPI")
	Set outlookReportFolderRef = outlookAppRef.GetDefaultFolder(6).Parent.Folders(outlookFolderName)
	Set outlookDraftsFolderRef = outlookAppRef.GetDefaultFolder(16)

	For each conversation in tbl_MAILS.ListColumns(GetConversationMailColumnName()).DataBodyRange.Cells
		If Not outlookReportFolderRef.Items.Restrict("[Subject] = '" & conversation.Value & "'").Count > 0 Then
			MsgBox "La conversación: '" & conversation.Value & "' no existe."
			Exit Function
		End If
	Next conversation

	IsConversationColumnCorrect = True
End Function