Attribute VB_Name = "ModInputValidation"
Function IsInputValidationCorrect() As Boolean
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
			MsgBox InputValidationTableNotExistsMessage(CStr(table("name")))
			Exit Function
		End If

		For Each column in table("columns")
			On Error Resume Next
			Set columnObject = Nothing
			Set columnObject = tableObject.ListColumns(column("name"))
			On Error GoTo 0

			If columnObject Is Nothing Then
				MsgBox InputValidationColumnNotExistsMessage(CStr(column("name")), CStr(table("name")))
				Exit Function
			End If

			If IsNull(column("rows")) Then Goto continueLoop

			For Each row in column("rows")
				On Error Resume Next
				Dim matchResult As Variant
				matchResult = Application.Match(row, Range(table("name") & "[" & column("name") & "]"), 0)

				If IsError(matchResult) Then
					MsgBox InputValidationValueNotExistsMessage(CStr(row), CStr(column("name")), CStr(table("name")))
					Exit Function
				End If
			Next row

		continueLoop:
		Next column
	Next table

	IsBasicTableStructureCorrent = True
End Function

Function IsParameterValidationCorrect() As Boolean
	Dim parameterName As String
	Dim parameterValue As String

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


	nameParameterColumnName = tbl_PARAMETERS.ListColumns(1).Name
	valueParameterColumnName = tbl_PARAMETERS.ListColumns(2).Name

	startProcessDateParameterName = tbl_PARAMETERS.ListRows(2).Range.Cells(1).Value
	endProcessDateParameterName = tbl_PARAMETERS.ListRows(3).Range.Cells(1).Value
	maxTimeoutInSecondsParameterName = tbl_PARAMETERS.ListRows(4).Range.Cells(1).Value
	filesBaseFolderParameterName = tbl_PARAMETERS.ListRows(5).Range.Cells(1).Value
	generateLogsParameterName = tbl_PARAMETERS.ListRows(6).Range.Cells(1).Value
	logsFileFolderParameterName = tbl_PARAMETERS.ListRows(7).Range.Cells(1).Value
	outlookFolderParameterName = tbl_PARAMETERS.ListRows(8).Range.Cells(1).Value
	dateFormatParameterName = tbl_PARAMETERS.ListRows(9).Range.Cells(1).Value
	scheduleTimeParameterName = tbl_PARAMETERS.ListRows(10).Range.Cells(1).Value

	For Each row In tbl_PARAMETERS.ListRows
		parameterName = row.Range.Cells(1).Value
		parameterValue = row.Range.Cells(2).Value

		If (parameterName = startProcessDateParameterName Or parameterName = endProcessDateParameterName) And Not IsDate(parameterValue) Then
			MsgBox InputValidationParameterMustBeValidDateMessage(parameterName)
			Exit Function
		End If

		If parameterName = maxTimeoutInSecondsParameterName And Not IsNumeric(parameterValue) Then
			MsgBox InputValidationParameterMustBeNumberMessage(parameterName)
			Exit Function
		End If

		If parameterName = logsFileFolderParameterName And tbl_PARAMETERS.ListRows(6).Range.Cells(2).Value = Split(tbl_PARAMETERS.ListRows(6).Range.Cells(2).Validation.Formula1, ",")(1) Then GoTo continueLoop

		If parameterValue = "" Then
			MsgBox InputValidationParameterCannotBeEmptyMessage(parameterName)
			Exit Function
		End If

		If parameterName = logsFileFolderParameterName Or parameterName = filesBaseFolderParameterName Then
			If Dir(parameterValue, vbDirectory) = "" Then
				MsgBox InputValidationParameterDirectoryNotExistsMessage(parameterName)
				Exit Function
			End If

			If Right(parameterValue, 1) = "\" Then
				MsgBox InputValidationParameterDirectoryEndsWithSlashMessage(parameterValue)
				Exit Function
			End If
		End If

		If parameterName = scheduleTimeParameterName Then
			On Error Goto NotValidTime
			scheduleTime = TimeValue(parameterValue)
			GoTo continueLoop

			NotValidTime:
			MsgBox InputValidationExecutionTimeNotValidMessage(parameterValue)
			Exit Function
		End If

		continueLoop:
	Next row

	startProcessDate = CDate(tbl_PARAMETERS.ListRows(2).Range.Cells(2).Value)
	endProcessDate = CDate(tbl_PARAMETERS.ListRows(3).Range.Cells(2).Value)
	baseReportFolder = tbl_PARAMETERS.ListRows(5).Range.Cells(2).Value
	canGenerateLogs = tbl_PARAMETERS.ListRows(6).Range.Cells(2).Value = Split(GetYesNoInCurrentLanguage(), ",")(0)
	logsFileFolder = tbl_PARAMETERS.ListRows(7).Range.Cells(2).Value
	outlookFolderName = tbl_PARAMETERS.ListRows(8).Range.Cells(2).Value
	dateFormat = tbl_PARAMETERS.ListRows(9).Range.Cells(2).Value
	scheduleTime = TimeValue(tbl_PARAMETERS.ListRows(10).Range.Cells(2).Value)

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
		MsgBox InputValidationTableIsEmptyMessage(table.Name)
		Exit Function
	End If

	For Each column in table.ListColumns
		For Each cell in column.DataBodyRange
			If cell.Value = "" Then
				MsgBox InputValidationTableHasEmptyValuesMessage(table.Name)
				Exit Function
			End If

			If table.Name = "MAILS" Then
				If column.Name = tbl_MAILS.ListColumns(3).Name Then
					GoTo continueLoop
				End If

				If column.Name = tbl_MAILS.ListColumns(4).Name Then
					If cell.Value = Split(GetYesNoInCurrentLanguage(), ",")(0) Then
						atLeast1MailToGenerate = True
					End If

					GoTo continueLoop
				End If

			End If

			If (table.Name = "MAIL_FILES" And column.Name = tbl_MAIL_FILES.ListColumns(2).Name) Or table.Name = "FILE_REPORTS" Then GoTo continueLoop

			If Application.CountIf(column.DataBodyRange, cell.Value) > 1 Then
				MsgBox InputValidationColumnHasDuplicatesMessage(column.Name, table.Name)
				Exit Function
			End If

			If table.Name = "MAILS" And column.Name = tbl_MAILS.ListColumns(1).Name Then
				For Each mailName In tbl_MAIL_FILES.ListColumns(2).DataBodyRange
					If mailName.Value = cell.Value Then Goto continueLoop
				Next mailName

				MsgBox InputValidationEmailHasNoFilesMessage(cell.Value)

				Exit Function
			End If

			If table.Name = "MAIL_FILES" And column.Name = tbl_FILE_REPORTS.ListColumns(1).Name Then
				For Each mailFileName in tbl_FILE_REPORTS.ListColumns(2).DataBodyRange
					If mailFileName.Value = cell.Value Then Goto continueLoop
				Next mailFileName

				MsgBox InputValidationFileHasNoReportMessage(cell.Value)

				Exit Function
			End If
			continueLoop:
		Next cell
	Next column

	If table.Name = "MAILS" And Not atLeast1MailToGenerate Then
		MsgBox InputValidationAtLeastOneEmailMessage()
		Exit Function
	End If

	ValidateBasicTableContent = True
End Function

Function IsPowerQueryWorksheetAndTableValidationCorrect() As Boolean
	Dim Worksheet As Worksheet
	Dim table As ListObject
	Dim connection As WorkbookConnection
	Dim nameColumn As String

	For Each row In tbl_FILE_REPORTS.ListRows
		nameColumn = row.Range.Cells(1).Value

		On Error Resume Next
		Set Worksheet = ThisWorkbook.Worksheets(nameColumn)
		On Error GoTo 0
		If Err.Number <> 0 Then
			MsgBox InputValidationWorksheetNotExistsMessage(nameColumn)
			Exit Function
		End If

		On Error Resume Next
		Set table = Worksheet.ListObjects(nameColumn)
		On Error GoTo 0
		If Err.Number <> 0 Then
			MsgBox InputValidationTableNotFoundOnSheetMessage(nameColumn)
			Exit Function
		End If

		Set connection = Nothing
		On Error Resume Next
		Set connection = ThisWorkbook.Connections("Query - " & nameColumn)
		On Error GoTo 0
		If connection Is Nothing Then
			MsgBox FileGenerationQueryNotFoundMessage(nameColumn)
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

	For each conversation in tbl_MAILS.ListColumns(GetMailConversationColumnName()).DataBodyRange.Cells
		If Not outlookReportFolderRef.Items.Restrict("[Subject] = '" & conversation.Value & "'").Count > 0 Then
			MsgBox InputValidationConversationNotExistsMessage(conversation.Value)
			Exit Function
		End If
	Next conversation

	IsConversationColumnCorrect = True
End Function