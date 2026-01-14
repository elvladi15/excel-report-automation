Attribute VB_Name = "ModLanguageSupport"
Sub UpdateApplicationLanguage()
    Dim isOneFilePerRangeMailColumnName As String
    Dim generateMailColumnName As String

    Dim mailFilesNameColumnName As String
    Dim mailFilesMailColumnName As String
    
    Dim fileReportsFileColumnName As String

    Dim previousYesNoInCurrentLanguage As String
    Dim currentYesNoInCurrentLanguage As String

    previousYesNoInCurrentLanguage = GetYesNoInCurrentLanguage()

    currentLanguage = GetLanguageByLanguageName(tbl_PARAMETERS.ListRows(1).Range.Cells(2).Value)

    currentYesNoInCurrentLanguage = GetYesNoInCurrentLanguage()

    PARAMETERS.Name =  GetParameterWorksheetName()

    'BUTTON CAPTIONS
    PARAMETERS.Buttons("btnRefreshAll").Caption = GetBtnRefreshAllCaption()
    PARAMETERS.Buttons("btnCreateMailFiles").Caption = GetBtnCreateMailFilesCaption()
    PARAMETERS.Buttons("btnCreateDrafts").Caption = GetBtnCreateDraftsCaption()
    PARAMETERS.Buttons("btnSendAllDrafts").Caption = GetBtnSendAllDraftsCaption()
    PARAMETERS.Buttons("btnScheduleFileGeneration").Caption = GetBtnScheduleFileGenerationCaption()
    PARAMETERS.Buttons("btnScheduleMailSending").Caption = GetBtnScheduleMailSendingCaption()

    'PARAMETERS TABLE
    tbl_PARAMETERS.HeaderRowRange.Columns(1).Offset(-1, 0).Value = GetParameterTableName()

    tbl_PARAMETERS.ListColumns(1).Name = GetParameterNameColumnName()
    tbl_PARAMETERS.ListColumns(2).Name = GetParameterValueColumnName()

    isSilentChange = True
    tbl_PARAMETERS.ListRows(1).Range.Cells(2).Value = GetLanguageNameByLanguage()
    isSilentChange = False

    currentLanguageName = tbl_PARAMETERS.ListRows(1).Range.Cells(2).Value

    With tbl_PARAMETERS.ListRows(1).Range.Cells(2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=GetAllLanguageNamesString()
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    tbl_PARAMETERS.ListRows(1).Range.Cells(1).Value = GetParameterApplicationLanguageName()
    tbl_PARAMETERS.ListRows(2).Range.Cells(1).Value = GetParameterStartProcessDateName()
    tbl_PARAMETERS.ListRows(3).Range.Cells(1).Value = GetParameterEndProcessDateName()
    tbl_PARAMETERS.ListRows(4).Range.Cells(1).Value = GetParameterMaxTimeoutInSecondsName()
    tbl_PARAMETERS.ListRows(5).Range.Cells(1).Value = GetParameterFilesBaseFolderName()
    tbl_PARAMETERS.ListRows(6).Range.Cells(1).Value = GetParameterGenerateLogsName()
    tbl_PARAMETERS.ListRows(7).Range.Cells(1).Value = GetParameterLogFilesFolderName()
    tbl_PARAMETERS.ListRows(8).Range.Cells(1).Value = GetParameterOutlookFolderName()
    tbl_PARAMETERS.ListRows(9).Range.Cells(1).Value = GetParameterDateFormatName()
    tbl_PARAMETERS.ListRows(10).Range.Cells(1).Value = GetParameterScheduleTimeName()

    With tbl_PARAMETERS.ListRows(6).Range.Cells(1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=currentYesNoInCurrentLanguage
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    If Split(previousYesNoInCurrentLanguage, ",")(0) = tbl_PARAMETERS.ListRows(6).Range.Cells(2).Value Then
        tbl_PARAMETERS.ListRows(6).Range.Cells(2).Value = Split(currentYesNoInCurrentLanguage, ",")(0)
    Else
        tbl_PARAMETERS.ListRows(6).Range.Cells(2).Value = Split(currentYesNoInCurrentLanguage, ",")(1)
    End If  

    'MAILS TABLE
    isOneFilePerRangeMailColumnName = GetMailIsOneFilePerRangeColumnName()
    generateMailColumnName = GetMailGenerateMailColumnName()

    tbl_MAILS.HeaderRowRange.Columns(1).Offset(-1, 0).Value = GetMailsTableName()

    tbl_MAILS.ListColumns(1).Name = GetMailNameColumnName()
    tbl_MAILS.ListColumns(2).Name = GetMailConversationColumnName()
    tbl_MAILS.ListColumns(3).Name = isOneFilePerRangeMailColumnName
    tbl_MAILS.ListColumns(4).Name = generateMailColumnName

    With tbl_MAILS.ListColumns(3).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=currentYesNoInCurrentLanguage
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    For Each cell In tbl_MAILS.ListColumns(3).DataBodyRange.Cells
        If Split(previousYesNoInCurrentLanguage, ",")(0) = cell.Value Then
            cell.Value = Split(currentYesNoInCurrentLanguage, ",")(0)
        Else
            cell.Value = Split(currentYesNoInCurrentLanguage, ",")(1)
        End If 
    Next cell

    With tbl_MAILS.ListColumns(4).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=currentYesNoInCurrentLanguage
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    For Each cell In tbl_MAILS.ListColumns(4).DataBodyRange.Cells
        If Split(previousYesNoInCurrentLanguage, ",")(0) = cell.Value Then
            cell.Value = Split(currentYesNoInCurrentLanguage, ",")(0)
        Else
            cell.Value = Split(currentYesNoInCurrentLanguage, ",")(1)
        End If 
    Next cell

    'MAIL FILES TABLE
    mailFilesMailColumnName = GetMailFilesMailColumnName()

    tbl_MAIL_FILES.HeaderRowRange.Columns(1).Offset(-1, 0).Value = GetMailFilesTableName()

    tbl_MAIL_FILES.ListColumns(1).Name = GetMailFilesNameColumnName()
    tbl_MAIL_FILES.ListColumns(2).Name = mailFilesMailColumnName

    With tbl_MAIL_FILES.ListColumns(2).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=INDIRECT(""MAILS[" & GetMailNameColumnName() & "]"")"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    'FILE REPORTS TABLE
    fileReportsFileColumnName = GetFileReportsFileColumnName()

    tbl_FILE_REPORTS.HeaderRowRange.Columns(1).Offset(-1, 0).Value = GetFileReportsTableName()

    tbl_FILE_REPORTS.ListColumns(1).Name = GetFileReportsNameColumnName()
    tbl_FILE_REPORTS.ListColumns(2).Name = fileReportsFileColumnName

    With tbl_FILE_REPORTS.ListColumns(2).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=INDIRECT(""MAIL_FILES[" & GetMailFilesNameColumnName() & "]"")"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With
End Sub

Function GetLanguageByLanguageName(languageName As String) As String
    For Each language in languageStructure("languages")
        If language("languageName") = languageName Then
                GetLanguageByLanguageName = language("name")
                Exit Function
        End If

        For Each newLanguageName in language("languageNames")
            If newLanguageName("name") = languageName Then
                GetLanguageByLanguageName = newLanguageName("language")
                Exit Function
            End If
        Next newLanguageName
    Next language
End Function

Function GetLanguageNameByLanguage() As String
    For Each language in languageStructure("languages")
        If language("name") = currentLanguage Then
            GetLanguageNameByLanguage = language("languageName")
            Exit Function
        End If
    Next language
End Function

Function GetAllLanguageNamesString() As String
    Dim languageNames As String
    languageNames = ""

    For Each language in languageStructure("languages")
        If language("name") <> currentLanguage Then Goto continueLoop
        
        languageNames = languageNames & language("languageName") & ","

        For Each languageName in language("languageNames")
            languageNames = languageNames & languageName("name") & ","
        Next languageName

        continueLoop:
    Next language

    languageNames = Left(languageNames, Len(languageNames) - 1)

    GetAllLanguageNamesString = languageNames
End Function

Function GetYesNoInCurrentLanguage() As String
    If currentLanguage = "SPANISH" Then
        GetYesNoInCurrentLanguage = "SI,NO"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetYesNoInCurrentLanguage = "YES,NO"
    End If
End Function

Function GetParameterWorksheetName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterWorksheetName = "PARÁMETROS"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterWorksheetName = "PARAMETERS"
    End If
End Function

'BUTTON CAPTION NAMES
Function GetBtnRefreshAllCaption() As String
    If currentLanguage = "SPANISH" Then
        GetBtnRefreshAllCaption = "REFRESCAR HOJAS"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetBtnRefreshAllCaption = "REFRESH WORKSHEETS"
    End If
End Function

Function GetBtnCreateMailFilesCaption() As String
    If currentLanguage = "SPANISH" Then
        GetBtnCreateMailFilesCaption = "GENERAR ARCHIVOS"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetBtnCreateMailFilesCaption = "CREATE MAIL FILES"
    End If
End Function

Function GetBtnCreateDraftsCaption() As String
    If currentLanguage = "SPANISH" Then
        GetBtnCreateDraftsCaption = "CREAR BORRADORES"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetBtnCreateDraftsCaption = "CREATE MAIL DRAFTS"
    End If
End Function

Function GetBtnSendAllDraftsCaption() As String
    If currentLanguage = "SPANISH" Then
        GetBtnSendAllDraftsCaption = "ENVIAR BORRADORES"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetBtnSendAllDraftsCaption = "SEND ALL DRAFTS"
    End If
End Function

Function GetBtnScheduleFileGenerationCaption() As String
    If currentLanguage = "SPANISH" Then
        GetBtnScheduleFileGenerationCaption = "PROGRAMAR GENERACIÓN DE ARCHIVOS"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetBtnScheduleFileGenerationCaption = "SCHEDULE FILE GENERATION"
    End If
End Function

Function GetBtnScheduleMailSendingCaption() As String
    If currentLanguage = "SPANISH" Then
        GetBtnScheduleMailSendingCaption = "PROGRAMAR ENVÍO DE CORREOS"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetBtnScheduleMailSendingCaption = "SCHEDULE MAIL SENDING"
    End If
End Function

'PARAMETERS TABLE
Function GetParameterTableName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterTableName = "PARÁMETROS"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterTableName = "PARAMETERS"
    End If
End Function

Function GetParameterNameColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterNameColumnName = "NOMBRE"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterNameColumnName = "NAME"
    End If
End Function

Function GetParameterValueColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterValueColumnName = "VALOR"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterValueColumnName = "VALUE"
    End If
End Function

Function GetParameterApplicationLanguageName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterApplicationLanguageName = "Idioma de la aplicación"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterApplicationLanguageName = "Application language"
    End If
End Function

Function GetParameterStartProcessDateName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterStartProcessDateName = "Fecha de proceso inicial"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterStartProcessDateName = "Start process date"
    End If
End Function

Function GetParameterEndProcessDateName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterEndProcessDateName = "Fecha de proceso final"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterEndProcessDateName = "End process date"
    End If
End Function

Function GetParameterMaxTimeoutInSecondsName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterMaxTimeoutInSecondsName = "Timeout máximo en segundos"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterMaxTimeoutInSecondsName = "Maximum timeout in seconds"
    End If
End Function

Function GetParameterFilesBaseFolderName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterFilesBaseFolderName = "Directorio base archivos"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterFilesBaseFolderName = "Files base directory"
    End If
End Function

Function GetParameterGenerateLogsName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterGenerateLogsName = "Generar logs?"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterGenerateLogsName = "Generate logs?"
    End If
End Function

Function GetParameterLogFilesFolderName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterLogFilesFolderName = "Directorio archivos de logs"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterLogFilesFolderName = "Log files directory"
    End If
End Function

Function GetParameterOutlookFolderName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterOutlookFolderName = "Carpeta de Outlook"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterOutlookFolderName = "Outlook folder"
    End If
End Function

Function GetParameterDateFormatName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterDateFormatName = "Formato de fechas"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterDateFormatName = "Date format"
    End If
End Function

Function GetParameterScheduleTimeName() As String
    If currentLanguage = "SPANISH" Then
        GetParameterScheduleTimeName = "Hora de ejecución"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetParameterScheduleTimeName = "Execution time"
    End If
End Function

' MAILS TABLE
Function GetMailsTableName() As String
    If currentLanguage = "SPANISH" Then
        GetMailsTableName = "CORREOS"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetMailsTableName = "MAILS"
    End If
End Function

Function GetMailNameColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetMailNameColumnName = "NOMBRE"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetMailNameColumnName = "NAME"
    End If
End Function

Function GetMailConversationColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetMailConversationColumnName = "CONVERSACIÓN"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetMailConversationColumnName = "CONVERSATION"
    End If
End Function

Function GetMailIsOneFilePerRangeColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetMailIsOneFilePerRangeColumnName = "UN ARCHIVO POR RANGO?"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetMailIsOneFilePerRangeColumnName = "ONE FILE PER RANGE?"
    End If
End Function

Function GetMailGenerateMailColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetMailGenerateMailColumnName = "GENERAR CORREO?"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetMailGenerateMailColumnName = "GENERATE MAIL?"
    End If
End Function

'MAIL FILES TABLE
Function GetMailFilesTableName() As String
    If currentLanguage = "SPANISH" Then
        GetMailFilesTableName = "ARCHIVOS"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetMailFilesTableName = "MAIL FILES"
    End If
End Function

Function GetMailFilesNameColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetMailFilesNameColumnName = "NOMBRE"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetMailFilesNameColumnName = "NAME"
    End If
End Function

Function GetMailFilesMailColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetMailFilesMailColumnName = "CORREO"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetMailFilesMailColumnName = "MAIL"
    End If
End Function

'FILE REPORTS TABLE
Function GetFileReportsTableName() As String
    If currentLanguage = "SPANISH" Then
        GetFileReportsTableName = "REPORTES"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetFileReportsTableName = "FILE REPORTS"
    End If
End Function

Function GetFileReportsNameColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetFileReportsNameColumnName = "NOMBRE"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetFileReportsNameColumnName = "NAME"
    End If
End Function

Function GetFileReportsFileColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetFileReportsFileColumnName = "ARCHIVO"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetFileReportsFileColumnName = "MAIL_FILE"
    End If
End Function

'INPUT VALIDATION MESSAGES

'BASIC TABLE CONTENT VALIDATION MESSAGES
Function InputValidationTableNotExistsMessage(tableName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationTableNotExistsMessage = "La tabla: '" & tableName & "' no existe. Favor revisar nombres internos de las tablas."
    ElseIf currentLanguage = "ENGLISH" Then 
        InputValidationTableNotExistsMessage = "The table: '" & tableName & "' does not exist. Please check internal table names."
    End If
End Function

Function InputValidationColumnNotExistsMessage(columnName As String, tableName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationColumnNotExistsMessage = "La columna: '" & columnName & "' de la tabla: '" & tableName & "' no existe. Favor revisar nombres."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationColumnNotExistsMessage = "The column: '" & columnName & "' in table: '" & tableName & "' does not exist. Please check names."
    End If
End Function

Function InputValidationValueNotExistsMessage(rowValue As String, columnName As String, tableName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationValueNotExistsMessage = "El valor: '" & rowValue & "', columna: '" & columnName & "', tabla: '" & tableName & "' no existe. Favor revisar nombres."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationValueNotExistsMessage = "The value: '" & rowValue & "', column: '" & columnName & "', table: '" & tableName & "' does not exist. Please check names."
    End If
End Function

' PARAMETERS TABLE VALIDATION MESSAGES
Function InputValidationParameterMustBeValidDateMessage(parameterName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationParameterMustBeValidDateMessage = "El valor del parámetro: '" & parameterName & "' debe ser una fecha válida."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationParameterMustBeValidDateMessage = "The value of parameter: '" & parameterName & "' must be a valid date."
    End If
End Function

Function InputValidationParameterMustBeNumberMessage(parameterName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationParameterMustBeNumberMessage = "El valor del parámetro: '" & parameterName & "' debe ser un número."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationParameterMustBeNumberMessage = "The value of parameter: '" & parameterName & "' must be a number."
    End If
End Function

Function InputValidationParameterCannotBeEmptyMessage(parameterName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationParameterCannotBeEmptyMessage = "El valor del parámetro: '" & parameterName & "' no puede quedar vacío."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationParameterCannotBeEmptyMessage = "The value of parameter: '" & parameterName & "' cannot be empty."
    End If
End Function

Function InputValidationParameterDirectoryNotExistsMessage(parameterName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationParameterDirectoryNotExistsMessage = "El directorio del parámetro: '" & parameterName & "' no existe. Favor de validar ruta."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationParameterDirectoryNotExistsMessage = "The directory for parameter: '" & parameterName & "' does not exist. Please validate the path."
    End If
End Function

Function InputValidationParameterDirectoryEndsWithSlashMessage(parameterValue As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationParameterDirectoryEndsWithSlashMessage = "El directorio del parámetro: '" & parameterValue & "' contiene el caracter \ al final. Favor de remover."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationParameterDirectoryEndsWithSlashMessage = "The directory value: '" & parameterValue & "' ends with the character \. Please remove it."
    End If
End Function

Function InputValidationExecutionTimeNotValidMessage(parameterValue As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationExecutionTimeNotValidMessage = "La hora de ejecución: '" & parameterValue & "' no es una hora válida."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationExecutionTimeNotValidMessage = "The execution time: '" & parameterValue & "' is not a valid time."
    End If
End Function

Function InputValidationTableIsEmptyMessage(tableName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationTableIsEmptyMessage = "La tabla: '" & tableName & "' está vacía."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationTableIsEmptyMessage = "The table: '" & tableName & "' is empty."
    End If
End Function

Function InputValidationTableHasEmptyValuesMessage(tableName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationTableHasEmptyValuesMessage = "Hay valores vacíos en la tabla: '" & tableName & "'."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationTableHasEmptyValuesMessage = "There are empty values in the table: '" & tableName & "'."
    End If
End Function

Function InputValidationColumnHasDuplicatesMessage(columnName As String, tableName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationColumnHasDuplicatesMessage = "Hay valores duplicados en la columna: '" & columnName & "' de la tabla: '" & tableName & "'."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationColumnHasDuplicatesMessage = "There are duplicate values in the column: '" & columnName & "' of the table: '" & tableName & "'."
    End If
End Function

Function InputValidationEmailHasNoFilesMessage(emailValue As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationEmailHasNoFilesMessage = "El correo: '" & emailValue & "' no tiene ningún archivo asociado."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationEmailHasNoFilesMessage = "The email: '" & emailValue & "' has no associated files."
    End If
End Function

Function InputValidationFileHasNoReportMessage(fileValue As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationFileHasNoReportMessage = "El archivo: '" & fileValue & "' no tiene ningún reporte asociado."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationFileHasNoReportMessage = "The file: '" & fileValue & "' has no associated report."
    End If
End Function

Function InputValidationAtLeastOneEmailMessage() As String
    If currentLanguage = "SPANISH" Then
        InputValidationAtLeastOneEmailMessage = "Debe haber al menos 1 correo a generar."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationAtLeastOneEmailMessage = "There must be at least 1 email to generate."
    End If
End Function

Function InputValidationWorksheetNotExistsMessage(sheetName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationWorksheetNotExistsMessage = "La hoja de cálculo: '" & sheetName & "' no existe. Favor crearla junto a su tabla de Power Query."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationWorksheetNotExistsMessage = "The worksheet: '" & sheetName & "' does not exist. Please create it next to its Power Query table."
    End If
End Function

Function InputValidationTableNotFoundOnSheetMessage(tableName As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationTableNotFoundOnSheetMessage = "La tabla: '" & tableName & "' no fue encontrada en su respectiva hoja de cálculo. Favor crear."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationTableNotFoundOnSheetMessage = "The table: '" & tableName & "' was not found on its corresponding worksheet. Please create it."
    End If
End Function

Function InputValidationConversationNotExistsMessage(conversationValue As String) As String
    If currentLanguage = "SPANISH" Then
        InputValidationConversationNotExistsMessage = "La conversación: '" & conversationValue & "' no existe."
    ElseIf currentLanguage = "ENGLISH" Then
        InputValidationConversationNotExistsMessage = "The conversation: '" & conversationValue & "' does not exist."
    End If
End Function

' FILE GENERATION MESSAGES

Function FileGenerationErrorWhenFetchingReportMessage(fileReportName As String) As String
    If currentLanguage = "SPANISH" Then
        FileGenerationErrorWhenFetchingReportMessage = "Hubo un error al consultar el reporte " & fileReportName & " desde la base de datos."
    ElseIf currentLanguage = "ENGLISH" Then
        FileGenerationErrorWhenFetchingReportMessage = "There was an error querying the report " & fileReportName & " from the database."
    End If
End Function

Function FileGenerationReportReturnedNoRowsMessage(fileReportName As String) As String
    If currentLanguage = "SPANISH" Then
        FileGenerationReportReturnedNoRowsMessage = "El reporte " & fileReportName & " no trajo registros."
    ElseIf currentLanguage = "ENGLISH" Then
        FileGenerationReportReturnedNoRowsMessage = "The report " & fileReportName & " returned no records."
    End If
End Function

Function FileGenerationReportNotUpdatedMessage(fileReportName As String) As String
    If currentLanguage = "SPANISH" Then
        FileGenerationReportNotUpdatedMessage = "El reporte " & fileReportName & " no se actualizó."
    ElseIf currentLanguage = "ENGLISH" Then
        FileGenerationReportNotUpdatedMessage = "The report " & fileReportName & " was not updated."
    End If
End Function

Function FileGenerationMissingProcessDateColumnMessage(fileReportName As String) As String
    If currentLanguage = "SPANISH" Then
        FileGenerationMissingProcessDateColumnMessage = "No se encontró la columna PROCESS_DATE_FOR_RANGE en el reporte " & fileReportName & "."
    ElseIf currentLanguage = "ENGLISH" Then
        FileGenerationMissingProcessDateColumnMessage = "The column PROCESS_DATE_FOR_RANGE was not found in the report " & fileReportName & "."
    End If
End Function

Function FileGenerationGenericErrorMessage(fileReportName As String) As String
    If currentLanguage = "SPANISH" Then
        FileGenerationGenericErrorMessage = "Ha ocurrido un error al generar el reporte " & fileReportName & "."
    ElseIf currentLanguage = "ENGLISH" Then
        FileGenerationGenericErrorMessage = "An error occurred while generating the report " & fileReportName & "."
    End If
End Function


' MISCELLANIOUS MESSAGES
Function GetLanguageChangePromptMessage() As String
    If currentLanguage = "SPANISH" Then
        GetLanguageChangePromptMessage = "¿Seguro que desea cambiar el idioma de la aplicación?"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetLanguageChangePromptMessage = "Are you sure you want to change the application language?"
    End If
End Function

Function RefreshAllUpdatingReportsMessage() As String
    If currentLanguage = "SPANISH" Then
        RefreshAllUpdatingReportsMessage = "Actualizando reportes..."
    ElseIf currentLanguage = "ENGLISH" Then
        RefreshAllUpdatingReportsMessage = "Updating reports..."
    End If
End Function

Function ExcelSheetsUpdatedMessage() As String
    If currentLanguage = "SPANISH" Then
        ExcelSheetsUpdatedMessage = "Hojas de Excel actualizadas."
    ElseIf currentLanguage = "ENGLISH" Then
        ExcelSheetsUpdatedMessage = "Excel sheets updated."
    End If
End Function
