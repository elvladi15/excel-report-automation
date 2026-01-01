Attribute VB_Name = "ModLanguageSupport"
Sub UpdateApplicationLanguage()
    Dim isOneFilePerRangeMailColumnName As String
    Dim generateMailColumnName As String

    Dim mailFilesNameColumnName As String
    Dim mailFilesMailColumnName As String
    
    Dim fileReportsFileColumnName As String

    'TEMP
    Set languageStructure = GetLanguageStructure()
    Set tbl_PARAMETERS = PARAMETERS.ListObjects("PARAMETERS")
	Set tbl_MAILS = PARAMETERS.ListObjects("MAILS")
	Set tbl_MAIL_FILES = PARAMETERS.ListObjects("MAIL_FILES")
	Set tbl_FILE_REPORTS = PARAMETERS.ListObjects("FILE_REPORTS")

    currentLanguage = GetLanguageByLanguageName(Range("B3").Value)

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

    tbl_PARAMETERS.ListColumns(1).Name = GetNameParameterColumnName()
    tbl_PARAMETERS.ListColumns(2).Name = GetValueParameterColumnName()

    isSilentChange = True
    Range("B3").Value = GetLanguageNameByLanguage()
    isSilentChange = False

    With Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=GetAllLanguageNamesString()
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    Dim excelRow As Long

    excelRow = 3

    Range("A" & excelRow).Value = GetApplicationLanguageParameterName()
        excelRow = excelRow + 1
    Range("A" & excelRow).Value = GetStartProcessDateParameterName()
        excelRow = excelRow + 1
    Range("A" & excelRow).Value = GetEndProcessDateParameterName()
        excelRow = excelRow + 1
    Range("A" & excelRow).Value = GetMaxTimeoutInSecondsParameterName()
        excelRow = excelRow + 1
    Range("A" & excelRow).Value = GetFilesBaseFolderParameterName()
        excelRow = excelRow + 1
    Range("A" & excelRow).Value = GetGenerateLogsParameterName()
        excelRow = excelRow + 1
    Range("A" & excelRow).Value = GetLogFilesFolderParameterName()
        excelRow = excelRow + 1
    Range("A" & excelRow).Value = GetOutlookFolderParameterName()
        excelRow = excelRow + 1
    Range("A" & excelRow).Value = GetDateFormatParameterName()
        excelRow = excelRow + 1
    Range("A" & excelRow).Value = GetScheduleTimeParameterName()

    'MAILS TABLE
    isOneFilePerRangeMailColumnName = GetIsOneFilePerRangeMailColumnName()
    generateMailColumnName = GetGenerateMailColumnName()

    tbl_MAILS.HeaderRowRange.Columns(1).Offset(-1, 0).Value = GetMailsTableName()

    tbl_MAILS.ListColumns(1).Name = GetNameMailColumnName()
    tbl_MAILS.ListColumns(2).Name = GetConversationMailColumnName()
    tbl_MAILS.ListColumns(3).Name = isOneFilePerRangeMailColumnName
    tbl_MAILS.ListColumns(4).Name = generateMailColumnName

    With Range("MAILS[" & isOneFilePerRangeMailColumnName & "]").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=GetYesNoInCurrentLanguage()
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    With Range("MAILS[" & generateMailColumnName & "]").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=GetYesNoInCurrentLanguage()
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    'MAIL FILES TABLE
    mailFilesMailColumnName = GetMailFilesMailColumnName()

    tbl_MAIL_FILES.HeaderRowRange.Columns(1).Offset(-1, 0).Value = GetMailFilesTableName()

    tbl_MAIL_FILES.ListColumns(1).Name = GetMailFilesNameColumnName()
    tbl_MAIL_FILES.ListColumns(2).Name = mailFilesMailColumnName
    
    With Range("MAIL_FILES[" & mailFilesMailColumnName & "]").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=INDIRECT(""MAILS[" & GetNameMailColumnName() & "]"")"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    'FILE REPORTS TABLE
    fileReportsFileColumnName = GetFileReportsFileColumnName()

    tbl_FILE_REPORTS.HeaderRowRange.Columns(1).Offset(-1, 0).Value = GetFileReportsTableName()

    tbl_FILE_REPORTS.ListColumns(1).Name = GetFileReportsNameColumnName()
    tbl_FILE_REPORTS.ListColumns(2).Name = fileReportsFileColumnName

    With Range("FILE_REPORTS[" & fileReportsFileColumnName & "]").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=INDIRECT(""MAIL_FILES[" & GetMailFilesNameColumnName() & "]"")"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With
End Sub

Function GetLanguageByLanguageName(languageName As String) As String
    For Each language in languageStructure("languages")
        For Each currentLanguageName in language("languageNames")
            If currentLanguageName("name") = languageName Then
                GetLanguageByLanguageName = currentLanguageName("language")
                Exit Function
            End If
        Next currentLanguageName
    Next language
End Function

Function GetLanguageNameByLanguage() As String
    For Each currentLanguageIterator in languageStructure("languages")
        If currentLanguageIterator("name") <> currentLanguage Then Goto continueLoop

        For Each languageName in currentLanguageIterator("languageNames")
            If languageName("language") = currentLanguage Then
                GetLanguageNameByLanguage = languageName("name")
                Exit Function
            End If
        Next languageName

        continueLoop:
    Next currentLanguageIterator
End Function

Function GetAllLanguageNamesString() As String
    Dim languageNames As String
    languageNames = ""

    For Each language in languageStructure("languages")
        If language("name") <> currentLanguage Then Goto continueLoop

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

Function GetNameParameterColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetNameParameterColumnName = "NOMBRE"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetNameParameterColumnName = "NAME"
    End If
End Function

Function GetValueParameterColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetValueParameterColumnName = "VALOR"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetValueParameterColumnName = "VALUE"
    End If
End Function

Function GetApplicationLanguageParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetApplicationLanguageParameterName = "Idioma de la aplicación"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetApplicationLanguageParameterName = "Application language"
    End If
End Function

Function GetStartProcessDateParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetStartProcessDateParameterName = "Fecha de proceso inicial"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetStartProcessDateParameterName = "Start process date"
    End If
End Function

Function GetEndProcessDateParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetEndProcessDateParameterName = "Fecha de proceso final"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetEndProcessDateParameterName = "End process date"
    End If
End Function

Function GetMaxTimeoutInSecondsParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetMaxTimeoutInSecondsParameterName = "Timeout máximo en segundos"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetMaxTimeoutInSecondsParameterName = "Maximum timeout in seconds"
    End If
End Function

Function GetFilesBaseFolderParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetFilesBaseFolderParameterName = "Directorio base archivos"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetFilesBaseFolderParameterName = "Files base directory"
    End If
End Function

Function GetGenerateLogsParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetGenerateLogsParameterName = "Generar logs?"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetGenerateLogsParameterName = "Generate logs?"
    End If
End Function

Function GetLogFilesFolderParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetLogFilesFolderParameterName = "Directorio archivos de logs"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetLogFilesFolderParameterName = "Log files directory"
    End If
End Function

Function GetOutlookFolderParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetOutlookFolderParameterName = "Carpeta de Outlook"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetOutlookFolderParameterName = "Outlook folder"
    End If
End Function

Function GetDateFormatParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetDateFormatParameterName = "Formato de fechas"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetDateFormatParameterName = "Date format"
    End If
End Function

Function GetScheduleTimeParameterName() As String
    If currentLanguage = "SPANISH" Then
        GetScheduleTimeParameterName = "Hora de ejecución"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetScheduleTimeParameterName = "Execution time"
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

Function GetNameMailColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetNameMailColumnName = "NOMBRE"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetNameMailColumnName = "NAME"
    End If
End Function

Function GetConversationMailColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetConversationMailColumnName = "CONVERSACIÓN"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetConversationMailColumnName = "CONVERSATION"
    End If
End Function

Function GetIsOneFilePerRangeMailColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetIsOneFilePerRangeMailColumnName = "UN ARCHIVO POR RANGO?"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetIsOneFilePerRangeMailColumnName = "ONE FILE PER RANGE?"
    End If
End Function

Function GetGenerateMailColumnName() As String
    If currentLanguage = "SPANISH" Then
        GetGenerateMailColumnName = "GENERAR CORREO?"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetGenerateMailColumnName = "GENERATE MAIL?"
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