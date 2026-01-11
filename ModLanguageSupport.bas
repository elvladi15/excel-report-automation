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

    With Range("MAILS[" & isOneFilePerRangeMailColumnName & "]").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=currentYesNoInCurrentLanguage
        .IgnoreBlank = False
        .InCellDropdown = True
    End With


    For Each cell In Range("MAILS[" & isOneFilePerRangeMailColumnName & "]").Cells
        If Split(previousYesNoInCurrentLanguage, ",")(0) = cell.Value Then
            cell.Value = Split(currentYesNoInCurrentLanguage, ",")(0)
        Else
            cell.Value = Split(currentYesNoInCurrentLanguage, ",")(1)
        End If 
    Next cell


    With Range("MAILS[" & generateMailColumnName & "]").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=currentYesNoInCurrentLanguage
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    For Each cell In Range("MAILS[" & generateMailColumnName & "]").Cells
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
    
    With Range("MAIL_FILES[" & mailFilesMailColumnName & "]").Validation
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

Function GetLanguageChangePromptMessage() As String
    If currentLanguage = "SPANISH" Then
        GetLanguageChangePromptMessage = "¿Seguro que desea cambiar el idioma de la aplicación?"
    ElseIf currentLanguage = "ENGLISH" Then 
        GetLanguageChangePromptMessage = "Are you sure you want to change the application language?"
    End If
End Function