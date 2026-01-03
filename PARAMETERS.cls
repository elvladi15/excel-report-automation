Sub Worksheet_Change(ByVal Target As Range)
    If Not isSilentChange And Not Intersect(Target, Range("B3")) Is Nothing And currentLanguageName <> Range("B3").Value Then
        If MsgBox("Are you sure you want to change the application language?", vbYesNo + vbQuestion) = vbYes Then
            UpdateApplicationLanguage
        Else
            Range("B3").Value = currentLanguageName
        End If
    End If
End Sub

Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Temp
    Set languageStructure = GetLanguageStructure()
    Set tbl_PARAMETERS = PARAMETERS.ListObjects("PARAMETERS")
	Set tbl_MAILS = PARAMETERS.ListObjects("MAILS")
	Set tbl_MAIL_FILES = PARAMETERS.ListObjects("MAIL_FILES")
	Set tbl_FILE_REPORTS = PARAMETERS.ListObjects("FILE_REPORTS")

    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        currentLanguageName = Range("B3").Value
        currentLanguage = GetLanguageByLanguageName(Range("B3").Value)
    End If
End Sub