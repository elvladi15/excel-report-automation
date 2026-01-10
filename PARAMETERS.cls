Sub Worksheet_Change(ByVal Target As Range)
    If Not isSilentChange And Not Intersect(Target, Range("B3")) Is Nothing And currentLanguageName <> Range("B3").Value Then
        If MsgBox(GetLanguageChangePromptMessage(), vbYesNo + vbQuestion) = vbYes Then
            UpdateApplicationLanguage
        Else
            tbl_PARAMETERS.ListRows(1).Range.Cells(1, 2).Value = currentLanguageName
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

    If Not Intersect(Target, tbl_PARAMETERS.ListRows(1).Range.Cells(1, 2)) Is Nothing Then
        currentLanguageName = tbl_PARAMETERS.ListRows(1).Range.Cells(1, 2).Value
        currentLanguage = GetLanguageByLanguageName(tbl_PARAMETERS.ListRows(1).Range.Cells(1, 2).Value)
    End If
End Sub