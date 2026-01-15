Sub Worksheet_Change(ByVal Target As Range)
    If Not isSilentChange And Not Intersect(Target, tbl_PARAMETERS.ListRows(1).Range.Cells(2)) Is Nothing And currentLanguageName <> tbl_PARAMETERS.ListRows(1).Range.Cells(2).Value Then
        If MsgBox(GetLanguageChangePromptMessage(), vbYesNo + vbQuestion) = vbYes Then
            UpdateApplicationLanguage
        Else
            tbl_PARAMETERS.ListRows(1).Range.Cells(2).Value = currentLanguageName
        End If
    End If
End Sub

Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Target, tbl_PARAMETERS.ListRows(1).Range.Cells(2)) Is Nothing Then
        currentLanguageName = tbl_PARAMETERS.ListRows(1).Range.Cells(2).Value
        currentLanguage = GetLanguageByLanguageName(tbl_PARAMETERS.ListRows(1).Range.Cells(2).Value)
    End If
End Sub