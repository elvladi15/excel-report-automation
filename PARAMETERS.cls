Sub Worksheet_Change(ByVal Target As Range)
    If Not isSilentChange And Not Intersect(Target, Me.Range("B2")) Is Nothing Then
        UpdateApplicationLanguage
    End If
End Sub