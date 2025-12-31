Sub Workbook_Open()
    Range("H1").Value = "v" & version

    Set languageStructure = GetLanguageStructure()

    currentLanguage = GetLanguageByLanguageName(Range("B2").Value)

    isSilentChange = False
End Sub