Sub Workbook_Open()
    Set languageStructure = GetLanguageStructure()

    Set tbl_PARAMETERS = PARAMETERS.ListObjects("PARAMETERS")
    Set tbl_MAILS = PARAMETERS.ListObjects("MAILS")
    Set tbl_MAIL_FILES = PARAMETERS.ListObjects("MAIL_FILES")
    Set tbl_FILE_REPORTS = PARAMETERS.ListObjects("FILE_REPORTS")

    Range("H1").Value = "v" & version

    currentLanguage = GetLanguageByLanguageName(tbl_PARAMETERS.ListRows(1).Range.Cells(1, 2).Value)

    isSilentChange = False
End Sub