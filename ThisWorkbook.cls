Sub Workbook_Open()
    Range("H1").Value = "v" & version

    Set languageStructure = GetLanguageStructure()

    Set tbl_PARAMETERS = PARAMETERS.ListObjects("PARAMETERS")
	Set tbl_MAILS = PARAMETERS.ListObjects("MAILS")
	Set tbl_MAIL_FILES = PARAMETERS.ListObjects("MAIL_FILES")
	Set tbl_FILE_REPORTS = PARAMETERS.ListObjects("FILE_REPORTS")

    currentLanguage = GetLanguageByLanguageName(Range("B2").Value)

    isSilentChange = False
End Sub