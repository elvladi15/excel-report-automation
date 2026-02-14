Attribute VB_Name = "ModMain"

Public Const version As String = "1.1.1"

Public currentLanguageName As String
Public currentLanguage As String

Public languageStructure As Object
Public isSilentChange As Boolean

Public basicTableStructure As Object

Public executionMode As String

Public attemptMaxCount As Long

Public tbl_PARAMETERS As ListObject
Public tbl_MAILS As ListObject
Public tbl_MAIL_FILES As ListObject
Public tbl_FILE_REPORTS As ListObject

Public startProcessDate As Date
Public endProcessDate As Date
Public baseReportFolder As String
Public outlookFolderName As String
Public canGenerateLogs As Boolean
Public logsFileFolder As String
Public dateFormat As String
Public scheduleTime As Date

Public isFirstWorksheet As Boolean

Public currentProcessDate As Variant

Public mailFilesNotGenerated As Collection
Public reportsNotGenerated As Collection
Public draftsNotGenerated As Collection
Public conversationsNotSent As collection

Public outlookAppRef As Object
Public outlookReportFolderRef As Object
Public outlookDraftsFolderRef As Object
Public sendMails As Boolean

Sub Main()
	If Not IsInputValidationCorrect Then Exit Sub

	CloseAllOtherWorkbooks

	executionMode = "MANUAL"
	attemptMaxCount = 3
	Set mailFilesNotGenerated = New Collection
	Set reportsNotGenerated = New Collection
	Set draftsNotGenerated = New Collection
	Set conversationsNotSent = New Collection

	If Application.Caller = "btnRefreshAll" Then
		RefreshAll
	ElseIf Application.Caller = "btnCreateMailFiles" Then
		CreateMailFiles
	ElseIf Application.Caller = "btnCreateDrafts" Then
		CreateDrafts
	ElseIf Application.Caller = "btnSendAllDrafts" Then
		SendAllDrafts
	ElseIf Application.Caller = "btnScheduleMailSending" Then
		sendMails = True
		ScheduleAutomaticRun
	ElseIf Application.Caller = "btnScheduleFileGeneration" Then
		sendMails = False
		ScheduleAutomaticRun
	End If
	Application.DisplayAlerts = True

	executionMode = "AUTOMATIC"
End Sub