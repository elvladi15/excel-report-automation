Attribute VB_Name = "ModMain"

Public Const version As String = "1.1.0"

Public currentLanguage As String
Public languageStructure As Object
Public isSilentChange As Boolean

Public basicTableStructure As Object

Public executionMode As String

Public attemptMaxCount As Long

Public tbl_PARAMETROS As ListObject
Public tbl_CORREOS As ListObject
Public tbl_ARCHIVOS As ListObject
Public tbl_REPORTES As ListObject

Public dictParameters As Object

Public startProcessDate As Date
Public endProcessDate As Date
Public baseReportFolder As String
Public outlookFolderName As String
Public canGenerateLogs As Boolean
Public logsFileFolder As String
Public dateFormat As String
Public scheduleTime As Date

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
	ElseIf Application.Caller = "btnScheduleMailGeneration" Then
		sendMails = False
		ScheduleAutomaticRun
	End If
	Application.DisplayAlerts = True
	Exit Sub
End Sub