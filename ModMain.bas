Attribute VB_Name = "ModMain"

Public wsPARAMETROS As Worksheet
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
Public errorReport As String

Public outlookAppRef As Object
Public outlookReportFolderRef As Object
Public outlookDraftsFolderRef As Object

Public allDraftsCreated As Boolean
Public allFilesCreated As Boolean

Public continueExecution As Boolean
Public sendMails As Boolean

Sub Main()
	If Not isInputValidationCorrect Then Exit Sub

	CloseAllOtherWorkbooks

	executionMode = "MANUAL"
	attemptMaxCount = 3
	allDraftsCreated = True
	allFilesCreated = True
	continueExecution = True

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
	Else
		MsgBox "Botón no reconocido."
	End If
	Application.DisplayAlerts = True
	Exit Sub
End Sub