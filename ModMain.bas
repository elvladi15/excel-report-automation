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
Public logsFileFolder As String
Public selectedReport As String
Public dateFormat As String
Public canGenerateLogs As Boolean

Public canMailBeSent As Boolean
Public currentProcessDate As Variant
Public errorReport As String

Public outlookAppRef As Object
Public outlookReportFolderRef As Object
Public outlookDraftsFolderRef As Object

Sub Main()
	If Not isInputValidationCorrect Then Exit Sub

	CloseAllOtherWorkbooks

	executionMode = "MANUAL"
	attemptMaxCount = 3
	canMailBeSent = True

	On Error GoTo ErrorHandler
	If Application.Caller = "btnRefreshAll" Then
		RefreshAll
	ElseIf Application.Caller = "btnCreateMailFiles" Then
		CreateMailFiles
	ElseIf Application.Caller = "btnCreateDrafts" Then
		CreateDrafts
	ElseIf Application.Caller = "btnSendAllDrafts" Then
		SendAllDrafts
	ElseIf Application.Caller = "btnScheduleMailSending" Then
		ScheduleMailSending
	ElseIf Application.Caller = "btnScheduleMailGeneration" Then
		ScheduleMailGeneration
	Else
		MsgBox "Botón no reconocido."
	End If

	Exit Sub
	ErrorHandler:
		MsgBox "No correr desde el código. Usar algún botón."
End Sub