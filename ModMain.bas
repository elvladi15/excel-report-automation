Attribute VB_Name = "ModMain"
Public executionMode As String

Public currentAttempt As Long
Public attemptMaxCount As Long

Public tbl_PARAMETROS As ListObject
Public tbl_CORREOS As ListObject
Public tbl_ARCHIVOS As ListObject
Public tbl_REPORTES As ListObject

Public dictParameters As Object

Public startProcessDate As Date
Public endProcessDate As Date
Public baseReportFolder As String
Public outlookFolder As String
Public logsFileFolder As String
Public selectedReport As String
Public dateFormat As String
Public canGenerateLogs As Boolean

Public canMailBeSent As Boolean
Public currentProcessDate As Variant
Public errorReport As String

'Public OutlookApp As Object
'Public OutlookNamespace As Object
Public Inbox As Object
Public Items As Object
Public MailItem As Object
Public Reply As Object

Sub Main()
	If isInputValidationCorrect = False Then Exit Sub

	executionMode = getExecutionMode()
	currentAttempt = 1
	attemptMaxCount = 3

	startProcessDate = CDate(dictParameters("START_PROCESS_DATE"))
	endProcessDate = CDate(dictParameters("END_PROCESS_DATE"))
	baseReportFolder = dictParameters("Directorio base reportes")
	logsFileFolder = dictParameters("Directorio archivos de logs")
	outlookFolder = dictParameters("Carpeta de Outlook")
	dateFormat = dictParameters("Formato de fechas")
	canGenerateLogs = dictParameters("Generar logs") = "SI"

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
	Else
		MsgBox "Botón no reconocido."
	End If

	Exit Sub
	ErrorHandler:
		MsgBox "No correr desde el código. Usar algún botón."
End Sub

Sub test()
	'Set OutlookApp = CreateObject("Outlook.Application")
	'Set OutlookNamespace = CreateObject("Outlook.Application").GetNamespace("MAPI")
	Set Inbox = CreateObject("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6).Parent.Folders(outlookFolder)

	Set Items = Inbox.Items.Restrict("[Subject] = '" & conversationSubject & "'")
	Items.Sort "ReceivedTime", True

	If Items.Count > 0 Then

	Else
		AppendToLogsFile ("No se pudo encontrar la cadena de correos: " & conversationSubject)
	End If
End Sub