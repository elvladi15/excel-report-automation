Attribute VB_Name = "ModDraftAndEmailSending"
Sub Manual_CreateDrafts()
	Call CreateDrafts("MANUAL")
End Sub
Sub Automatic_CreateDrafts()
	Call CreateDrafts("AUTOMATICO")
End Sub
Sub CreateDrafts(executionMode As String)
	If isInputValidationCorrect = False Then Exit Sub

	InitializeGlobals
	
	Dim mails As Variant
	
	mails = Range("CORREOS[NOMBRE]")
	
	If selectedReport = "Todos" Then
		For Each item In mails
			Call CreateDraft(CStr(item))
		Next item
		
		Exit Sub
	End If
	
	Call CreateDraft(selectedReport)

	If executionMode = "MANUAL" Then MsgBox "Borradores creados correctamente."
End Sub

Sub CreateDraft(mailName As String)
	If isInputValidationCorrect = False Then Exit Sub
	
	InitializeGlobals
	
	On Error GoTo ErrorHandler

	Call AppendToLogsFile("Creando borrador: " & mailName & "...")
	
	Dim OutlookApp As Object
	Dim OutlookNamespace As Object
	Dim Inbox As Object
	Dim Items As Object
	Dim MailItem As Object
	Dim Reply As Object
	
	Dim mailFiles As Variant
	Dim conversationSubject As String
	Dim foldersToSearch As New collection
	Dim fileEndings As New collection
	Dim fileFolder As String
	Dim isOneFilePerRange As Boolean
	Dim mailFileCount As Long
	Dim dateValue As Date
	Dim fileEndingFound As Boolean
	
	fileFolder = baseReportFolder & "\" & mailName & "\"
	
	isOneFilePerRange = ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""" & mailName & """, CORREOS[NOMBRE], CORREOS[UN ARCHIVO POR RANGO?])") = "SI"
	
	conversationSubject = ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""" & mailName & """, CORREOS[NOMBRE], CORREOS[CONVERSACION])")
	
	mailFiles = ThisWorkbook.ActiveSheet.Evaluate("FILTER(ARCHIVOS[NOMBRE], ARCHIVOS[CORREO] = """ & mailName & """)")
		
	mailFileCount = UBound(mailFiles) - LBound(mailFiles) + 1
	
	If mailFileCount > 1 Then
		If isOneFilePerRange Then
			If startProcessDate = endProcessDate Then
				fileEndings.Add Format(endProcessDate, dateFormat)
				foldersToSearch.Add fileFolder & Format(endProcessDate, dateFormat) & "\"
			Else
				fileEndings.Add Format(startProcessDate, "dd") & "-" & Format(endProcessDate, "dd")
				foldersToSearch.Add fileFolder & Format(startProcessDate, "dd") & "-" & Format(endProcessDate, "dd") & "\"
			End If
		Else
			For dateValue = startProcessDate To endProcessDate
				fileEndings.Add Format(dateValue, dateFormat)
				foldersToSearch.Add fileFolder & Format(dateValue, dateFormat) & "\"
			Next dateValue
		End If
	Else
		If isOneFilePerRange Then
			If startProcessDate = endProcessDate Then
				fileEndings.Add Format(endProcessDate, dateFormat)
			Else
				fileEndings.Add Format(startProcessDate, "dd") & "-" & Format(endProcessDate, "dd")
			End If
		Else
			For dateValue = startProcessDate To endProcessDate
				fileEndings.Add Format(dateValue, dateFormat)
			Next dateValue
		End If
			
		foldersToSearch.Add fileFolder
	End If
	
	Set OutlookApp = CreateObject("Outlook.Application")
	Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
	Set Inbox = OutlookNamespace.GetDefaultFolder(6).Parent.Folders(outlookFolder)
	
	Set Items = Inbox.Items.Restrict("[Subject] = '" & conversationSubject & "'")
	Items.Sort "ReceivedTime", True
	
	If Items.Count > 0 Then
		Set MailItem = Items.item(1)
		Set ReplyAll = MailItem.ReplyAll
		
		For Each folder In foldersToSearch
			For Each fileEnding In fileEndings
				fileEndingFound = False
				
				filePath = Dir(folder & "*.*")
				
				Do While filePath <> ""
					If InStr(filePath, CStr(fileEnding)) > 0 Then
						ReplyAll.Attachments.Add folder & filePath
						
						fileEndingFound = True
					End If
					
					filePath = Dir()
				Loop
				If (fileEndingFound = False) Then
					AppendToLogsFile ("No se puede crear el borrador: " & mailName & ". Faltan archivos por generar.")
					
					Exit Sub
				End If
			Next fileEnding
		Next folder

		ReplyAll.Body = "MENSAJE AUTOMÁTICO. Anexo reporte. Saludos"
		
		ReplyAll.Save
		
		AppendToLogsFile ("El borrador: " & mailName & " fue creado exitosamente.")
	Else
		AppendToLogsFile ("No se pudo encontrar la cadena de correos: " & conversationSubject)
	End If
	
	Exit Sub
ErrorHandler:
	AppendToLogsFile ("Ha ocurrido un error al crear el borrador: " & mailName)
End Sub

Sub Manual_SendAllDrafts()
	If isInputValidationCorrect = False Then Exit Sub
	
	InitializeGlobals

	Call SendAllDrafts(3, "MANUAL")
End Sub
Sub Automatic_SendAllDrafts()
	If isInputValidationCorrect = False Then Exit Sub
	
	InitializeGlobals

	Call SendAllDrafts(3, "AUTOMATICO")
End Sub
Sub SendAllDrafts(attemptCount As Long, executionMode)
	On Error GoTo ErrHandler
	
	Call AppendToLogsFile("Enviando borradores...")
	
	Dim olApp As Object
	Dim ns As Object
	Dim drafts As Object
	Dim itms As Object
	Dim i As Long
	Dim mi As Object

	Const olFolderDrafts As Long = 16

	Set olApp = GetObject("", "Outlook.Application")
	If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")

	Set ns = olApp.GetNamespace("MAPI")
	Set drafts = ns.GetDefaultFolder(olFolderDrafts)

	Set itms = drafts.Items
	
	On Error Resume Next
	itms.Sort "[LastModificationTime]", True
	On Error GoTo ErrHandler

	For i = itms.Count To 1 Step -1
		Set mi = itms(i)

		If Not mi Is Nothing Then
			Dim msgClass As String, isSent As Boolean
			On Error Resume Next
			msgClass = LCase$(CStr(mi.MessageClass))
			isSent = False
			isSent = mi.Sent
			On Error GoTo ErrHandler

			If msgClass = "ipm.note" And isSent = False Then
				mi.Display False
				DoEvents

				Dim allRecipients As String
				On Error Resume Next
				allRecipients = Trim$(CStr(mi.To) & CStr(mi.CC) & CStr(mi.BCC))
				On Error GoTo ErrHandler

				If Len(allRecipients) > 0 Then
					mi.Send
				End If
			End If
		End If
	Next i
	
	Call AppendToLogsFile("Correos enviados exitosamente")

	If executionMode = "MANUAL" Then MsgBox "Correos enviados exitosamente."
	
	Automatic_ScheduleMailSending
	Exit Sub
ErrHandler:
	Call AppendToLogsFile("Ha ocurrido un error al enviar los borradores. " & Err.Number & " " & Err.Description)
	
	If attemptCount = 1 Then
		Call AppendToLogsFile("Intentos agotados. Envío de correos abortado")

		If executionMode = "MANUAL" Then MsgBox "Ha ocurrido un error al enviar los correos."

		Exit Sub
	End If
	
	Call SendAllDrafts(attemptCount - 1, executionMode)
End Sub