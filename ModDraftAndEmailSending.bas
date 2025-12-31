Attribute VB_Name = "ModDraftAndEmailSending"
Sub CreateDrafts()
	Dim generateReportColumn As String
	Dim nameColumn As String

	If Not IsConversationColumnCorrect Then Exit Sub

	For Each mailName In PARAMETERS.Evaluate("FILTER([NOMBRE], [GENERAR CORREO?] = """ & Split(GetYesNoInCurrentLanguage(), ",")(0) & """)")
		Call CreateDraft(CStr(mailName))
	Next mailName

	If executionMode = "MANUAL" Then
		If draftsNotGenerated.Count = 0 Then
				outputMesssage = outputMesssage & "Borradores generados exitosamente. "
			Else
				outputMesssage = "Los borradores:" & vbCrLf & vbCrLf

				For Each draft In draftsNotGenerated
					outputMesssage = outputMesssage & draft & vbCrLf
				Next draft
				
				outputMesssage = outputMesssage & vbCrLf

				outputMesssage = outputMesssage & " no se pudieron crear porque sus archivos no se crearon." & vbCrLf & vbCrLf
		End If
		
		MsgBox outputMesssage
	End If
End Sub

Sub CreateDraft(mailName As String)
	On Error GoTo ErrorHandler

	Call AppendToLogsFile("Creando borrador: '" & mailName & "'...")

	Dim conversation As Object


	Dim mailFiles As Variant
	Dim conversationSubject As String
	Dim foldersToSearch As New collection
	Dim fileEndings As New collection
	Dim fileFolder As String
	Dim isOneFilePerRange As Boolean
	Dim mailFileCount As Long
	Dim dateValue As Date
	Dim quantityOfFilesFound As Long

	fileFolder = baseReportFolder & "\" & mailName & "\"

	isOneFilePerRange = PARAMETERS.Evaluate("XLOOKUP(""" & mailName & """, MAILS[" & GetNameMailColumnName() & "], MAILS[" & GetIsOneFilePerRangeMailColumnName() & "])") = Split(GetYesNoInCurrentLanguage(), ",")(0)
	conversationSubject = PARAMETERS.Evaluate("XLOOKUP(""" & mailName & """, MAILS[" & GetNameMailColumnName() & "], MAILS[" & GetConversationMailColumnName() & "])")
	mailFiles = PARAMETERS.Evaluate("FILTER(MAIL_FILES[NOMBRE], MAIL_FILES[CORREO] = """ & mailName & """)")
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

	Set conversation = outlookReportFolderRef.Items.Restrict("[Subject] = '" & conversationSubject & "'").item(1).ReplyAll

	For Each folder In foldersToSearch
		For Each fileEnding In fileEndings
			quantityOfFilesFound = 0

			filePath = Dir(folder & "*.*")

			Do While filePath <> ""
				If InStr(filePath, CStr(fileEnding)) > 0 Then
					conversation.Attachments.Add folder & filePath

					quantityOfFilesFound = quantityOfFilesFound + 1
				End If

				filePath = Dir()
			Loop
			If quantityOfFilesFound = 0 Then
				draftsNotGenerated.Add mailName

				AppendToLogsFile ("No se puede crear el borrador: '" & mailName & "' porque no hay archivos a generar.")

				Exit Sub
			End If
		Next fileEnding
	Next folder

	conversation.Body = "MENSAJE " & executionMode & ". Anexo reporte. Saludos"

	conversation.Save

	AppendToLogsFile ("El borrador: '" & mailName & "' fue creado exitosamente.")
	Exit Sub

	ErrorHandler:
	AppendToLogsFile ("Ha ocurrido un error al crear el borrador: '" & mailName & "'.")
End Sub

Sub SendAllDrafts()
	If executionMode = "MANUAL" Then
		If Not IsConversationColumnCorrect Then Exit Sub
	End If
	
	Call AppendToLogsFile("Enviando borradores...")

	SendAllDraftsRecursive(1)

	If executionMode = "MANUAL" Then
		If IsNull(conversationsNotSent) Then
			outputMesssage = outputMesssage & "Ha ocurrido un error durante el envío de correos."
		ElseIf conversationsNotSent.Count = 0 Then
			outputMesssage = outputMesssage & "Correos enviados exitosamente."
		Else
			outputMesssage = "Los correos con asunto:" & vbCrLf & vbCrLf

			For Each conversation In conversationsNotSent
				outputMesssage = outputMesssage & conversation & vbCrLf
			Next conversation
			
			outputMesssage = outputMesssage & vbCrLf

			outputMesssage = outputMesssage & " no se pudieron enviar." & vbCrLf & vbCrLf
		End If

		MsgBox outputMesssage
	End If
End Sub

Sub SendAllDraftsRecursive(attemptCount As Long)
	Dim mailItem As Object
	Dim i As Long
	
	On Error GoTo ErrHandler

	If outlookDraftsFolderRef.Items.Count = 0 Then
		Call AppendToLogsFile("No hay borradores que enviar.")
		If executionMode = "MANUAL" Then MsgBox "No hay borradores que enviar."
		Exit Sub
	End If

	For Each conversation In PARAMETERS.Evaluate("FILTER(MAIL[" & GetConversationMailColumnName() & "], MAIL[" & GetGenerateMailColumnName() & "] = """ & Split(GetYesNoInCurrentLanguage(), ",")(0) & """)")
		On Error Goto mailItemNotFound
		Set mailItem = outlookDraftsFolderRef.Items.Restrict("[Subject] = '" & CStr(conversation) & "'").item(1)

		If mailItem.MessageClass = "IPM.Note" And Not mailItem.Sent Then
			mailItem.Display False
			DoEvents
			mailItem.Send
		End If
		Goto continueLoop

		mailItemNotFound:
		Call AppendToLogsFile("La conversación: '" & CStr(conversation) & "' no fue encontrada.")
		If executionMode = "MANUAL" Then MsgBox "La conversación: '" & CStr(conversation) & "' no fue encontrada."

		conversationsNotSent.Add conversation

		continueLoop:
	Next conversation

	Application.Wait Now + TimeValue("00:00:30")

	If outlookDraftsFolderRef.Items.Count > 0 Then
		for i = outlookDraftsFolderRef.Items.Count To 1 Step -1
			outlookDraftsFolderRef.Items(i).Delete
		Next
	End If

	Exit Sub

	ErrHandler:
	If attemptCount = attemptMaxCount Then
		Call AppendToLogsFile("El intento número " & attemptCount & " ha sido agotado. Envío de correos abortado.")

		Set conversationsNotSent = Null
		Exit Sub
	End If
	
	Call AppendToLogsFile("Ha ocurrido un error al enviar los borradores en el intento número " & attemptCount & ".")

	Call SendAllDraftsRecursive(attemptCount + 1)
End Sub