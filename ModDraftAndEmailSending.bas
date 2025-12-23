Attribute VB_Name = "ModDraftAndEmailSending"
Sub CreateDrafts()
	Dim colGENERAR_REPORTE As String
	Dim colNOMBRE As String

	If Not isConversationColumnCorrect Then Exit Sub

	For Each row In tbl_CORREOS.DataBodyRange.Rows
		colGENERAR_REPORTE = row.Cells(1, tbl_CORREOS.ListColumns("GENERAR CORREO?").Index).Value
		colNOMBRE = row.Cells(1, tbl_CORREOS.ListColumns("NOMBRE").Index).Value

		If colGENERAR_REPORTE = "SI" Then
			Call CreateDraft(colNOMBRE)
		End If
	Next row

	If executionMode = "MANUAL" Then MsgBox "Borradores creados correctamente."
End Sub

Sub CreateDraft(mailName As String)
	On Error GoTo ErrorHandler

	Call AppendToLogsFile("Creando borrador: " & mailName & "...")

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
				AppendToLogsFile ("No se puede crear el borrador: " & mailName & " porque no hay archivos a generar.")

				Exit Sub
			End If
		Next fileEnding
	Next folder

	conversation.Body = "MENSAJE " & executionMode & ". Anexo reporte. Saludos"

	conversation.Save

	AppendToLogsFile ("El borrador: " & mailName & " fue creado exitosamente.")
	Exit Sub
ErrorHandler:
	AppendToLogsFile ("Ha ocurrido un error al crear el borrador: " & mailName)

	continueExecution = False
End Sub

Sub SendAllDrafts()
	Call AppendToLogsFile("Enviando borradores...")

	If executionMode = "AUTOMÁTICO" Then OpenOutlookIfNotRunning
	SendAllDraftsRecursive(1)
End Sub

Sub SendAllDraftsRecursive(attemptCount As Long)
	On Error GoTo ErrHandler

	For Each item In outlookDraftsFolderRef.Items
		If item.MessageClass = "IPM.Note" And Not item.Sent Then
			item.Display False
			DoEvents
			item.Send
		End If
	Next item

	Call AppendToLogsFile("Correos enviados exitosamente.")

	If executionMode = "MANUAL" Then MsgBox "Correos enviados exitosamente."

	Exit Sub
ErrHandler:
	If attemptCount = attemptMaxCount Then
		Call AppendToLogsFile("El intento " & attemptCount & " ha sido agotado. Envío de correos abortado.")

		If executionMode = "MANUAL" Then MsgBox "Ha ocurrido un error al enviar los correos."
		
		continueExecution = False
		Exit Sub
	End If
	
	Call AppendToLogsFile("Ha ocurrido un error al enviar los borradores en el intento " & attemptCount & ".")

	Call SendAllDraftsRecursive(attemptCount + 1)
End Sub