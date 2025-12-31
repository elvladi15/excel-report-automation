Attribute VB_Name = "ModAutomationProcess"
Sub ScheduleAutomaticRun()
	Dim scheduleDateTime As Date
	Dim mails As Variant
	Dim mailCount As Long

	scheduleDateTime = Date + 1 + ScheduleTime

	If sendMails Then
		If Not IsConversationColumnCorrect Then Exit Sub
	End If

	Call ScheduleProcedure("AutomaticRun", scheduleDateTime)

	If executionMode = "MANUAL" Then	
		mails = PARAMETERS.Evaluate("FILTER(MAILS[NOMBRE], MAILS[GENERAR CORREO?] = ""SI"")")
		mailCount = UBound(mails) - LBound(mails) + 1

		If sendMails Then
			MsgBox "Programación de envío de correos exitosa. Se enviarán " & mailCount & " correos. Próxima corrida: " & Format(scheduleDateTime, dateFormat & " hh:mm:ss")
		Else
			MsgBox "Programación de genereración de reportes exitosa. Se generarán los archivos de " & mailCount & " correos. Próxima corrida: " & Format(scheduleDateTime, dateFormat & " hh:mm:ss")
		End If
		executionMode = "AUTOMATIC"
	End If
End Sub

Sub AutomaticRun()
	Call AppendToLogsFile("Cerrando los demás libros de Excel...")
	CloseAllOtherWorkbooks

	Call AppendToLogsFile("Refrescando hoja de cálculo...")
	PARAMETERS.Calculate

	startProcessDate = CDate(CStr(PARAMETERS.Evaluate("XLOOKUP(""" & GetStartProcessDateParameterName() & """, PARAMETERS[" & GetNameParameterColumnName() & "], PARAMETERS[" & GetValueParameterColumnName() & "])")))
	endProcessDate = CDate(CStr(PARAMETERS.Evaluate("XLOOKUP(""" & GetEndProcessDateParameterName() & """, PARAMETERS[" & GetNameParameterColumnName() & "], PARAMETERS[" & GetValueParameterColumnName() & "])")))

	RefreshAll

	CreateMailFiles

	If sendMails Then
		CreateDrafts

		OpenOutlookIfNotRunning

		SendAllDrafts
	End If

	scheduleNextRun:
	Call ScheduleProcedure("AutomaticRun", Date + 1 + scheduleTime)
End Sub

Sub ScheduleProcedure(procedure As String, time As Date)
	On Error GoTo Schedule
	Application.OnTime EarliestTime:=time, procedure:=procedure, Schedule:=False
	
	Schedule:
	Application.OnTime EarliestTime:=time, procedure:=procedure, Schedule:=True

	Call AppendToLogsFile(Format(Now, "yyyy-MM-dd hh:mm:ss") & " - Procedimiento " & procedure & " programado exitosamente para " & Format(time, dateFormat & " hh:mm:ss"))
End Sub