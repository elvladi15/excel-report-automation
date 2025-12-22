Attribute VB_Name = "ModAutomationProcess"
Sub ScheduleMailSending()
	Dim scheduleDayIncrement As Long
	Dim scheduleDate As Date

	If Not isConversationColumnCorrect Then Exit Sub

	scheduleDayIncrement = 1

	scheduleDate = Date + scheduleDayIncrement

	Call ScheduleProcedure("RefreshAll", scheduleDate + TimeValue("06:45:00"))
	Call ScheduleProcedure("CreateMailFiles", scheduleDate + TimeValue("06:50:00"))
	Call ScheduleProcedure("CreateDrafts", scheduleDate + TimeValue("06:55:00"))
	Call ScheduleProcedure("SendAllDrafts", scheduleDate + TimeValue("07:00:00"))

	If executionMode = "AUTOMÁTICO" Then
		Call ScheduleProcedure("ScheduleMailSending", scheduleDate + TimeValue("07:02:00"))
	ElseIf executionMode = "MANUAL" Then
		MsgBox "Programación de envío de correos exitosa. Próxima corrida: " & Format(scheduleDate, dateFormat)
		executionMode = "AUTOMÁTICO"
	End If
End Sub

Sub ScheduleMailGeneration()
	Dim scheduleDayIncrement As Long
	Dim scheduleDate As Date

	scheduleDayIncrement = 1

	scheduleDate = Date + scheduleDayIncrement

	Call ScheduleProcedure("RefreshAll", scheduleDate + TimeValue("06:55:00"))
	Call ScheduleProcedure("CreateMailFiles", scheduleDate + TimeValue("07:00:00"))

	If executionMode = "AUTOMÁTICO" Then
		Call ScheduleProcedure("ScheduleMailSending", scheduleDate + TimeValue("07:02:00"))
	ElseIf executionMode = "MANUAL" Then
		MsgBox "Programación de generación de correo exitosa. Próxima corrida: " & Format(scheduleDate, dateFormat)
		executionMode = "AUTOMÁTICO"
	End If
End Sub

Sub ScheduleProcedure(procedure As String, time As Date)
	On Error GoTo Schedule
	Application.OnTime EarliestTime:=time, procedure:=procedure, Schedule:=False
Schedule:
	Application.OnTime EarliestTime:=time, procedure:=procedure, Schedule:=True

	Call AppendToLogsFile(Format(Now, "yyyy-MM-dd hh:mm:ss") & " - Procedimiento " & procedure & " programado exitosamente para " & Format(time, dateFormat))
End Sub