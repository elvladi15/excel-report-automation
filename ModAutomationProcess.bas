Attribute VB_Name = "ModAutomationProcess"
Sub Manual_ScheduleMailSending()
	Call ScheduleMailSending("MANUAL")
End Sub
Sub Automatic_ScheduleMailSending()
	Call ScheduleMailSending("AUTOMATICO")
End Sub
Sub ScheduleMailSending(executionMode As String)
	If isInputValidationCorrect = False Then Exit Sub
	
	InitializeGlobals
	
	Dim scheduleDayIncrement As Long
	Dim scheduleDate As Date
	
	scheduleDayIncrement = 0
	
	If TimeValue(Now) >= TimeValue("06:00:00") Then
		scheduleDayIncrement = 1
	End If
	
	If Weekday(Now, vbMonday) = 5 Then
		scheduleDayIncrement = 3
	End If
	
	scheduleDate = Date + scheduleDayIncrement
	
	Call ScheduleProcedure("Automatic_RefreshAll", scheduleDate + TimeValue("06:45:00"))
	Call ScheduleProcedure("Automatic_CreateMailFiles", scheduleDate + TimeValue("06:50:00"))
	Call ScheduleProcedure("Automatic_CreateDrafts", scheduleDate + TimeValue("06:55:00"))
	Call ScheduleProcedure("OpenOutlookIfNotRunning", scheduleDate + TimeValue("07:00:00"))
	Call ScheduleProcedure("Automatic_SendAllDrafts", scheduleDate + TimeValue("07:05:00"))
	
	If executionMode = "MANUAL" Then MsgBox "Programación exitosa. Próxima corrida: " & Format(scheduleDate, dateFormat)
End Sub

Sub ScheduleProcedure(procedure As String, time As Date)
	On Error GoTo Schedule
	Application.OnTime EarliestTime:=time, procedure:=procedure, Schedule:=False
	
Schedule:
	Application.OnTime EarliestTime:=time, procedure:=procedure, Schedule:=True

	Call AppendToLogsFile(Format(Now, "yyyy-MM-dd hh:mm:ss") & " - Procedimiento " & procedure & " programado exitosamente para " & Format(time, dateFormat))
End Sub