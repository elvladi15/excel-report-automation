Attribute VB_Name = "ModAutomationProcess"
Sub ScheduleMailSending()
	Dim scheduleDayIncrement As Long
	Dim scheduleDate As Date

	scheduleDayIncrement = 0

	If TimeValue(Now) >= TimeValue("06:00:00") Then
		scheduleDayIncrement = 1
	End If

	scheduleDate = Date + scheduleDayIncrement

	Call ScheduleProcedure("RefreshAll", scheduleDate + TimeValue("06:45:00"))
	Call ScheduleProcedure("CreateMailFiles", scheduleDate + TimeValue("06:50:00"))
	Call ScheduleProcedure("CreateDrafts", scheduleDate + TimeValue("06:55:00"))
	Call ScheduleProcedure("SendAllDrafts", scheduleDate + TimeValue("07:00:00"))

	If executionMode = "MANUAL" Then MsgBox "Programación exitosa. Próxima corrida: " & Format(scheduleDate, dateFormat)
End Sub

Sub ScheduleProcedure(procedure As String, time As Date)
	On Error GoTo Schedule
	Application.OnTime EarliestTime:=time, procedure:=procedure, Schedule:=False

Schedule:
	Application.OnTime EarliestTime:=time, procedure:=procedure, Schedule:=True

	Call AppendToLogsFile(Format(Now, "yyyy-MM-dd hh:mm:ss") & " - Procedimiento " & procedure & " programado exitosamente para " & Format(time, dateFormat))
End Sub