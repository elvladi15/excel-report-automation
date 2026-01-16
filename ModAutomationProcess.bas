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
		mails = PARAMETERS.Evaluate("FILTER(" & tbl_MAILS.ListColumns(1).DataBodyRange.Address & ", " & tbl_MAILS.ListColumns(3).DataBodyRange.Address & " = """ & Split(tbl_MAILS.ListColumns(3).DataBodyRange.Validation.Formula1, ",")(0) & """)")
		mailCount = UBound(mails) - LBound(mails) + 1

		If sendMails Then
			MsgBox AutomationProcessMailScheduleSuccessMessage(CStr(mailCount), Format(scheduleDateTime, dateFormat & " hh:mm:ss"))
		Else
			MsgBox AutomationProcessReportScheduleSuccessMessage(CStr(mailCount), Format(scheduleDateTime, dateFormat & " hh:mm:ss"))
		End If
		executionMode = "AUTOMATIC"
	End If
End Sub

Sub AutomaticRun()
	Call AppendToLogsFile(AutomationProcessClosingOtherWorkbooksMessage())
	CloseAllOtherWorkbooks

	Call AppendToLogsFile(AutomationProcessRefreshingWorksheetMessage() & "...")
	PARAMETERS.Calculate

	startProcessDate = CDate(tbl_PARAMETERS.ListRows(2).Range.Cells(2).Value)
	endProcessDate = CDate(tbl_PARAMETERS.ListRows(3).Range.Cells(2).Value)

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

	Call AppendToLogsFile(Format(Now, "yyyy-MM-dd hh:mm:ss") & " - " & AutomationProcessProcedureScheduledMessage(procedure, Format(time, dateFormat & " hh:mm:ss")))
End Sub