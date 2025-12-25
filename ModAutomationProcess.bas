Attribute VB_Name = "ModAutomationProcess"
Sub ScheduleAutomaticRun()
	Dim scheduleDateTime As Date

	scheduleDateTime = Date + 1 + ScheduleTime

	If sendMails Then
		If Not isConversationColumnCorrect Then Exit Sub
	End If

	Call ScheduleProcedure("AutomaticRun", scheduleDateTime)

	If executionMode = "MANUAL" Then
		If sendMails Then
			MsgBox "Programación de envío de correos exitosa. Próxima corrida: " & Format(scheduleDateTime, dateFormat & " hh:mm:ss")
		Else
			MsgBox "Programación de genereración de reportes exitosa. Próxima corrida: " & Format(scheduleDateTime, dateFormat & " hh:mm:ss")
		End If
		executionMode = "AUTOMÁTICO"
	End If
End Sub

Sub AutomaticRun()
	Call AppendToLogsFile("Cerrando los demás libros de Excel...")
	CloseAllOtherWorkbooks

	Call AppendToLogsFile("Refrescando hoja de cálculo...")
	ThisWorkbook.Sheets("PARAMETROS").Calculate

	Set wsPARAMETROS = ThisWorkbook.Sheets("PARAMETROS")
	startProcessDate = CDate(CStr(ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""START_PROCESS_DATE"", PARAMETROS[NOMBRE], PARAMETROS[VALOR])")))
	endProcessDate = CDate(CStr(ThisWorkbook.ActiveSheet.Evaluate("XLOOKUP(""END_PROCESS_DATE"", PARAMETROS[NOMBRE], PARAMETROS[VALOR])")))

	RefreshAll
	'If Not continueExecution Then
	'	Call AppendToLogsFile("Se ha abortado la ejecución debido a un error en el proceso de refrescar hoja de cálculo.")
	'	Goto scheduleNextRun
	'End If

	CreateMailFiles
	'If Not continueExecution Then
	'	Call AppendToLogsFile("Se ha abortado la ejecución debido a un error en el proceso de crear los archivos.")
	'	Goto scheduleNextRun
	'End If

	If sendMails Then
		CreateDrafts
	'	If Not continueExecution Then
	'		Call AppendToLogsFile("Se ha abortado la ejecución debido a un error en el proceso de creación de borradores.")
	'		Goto scheduleNextRun
	'	End If

		OpenOutlookIfNotRunning

		SendAllDrafts
	'	If Not continueExecution Then
	'		Call AppendToLogsFile("Ha ocurrido un error en el proceso de envío de correos.")
	'	End If
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