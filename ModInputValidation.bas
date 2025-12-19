Attribute VB_Name = "ModInputValidation"
Function isInputValidationCorrect() As Boolean
	Set dictParameters = CreateObject("Scripting.Dictionary")

	Set tbl_PARAMETROS = PARAMETROS.ListObjects("PARAMETROS")
	Set tbl_CORREOS = PARAMETROS.ListObjects("CORREOS")
	Set tbl_ARCHIVOS = PARAMETROS.ListObjects("ARCHIVOS")
	Set tbl_REPORTES = PARAMETROS.ListObjects("REPORTES")

	If isWorksheetAndTableValidationCorrect = False Then
		isInputValidationCorrect = False
		Exit Function
	ElseIf isParameterValidationCorrect = False Then
		isInputValidationCorrect = False
		Exit Function
	End If
	isInputValidationCorrect = True
End Function

Function isWorksheetAndTableValidationCorrect() As Boolean
	Dim reportNames As Variant
	Dim Worksheet As Worksheet
	Dim table As ListObject
	Dim columnExists As Boolean
	Dim colNOMBRE As String

	reportNames = Range("REPORTES[NOMBRE]")

	For Each row In tbl_REPORTES.DataBodyRange.Rows
		colNOMBRE = row.Cells(1, tbl_REPORTES.ListColumns("NOMBRE").Index).Value

		On Error GoTo worksheetNotFound
		Set Worksheet = ThisWorkbook.Worksheets(colNOMBRE)

		On Error GoTo tableNotFound
		Set table = Worksheet.ListObjects(colNOMBRE)

		On Error GoTo columnNotFound
		colNOMBRE = table.ListColumns("PROCESS_DATE_FOR_RANGE")

		GoTo continueLoop

		worksheetNotFound:
		MsgBox "La hoja de cálculo " & colNOMBRE & " no existe. Favor crearla junto a su tabla de Power Query."
		isWorksheetAndTableValidationCorrect = False
		Exit Function

		tableNotFound:
		MsgBox "La tabla " & colNOMBRE & " no fue encontrada en su respectiva hoja de cálculo. Favor crear."
		isWorksheetAndTableValidationCorrect = False
		Exit Function

		columnNotFound:
		MsgBox "La columna PROCESS_DATE_FOR_RANGE no fue encontrada en la tabla " & colNOMBRE & ". Favor crear."
		isWorksheetAndTableValidationCorrect = False
		Exit Function

		continueLoop:
	Next row
	isWorksheetAndTableValidationCorrect = True
End Function

Function isParameterValidationCorrect() As Boolean
	Dim colNOMBRE As String
	Dim colVALOR As String

	For Each row In tbl_PARAMETROS.DataBodyRange.Rows
		colNOMBRE = row.Cells(1, tbl_PARAMETROS.ListColumns("NOMBRE").Index).Value
		colVALOR = row.Cells(1, tbl_PARAMETROS.ListColumns("VALOR").Index).Value

		dictParameters.Add colNOMBRE, colVALOR
	Next row

	For Each row In tbl_PARAMETROS.DataBodyRange.Rows
		colNOMBRE = row.Cells(1, tbl_PARAMETROS.ListColumns("NOMBRE").Index).Value
		colVALOR = row.Cells(1, tbl_PARAMETROS.ListColumns("VALOR").Index).Value

		If colNOMBRE = "Directorio archivos de logs" And dictParameters("Generar logs") = "NO" Then GoTo continueLoop

		If colVALOR = "" Then
			MsgBox "El valor del parámetro " & colNOMBRE & " no puede quedar vacío."
			isParameterValidationCorrect = False
			Exit Function
		End If


		If colNOMBRE Like "Directorio*" Then
			If Dir(colVALOR, vbDirectory) = "" Then
				MsgBox "El directorio del parámetro " & colNOMBRE & " no existe. Favor de validar ruta."
				isParameterValidationCorrect = False
				Exit Function
			End If

			If Right(colVALOR, 1) = "\" Then
				MsgBox "El directorio del parámetro " & colVALOR & " contiene el caracter \ al final. Favor de remover."

				isParameterValidationCorrect = False

				Exit Function
			End If
		End If

		continueLoop:
	Next row

	isParameterValidationCorrect = True
End Function

Function getExecutionMode() As String
	On Error GoTo Testing
	If(Application.Caller = "btnScheduleMailSending") Then
		getExecutionMode = "AUTOMÁTICO"
	Else
		getExecutionMode = "MANUAL"
	End If
	Testing:
	getExecutionMode = "MANUAL"
End Function