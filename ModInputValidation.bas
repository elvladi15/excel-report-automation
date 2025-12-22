Attribute VB_Name = "ModInputValidation"
Function isInputValidationCorrect() As Boolean
	Set dictParameters = CreateObject("Scripting.Dictionary")

	On Error Resume Next
	Set wsPARAMETROS = ThisWorkbook.Sheets("PARAMETROS")
	On Error GoTo 0

	If Err.Number <> 0 Then
		MsgBox "La hoja de cálculo PARÁMETROS no existe. Favor revisar nombres de las hojas."
		Exit Function
	End If

	If Not isBasicTableStructureCorrent() Then Exit Function
	If Not isParameterValidationCorrect() Then Exit Function
	If Not validateAllBasicTableContents() Then Exit Function
	If Not isPowerQueryWorksheetAndTableValidationCorrect() Then Exit Function

	isInputValidationCorrect = True
End Function

Function isBasicTableStructureCorrent() As Boolean
	Dim tableObject As ListObject
	Dim columnObject As ListColumn
	
	Set basicTableStructure = GetBasicTableStructure()

	For Each table in basicTableStructure("tables")
		On Error Resume Next
		Set tableObject = wsPARAMETROS.ListObjects(table("name"))
		On Error GoTo 0

		If Err.Number <> 0 Then
			MsgBox "La tabla " & table("name") & " no existe. Favor revisar nombres internos de las tablas."
			Exit Function
		End If

		For Each column in table("columns")
			On Error Resume Next
			Set columnObject = tableObject.ListColumns(column("name"))
			On Error GoTo 0

			If Err.Number <> 0 Then
				MsgBox "La columna " & column("name") & " de la tabla " & table("name") & " no existe. Favor revisar nombres."
				Exit Function
			End If

				For Each row in column("rows")
					If IsError(Application.Match(row, Range(table("name") & "[" & column("name") & "]"), 0)) Then
						MsgBox "El valor " & row & ", columna " & column("name") & ", tabla " & table("name") & " no existe. Favor revisar nombres."
						Exit Function
					End If
				Next row
		Next column
	Next table

	isBasicTableStructureCorrent = True
End Function

Function isParameterValidationCorrect() As Boolean
	Dim colNOMBRE As String
	Dim colVALOR As String

	Set tbl_PARAMETROS = wsPARAMETROS.ListObjects("PARAMETROS")
	Set tbl_CORREOS = wsPARAMETROS.ListObjects("CORREOS")
	Set tbl_ARCHIVOS = wsPARAMETROS.ListObjects("ARCHIVOS")
	Set tbl_REPORTES = wsPARAMETROS.ListObjects("REPORTES")

	For Each row In tbl_PARAMETROS.DataBodyRange.Rows
		colNOMBRE = row.Cells(1, tbl_PARAMETROS.ListColumns("NOMBRE").Index).Value
		colVALOR = row.Cells(1, tbl_PARAMETROS.ListColumns("VALOR").Index).Value

		dictParameters.Add colNOMBRE, colVALOR
	Next row

	For Each row In tbl_PARAMETROS.DataBodyRange.Rows
		colNOMBRE = row.Cells(1, tbl_PARAMETROS.ListColumns("NOMBRE").Index).Value
		colVALOR = row.Cells(1, tbl_PARAMETROS.ListColumns("VALOR").Index).Value

		If (colNOMBRE = "START_PROCESS_DATE" Or colNOMBRE = "END_PROCESS_DATE") And Not IsDate(colVALOR) Then
			MsgBox "El valor del parámetro " & colNOMBRE & " debe ser una fecha válida."
			Exit Function
		End If

		If colNOMBRE = "Directorio archivos de logs" And dictParameters("Generar logs") = "NO" Then GoTo continueLoop

		If colVALOR = "" Then
			MsgBox "El valor del parámetro " & colNOMBRE & " no puede quedar vacío."
			Exit Function
		End If

		If colNOMBRE Like "Directorio*" Then
			If Dir(colVALOR, vbDirectory) = "" Then
				MsgBox "El directorio del parámetro " & colNOMBRE & " no existe. Favor de validar ruta."
				Exit Function
			End If

			If Right(colVALOR, 1) = "\" Then
				MsgBox "El directorio del parámetro " & colVALOR & " contiene el caracter \ al final. Favor de remover."
				Exit Function
			End If
		End If

		continueLoop:
	Next row

	startProcessDate = CDate(dictParameters("START_PROCESS_DATE"))
	endProcessDate = CDate(dictParameters("END_PROCESS_DATE"))
	baseReportFolder = dictParameters("Directorio base reportes")
	logsFileFolder = dictParameters("Directorio archivos de logs")
	outlookFolderName = dictParameters("Carpeta de Outlook")
	dateFormat = dictParameters("Formato de fechas")
	canGenerateLogs = dictParameters("Generar logs?") = "SI"

	isParameterValidationCorrect = True
End Function

Function validateAllBasicTableContents() As Boolean
	If Not validateBasicTableContent(tbl_CORREOS) Then Exit Function
	If Not validateBasicTableContent(tbl_ARCHIVOS) Then Exit Function
	If Not validateBasicTableContent(tbl_REPORTES) Then Exit Function

	validateAllBasicTableContents = True
End Function

Function validateBasicTableContent(table As ListObject)
	Dim atLeast1MailToGenerate As Boolean

	atLeast1MailToGenerate = False

	If table.ListRows.Count = 0 Then
		MsgBox "La tabla " & table.Name & " está vacía."
		Exit Function
	End If

	For Each column in table.ListColumns
		For Each cell in column.DataBodyRange
			If cell.Value = "" Then
				MsgBox "Hay valores vacíos en la tabla " & table.Name & "."
				Exit Function
			End If


			If table.Name = "CORREOS" Then
				If column.Name = "UN ARCHIVO POR RANGO?" Then
					GoTo continueLoop
				End If

				If column.Name = "GENERAR CORREO?" Then
					If cell.Value = "SI" Then
						atLeast1MailToGenerate = True
					End If

					GoTo continueLoop
				End If
			End If

			If table.Name = "REPORTES" Then GoTo continueLoop

			If Application.CountIf(column.DataBodyRange, cell.Value) > 1 Then
				MsgBox "Hay valores duplicados en la columna " & column.Name & " de la tabla " & table.Name & "."
				Exit Function
			End If
			continueLoop:
		Next cell
	Next column
	
	If table.Name = "CORREOS" And Not atLeast1MailToGenerate Then
		MsgBox "Debe haber al menos 1 correo a generar."
		Exit Function
	End If

	validateBasicTableContent = True
End Function

Function isPowerQueryWorksheetAndTableValidationCorrect() As Boolean
	Dim Worksheet As Worksheet
	Dim table As ListObject
	Dim columnExists As Boolean
	Dim colNOMBRE As String

	For Each row In tbl_REPORTES.DataBodyRange.Rows
		colNOMBRE = row.Cells(1, tbl_REPORTES.ListColumns("NOMBRE").Index).Value

		On Error Resume Next
		Set Worksheet = ThisWorkbook.Worksheets(colNOMBRE)
		On Error GoTo 0
		If Err.Number <> 0 Then
			MsgBox "La hoja de cálculo " & colNOMBRE & " no existe. Favor crearla junto a su tabla de Power Query."
			Exit Function
		End If

		On Error Resume Next
		Set table = Worksheet.ListObjects(colNOMBRE)
		On Error GoTo 0
		If Err.Number <> 0 Then
			MsgBox "La tabla " & colNOMBRE & " no fue encontrada en su respectiva hoja de cálculo. Favor crear."
			Exit Function
		End If

		On Error Resume Next
		colNOMBRE = table.ListColumns("PROCESS_DATE_FOR_RANGE")
		On Error GoTo 0
		If Err.Number <> 0 Then
			MsgBox "La columna PROCESS_DATE_FOR_RANGE no fue encontrada en la tabla " & colNOMBRE & ". Favor crear."
			Exit Function
		End If
	Next row
	isPowerQueryWorksheetAndTableValidationCorrect = True
End Function

Function isConversationColumnCorrect() As Boolean
	Dim colCONVERSACION As String

	Set outlookAppRef = CreateObject("Outlook.Application").GetNamespace("MAPI")
	Set outlookReportFolderRef = outlookAppRef.GetDefaultFolder(6).Parent.Folders(outlookFolderName)
	Set outlookDraftsFolderRef = outlookAppRef.GetDefaultFolder(16)

	For each conversation in tbl_CORREOS.ListColumns("CONVERSACION").DataBodyRange.Cells
		If Not outlookReportFolderRef.Items.Restrict("[Subject] = '" & conversation.Value & "'").Count > 0 Then
			MsgBox "La conversación " & conversation.Value & " no existe."
			Exit Function
		End If
	Next conversation

	isConversationColumnCorrect = True
End Function