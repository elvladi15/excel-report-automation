Attribute VB_Name = "ModInputValidation"
Function isInputValidationCorrect() As Boolean
	'Dim wsPARAMETROS As wsPARAMETROS

	Set dictParameters = CreateObject("Scripting.Dictionary")

	On Error Resume Next
	Set wsPARAMETROS = ThisWorkbook.Sheets("PARAMETROS")

	If Err.Number <> 0 Then
		MsgBox "La hoja de cálculo PARÁMETROS no existe. Favor revisar nombres de las hojas."
		isInputValidationCorrect = False
		Exit Function
	End If

	'Set tbl_PARAMETROS = wsPARAMETROS.ListObjects("PARAMETROS")
	'Set tbl_CORREOS = wsPARAMETROS.ListObjects("CORREOS")
	'Set tbl_ARCHIVOS = wsPARAMETROS.ListObjects("ARCHIVOS")
	'Set tbl_REPORTES = wsPARAMETROS.ListObjects("REPORTES")

	If isBasicTableStructureCorrent = False Then
		isInputValidationCorrect = False
		Exit Function
	ElseIf isParameterValidationCorrect = False Then
		isInputValidationCorrect = False
		Exit Function
	ElseIf isWorksheetAndTableValidationCorrect = False Then
		isInputValidationCorrect = False
		Exit Function
	End If
	isInputValidationCorrect = True
End Function

Function isBasicTableStructureCorrent() As Boolean
	'Dim basicTableStructure As Object
	Dim tableObject As ListObject
	Dim columnObject As ListColumn
	
	Set basicTableStructure = GetBasicTableStructure()

	For Each table in basicTableStructure("tables")
		On Error Resume Next
		Set tableObject = wsPARAMETROS.ListObjects(table("key"))

		If Err.Number <> 0 Then
			MsgBox "La tabla " & table("key") & " no existe. Favor revisar nombres internos de las tablas."
			isBasicTableStructureCorrent = False
			Exit Function
		End If

		For Each column in table("columns")
			On Error Resume Next
			Set columnObject = tableObject.ListColumns(column("key"))

			If Err.Number <> 0 Then
				MsgBox "La columna " & column("key") & " de la tabla " & table("key") & " no existe. Favor revisar nombres."
				isBasicTableStructureCorrent = False
				Exit Function
			End If

			If column("rows").Count > 0 Then
				For Each row in column("rows")
					If IsError(Application.Match(row("key"), Range(table("key") & "[" & column("key") & "]"), 0)) Then
						MsgBox "El valor " & row("key") & ", columna " & column("key") & ", tabla " & table("key") & " no existe. Favor revisar nombres."
						isBasicTableStructureCorrent = False
						Exit Function
					End If
				Next row
			Else
				basicTableStructure("tables")(table(key))("columns")(column("key"))("rows").Add CreateObject("Scripting.Dictionary")
				basicTableStructure("tables")(table(key))("columns")(column("key"))("rows")("key") = 
			End If

		Next column
	Next table

	isBasicTableStructureCorrent = True
End Function

Function isParameterValidationCorrect() As Boolean
	Dim colNOMBRE As String
	Dim colVALOR As String

	Dim tableIndex As Long
	Dim columnIndex As Long
	Dim rowIndex As Long

	Set tableIndex = GetBasicTableStructureTableIndex("PARAMETROS")
	Set columnIndex = GetBasicTableStructureColumnIndex(tableIndex, "NOMBRE")

	For Each row In tbl_PARAMETROS.DataBodyRange.Rows
		colNOMBRE = row.Cells(1, tbl_PARAMETROS.ListColumns("NOMBRE").Index).Value
		colVALOR = row.Cells(1, tbl_PARAMETROS.ListColumns("VALOR").Index).Value

		rowIndex = GetBasicTableStructureRowIndex(tableIndex, columnIndex, colNOMBRE)

		'dictParameters.Add colNOMBRE, colVALOR
		basicTableStructure("tables")(tableIndex)("columns")(columnIndex)("rows")(rowIndex)("value") = colVALOR
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

Function isWorksheetAndTableValidationCorrect() As Boolean
	'Dim reportNames As Variant
	Dim Worksheet As Worksheet
	Dim table As ListObject
	Dim columnExists As Boolean
	Dim colNOMBRE As String

	'reportNames = Range("REPORTES[NOMBRE]")

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



Sub test()
	Dim conversationSubject As String
	conversationSubject = "something"
	Set Items = CreateObject("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6).Parent.Folders(outlookFolder).Items.Restrict("[Subject] = '" & conversationSubject & "'")
	'Set OutlookApp = CreateObject("Outlook.Application")
	'Set OutlookNamespace = CreateObject("Outlook.Application").GetNamespace("MAPI")
	Set Inbox = CreateObject("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6).Parent.Folders(outlookFolder)

	Set Items = Inbox.Items.Restrict("[Subject] = '" & conversationSubject & "'")
	Items.Sort "ReceivedTime", True

	If Items.Count > 0 Then
		MsgBox "ENCONTRADO: " & conversationSubject
	Else
		MsgBox "No se pudo encontrar la cadena de correos: " & conversationSubject
	End If
End Sub