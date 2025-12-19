Attribute VB_Name = "ModInitializeGlobals"
Public dictParameters As Object

Public startProcessDate As Date
Public endProcessDate As Date
Public baseReportFolder As String
Public outlookFolder As String
Public logsFileFolder As String
Public selectedReport As String
Public dateFormat As String
Public canGenerateLogs As Boolean

Public canMailBeSent As Boolean
Public currentProcessDate As Variant
Public errorReport As String

Sub InitializeGlobals()
    canMailBeSent = True

    startProcessDate = CDate(dictParameters("START_PROCESS_DATE"))
    endProcessDate = CDate(dictParameters("END_PROCESS_DATE"))
    baseReportFolder = dictParameters("Directorio base reportes")
    logsFileFolder = dictParameters("Directorio archivos de logs")
    outlookFolder = dictParameters("Carpeta de Outlook")
    selectedReport = dictParameters("Reporte a generar")
    dateFormat = dictParameters("Formato de fechas")
    canGenerateLogs = dictParameters("Generar logs") = "SI"
End Sub

Function isInputValidationCorrect() As Boolean
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
    
    reportNames = Range("REPORTES[NOMBRE]")

    For Each item In reportNames
        On Error GoTo worksheetNotFound
        Set Worksheet = ThisWorkbook.Worksheets(item)

        On Error GoTo tableNotFound
        Set table = Worksheet.ListObjects(item)
        
        columnExists = False

        For Each col in table.ListColumns
            If col.Name = "PROCESS_DATE_FOR_RANGE" Then columnExists = True
        Next col

        If columnExists = False Then GoTo columnNotFound

        GoTo continueLoop

        worksheetNotFound:
        MsgBox "La hoja de cálculo " & item & " no existe. Favor crearla junto a su tabla de Power Query."
        isWorksheetAndTableValidationCorrect = False
        Exit Function

        tableNotFound:
        MsgBox "La tabla " & item & " no fue encontrada en su respectiva hoja de cálculo. Favor crear."
        isWorksheetAndTableValidationCorrect = False
        Exit Function

        columnNotFound:
        MsgBox "La columna PROCESS_DATE_FOR_RANGE no fue encontrada en la tabla " & item & ". Favor crear."
        isWorksheetAndTableValidationCorrect = False
        Exit Function

        continueLoop:
    Next item
    isWorksheetAndTableValidationCorrect = True
End Function

Function isParameterValidationCorrect() As Boolean
    Set dictParameters = CreateObject("Scripting.Dictionary")
    
    Dim keyArr As Variant
    Dim valueArr As Variant
    
    keyArr = PARAMETROS.ListObjects("PARAMETROS").ListColumns("NOMBRE").DataBodyRange.Value
    valueArr = PARAMETROS.ListObjects("PARAMETROS").ListColumns("VALOR").DataBodyRange.Value

    Dim i As Long
    For i = 1 To UBound(keyArr, 1)
        If valueArr(i, 1) = "" Then
            MsgBox "El valor del parámetro " & keyArr(i, 1) & " no puede quedar vacío."
            isParameterValidationCorrect = False
            Exit Function
        End If

        If keyArr(i, 1) Like "Directorio*" Then
            If Dir(CStr(valueArr(i, 1)), vbDirectory) = "" Then
                MsgBox "El directorio del parámetro " & keyArr(i, 1) & " no existe. Favor de validar ruta."

                isParameterValidationCorrect = False

                Exit Function
            End If

            If Right(valueArr(i, 1), 1) = "\" Then
                MsgBox "El directorio del parámetro " & keyArr(i, 1) & " contiene el caracter \ al final. Favor de remover."

                isParameterValidationCorrect = False
                
                Exit Function
            End If
        End If
        
        dictParameters.Add keyArr(i, 1), valueArr(i, 1)
    Next i

    isParameterValidationCorrect = True
End Function