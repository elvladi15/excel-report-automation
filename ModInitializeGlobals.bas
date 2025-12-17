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
    Set dictParameters = CreateObject("Scripting.Dictionary")
    
    Dim keyArr As Variant
    Dim valueArr As Variant
    
    keyArr = PARAMETROS.ListObjects("PARAMETROS").ListColumns("NOMBRE").DataBodyRange.Value
    valueArr = PARAMETROS.ListObjects("PARAMETROS").ListColumns("VALOR").DataBodyRange.Value

    Dim i As Long
    For i = 1 To UBound(keyArr, 1)
        If valueArr(i, 1) = "" Then
            MsgBox "El valor del parámetro " & keyArr(i, 1) & " no puede quedar vacío."
            isInputValidationCorrect = False
            Exit Function
        End If

        If keyArr(i, 1) Like "Directorio*" And Dir(CStr(valueArr(i, 1)), vbDirectory) = "" Then
            MsgBox "El directorio del parámetro " & keyArr(i, 1) & " no existe. Favor de validar ruta."
            isInputValidationCorrect = False
            Exit Function
        End If
        
        dictParameters.Add keyArr(i, 1), valueArr(i, 1)
    Next i

    isInputValidationCorrect = True
End Function

