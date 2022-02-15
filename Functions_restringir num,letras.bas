Attribute VB_Name = "Functions"
Option Private Module
'---------------------------------------------------------------------------------------
' Module    : Functions
' Author    : MVP, Sergio Alejandro Campos
' Date      : 21/sep/2015
' Purpose   : Funciones para permitir sólo texto, número y números con decimales
'---------------------------------------------------------------------------------------
'
Sub MostrarFormulario()
'
    frmValidaciones.Show
    '
End Sub
'
Function SoloTexto(Texto As Variant)
'
    Dim Caracter As Variant
    Dim Largo As String
    On Error Resume Next
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = CInt(Mid(Texto, i, 1))
        '
        If Caracter <> "" Then
            If Not Application.WorksheetFunction.IsText(Caracter) Then
                Texto = Replace(Texto, Caracter, "")
                SoloTexto = Texto
            Else
            End If
        End If
        '
    Next i
    '
    SoloTexto = Texto
    On Error GoTo 0
    '
End Function
'
Function SoloNumero(Texto As Variant)
'
    Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = Mid(CStr(Texto), i, 1)
        '
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Texto = Replace(Texto, Caracter, "")
                SoloNumero = Texto
            Else
            End If

        End If
        '
    Next i
    '
    SoloNumero = Texto
    On Error GoTo 0
    '
End Function
'
Function SoloNumeroDecimal(Texto As Variant)
'
    Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    Punto = 0
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = Mid(CStr(Texto), i, 1)
        If Caracter <> "" Then
            '
            If Caracter = Chr(46) Then
                Punto = Punto + 1
                If Punto > 1 Then
                    Texto = WorksheetFunction.Replace(Texto, i, 1, "")
                    SoloNumeroDecimal = Texto
                    Punto = 0
                End If
            Else
                If Caracter < Chr(48) Or Caracter > Chr(57) Then
                    Texto = Replace(Texto, Caracter, "")
                    SoloNumeroDecimal = Texto
                Else
                End If
                '
            End If
            '
        End If
    Next i
    '
    SoloNumeroDecimal = Texto
    On Error GoTo 0
    '
End Function

