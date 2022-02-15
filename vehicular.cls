VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CheckBox1_Click()
 
If Me.CheckBox1 = True Then
    Call COLOR
    Else
    Call blank
End If
    
 
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

Dim HojaLista As Worksheet
Dim RangoLista As Range
Dim MiRangoLista As Range
Dim RangoCon As Range
Dim Fac As Range
Dim RangoBus As Range
Dim RANGOCHOF As Range



Set Fac = ThisWorkbook.Sheets("Datos_hm").Range("D:G") ' buscarv
Set HojaLista = ThisWorkbook.Worksheets("PAQS")
Set RangoLista = HojaLista.Range("Paqs[PAQUETERIAS]")
Set RANGOCHOF = HojaLista.Range("Choferes[Choferes]")
Set MiRangoLista = Intersect(Target, Sheets("HM").ListObjects("HM").ListColumns(7).Range) ' qty embarcada
Set RangoCon = Intersect(Target, Sheets("HM").ListObjects("HM").ListColumns(12).Range)
Set RangoBus = Intersect(Target, Sheets("HM").ListObjects("HM").ListColumns(3).Range)
Set foto = Intersect(Target, Sheets("HM").ListObjects("HM").ListColumns(1).Range)



'buscarv
If Not RangoBus Is Nothing Then
   On Error Resume Next
    If Target = Empty Then
    Cells(Target.Row, 4).Value = Empty: Cells(Target.Row, 5).Value = Empty: Cells(Target.Row, 1).Value = Empty
        Else
        If Len(Target) > 110000 Then
        Cells(Target.Row, 4).Value = "N/A": Cells(Target.Row, 5).Value = "TRANSFERENCIA": _
        Application.ErrorCheckingOptions.OmittedCells = False: Cells(Target.Row, 1).Value = "=SUBTOTAL(3,R2C3:RC[2])"
            Else
            If Len(Target) = 7 Then
            Cells(Target.Row, 4).Value = "N/A": Cells(Target.Row, 5).Value = "MUSIC CLUB INTERNATIONAL": _
            Application.ErrorCheckingOptions.OmittedCells = False: Cells(Target.Row, 1).Value = "=SUBTOTAL(3,R2C3:RC[2])"
                Else
                If Target.Value <= 20000 Then
                Cells(Target.Row, 4).Value = "N/A": Cells(Target.Row, 5).Value = "TRANSFERENCIA HM": _
                Application.ErrorCheckingOptions.OmittedCells = False: Cells(Target.Row, 1).Value = "=SUBTOTAL(3,R2C3:RC[2])"
                    Else
                    On Error Resume Next
                    Cells(Target.Row, 4).Value = Application.WorksheetFunction.VLookup(Target, Fac, 2, False): _
                    Application.ErrorCheckingOptions.OmittedCells = False: Cells(Target.Row, 1).Value = "=SUBTOTAL(3,R2C3:RC[2])"
                    If Err.Number > 0 Then Cells(Target.Row, 4).Value = Empty
        
                    Cells(Target.Row, 5).Value = Application.WorksheetFunction.VLookup(Target, Fac, 4, False): _
                    Application.ErrorCheckingOptions.OmittedCells = False: Cells(Target.Row, 1).Value = "=SUBTOTAL(3,R2C3:RC[2])"
                    If Err.Number > 0 Then Cells(Target.Row, 5).Value = Empty

            Exit Sub
            End If
        End If
    End If
 End If
End If

'tabla
If Not MiRangoLista Is Nothing Then
        On Error Resume Next
        If Target = Empty Then
        Cells(Target.Row, 9).Value = Empty: Cells(Target.Row, 9).Validation.Delete
        Cells(Target.Row, 8).Value = Empty: Cells(Target.Row, 8).Validation.Delete
        Else
            With Cells(Target.Row, 9).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="='" & HojaLista.Name & "'!" & RangoLista.Address
            End With
             With Cells(Target.Row, 8).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="='" & HojaLista.Name & "'!" & RANGOCHOF.Address
            End With
        End If
End If

'concatenar
If Not RangoCon Is Nothing Then
    On Error Resume Next
    Select Case Target
        Case Is <> Empty
        Cells(Target.Row, 13).Value = Application.WorksheetFunction.Concat _
        (Cells(Target.Row, 11).Value, "_", Cells(Target.Row, 12).Value)
        ', "_", VBA.Format(Cells(Target.Row, 2).Value, "dd-mmm-yy"))
            With Cells(Target.Row, 9).Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
            :=xlBetween
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
            End With
            With Cells(Target.Row, 8).Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
            :=xlBetween
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
            End With
        Case Empty
        Cells(Target.Row, 13).Value = Empty
    End Select
    
End If
'If Not foto Is Nothing Then



    
End Sub


