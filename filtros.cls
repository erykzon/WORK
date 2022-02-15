VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

Dim PROP As Range
Dim filtro As Range

Set PROP = Intersect(Target, Sheets("PRO").ListObjects("NOM_PRO").ListColumns(8).Range)
Set filtro = Intersect(Target, Sheets("PRO").ListObjects("NOM_PRO").ListColumns(15).Range)

If Not filtro Is Nothing Then
    With Range("NOM_PRO")
    .Sort key1:=.Cells(1, 3), Header:=xlYes
    End With
        Else
        If Target.Column = 8 Then
        On Error Resume Next
        Cells(Target.Row, 4).Value = Application.WorksheetFunction.Proper(Cells(Target.Row, 4)): _
        Cells(Target.Row, 7).Value = Application.WorksheetFunction.Proper(Cells(Target.Row, 7))
            Else
            If Target.Column = 13 Then
            On Error Resume Next
            Cells(Target.Row, 4).Value = Application.WorksheetFunction.Proper(Cells(Target.Row, 4)): _
            Cells(Target.Row, 7).Value = Application.WorksheetFunction.Proper(Cells(Target.Row, 7)): _
            Cells(Target.Row, 10).Value = Application.WorksheetFunction.Proper(Cells(Target.Row, 10)): _
            Cells(Target.Row, 11).Value = Application.WorksheetFunction.Proper(Cells(Target.Row, 11))
                Else
                If Target.Row = 2 And (Target.Column >= 4 And Target.Column <= 7) Then
                Application.ScreenUpdating = False
                Call MIFILTRO_pro
                Application.ScreenUpdating = True
                End If
            End If
        End If
End If

End Sub

