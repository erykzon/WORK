VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmValidaciones 
   Caption         =   "EXCELeINFO"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3885
   OleObjectBlob   =   "frmValidaciones.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmValidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
'
    Unload Me
    '
End Sub

Private Sub txtFilas_Change()

End Sub

Private Sub txtNumero_Change()
'
    Me.txtNumero.Value = SoloNumero(Me.txtNumero.Value)
    '
End Sub

Private Sub txtNumeroDecimal_Change()

    Me.txtNumeroDecimal.Value = SoloNumeroDecimal(Me.txtNumeroDecimal.Value)
    '
End Sub

Private Sub txtTexto_Change()
'
    Me.txtTexto.Value = SoloTexto(Me.txtTexto.Value)
    '
End Sub

'
Private Sub UserForm_Initialize()
'
    Call FormDesign
    '
End Sub
