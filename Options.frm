VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Options 
   Caption         =   "UserForm1"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3465
   OleObjectBlob   =   "Options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub default_Click()
    topAsBase.Value = False
    exportChart.Value = True
    scale_pixel.Value = ""
    scale_um.Value = ""
    scale_pixel.SetFocus
    Me.Hide
End Sub

Private Sub ok_Click()
    Me.Hide
End Sub


Private Sub scale_pixel_Change()
    If !IsNumeric(scale_pixel.Text) Then ok.Enabled = False
    ok.Enabled = True
End Sub


Private Sub scale_um_Change()
    If !IsNumeric(scale_um.Text) Then ok.Enabled = False
    ok.Enabled = True
End Sub

Private Sub UserForm_Initialize()
    topAsBase.Value = False
    exportChart.Value = True
    scale_pixel.Value = ""
    scale_um.Value = ""
    scale_pixel.SetFocus
End Sub
