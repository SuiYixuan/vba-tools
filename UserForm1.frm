VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ペースト模式"
   ClientHeight    =   1200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2355
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Call EndMonitor
    Call StartMonitor
End Sub

Private Sub CommandButton2_Click()
    Call EndMonitor
    MsgBox "Clipboard Monitor End", vbInformation
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If isMonitoring Then
            Cancel = True
            Me.Hide
        Else
            Cancel = False
        End If
    End If
End Sub
