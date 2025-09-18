VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "RenameAllSheets"
   ClientHeight    =   1065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3060
   OleObjectBlob   =   "RenameAllSheetsUserForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Call undoRename
    'Call saveUndoStackToStorage
End Sub

Private Sub CommandButton2_Click()
    Call renameAllSheets(TextBox1.Text)
    'Call saveUndoStackToStorage
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 0 Then
         Call renameAllSheets(TextBox1.Text)
         'Call saveUndoStackToStorage
    End If
End Sub
