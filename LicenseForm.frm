VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LicenseForm 
   Caption         =   "Missing License"
   ClientHeight    =   1470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LicenseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LicenseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ls As String

Private Sub CommandButton1_Click()
ls = Trim(TextBox1)
Me.Hide
End Sub

Private Sub CommandButton2_Click()
Unload Me

End Sub

Public Property Get GetLicense() As String
    GetLicense = ls
End Property


Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        KeyCode = 0
        ls = Trim(TextBox1)
        Me.Hide
End Select
End Sub

Private Sub UserForm_Click()

End Sub
