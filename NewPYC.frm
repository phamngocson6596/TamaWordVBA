VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewPYC 
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5550
   OleObjectBlob   =   "NewPYC.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewPYC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim docNew As Document

Private Sub runbt_Click()

    
    
   ' Application.Visible = True

   
    Documents.Open fileName:="Z:\z.kh\PYC_BM.docx"

    
    Set docNew = Documents("PYC_BM.docx")

    Call xulythongtin
    
    docNew.PrintOut Copies:=soluongtb.Text
    
    docNew.Close False
    Set docNew = Nothing

    
    Unload Me
    
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  ' Set docNew = Documents("PYC - LHA - Master.docx")
    
End Sub


Private Sub UserForm_Initialize()

NoidungTextbox.SetFocus
NoidungTextbox.SelStart = 1

End Sub

Private Sub xulythongtin()

Dim a As Bookmark

For Each a In docNew.Bookmarks

    Select Case a.Name
    
        Case "Ten"
            If Not TenTextbox = "" Then
            a.Range.Font.Bold = vbRed
            a.Range.Text = TenTextbox
            End If
        Case "Diachi"
            If Not DCTextbox = "" Then a.Range.Text = DCTextbox
        Case "SDT"
            If Not SDTTextbox = "" Then a.Range.Text = SDTTextbox
        Case "Noidung"
            If Not NoidungTextbox = "" Then a.Range.Text = NoidungTextbox
    End Select

Next

End Sub

