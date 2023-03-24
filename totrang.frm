VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} totrang 
   Caption         =   "UserForm1"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "totrang.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "totrang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim toEnter As Byte


Private Sub CheckBox1_Change()


If CheckBox1.Value = True Then
    
    Label_startwith.Visible = True
    TextBox_startwith.Visible = True
    SpinButton_startwith.Visible = True
    
    TextBox_startwith.Value = txt_trang.Value
    

Else

    Label_startwith.Visible = False
    TextBox_startwith.Visible = False
    SpinButton_startwith.Visible = False
    
    TextBox_startwith.Value = txt_trang.Value

End If




End Sub

Private Sub CommandButton2_Click()

If txt_to.Value = "" Or txt_trang.Value = "" Then GoTo kethuc


Dim a As String
Dim b As String
Dim ketqua As String

a = Trim(rule(txt_to.Text))
b = Trim(rule(txt_trang.Text))

    If Len(txt_to) = 1 Then txt_to.Text = "0" & txt_to.Text
    If Len(txt_trang) = 1 Then txt_trang.Text = "0" & txt_trang.Text
    
ketqua = txt_to.Text & " (" & a & ") t" & ChrW(7901) & ", " & txt_trang.Text & " (" & b & ") trang"

    Call SearchDocForPattern("V.n.b.n.*l.p.thành.*\(.*\).*\(.*\).*\(.*\).*l.u")
    Call ReplaceSelectionTextWithRegex(ketqua, "[\d]{2}.\(.{1,15}\).t.,.*[\d]{2}.\(.{1,15}\).trang")

If CheckBox1.Value = True Then Call thechap(TextBox_startwith.Value)


kethuc:

Unload Me



End Sub


Private Sub abc()
Dim i As Integer

    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "V.n.*b.n.*\(.*\).*\(.*\).*\(.*\)"
     

    Dim temp_para As Paragraph
    Dim AdjustedParagraph As String
    
    For Each temp_para In ActiveDocument.Paragraphs
    
        If .Test(temp_para.Range.Text) Then
        
            .pattern = "[\d]{2}.\(.{1,15}\).t.,.*[\d]{2}.\(.{1,15}\).trang"
            AdjustedParagraph = .Replace(temp_para.Range.Text, ketqua)
            
            .pattern = "\s$"
            AdjustedParagraph = .Replace(AdjustedParagraph, vbNullString)

            
            temp_para.Range.Select
            Selection.TypeText AdjustedParagraph
            
            
            Unload Me
            Exit Sub
            
        End If
    
    Next
    
End With

MsgBox "Insert manual"
Clipboard ketqua






kethuc:

Unload Me

End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub
Private Sub SpinButton_startwith_SpinDown()
On Error Resume Next
    TextBox_startwith.Value = TextBox_startwith.Value - 1
End Sub
Private Sub SpinButton_startwith_SpinUp()
On Error Resume Next
    TextBox_startwith.Value = TextBox_startwith.Value + 1
End Sub

Private Sub txt_to_Change()
If toEnter = 1 Then Exit Sub
 txt_trang.Text = txt_to.Text

End Sub

Private Sub txt_trang_Change()
 TextBox_startwith.Text = txt_trang.Text
End Sub

Private Sub txt_trang_Enter()
toEnter = 1
End Sub

Private Sub txt_to_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(txt_to) = 1 Then txt_to.Text = "0" & txt_to.Text
End Sub
Private Sub txt_trang_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(txt_trang) = 1 Then txt_trang.Text = "0" & txt_trang.Text
End Sub


Private Sub txt_to_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii >= 48 And KeyAscii <= 57 Then

Else
    
    KeyAscii = 0
End If
End Sub
Private Sub CheckBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
         CommandButton2_Click
    End If
End Sub
Private Sub txt_to_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 If KeyCode = 13 Then
         CommandButton2_Click
    End If
End Sub
Private Sub txt_trang_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 If KeyCode = 13 Then
         CommandButton2_Click
    End If
End Sub

Private Sub txt_trang_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii >= 48 And KeyAscii <= 57 Then

Else
    
    KeyAscii = 0
End If
End Sub


Private Sub UserForm_Initialize()
Me.Caption = "Numberring"
toEnter = 0

End Sub
Private Sub thechap(Optional ByVal DefaultNumbering As Integer = 1)

       
    With ActiveDocument.Sections(1) _
     .Footers(wdHeaderFooterPrimary).PageNumbers
     .NumberStyle = wdPageNumberStyleArabic
     .IncludeChapterNumber = False
     .RestartNumberingAtSection = True
     .StartingNumber = DefaultNumbering
     .Add PageNumberAlignment:=wdAlignPageNumberRight, _
     FirstPage:=True
    End With

End Sub
