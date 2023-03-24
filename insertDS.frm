VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} insertDS 
   Caption         =   "Find"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   OleObjectBlob   =   "insertDS.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "insertDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Function MessageBoxW Lib "User32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long


Dim mainRs As Object
Dim mainCn As Object

Dim TotalRecordCount As Single
'Const DataLocation = "G:\My Drive\Rice Shirt Rice Money\New Generation\KH.accdb"
Const DataLocation = "\\192.168.1.30\ho so chung\z.kh\KH.accdb"

Const adStateOpen = 1
Const adOpenStatic = 1
Const adOpenDynamic = 2
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adCmdText = 1



Private Sub AddButton_Click()
If Not (mainRs Is Nothing) Then
    If (mainRs.State And adStateOpen) = adStateOpen Then
    Else
        Exit Sub
    End If
End If
If mainRs.RecordCount > 0 Then
    Label4.Caption = "Fail!"
    Exit Sub
End If

infoDS!Var_Gt.Value = "Ông/Bà"
infoDS!Var_Ten.Value = UCase(Find_Name.Value)



infoDS.Show

End Sub

Private Sub DeleteButton_Click()
If Not (mainRs Is Nothing) Then
    If (mainRs.State And adStateOpen) = adStateOpen Then
    Else
        Exit Sub
    End If
End If

If mainRs.RecordCount = 0 Then
Label4.Caption = "Fail!"
Exit Sub
End If


Select Case MsgBoxW("Xoá: " & mainRs("Ten") & " ra kh" & ChrW(7887) & "i c" & ChrW(417) & " s" & _
         ChrW(7903) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u?", vbOKCancel, "Deleting...")

Case Is = vbOK
    mainRs.Delete
    mainRs.Requery
    
    Call FillTheListbox
End Select

End Sub


Private Sub DSListbox_Click()
If Not (mainRs Is Nothing) Then
    If (mainRs.State And adStateOpen) = adStateOpen Then
    Else
        Exit Sub
    End If
End If

mainRs.MoveFirst
mainRs.Move DSListbox.ListIndex
Label4.Caption = mainRs.Fields("Ten")


End Sub

Private Sub DSListbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Not (mainRs Is Nothing) Then
    If (mainRs.State And adStateOpen) = adStateOpen Then
    Else
        Exit Sub
    End If
End If

If mainRs.EOF Then mainRs.MoveFirst

Dim iCaption As String, iField As Variant
For Each iField In mainRs.Fields
    iCaption = iCaption & iField.Value & "; "
Next

            Dim objRegex
            Set objRegex = CreateObject("vbscript.regexp")
            With objRegex
             .Global = True
             .pattern = "(;\s)+"
            End With
    iCaption = objRegex.Replace(iCaption, "; ")

ttDS!Label1.Caption = iCaption
ttDS.Show
End Sub
Private Sub DSListbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        KeyCode = 0
        If Not (mainRs Is Nothing) Then
            If (mainRs.State And adStateOpen) = adStateOpen Then
                Call ExportButton_Click
            Else
            End If
        End If
    Case vbKeyEscape
        KeyCode = 0
        Unload Me
    Case vbKeyDown
        KeyCode = 0
        If DSListbox.ListCount = 0 Then Exit Sub
        If Not DSListbox.ListIndex = DSListbox.ListCount - 1 Then DSListbox.ListIndex = DSListbox.ListIndex + 1
    Case vbKeyUp
        KeyCode = 0
        If DSListbox.ListCount = 0 Then Exit Sub
        If DSListbox.ListIndex = -1 Then DSListbox.ListIndex = 0
        If Not DSListbox.ListIndex = 0 Then DSListbox.ListIndex = DSListbox.ListIndex - 1
End Select

End Sub

Private Sub EditButton_Click()
If Not (mainRs Is Nothing) Then
    If (mainRs.State And adStateOpen) = adStateOpen Then
    Else
        Exit Sub
    End If
End If

Dim iID As String

With mainRs

Dim iField As Variant
For Each iField In .Fields
    If Not iField.Value = "" Then
        infoDS.Controls("Var_" & iField.Name) = iField
    End If
Next

End With

infoDS.InsertButton.Visible = False
infoDS.UpdateButton.top = infoDS.InsertButton.top

infoDS.Show

End Sub

Private Sub ExportButton_Click()
If Not (mainRs Is Nothing) Then
    If (mainRs.State And adStateOpen) = adStateOpen Then
    Else
        Exit Sub
    End If
End If

If mainRs.BOF Then
    Label4.Caption = "Not Available. Add new please!"
    Exit Sub
End If

Dim exportString As String

With mainRs
exportString = .Fields("Gt") & vbTab & ": " & .Fields("Ten") & vbCr _
            & Title_Sn & vbTab & ": " & .Fields("Sn") & vbCr



Dim cmtThings
cmtThings = Array("CCCD", "CMND", "HC", "CMSQ", "SDDCN")
Dim iCMT As String

For Each item In cmtThings
    If Not .Fields(item) = "" Then
        exportString = exportString _
        & DSPage("Details").Controls("Title_" & item).Value & vbTab & ": " _
        & InputSpaceID(.Fields(item)) & vbCr
    End If
Next

exportString = exportString & Title_TT & vbTab & ": " & .Fields("TT")

End With

    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(5.25), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

Selection.TypeText exportString


End Sub

Private Sub Find_ID_Change()
Call CheckButton_Click
End Sub

Private Sub Find_ID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Select Case KeyCode
    Case vbKeyReturn
        KeyCode = 0
        If Not (mainRs Is Nothing) Then
            If (mainRs.State And adStateOpen) = adStateOpen Then
                Call ExportButton_Click
            Else
            End If
        End If
    Case vbKeyEscape
        KeyCode = 0
        Unload Me
    Case vbKeyDown
        KeyCode = 0
        If DSListbox.ListCount = 0 Then Exit Sub
        If Not DSListbox.ListIndex = DSListbox.ListCount - 1 Then DSListbox.ListIndex = DSListbox.ListIndex + 1
    Case vbKeyUp
        KeyCode = 0
        If DSListbox.ListCount = 0 Then Exit Sub
        If DSListbox.ListIndex = -1 Then DSListbox.ListIndex = 0
        If Not DSListbox.ListIndex = 0 Then DSListbox.ListIndex = DSListbox.ListIndex - 1
End Select
End Sub


Private Sub Find_Name_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        KeyCode = 0
        If Not (mainRs Is Nothing) Then
            If (mainRs.State And adStateOpen) = adStateOpen Then
                Call ExportButton_Click
            Else
            End If
        End If
    Case vbKeyEscape
        KeyCode = 0
        Unload Me
    Case vbKeySpace
        Call CheckButton_Click
    Case vbKeyDown
        KeyCode = 0
        If DSListbox.ListCount = 0 Then Exit Sub
        If Not DSListbox.ListIndex = DSListbox.ListCount - 1 Then DSListbox.ListIndex = DSListbox.ListIndex + 1
    Case vbKeyUp
        KeyCode = 0
        If DSListbox.ListCount = 0 Then Exit Sub
        If DSListbox.ListIndex = -1 Then DSListbox.ListIndex = 0
        If Not DSListbox.ListIndex = 0 Then DSListbox.ListIndex = DSListbox.ListIndex - 1
End Select

End Sub


Private Sub InsertButton_Click()

Dim conectionLine As String

conectionLine = "SELECT * FROM [KH (3)$] WHERE" _
        & "[Ten] = '" & Trim(Var_Ten) & "'" _
        & "AND" _
        & "([CMND] = '" & RemoveSpaceBeta(Var_CMND) & "'" _
        & "OR [CCCD] = '" & RemoveSpaceBeta(Var_CCCD) & "'" _
        & "OR [HC] = '" & RemoveSpaceBeta(Var_HC) & "'" _
        & "OR [CMSQ] = '" & RemoveSpaceBeta(Var_CMSQ) & "'" _
        & "OR [SDDCN] = '" & RemoveSpaceBeta(Var_SDDCN) & "')"
mainRs.Open conectionLine, mainCn, adOpenStatic
If mainRs.RecordCount > 0 Then
Label4.Caption = "Already Available!"
Exit Sub
End If
mainRs.Close

    conectionLine = "INSERT INTO [KH (3)$] (Gt, Ten, Sn, CCCD, CMND, HC, CMSQ, SDDCN, TT) VALUES ('" _
& Trim(Var_Gt) & "','" & Trim(Var_Ten) & "','" & Trim(Var_Sn) & "','" _
& RemoveSpaceBeta(Var_CCCD) & "','" & RemoveSpaceBeta(Var_CMND) & "','" & RemoveSpaceBeta(Var_HC) & "','" _
& RemoveSpaceBeta(Var_CMSQ) & "','" & RemoveSpaceBeta(Var_SDDCN) & "','" & Trim(Var_TT) & "')"
mainRs.Open conectionLine, mainCn

conectionLine = "SELECT * FROM [KH (3)$] WHERE" _
        & "[Ten] = '" & Trim(Var_Ten) & "'" _
        & "AND" _
        & "([CMND] = '" & RemoveSpaceBeta(Var_CMND) & "'" _
        & "OR [CCCD] = '" & RemoveSpaceBeta(Var_CCCD) & "'" _
        & "OR [HC] = '" & RemoveSpaceBeta(Var_HC) & "'" _
        & "OR [CMSQ] = '" & RemoveSpaceBeta(Var_CMSQ) & "'" _
        & "OR [SDDCN] = '" & RemoveSpaceBeta(Var_SDDCN) & "')"
        
        
mainRs.Open conectionLine, mainCn, adOpenStatic
If mainRs.RecordCount = 1 Then Label4.Caption = "Successfully"
mainRs.Close
End Sub

Public Sub CheckButton_Click()

If Not (mainRs Is Nothing) Then
    If (mainRs.State And adStateOpen) = adStateOpen Then
        mainRs.Close
    End If
End If


Dim tempTenValue As String
Dim tempCMTValue As String
tempTenValue = Trim(Find_Name)
tempCMTValue = RemoveSpaceBeta(Find_ID)

If tempTenValue = "" And tempCMTValue = "" Then Exit Sub

Dim cntString As String
cntString = "SELECT * FROM [KH (3)] WHERE" _
        & "[Ten] LIKE '%" & tempTenValue & "%'" _
        & "AND" _
        & "([CMND] LIKE '%" & tempCMTValue & "%'" _
        & "OR [CCCD] LIKE '%" & tempCMTValue & "%'" _
        & "OR [HC] LIKE '%" & tempCMTValue & "%'" _
        & "OR [CMSQ] LIKE '%" & tempCMTValue & "%'" _
        & "OR [SDDCN] LIKE '%" & tempCMTValue & "%')"

mainRs.Open cntString, mainCn, adOpenStatic, adLockOptimistic

'___________________
Call FillTheListbox


End Sub

Private Sub PYCButton_Click()
If Not (mainRs Is Nothing) Then
    If (mainRs.State And adStateOpen) = adStateOpen Then
    Else
        Exit Sub
    End If
End If

NewPYC!TenTextbox.Value = mainRs.Fields("Ten")
NewPYC!DCTextbox.Value = mainRs.Fields("TT")
Unload Me
NewPYC.Show

End Sub

Private Sub ResetButton_Click()

Dim iControl As Control
Me.Controls("Find_Name").Value = ""
Me.Controls("Find_ID").Value = ""

      
Label4.Caption = ""
DSListbox.Clear

On Error Resume Next
mainRs.Close

End Sub

Private Sub UserForm_Initialize()

Set mainCn = CreateObject("ADODB.Connection")
Set mainRs = CreateObject("ADODB.Recordset")

mainCn.Provider = "Microsoft.ACE.OLEDB.12.0"
mainCn.Open DataLocation

'mainRs.Open "SELECT * FROM [KH (3)]", mainCn, adOpenStatic
'TotalRecordCount = mainRs.RecordCount
'mainRs.Close

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    'Giai phong bo nho
    On Error Resume Next

    mainRs.Close
    Set mainRs = Nothing
    mainCn.Close
    
    Unload Me
    Unload infoDS
    Unload ttDS
    
End Sub
Sub GoRequery()

End Sub

Private Sub FillTheListbox()
If Not (mainRs Is Nothing) Then
    If (mainRs.State And adStateOpen) = adStateOpen Then
    Else
        Exit Sub
    End If
End If

DSListbox.Clear

Dim cmtThings: cmtThings = Array("CCCD", "CMND", "HC", "CMSQ", "SDDCN")

    Dim listRow As Integer
    
    Do Until mainRs.EOF
        DSListbox.AddItem mainRs.Fields("Ten")
        
        DSListbox.List(listRow, 1) = mainRs.Fields("Sn")
        
        For Each item In cmtThings
            If Not mainRs.Fields(item) = "" Then
                DSListbox.List(listRow, 2) = mainRs.Fields(item)
                Exit For
            End If
        Next
        
        listRow = listRow + 1
        mainRs.MoveNext
    Loop
If mainRs.RecordCount > 0 Then
mainRs.MoveFirst
Label4.Caption = ""
End If

End Sub
Private Function InputSpaceID(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = False
     .pattern = "\d{9}"
    If Not .Test(strIn) Then
        InputSpaceID = strIn
        Exit Function
    End If
    .pattern = "\d{12}"
    
    If .Test(strIn) Then
        InputSpaceID = Mid(strIn, 1, 3) & " " & Mid(strIn, 4, 3) & " " & Mid(strIn, 7, 3) & " " & Mid(strIn, 10, 3)
    Else
        InputSpaceID = Mid(strIn, 1, 3) & " " & Mid(strIn, 4, 3) & " " & Mid(strIn, 7, 3)
    End If
    
    End With
End Function

Private Function NumberOnlyString(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "\D"
    NumberOnlyString = .Replace(strIn, vbNullString)
    
    End With
End Function

Private Function RemoveSpaceBeta(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "\s+"
    RemoveSpaceBeta = .Replace(strIn, vbNullString)
    
    End With
    
End Function

Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "Microsoft Access") As VbMsgBoxResult
    MsgBoxW = MessageBoxW(Application.ActiveWindow.hWnd, StrPtr(Prompt), StrPtr(Title), Buttons)
End Function

