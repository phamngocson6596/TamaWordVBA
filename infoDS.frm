VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} infoDS 
   Caption         =   "Information"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   OleObjectBlob   =   "infoDS.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "infoDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Function MessageBoxW Lib "User32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

'Const DataLocation = "G:\My Drive\Rice Shirt Rice Money\New Generation\KH.accdb"
Const DataLocation = "\\192.168.1.30\ho so chung\z.kh\KH.accdb"

Const adOpenStatic = 1

Private Sub ExportButton_Click()

Dim exportString As String

exportString = Var_Gt & vbTab & ": " & Var_Ten & vbCr _
            & Title_Sn & vbTab & ": " & Var_Sn & vbCr



Dim cmtThings
cmtThings = Array("CCCD", "CMND", "HC", "CMSQ", "SDDCN")
Dim iCMT As String

For Each item In cmtThings
    If Not Var_Frame.Controls("Var_" & item) = "" Then
        exportString = exportString _
        & Title_Frame.Controls("Title_" & item) & vbTab & ": " _
        & InputSpaceID(Var_Frame.Controls("Var_" & item)) & vbCr
    End If
Next

exportString = exportString & Title_TT & vbTab & ": " & Var_TT


    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(5.25), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

Selection.TypeText exportString

End Sub

Private Sub InsertButton_Click()

If Trim(Var_Ten) = "" Or (Trim(Var_CCCD) = "" And Trim(Var_CMND) = "" And Trim(Var_HC) = "" And Trim(Var_CMSQ) = "" _
And Trim(Var_SDDCN) = "") Then
Label4.Caption = "Fail!!"
Exit Sub
End If

Dim conectionLine As String
Set mainCn = CreateObject("ADODB.Connection")
Set mainRs = CreateObject("ADODB.Recordset")

mainCn.Provider = "Microsoft.ACE.OLEDB.12.0"
mainCn.Open DataLocation

Dim cmtThings
cmtThings = Array("CCCD", "CMND", "HC", "CMSQ", "SDDCN")


conectionLine = "SELECT * FROM [KH (3)] WHERE" _
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

conectionLine = "INSERT INTO [KH (3)] (Gt, Ten, Sn, CCCD, CMND, HC, CMSQ, SDDCN, TT) VALUES ('" _
& Trim(Var_Gt) & "','" & Trim(Var_Ten) & "','" & Trim(Var_Sn) & "','" _
& RemoveSpaceBeta(Var_CCCD) & "','" & RemoveSpaceBeta(Var_CMND) & "','" & RemoveSpaceBeta(Var_HC) & "','" _
& RemoveSpaceBeta(Var_CMSQ) & "','" & RemoveSpaceBeta(Var_SDDCN) & "','" & Trim(Var_TT) & "')"

mainRs.Open conectionLine, mainCn

conectionLine = "SELECT * FROM [KH (3)] WHERE" _
        & "[Ten] = '" & Trim(Var_Ten) & "'" _
        & "AND" _
        & "([CMND] = '" & RemoveSpaceBeta(Var_CMND) & "'" _
        & "OR [CCCD] = '" & RemoveSpaceBeta(Var_CCCD) & "'" _
        & "OR [HC] = '" & RemoveSpaceBeta(Var_HC) & "'" _
        & "OR [CMSQ] = '" & RemoveSpaceBeta(Var_CMSQ) & "'" _
        & "OR [SDDCN] = '" & RemoveSpaceBeta(Var_SDDCN) & "')"
        
        
mainRs.Open conectionLine, mainCn, adOpenStatic
    If mainRs.RecordCount = 1 Then
    Label4.Caption = "Successfully"
    Else
    Label4.Caption = mainRs.RecordCount
    End If
mainRs.Close


End Sub

Private Sub PYCButton_Click()

NewPYC!TenTextbox.Value = Var_Ten
NewPYC!DCTextbox.Value = Var_TT
Unload Me
NewPYC.Show

End Sub

Private Sub UpdateButton_Click()

Dim conectString As String
conectString = "UPDATE [KH (3)] SET " _
& "[GT] = '" & Trim(Var_Gt) & "'," _
& "[Ten] = '" & Trim(Var_Ten) & "'," _
& "[Sn] = '" & RemoveSpaceBeta(Var_Sn) & "'," _
& "[CCCD] = '" & RemoveSpaceBeta(Var_CCCD) & "'," _
& "[CMND] = '" & RemoveSpaceBeta(Var_CMND) & "'," _
& "[HC] = '" & RemoveSpaceBeta(Var_HC) & "'," _
& "[CMSQ] = '" & RemoveSpaceBeta(Var_CMSQ) & "'," _
& "[SDDCN] = '" & RemoveSpaceBeta(Var_SDDCN) & "'," _
& "[TT] = '" & Trim(Var_TT) & "' " _
& "WHERE [ID] =" & Var_ID

Set mainCn = CreateObject("ADODB.Connection")
Set mainRs = CreateObject("ADODB.Recordset")

mainCn.Provider = "Microsoft.ACE.OLEDB.12.0"
mainCn.Open DataLocation

mainRs.Open conectString, mainCn

Me.Hide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'Set mainCn = Nothing
'Set mainRs = Nothing
Unload Me

End Sub

Private Function InputSpaceID(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = False
     .pattern = "\d{9}"
    If Not .Test(strIn) Then Exit Function
    .pattern = "\d{12}"
    
    If .Test(strIn) Then
        InputSpaceID = Mid(strIn, 1, 3) & " " & Mid(strIn, 4, 3) & " " & Mid(strIn, 7, 3) & " " & Mid(strIn, 10, 3)
    Else
        InputSpaceID = Mid(strIn, 1, 3) & " " & Mid(strIn, 4, 3) & " " & Mid(strIn, 7, 3)
    End If
    
    End With
End Function


Public Sub analyzeSlectionParagraph()
If Not IsLicenseValid Then Exit Sub

Dim TheOne As String: TheOne = OneLineDS
Dim TheOneW0S As String: TheOneW0S = RemoveSpaceBeta(TheOne)

On Error GoTo ketthuc

            Dim RX1, RXns, RXname, rxcmt, RXtt
    
'ongba: ho va ten
    
    Set RX1 = CreateObject("vbscript.regexp")
    With RX1
     .Global = False
     .IgnoreCase = True
     .pattern = "Ông/Bà"
    End With
        
    Select Case RX1.Test(TheOne)
    Case True
        Me.Var_Gt.Value = "Ông/Bà"
        
        RX1.pattern = "Ông/Bà\s*?:\s?(.*?);"
        Dim MatchesName
        Set MatchesName = RX1.Execute(TheOne)
        Me.Var_Ten = MatchesName(0).SubMatches(0)
    Case False
        RX1.pattern = "^Ông"
        Select Case RX1.Test(TheOne)
        Case True
            Me.Var_Gt.Value = "Ông"
            RX1.pattern = "Ông\s*?:\s?(.*?);"
            Set MatchesName = RX1.Execute(TheOne)
            Me.Var_Ten = MatchesName(0).SubMatches(0)
        Case False
            RX1.pattern = "^Bà"
            Select Case RX1.Test(TheOne)
                Case True
                    Me.Var_Gt.Value = "Bà"
                    RX1.pattern = "Bà\s*?:\s?(.*?);"
                    Set MatchesName = RX1.Execute(TheOne)
                    Me.Var_Ten = MatchesName(0).SubMatches(0)
                Case False
                    RX1.pattern = "là Bà"
                    Select Case RX1.Test(TheOne)
                        Case True
                            Me.Var_Gt.Value = "Bà"
                            RX1.pattern = "Bà\s*?:\s?(.*?);"
                            Set MatchesName = RX1.Execute(TheOne)
                            Me.Var_Ten = MatchesName(0).SubMatches(0)
                        Case False
                            RX1.pattern = "là Ông"
                            Select Case RX1.Test(TheOne)
                                Case True
                                    Me.Var_Gt.Value = "Ông"
                                    RX1.pattern = "Ông\s*?:\s?(.*?);"
                                    Set MatchesName = RX1.Execute(TheOne)
                                    Me.Var_Ten = MatchesName(0).SubMatches(0)
                                Case False
                                    RX1.pattern = "Do Bà"
                                    Select Case RX1.Test(TheOne)
                                        Case True
                                            Me.Var_Gt.Value = "Bà"
                                            RX1.pattern = "Bà\s*?:\s?(.*?);"
                                            Set MatchesName = RX1.Execute(TheOne)
                                            Me.Var_Ten = MatchesName(0).SubMatches(0)
                                        Case False
                                            RX1.pattern = "Do Ông"
                                            Select Case RX1.Test(TheOne)
                                                Case True
                                                    Me.Var_Gt.Value = "Ông"
                                                    RX1.pattern = "Ông\s*?:\s?(.*?);"
                                                    Set MatchesName = RX1.Execute(TheOne)
                                                    Me.Var_Ten = MatchesName(0).SubMatches(0)
                                                Case False
                                                
                                            End Select
                                    End Select
                            End Select
                    End Select
                End Select
        End Select
    End Select
    Set RX1 = Nothing
    

'nam sinh


    Set RXns = CreateObject("vbscript.regexp")
    With RXns
     .Global = False
     .pattern = "[:\/](\d\d\d\d);"
    End With
    
    Dim MatchesNS
    If RXns.Test(TheOneW0S) Then
        Set MatchesNS = RXns.Execute(TheOneW0S)
        Me.Var_Sn.Value = MatchesNS(0).SubMatches(0)
    End If
    Set RXns = Nothing

'so cmnd

    Dim MatchesCMT
    Set rxcmt = CreateObject("vbscript.regexp")
    With rxcmt
     .Global = False
     .IgnoreCase = True
     .pattern = ":(\d{12})[;\)\(]"
    End With
    If rxcmt.Test(TheOneW0S) Then
        Set MatchesCMT = rxcmt.Execute(TheOneW0S)
        Me.Var_CCCD = Trim(MatchesCMT(0).SubMatches(0))
    End If

    rxcmt.pattern = ":(\d{9})[;\)\(]"
    If rxcmt.Test(TheOneW0S) Then
        Set MatchesCMT = rxcmt.Execute(TheOneW0S)
        Me.Var_CMND = Trim(MatchesCMT(0).SubMatches(0))
    End If
    
    rxcmt.pattern = "[;\(]H.chi.us?.?:(.*?)[;\)\(]"
    If rxcmt.Test(TheOneW0S) Then
        Set MatchesCMT = rxcmt.Execute(TheOneW0S)
        Me.Var_HC = Trim(MatchesCMT(0).SubMatches(0))
    End If
    
    rxcmt.pattern = "[;\(]Ch.ngminhs.quans?.?:(.*?)[;\)\(]"
    If rxcmt.Test(TheOneW0S) Then
        Set MatchesCMT = rxcmt.Execute(TheOneW0S)
        Me.Var_CMSQ = Trim(MatchesCMT(0).SubMatches(0))
    End If
    Set rxcmt = Nothing
    
    
'thuong tru

    Set RXtt = CreateObject("vbscript.regexp")
    With RXtt
     .Global = False
     .pattern = "th..ngtr.:"
     .IgnoreCase = True
    End With

    Select Case RXtt.Test(TheOneW0S)
    Case True
    RXtt.pattern = "th..ng tr.*?:\s*?(.*?);"
        Set MatchesCMT = RXtt.Execute(TheOne)
        Me.Var_TT = Trim(MatchesCMT(0).SubMatches(0))
    Case False
        RXtt.pattern = "c.tr.:"
        Select Case RXtt.Test(TheOneW0S)
        Case True
            RXtt.pattern = "c. tr..*?:\s*?(.*?);"
            Set MatchesCMT = RXtt.Execute(TheOne)
            Me.Var_TT = Trim(MatchesCMT(0).SubMatches(0))
        Case False
        End Select
    End Select

Exit Sub

ketthuc:
MsgBox "Sai dinh dang"
End Sub




Private Function OneLineDS()

Dim temp_para As Paragraph
Dim big_para As String

For Each temp_para In Selection.Paragraphs
    big_para = big_para & Trim(RemovespaceEnd(temp_para.Range.Text)) & ";"
Next

OneLineDS = big_para

End Function

Private Function RemovespaceEnd(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "\s$"
    RemovespaceEnd = .Replace(strIn, vbNullString)
    .pattern = "\t"
    RemovespaceEnd = .Replace(RemovespaceEnd, vbNullString)
    .pattern = ";\s"
    RemovespaceEnd = .Replace(RemovespaceEnd, ";")
    
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

