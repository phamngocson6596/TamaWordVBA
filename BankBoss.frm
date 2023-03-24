VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BankBoss 
   Caption         =   "Dang ky chu ky LHA"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4545
   OleObjectBlob   =   "BankBoss.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BankBoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If lstFiles.ListCount = 1 Then lstFiles.ListIndex = 0
    
    Dim folderPath As String
    folderPath = "\\192.168.1.30\ho so chung\x.DKCK\" ' specify the folder path here

    If lstFiles.ListIndex = -1 Then Exit Sub
    
    Dim fileName As String
    fileName = Me.lstFiles.Value

    Shell "rundll32.exe url.dll,FileProtocolHandler " & folderPath & fileName, vbNormalFocus

End Sub

Private Sub txtSearch_Change()

    Dim File As String
    Dim Path As String
    Path = "\\192.168.1.30\ho so chung\x.DKCK\" 'update this with the correct path
    File = Dir(Path & "*.pdf")
    lstFiles.Clear
    Do While File <> ""
        If InStr(UCase(File), UCase(txtSearch.text)) > 0 Then
            If lstFiles.ListCount = 0 Then
                lstFiles.AddItem File
            Else
                Dim Found As Boolean
                Found = False
                For i = 0 To lstFiles.ListCount - 1
                    If lstFiles.List(i) = File Then
                        Found = True
                    End If
                Next i
                If Found = False Then
                    lstFiles.AddItem File
                End If
            End If
        End If
        File = Dir
    Loop
End Sub


Private Sub UserForm_Initialize()
    Dim folderPath As String
    folderPath = "\\192.168.1.30\ho so chung\x.DKCK\" ' specify the folder path here

    Dim File As String
    File = Dir(folderPath & "*.pdf")

    Do While File <> ""
        Me.lstFiles.AddItem File
        File = Dir
    Loop
End Sub

