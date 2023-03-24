VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccessCCV 
   Caption         =   "Notary Public"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3810
   OleObjectBlob   =   "AccessCCV.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "AccessCCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Cancle_Button_Click()
Unload Me

End Sub

Private Sub CommandButton1_Click()
Acces_CCV (Space(45))

End Sub

Private Sub dbt_Click()
Acces_CCV ("D" & ChrW(431) & ChrW(416) & "NG BÍCH TUY" & _
        ChrW(7872) & "N")

End Sub

Private Sub DPHK_Click()
Acces_CCV ("D" & ChrW(431) & ChrW(416) & "NG PH" & ChrW(431) _
         & ChrW(7898) & "C HOÀNG KHÁNH")

End Sub

Private Sub vhnp_Click()
Acces_CCV ("V" & ChrW(431) & ChrW(416) & "NG HOÀNG NH" & _
        ChrW(7844) & "T PH" & ChrW(431) & ChrW(416) & "NG")
End Sub
Private Sub Acces_CCV(tenCCV As String)


Application.ScreenUpdating = False

    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("CCV")



    DientenCCV_Middle (tenCCV)
    DientenCCV_Bottom (tenCCV)
         
    objUndo.EndCustomRecord


Application.ScreenUpdating = True


End Sub
Private Sub DientenCCV_Middle(tenCCV As String)


    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "^T.i.*ch.ng.*vi.n.*trong.*ph.m.*vi.*tr.ch.*nhi.m.*c.a.*m.nh.*theo.*quy.*c.a.*ph.p"
    Debug.Print .Test(Selection.Text)
    
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
    
        If .Test(para.Range.Text) Then
        
            para.Range.Select
        
           Selection.TypeText Text:="Tôi "
           Selection.Font.Bold = True
           Selection.TypeText tenCCV
           Selection.Font.Bold = False
           Selection.TypeText ", công ch" & ChrW(7913) & _
        "ng viên, trong ph" & ChrW(7841) & "m vi trách nhi" & ChrW(7879) & "m c" & ChrW(7911) & _
        "a mình theo quy " & ChrW(273) & ChrW(7883) & "nh c" & ChrW(7911) & _
        "a pháp lu" & ChrW(7853) & "t,"
        
            Exit Sub
        
        End If
    
    Next
    
    End With

    MsgBox "Not find pattern 1"

End Sub
Private Sub DientenCCV_Bottom(tenCCV As String)

    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "\s$"
     Dim a As String
     a = .Replace(Selection.Text, vbNullString)
     
     .pattern = "^(C.NG).*CH.NG.*VI.N"

    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
    
        If .Test(para.Range.Text) Then
        
            para.Range.Select
            
            On Error Resume Next
                Selection.Tables(1).Delete
            On Error GoTo 0

            Set tblNew = ActiveDocument.Tables.Add(Selection.Range, 1, 2)
            
             With tblNew
             .Cell(1, 1).Range.Select
                 Selection.Font.Bold = False
                 Selection.Font.Italic = True
                 
                 Selection.Font.Size = Me.tenthukyFontSize.Caption
                 Selection.Font.Name = Me.tenthukyFontName.Caption
                 Selection.TypeText Me.tenthuky.Caption
                 
                 With Selection.ParagraphFormat
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpaceSingle
                    .Alignment = wdAlignParagraphLeft
                    .FirstLineIndent = CentimetersToPoints(1.25)
                 End With
            
             .Cell(1, 2).Range.Select
                 With Selection.ParagraphFormat
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(7)
                    .Alignment = wdAlignParagraphCenter
                End With
                 Selection.Font.Bold = True
                 Selection.Font.Italic = False
                 Selection.Font.Name = "Times New Roman"
                 Selection.Font.Size = "13"
                 Selection.TypeText "CÔNG CH" & ChrW(7912) & "NG VIÊN"
                 Selection.TypeParagraph
                 Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                 Selection.TypeText Text:=tenCCV
             End With

            
            Exit Sub
        
        End If
        
    Next
    
    End With


    MsgBox "Not find pattern 2"


End Sub


