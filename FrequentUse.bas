Attribute VB_Name = "FrequentUse"
Const tenthuky = "(Son)"
Const tenthukyFontName = "Freestyle Script"
Const tenthukyFontSize = 12

Sub ngaythangchu_Auto()

If Not IsLicenseValid Then Exit Sub

        Dim ngay, thang, nam As String
        ngay = Trim(rule(Day(Date)))
        thang = Trim(rule(Month(Date)))
        nam = Trim(rule(Year(Date)))

Dim pattern1 As String, pattern2 As String, replacetext As String
pattern1 = "^H.m.*nay.*ng.y.*\(ng.y.*th.ng.*n.m.*\)"

    
pattern2 = "H*\(*\)"

replacetext = "Hôm nay, ngày " & Date & " (ngày " & ngay & ", tháng " & thang & ", n" & ChrW( _
                259) & "m " & nam & ")"
                
Call SearchDocForPattern(pattern1)

'Call ReplaceSelectionTextWithRegex(replacetext, pattern)

Call ReplaceSelectionTextWithFindAndReplace(replacetext, pattern2)

Selection.MoveRight Unit:=wdCharacter, Count:=1

End Sub


Sub chukyAuto()
If Not IsLicenseValid Then Exit Sub

Application.ScreenUpdating = False
    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("Tamayama")

Call SearchDocForPattern("^(C.NG).*CH.NG.*VI.N")

On Error Resume Next
    Selection.Tables(1).Delete
On Error GoTo 0

Set tblNew = ActiveDocument.Tables.Add(Selection.Range, 1, 2)
 With tblNew
     .Cell(1, 1).Range.Select
     With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(1.25)
     End With
     Selection.Font.Bold = False
     Selection.Font.Italic = True
     Selection.Font.Size = tenthukyFontSize
     Selection.Font.Name = tenthukyFontName
     Selection.TypeText tenthuky
     
     
     .Cell(1, 2).Range.Select
     With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .Alignment = wdAlignParagraphCenter
    End With
     Selection.Range.Font.Bold = True
     Selection.Range.Font.Italic = False
     Selection.Range.Font.Italic = False
    Selection.Font.Size = "13"
    Selection.Range.Font.Name = "Times New Roman"
     Selection.TypeText "CÔNG CH" & ChrW(7912) & "NG VIÊN"
    
 End With
 
     objUndo.EndCustomRecord
    'Denotes the End of The Undo Record
 
 
 Application.ScreenUpdating = True

End Sub


Sub soquyen_Auto()
If Not IsLicenseValid Then Exit Sub

Application.ScreenUpdating = False

    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("So Quyen")



Dim pattern As String, replacetext As String
pattern = "\.{5,}.*?TP.CC"
                
Call SearchDocForPattern(pattern)

Dim SelectionText As String
SelectionText = Selection.Range.Text

    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "\d{2}.\d{4}"
     If Not .Test(SelectionText) Then Exit Sub
    SelectionText = .Replace(SelectionText, "ok")
    
    .pattern = "\s$"
    SelectionText = .Replace(SelectionText, vbNullString)

    End With

Dim temp_array
temp_array = Split(SelectionText, "ok")

Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
Selection.Range.Delete
Selection.TypeText temp_array(0)
Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""MM/yyyy"" ", PreserveFormatting:=True
Selection.TypeText temp_array(1)

     
     
     objUndo.EndCustomRecord
    'Denotes the End of The Undo Record
 
 
 Application.ScreenUpdating = True


End Sub


Sub MoneyToClipboard()
If Not IsLicenseValid Then Exit Sub

Dim iv As String

xx = Selection.Range.Text
For zxc = 1 To Len(xx)
    If IsNumeric(Mid(xx, zxc, 1)) Then iii = iii + Mid(xx, zxc, 1)
Next zxc
iv = Trim(rule(iii)) & " " & ChrW(273) & ChrW(7891) & "ng"

iv = Replace(iv, Left(iv, 1), UCase(Left(iv, 1)), 1, 1)

'Selection.Paragraphs(1).Range.Select
'Call ReplaceSelectionTextWithRegex(Selection.Range.Text, "(" & iv & ")", "\(.*?\)")



Clipboard iv

End Sub
Sub accessAccessCCV()
If Not IsLicenseValid Then Exit Sub


AccessCCV!tenthuky.Caption = tenthuky
AccessCCV!tenthukyFontSize.Caption = tenthukyFontSize
AccessCCV!tenthukyFontName.Caption = tenthukyFontName


AccessCCV.top = 200
AccessCCV.Left = 1000

AccessCCV.Show


End Sub


Sub chuky()
If Not IsLicenseValid Then Exit Sub

On Error GoTo ketthuc

Application.ScreenUpdating = False

    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("Tamayama")


Set tblNew = ActiveDocument.Tables.Add(Selection.Range, 1, 2)
 With tblNew
     .Cell(1, 1).Range.Select
     With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(1.25)
     End With
     Selection.Font.Bold = False
     Selection.Font.Italic = True
     Selection.Font.Size = tenthukyFontSize
     Selection.Font.Name = tenthukyFontName
     Selection.TypeText tenthuky
     
     
     .Cell(1, 2).Range.Select
     With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .Alignment = wdAlignParagraphCenter
    End With
     Selection.Range.Font.Bold = True
     Selection.Range.Font.Italic = False
     Selection.Range.Font.Italic = False
    Selection.Font.Size = "13"
    Selection.Range.Font.Name = "Times New Roman"
     Selection.TypeText "CÔNG CH" & ChrW(7912) & "NG VIÊN"
    
 End With
 
ketthuc:
 
     objUndo.EndCustomRecord
    'Denotes the End of The Undo Record
 
 
 Application.ScreenUpdating = True

End Sub




Sub ngaythangchu()
If Not IsLicenseValid Then Exit Sub

Dim ngay, thang, nam As String
ngay = Trim(rule(Day(Date)))
thang = Trim(rule(Month(Date)))
nam = Trim(rule(Year(Date)))

Selection.TypeText Text:="Hôm nay, ngày " & Date & " (ngày " & ngay & ", tháng " & thang & ", n" & ChrW( _
        259) & "m " & nam & ")"
End Sub
Sub chuangayloichung1()
If Not IsLicenseValid Then Exit Sub

Dim ngay, thang, nam As String
ngay = Trim(rule(Day(Date)))
thang = Trim(rule(Month(Date)))
nam = Trim(rule(Year(Date)))

Selection.TypeText Text:="Hôm nay, ngày " & "....................." & " (ngày ....................." & ", tháng " & thang & ", n" & ChrW( _
        259) & "m " & nam & ")"
End Sub

Sub ongba()
If Not IsLicenseValid Then Exit Sub

On Error Resume Next
    Selection.Tables(1).Delete


    Application.Templates( _
        "C:\Users\Tama\AppData\Roaming\Microsoft\Templates\Normal.dotm"). _
        BuildingBlockEntries("Ông bà").Insert Where:=Selection.Range, RichText:= _
        True
End Sub

Sub chuthich_decimal2()
If Not IsLicenseValid Then Exit Sub


Dim SelectionText As String
SelectionText = Selection.Range.Text
SelectionText = RemoveSpaceAlpha(SelectionText)

Dim temp_array
temp_array = Split(SelectionText, ",")

If UBound(temp_array) > 2 Then
    MsgBox "Sai dinh dang"
    Exit Sub
End If

Dim phantruoc As String, phansau As String
phantruoc = temp_array(0)
phantruoc = rule(NumberOnlyString(phantruoc))

On Error GoTo songuyen
phansau = temp_array(1)
On Error GoTo 0

    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    objRegex.pattern = "0+$"
    phansau = objRegex.Replace(phansau, "")

phansau = rule(NumberOnlyString(phansau))

Selection.TypeText Remove_subLetter(SelectionText) & " (" & Trim(phantruoc) & " ph" & ChrW(7849) & "y " & Trim(phansau) & ")"

Exit Sub

songuyen:

Selection.TypeText Remove_subLetter(SelectionText) & " (" & Trim(phantruoc) & ")"

End Sub


Sub allwhite()
If Not IsLicenseValid Then Exit Sub


    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("Black&White")
 


    Selection.WholeStory
    
    Selection.Range.HighlightColorIndex = wdNoHighlight
    
    Selection.Range.Font.ColorIndex = wdBlack
    
    Selection.Range.Font.Hidden = False
    
    ActiveDocument.Sections.First.Footers(wdHeaderFooterPrimary).Range.Font.ColorIndex = wdBlack
    
    ActiveDocument.Sections.First.Footers(wdHeaderFooterPrimary).Range.Font.Name = "Times New Roman"

    ActiveDocument.Sections.Last.Footers(wdHeaderFooterPrimary).Range.Font.ColorIndex = wdBlack
    
    ActiveDocument.Sections.Last.Footers(wdHeaderFooterPrimary).Range.Font.Name = "Times New Roman"

    Application.Selection.EndOf
    

    objUndo.EndCustomRecord
    'Denotes the End of The Undo Record


End Sub

Sub ngaythangnamtudong()
If Not IsLicenseValid Then Exit Sub


    Application.ScreenUpdating = False
  Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("Day Updating")

    Selection.TypeText Text:="ngày "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""dd"" ", PreserveFormatting:=True
    Selection.TypeText Text:=" tháng "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""MM"" ", PreserveFormatting:=True
    Selection.TypeText Text:=" n" & ChrW(259) & "m "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""yyyy"" ", PreserveFormatting:=True
        
        
    objUndo.EndCustomRecord
    Application.ScreenUpdating = True
End Sub



Sub danhsotrang()
If Not IsLicenseValid Then Exit Sub


    Dim DefaultNumbering As Integer
    DefaultNumbering = 1
       
    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("PageNumbering")
   
       
    With ActiveDocument.Sections.First _
     .Footers(wdHeaderFooterPrimary).PageNumbers
     .NumberStyle = wdPageNumberStyleArabic
     .IncludeChapterNumber = False
     .RestartNumberingAtSection = True
     .StartingNumber = DefaultNumbering
     .Add PageNumberAlignment:=wdAlignPageNumberRight, _
     FirstPage:=True
    End With
    
    With ActiveDocument.Sections.Last _
     .Footers(wdHeaderFooterPrimary).PageNumbers
     .Add PageNumberAlignment:=wdAlignPageNumberRight, _
     FirstPage:=True
    End With

    With Selection.PageSetup
        .HeaderDistance = CentimetersToPoints(0.8)
        .FooterDistance = CentimetersToPoints(0.8)
    End With


    objUndo.EndCustomRecord

End Sub


Sub Canle15()
If Not IsLicenseValid Then Exit Sub


    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("page1,5")

    
    With ActiveDocument.PageSetup
    
        .TopMargin = CentimetersToPoints(1.67)
        .BottomMargin = CentimetersToPoints(1.67)
        .LeftMargin = CentimetersToPoints(2.5)
        .RightMargin = CentimetersToPoints(1.67)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1)
        .FooterDistance = CentimetersToPoints(0.8)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
    
    End With
    
    objUndo.EndCustomRecord
    'Denotes the End of The Undo Record

End Sub
Sub Canle3()
If Not IsLicenseValid Then Exit Sub

    
    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("page3")

    With ActiveDocument.PageSetup
    
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(2)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1)
        .FooterDistance = CentimetersToPoints(0.8)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
    
    End With
    
    objUndo.EndCustomRecord
    'Denotes the End of The Undo Record
 
    
End Sub
Sub Candongsingle()
'
If Not IsLicenseValid Then Exit Sub

  Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("SingeLine")

    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    
   
        objUndo.EndCustomRecord
    'Denotes the End of The Undo Record
    
End Sub

Sub candong66()
If Not IsLicenseValid Then Exit Sub

'
On Error Resume Next

    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("66")
 
    With Selection.ParagraphFormat
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
    End With
    Selection.Range.Style.NoSpaceBetweenParagraphsOfSameStyle = False
    objUndo.EndCustomRecord
    
    
End Sub
Sub candong61()
If Not IsLicenseValid Then Exit Sub

'
On Error Resume Next

    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("61")

    With Selection.ParagraphFormat
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
    End With
    Selection.Range.Style.NoSpaceBetweenParagraphsOfSameStyle = False
    objUndo.EndCustomRecord
    
End Sub

Sub multiple1_2()
If Not IsLicenseValid Then Exit Sub


Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.2)
End Sub



Sub HXN_LHA()

Application.ScreenUpdating = False

    Dim objUndo As UndoRecord
    'Declares the variable objUndo of the UndoRecord type
    Set objUndo = Application.UndoRecord
    'sets the object to an actual item on the undo stack
    objUndo.StartCustomRecord ("HXN_LHA")

On Error Resume Next


    Dim oSec As Section
    Dim oHead As HeaderFooter
    Dim oFoot As HeaderFooter

    For Each oSec In ActiveDocument.Sections
        For Each oHead In oSec.Headers
            If oHead.Exists Then oHead.Range.Delete
        Next oHead

        For Each oFoot In oSec.Footers
            If oFoot.Exists Then oFoot.Range.Delete
        Next oFoot
    Next oSec
    
On Error GoTo buoc3
    
    Selection.HomeKey Unit:=wdStory
    Application.Templates( _
        "C:\Users\Tama\AppData\Roaming\Microsoft\Templates\Normal.dotm"). _
        BuildingBlockEntries("quoc hieu tieu ngu 2").Insert Where:=Selection. _
        Range, RichText:=True
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Hoàng Xuân Ng" & ChrW(7909)
        .Replacement.Text = "Lê Hùng Anh"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Application.ScreenUpdating = True
    
    Call danhsotrang

    objUndo.EndCustomRecord
    
    Exit Sub
    
buoc3:
    
    MsgBox "Da xay ra loi"
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Hoàng Xuân Ng" & ChrW(7909)
        .Replacement.Text = "Lê Hùng Anh"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
       
    objUndo.EndCustomRecord
   Application.ScreenUpdating = True

End Sub
Sub soquyencongchung()
If Not IsLicenseValid Then Exit Sub


    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""MM/yyyy"" ", PreserveFormatting:=True
        
End Sub



Sub sotosotrang()
If Not IsLicenseValid Then Exit Sub

    totrang.Show
    
End Sub
Sub phanmemtinhngay()
If Not IsLicenseValid Then Exit Sub

    tinhchenhlechngay.Show

End Sub
Sub hienpyc()
If Not IsLicenseValid Then Exit Sub

    NewPYC.Show

End Sub

Sub analyzeSlectionParagraph()
If Not IsLicenseValid Then Exit Sub

Dim a As infoDS
Set a = New infoDS
a.analyzeSlectionParagraph
a.Show
End Sub
Sub showinfoDS()
If Not IsLicenseValid Then Exit Sub

insertDS.Show
End Sub

Sub showBankBoss()
If Not IsLicenseValid Then Exit Sub

BankBoss.Show

End Sub

'__________________Function_____________________


Function rule(spl As Variant) As String

Dim sc: sc = Array("không", "m" & ChrW(7897) & "t", "hai", "ba", "b" & ChrW(7889) & "n", "n" & ChrW(259) & "m", "sáu", "b" & ChrW(7843) & "y", "tám", "chín", "m" & ChrW(432) & ChrW(7901) & "i", "l" & ChrW(259) & "m")
Dim a As Integer, b As Integer, C As Integer, k As Integer
Dim s1 As String, s2 As String, s3 As String, mns As String

a = Len(spl)
If (a Mod 3) <> 0 Then
    C = (a \ 3 + 1) * 3
    spl = Format(spl, String(C, "0"))
Else
    spl = Format(spl, String(a, "0"))
End If

C = Len(spl) / 3
k = 0

For i = C To 1 Step -1

    b = i * 3 - 2
    k = k + 1
    
    mns = Mid(spl, b, 3)
    s1 = Mid(mns, 1, 1): s2 = Mid(mns, 2, 1): s3 = Mid(mns, 3, 1)
    
        Select Case k
            Case 1: If (i <> C) Then rule = "t" & ChrW(7927) + Space(1) + Trim(rule)
            Case 2: If mns <> 0 Then rule = "ngh" & ChrW(236) & "n" + Space(1) + Trim(rule)
            Case 3: If mns <> 0 Then rule = "tri" & ChrW(7879) & "u" + Space(1) + Trim(rule)
        End Select

    Select Case s3
        Case 0
        Case 5
            Select Case s2
                Case 0: rule = sc(5) + Space(1) + Trim(rule)
                Case Else: rule = sc(11) + Space(1) + Trim(rule)
            End Select
        Case 1
            Select Case s2
                Case 0, 1: rule = sc(1) + Space(1) + Trim(rule)
                Case Else: rule = "m" & ChrW(7889) & "t" & Space(1) & Trim(rule)
            End Select
        Case Else
            rule = sc(s3) + Space(1) + Trim(rule)
    End Select
    
    Select Case s2
        Case 0
            Select Case s3
                Case Is <> 0
                    Select Case i
                        Case 1: Select Case s1: Case 0: Case Else: rule = "l" & ChrW(7867) + Space(1) + Trim(rule): End Select
                        Case Else: rule = "l" & ChrW(7867) + Space(1) + Trim(rule)
                    End Select
            End Select
        Case 1: rule = sc(10) + Space(1) + Trim(rule)
        Case Else: rule = sc(s2) + " " + "m" & ChrW(432) & ChrW(417) & "i" + Space(1) + Trim(rule)
    End Select
    
    Select Case s1
        Case 0
            If (Mid(spl, 1, b)) <> 0 And (s2 <> 0 Or s3 <> 0) Then rule = sc(0) + Space(1) + "tr" & ChrW(259) & "m" + Space(1) + Trim(rule)
        Case Else
            rule = sc(s1) + " " + "tr" & ChrW(259) & "m" + " " + Trim(rule)
    End Select

If k = 3 Then k = 0
    
Next i
               
End Function

Function Clipboard(Optional StoreText As Variant) As String
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)

Dim X As Variant

'Store as variant for 64-bit VBA support
  X = StoreText

'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
      Select Case True
        Case Len(StoreText)
          'Write to the clipboard
            .setData "text", X
        Case Else
          'Read from the clipboard (no variable passed through)
            Clipboard = .GetData("text")
      End Select
    End With
  End With

End Function


Sub SearchDocForPattern(pattern As String, Optional find_stop As Boolean)


    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
          
     .pattern = pattern
     
Dim temp_para As Paragraph
Dim findcount As Integer

For Each temp_para In ActiveDocument.Paragraphs

If .Test(temp_para.Range.Text) Then
    
    temp_para.Range.Select
    If find_stop Then Exit Sub
    findcount = findcount + 1
    
End If

Next

End With

If findcount > 1 Then
    MsgBox "Find counted: " & findcount
ElseIf findcount < 1 Then
    MsgBox "Not find anything. Move to End"
    Application.Selection.EndOf
End If


End Sub
Sub ReplaceSelectionTextWithRegex(TextReplace As String, pattern As String, Optional RemovespaceEnd As Boolean = True)

    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = pattern

    Dim FinText As String
    
    FinText = .Replace(Selection.Range.Text, TextReplace)
    
    If RemovespaceEnd Then
     .pattern = "\s$"
    FinText = .Replace(FinText, vbNullString)
    End If
    
    Selection.TypeText FinText

    End With

End Sub

Sub ReplaceSelectionTextWithFindAndReplace(ReplaceWithWhat As String, patternFindnReplaceFormat As String)

    
    Dim myrange
    Set myrange = Selection.Range
    myrange.Find.Execute FindText:=patternFindnReplaceFormat, replacewith:=ReplaceWithWhat, Forward:=True, MatchWildcards:=True
    
    'https://www.avantixlearning.ca/microsoft-word/how-to-use-wildcards-in-word-to-find-and-replace/
    
End Sub

Function Remove_subLetter(strIn As String) As String

    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "\(.*\)"
    Remove_subLetter = .Replace(strIn, vbNullString)
    
    
    
    End With
    
    
End Function

Function RemoveSpace(strIn As String, Optional EndOnly As Boolean = False) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "\s"
     
     If EndOnly Then .pattern = "\s$"
     
     
    RemoveSpace = .Replace(strIn, vbNullString)
    
    End With
    

End Function

Function NumberOnlyString(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "\D"
    NumberOnlyString = .Replace(strIn, vbNullString)
    
    End With
    
'https://www.w3schools.com/jsref/jsref_obj_regexp.asp (Ky tu Regex)
'https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
'https://software-solutions-online.com/vba-regex-guide/#Example_3_Execute_Operation

End Function

Function RemoveSpaceAlpha(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .pattern = "\s+"
    RemoveSpaceAlpha = .Replace(strIn, vbNullString)
    
    End With
    
End Function

Function IsLicenseValid() As Boolean
    IsLicenseValid = False

    On Error GoTo ketthuc

    Dim LicenseKey As String
    Dim fileName As String
    Dim fileNumber As Integer
    Dim fileLine As String
    Dim fso As Object
    Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = "\\192.168.1.30\ho so chung\z.kh\file.txt"

    ' Check if the file exists
    If fso.FileExists(fileName) Then
        fileNumber = FreeFile
        Open fileName For Input As #fileNumber

        ' Read the file line by line
        Do Until EOF(fileNumber)
            Line Input #fileNumber, fileLine
            If InStr(fileLine, "Hello Tama") > 0 Then
                IsLicenseValid = True
                Close #fileNumber
                Exit Function
            End If
        Loop
        Close #fileNumber
    End If

    Dim iLSBox As LicenseForm
    Set iLSBox = New LicenseForm
    iLSBox.Show
    
    If Trim(iLSBox.GetLicense) = "Hello Tama" Then
        Set ts = fso.CreateTextFile(fileName)
        ts.WriteLine iLSBox.GetLicense
        ts.Close
        IsLicenseValid = True
    End If
    Set iLSBox = Nothing
    
ketthuc:
End Function


