VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tinhchenhlechngay 
   Caption         =   "UserForm1"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4950
   OleObjectBlob   =   "tinhchenhlechngay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tinhchenhlechngay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim ngayhomnay As Date
Dim ngaybatdau As Date
Dim ngayketthuc As Date

Sub tinhngay()

    ngaybatdau = DateSerial(yearfrom, monthfrom, dayfrom)
    ngayketthuc = DateSerial(yearto, monthto, dayto)
    
    Dim ngaychenhlech As Long
    
    ngaychenhlech = DateDiff("d", ngaybatdau, ngayketthuc)
    
    thongbao1.Caption = "Ð" & ChrW(227) & " qua: " & ngaychenhlech & " ngày"
    
    
    Dim chenhlechngay As String
    
    chenhlechngay = YearsMonthsDays(ngaybatdau, ngayketthuc, False, False)
    
    thongbao2.Caption = "Ho" & ChrW(7863) & "c " & chenhlechngay




End Sub

Private Sub CommandButton1_Click()

    Call tinhngay

End Sub
Private Sub autotinhtoan()

    If Trim(dayfrom.Value & vbNullString) = vbNullString Then
    ElseIf Trim(monthfrom.Value & vbNullString) = vbNullString Then
    ElseIf Trim(yearfrom.Value & vbNullString) = vbNullString Then
    ElseIf Trim(dayto.Value & vbNullString) = vbNullString Then
    ElseIf Trim(monthto.Value & vbNullString) = vbNullString Then
    ElseIf Trim(yearto.Value & vbNullString) = vbNullString Then
    
    Else
    
    Call tinhngay
    
    End If

End Sub

Private Sub dayfrom_Change()

    Call autotinhtoan

End Sub

Private Sub dayto_Change()

    Call autotinhtoan

End Sub

Private Sub monthfrom_Change()

    Call autotinhtoan

End Sub

Private Sub monthto_Change()

    Call autotinhtoan

End Sub
Private Sub yearfrom_Change()

    Call autotinhtoan

End Sub

Private Sub yearto_Change()

    Call autotinhtoan

End Sub


Private Sub UserForm_Initialize()

ngayhomnay = Date

    dayfrom.Text = Day(ngayhomnay)
    monthfrom.Text = Month(ngayhomnay)
    yearfrom.Text = Year(ngayhomnay)

        If Len(dayfrom.Text) = 1 Then dayfrom.Text = "0" & dayfrom.Text
        If Len(monthfrom.Text) = 1 Then monthfrom.Text = "0" & monthfrom.Text
        
        dayto.Text = Day(ngayhomnay)
    
    monthto.Text = Month(ngayhomnay)
    yearto.Text = Year(ngayhomnay)

        If Len(dayto.Text) = 1 Then dayto.Text = "0" & dayto.Text
        If Len(monthto.Text) = 1 Then monthto.Text = "0" & monthto.Text

Me.Caption = "Tính ngày"
        
End Sub

Function YearsMonthsDays(Date1 As Date, _
                     Date2 As Date, _
                     Optional ShowAll As Boolean = False, _
                     Optional Grammar As Boolean = True, _
                     Optional MinusText As String = "Minus " _
                     ) As String
Dim dTempDate As Date
Dim iYears As Integer
Dim iMonths As Integer
Dim iDays As Integer
Dim sYears As String
Dim sMonths As String
Dim sDays As String
Dim sGrammar(-1 To 0) As String
Dim sMinusText As String

If Grammar = True Then
    sGrammar(0) = "s"
End If


If Date1 > Date2 Then
    dTempDate = Date1
    Date1 = Date2
    Date2 = dTempDate
    sMinusText = MinusText
End If

iYears = DateDiff("yyyy", Date1, Date2)
Date1 = DateAdd("yyyy", iYears, Date1)
If Date1 > Date2 Then
    iYears = iYears - 1
    Date1 = DateAdd("yyyy", -1, Date1)
End If

iMonths = DateDiff("M", Date1, Date2)
Date1 = DateAdd("M", iMonths, Date1)
If Date1 > Date2 Then
    iMonths = iMonths - 1
    Date1 = DateAdd("m", -1, Date1)
End If

iDays = DateDiff("d", Date1, Date2)

If ShowAll Or iYears > 0 Then
    sYears = iYears & " nãm" & sGrammar((iYears = 1)) & ", "
End If
If ShowAll Or iYears > 0 Or iMonths > 0 Then
    sMonths = iMonths & " tháng" & sGrammar((iMonths = 1)) & ", "
End If
sDays = iDays & " ngày" & sGrammar((iDays = 1))

YearsMonthsDays = sMinusText & sYears & sMonths & sDays
End Function


