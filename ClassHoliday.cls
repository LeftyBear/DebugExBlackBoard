VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormExBB")
Option Explicit
'@Ignore UseMeaningfulName
Private Const j“ú–@{s            As Date = #7/20/1948#
'@Ignore UseMeaningfulName
Private Const j“ú–@‰ü³_U‘Ö‹x“ú1  As Date = #4/12/1973#
'@Ignore UseMeaningfulName
Private Const j“ú–@‰ü³_U‘Ö‹x“ú2  As Date = #1/1/2007#
'@Ignore UseMeaningfulName
Private Const j“ú–@‰ü³_‘–¯‚Ì‹x“ú As Date = #12/27/1985#
Private Const LimitYear             As Long = 9999
'@Ignore UseMeaningfulName
Private Function ‘–¯‚Ìj“ú–¼() As Variant()
    ‘–¯‚Ìj“ú–¼ = Array("Œ³“ú", "¬l‚Ì“ú", "Œš‘‹L”O‚Ì“ú", "‚İ‚Ç‚è‚Ì“ú", "º˜a‚Ì“ú", _
        "Œ›–@‹L”O“ú", "‚±‚Ç‚à‚Ì“ú", "ŠC‚Ì“ú", "ƒXƒ|[ƒc‚Ì“ú", "R‚Ì“ú", _
        "Œh˜V‚Ì“ú", "H•ª‚Ì“ú", "‘Ìˆç‚Ì“ú", "•¶‰»‚Ì“ú", "‹Î˜JŠ´Ó‚Ì“ú", _
        "t•ª‚Ì“ú", "“Vc’a¶“ú", "‘¦ˆÊ—ç³“a‚Ì‹V", "“Vc‚Ì‘¦ˆÊ", "c‘¾q–¾me‰¤‚ÌŒ‹¥‚Ì‹V", _
        "c‘¾q“¿me‰¤‚ÌŒ‹¥‚Ì‹V", "º˜a“Vc‚Ì‘å‘r‚Ì—ç")
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function Œ³“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 1
    '@Ignore UseMeaningfulName
    Const “ú As Long = 1
    Select Case ”N
        Case Year(j“ú–@{s) To LimitYear
            RET = Œ“úw’è(”N, Œ, “ú)
            If RET < j“ú–@{s Then Exit Function
        Case Else
            Exit Function
    End Select
    Œ³“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function ¬l‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 1
    '@Ignore UseMeaningfulName
    Const “ú As Long = 15
    '@Ignore UseMeaningfulName
    Const T As Long = 2
    '@Ignore UseMeaningfulName
    Const —j“ú As Long = vbMonday
    Select Case ”N
        Case Year(j“ú–@{s) To 1999
            RET = Œ“úw’è(”N, Œ, “ú)
            If RET < j“ú–@{s Then Exit Function
        Case 2000 To LimitYear
            RET = ŒT—j“úw’è(”N, Œ, T, —j“ú)
        Case Else
            Exit Function
    End Select
    ¬l‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function Œš‘‹L”O‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 2
    '@Ignore UseMeaningfulName
    Const “ú As Long = 11
    Select Case ”N
        Case 1967 To LimitYear
            RET = Œ“úw’è(”N, Œ, “ú)
        Case Else
            Exit Function
    End Select
    Œš‘‹L”O‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function ‚İ‚Ç‚è‚Ì“ú(ByVal ”N As Long) As Date
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    Select Case ”N
        Case 1989 To 2006
            Œ = 4: “ú = 29
        Case 2007 To LimitYear
            Œ = 5: “ú = 4
        Case Else
            Exit Function
    End Select
    ‚İ‚Ç‚è‚Ì“ú = Œ“úw’è(”N, Œ, “ú)
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function º˜a‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 4
    '@Ignore UseMeaningfulName
    Const “ú As Long = 29
    Select Case ”N
        Case 2007 To LimitYear
            RET = Œ“úw’è(”N, Œ, “ú)
        Case Else
            Exit Function
    End Select
    º˜a‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function Œ›–@‹L”O“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 5
    '@Ignore UseMeaningfulName
    Const “ú As Long = 3
    Select Case ”N
        Case Year(j“ú–@{s) To LimitYear
            RET = Œ“úw’è(”N, Œ, “ú)
            If RET < j“ú–@{s Then Exit Function
        Case Else
            Exit Function
    End Select
    Œ›–@‹L”O“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function ‚±‚Ç‚à‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 5
    '@Ignore UseMeaningfulName
    Const “ú As Long = 5
    Select Case ”N
        Case Year(j“ú–@{s) To LimitYear
            RET = Œ“úw’è(”N, Œ, “ú)
            If RET < j“ú–@{s Then Exit Function
        Case Else
            Exit Function
    End Select
    ‚±‚Ç‚à‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function ŠC‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    '@Ignore UseMeaningfulName
    Dim T As Long
    '@Ignore UseMeaningfulName
    Dim —j“ú As Long
    Select Case ”N
        Case 1996 To 2002
            Œ = 7: “ú = 20
            RET = Œ“úw’è(”N, Œ, “ú)
        Case 2003 To 2019
            Œ = 7: T = 3: —j“ú = vbMonday
            RET = ŒT—j“úw’è(”N, Œ, T, —j“ú)
        Case 2020 To 2020
            Œ = 7: “ú = 23
            RET = Œ“úw’è(”N, Œ, “ú)
        Case 2021 To 2021
            Œ = 7: “ú = 22
            RET = Œ“úw’è(”N, Œ, “ú)
        Case 2022 To LimitYear
            Œ = 7: T = 3: —j“ú = vbMonday
            RET = ŒT—j“úw’è(”N, Œ, T, —j“ú)
        Case Else
            Exit Function
    End Select
    ŠC‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function ƒXƒ|[ƒc‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    '@Ignore UseMeaningfulName
    Dim T As Long
    '@Ignore UseMeaningfulName
    Dim —j“ú As Long
    Select Case ”N
        Case 2020 To 2020
            Œ = 7: “ú = 24
            RET = Œ“úw’è(”N, Œ, “ú)
        Case 2021 To 2021
            Œ = 7: “ú = 23
            RET = Œ“úw’è(”N, Œ, “ú)
        Case 2022 To LimitYear
            Œ = 10: T = 2: —j“ú = vbMonday
            RET = ŒT—j“úw’è(”N, Œ, T, —j“ú)
        Case Else
            Exit Function
    End Select
    ƒXƒ|[ƒc‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function R‚Ì“ú(ByVal ”N As Long) As Date
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    Select Case ”N
        Case 2016 To 2019
            Œ = 8: “ú = 11
        Case 2020 To 2020
            Œ = 8: “ú = 10
        Case 2021 To 2021
            Œ = 8: “ú = 8
        Case 2022 To LimitYear
            Œ = 8: “ú = 11
        Case Else
            Exit Function
    End Select
    R‚Ì“ú = Œ“úw’è(”N, Œ, “ú)
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function Œh˜V‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    '@Ignore UseMeaningfulName
    Dim T As Long
    '@Ignore UseMeaningfulName
    Dim —j“ú As Long
    Select Case ”N
        Case 1966 To 2002
            Œ = 9: “ú = 15
            RET = Œ“úw’è(”N, Œ, “ú)
        Case 2003 To LimitYear
            Œ = 9: T = 3: —j“ú = vbMonday
            RET = ŒT—j“úw’è(”N, Œ, T, —j“ú)
        Case Else
            Exit Function
    End Select
    Œh˜V‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function H•ª‚Ì“ú(ByVal ”N As Long) As Date
    '@Ignore UseMeaningfulName
    Dim “à•”•Ï” As Double
    Select Case ”N
        Case Year(j“ú–@{s) To 1979
            “à•”•Ï” = 23.2588
        Case 1980 To 2099
            “à•”•Ï” = 23.2488
        Case 2100 To 2150
            “à•”•Ï” = 24.2488
        Case Else
            Exit Function
    End Select
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    Œ = 9
    “ú = Int(“à•”•Ï” + 0.242194 * (”N - 1980) - Int((”N - 1980) / 4))
    H•ª‚Ì“ú = Œ“úw’è(”N, Œ, “ú)
    If H•ª‚Ì“ú < j“ú–@{s Then H•ª‚Ì“ú = 0
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function ‘Ìˆç‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    '@Ignore UseMeaningfulName
    Dim T As Long
    '@Ignore UseMeaningfulName
    Dim —j“ú As Long
    Select Case ”N
        Case 1966 To 1999
            Œ = 10: “ú = 10
            RET = Œ“úw’è(”N, Œ, “ú)
        Case 2000 To 2019
            Œ = 10: T = 2: —j“ú = vbMonday
            RET = ŒT—j“úw’è(”N, Œ, T, —j“ú)
        Case Else
            Exit Function
    End Select
    ‘Ìˆç‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function •¶‰»‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 11
    '@Ignore UseMeaningfulName
    Const “ú As Long = 3
    Select Case ”N
        Case Year(j“ú–@{s) To LimitYear
            RET = Œ“úw’è(”N, Œ, “ú)
            If RET < j“ú–@{s Then Exit Function
        Case Else
            Exit Function
    End Select
    •¶‰»‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function ‹Î˜JŠ´Ó‚Ì“ú(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 11
    '@Ignore UseMeaningfulName
    Const “ú As Long = 23
    Select Case ”N
        Case Year(j“ú–@{s) To LimitYear
            RET = Œ“úw’è(”N, Œ, “ú)
            If RET < j“ú–@{s Then Exit Function
        Case Else
            Exit Function
    End Select
    ‹Î˜JŠ´Ó‚Ì“ú = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function t•ª‚Ì“ú(ByVal ”N As Long) As Date
    '@Ignore UseMeaningfulName
    Dim “à•”•Ï” As Double
    Select Case ”N
        Case Year(j“ú–@{s) To 1979
            “à•”•Ï” = 20.8357
        Case 1980 To 2099
            “à•”•Ï” = 20.8431
        Case 2100 To 2150
            “à•”•Ï” = 21.851
        Case Else
            Exit Function
    End Select
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    Œ = 3
    “ú = Int(“à•”•Ï” + 0.242194 * (”N - 1980) - Int((”N - 1980) / 4))
    t•ª‚Ì“ú = Œ“úw’è(”N, Œ, “ú)
    If t•ª‚Ì“ú < j“ú–@{s Then t•ª‚Ì“ú = 0
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function “Vc’a¶“ú(ByVal ”N As Long) As Date
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    Select Case ”N
        Case Year(j“ú–@{s) To 1988
            Œ = 4: “ú = 29
        Case 1989 To 2018
            Œ = 12: “ú = 23
        Case 2020 To LimitYear
            Œ = 2: “ú = 23
        Case Else
            Exit Function
    End Select
    “Vc’a¶“ú = Œ“úw’è(”N, Œ, “ú)
    If “Vc’a¶“ú < j“ú–@{s Then “Vc’a¶“ú = 0
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function ‘¦ˆÊ—ç³“a‚Ì‹V(ByVal ”N As Long) As Date
    '@Ignore UseMeaningfulName
    Dim Œ As Long
    '@Ignore UseMeaningfulName
    Dim “ú As Long
    Select Case ”N
        Case 1990 To 1990
            Œ = 11: “ú = 12
        Case 2019 To 2019
            Œ = 10: “ú = 22
        Case Else
            Exit Function
    End Select
    ‘¦ˆÊ—ç³“a‚Ì‹V = Œ“úw’è(”N, Œ, “ú)
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function “Vc‚Ì‘¦ˆÊ(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 5
    '@Ignore UseMeaningfulName
    Const “ú As Long = 1
    Select Case ”N
        Case 2019 To 2019
            RET = Œ“úw’è(”N, Œ, “ú)
        Case Else
            Exit Function
    End Select
    “Vc‚Ì‘¦ˆÊ = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function c‘¾q–¾me‰¤‚ÌŒ‹¥‚Ì‹V(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 4
    '@Ignore UseMeaningfulName
    Const “ú As Long = 10
    Select Case ”N
        Case 1959 To 1959
            RET = Œ“úw’è(”N, Œ, “ú)
        Case Else
            Exit Function
    End Select
    c‘¾q–¾me‰¤‚ÌŒ‹¥‚Ì‹V = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function c‘¾q“¿me‰¤‚ÌŒ‹¥‚Ì‹V(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 6
    '@Ignore UseMeaningfulName
    Const “ú As Long = 9
    Select Case ”N
        Case 1993 To 1993
            RET = Œ“úw’è(”N, Œ, “ú)
        Case Else
            Exit Function
    End Select
    c‘¾q“¿me‰¤‚ÌŒ‹¥‚Ì‹V = RET
End Function
'@Ignore UseMeaningfulName, ProcedureNotUsed
Public Function º˜a“Vc‚Ì‘å‘r‚Ì—ç(ByVal ”N As Long) As Date
    Dim RET As Date
    '@Ignore UseMeaningfulName
    Const Œ As Long = 2
    '@Ignore UseMeaningfulName
    Const “ú As Long = 24
    Select Case ”N
        Case 1989 To 1989
            RET = Œ“úw’è(”N, Œ, “ú)
        Case Else
            Exit Function
    End Select
    º˜a“Vc‚Ì‘å‘r‚Ì—ç = RET
End Function
'@Ignore UseMeaningfulName
'@Ignore UseMeaningfulName
'@Ignore UseMeaningfulName
'@Ignore UseMeaningfulName
Private Function Œ“úw’è(ByVal ”N As Long, ByVal Œ As Long, ByVal “ú As Long) As Date
    If Not IsValidYears(”N) Then Exit Function
    If Not IsValidMonths(Œ) Then Exit Function
    If Not IsValidDays(“ú) Then Exit Function
    Œ“úw’è = DateSerial(”N, Œ, “ú)
End Function
'@Ignore UseMeaningfulName
'@Ignore UseMeaningfulName
'@Ignore UseMeaningfulName
'@Ignore UseMeaningfulName
'@Ignore UseMeaningfulName
Private Function ŒT—j“úw’è(ByVal ”N As Long, ByVal Œ As Long, ByVal T As Long, ByVal —j“ú As Long) As Date
    If Not IsValidWeeks(T) Then Exit Function
    If Not IsValidWeekday(—j“ú) Then Exit Function
    '@Ignore UseMeaningfulName
    Dim ‰“ú As Date
    '@Ignore UseMeaningfulName
    Dim ‰“ú‚Ì—j“ú As Long
    '@Ignore UseMeaningfulName
    Dim ‰“ú‚©‚çw’è—j“ú‚Ü‚Å‚Ì“ú” As Long
    ‰“ú = Œ“úw’è(”N, Œ, 1)
    ‰“ú‚Ì—j“ú = Weekday(‰“ú)
    ‰“ú‚©‚çw’è—j“ú‚Ü‚Å‚Ì“ú” = IIf(—j“ú >= ‰“ú‚Ì—j“ú, —j“ú - ‰“ú‚Ì—j“ú, —j“ú - ‰“ú‚Ì—j“ú + 7)
    '@Ignore UseMeaningfulName
    Dim ‰“ú‚©‚çw’èT‚Ì—j“ú‚Ü‚Å‚Ì“ú” As Long
    ‰“ú‚©‚çw’èT‚Ì—j“ú‚Ü‚Å‚Ì“ú” = ‰“ú‚©‚çw’è—j“ú‚Ü‚Å‚Ì“ú” + 7 * (T - 1)
    ŒT—j“úw’è = DateAdd("d", ‰“ú‚©‚çw’èT‚Ì—j“ú‚Ü‚Å‚Ì“ú”, ‰“ú)
End Function
'@Ignore UseMeaningfulName
Public Function AddHolidays(ByVal ”N As Long) As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim HolidayNames() As Variant
    HolidayNames = ‘–¯‚Ìj“ú–¼
    Dim i As Long
    For i = LBound(HolidayNames) To UBound(HolidayNames)
        Dim ReturnedDate As Date
        ReturnedDate = CallByName(Me, HolidayNames(i), VbMethod, ”N)
        If ReturnedDate > 0 Then RET.Add ReturnedDate, HolidayNames(i)
    Next
    Const SubstituteHolidayName As String = "U‘Ö‹x“ú"
    Dim SubstituteHoliday As Collection
    Set SubstituteHoliday = U‘Ö‹x“ú(RET)
    If Not SubstituteHoliday Is Nothing Then
        Dim HolidayDate As Variant
        For Each HolidayDate In SubstituteHoliday
            RET.Add HolidayDate, SubstituteHolidayName
        Next
    End If
    Const NationalHolidayName As String = "‘–¯‚Ì‹x“ú"
    Dim NationalHoliday As Collection
    Set NationalHoliday = ‘–¯‚Ì‹x“ú(RET)
    If Not NationalHoliday Is Nothing Then
        For Each HolidayDate In NationalHoliday
            RET.Add HolidayDate, NationalHolidayName
        Next
    End If
    Set AddHolidays = LIB.SortDictionary(RET, xlAscending)
End Function
'@Ignore UseMeaningfulName
Private Function U‘Ö‹x“ú(ByVal HolidayList As Dictionary) As Collection
    Dim RET As Collection
    Set RET = New Collection
    Dim HolidayDate As Variant
    For Each HolidayDate In HolidayList
        If Weekday(HolidayDate) = vbSunday Then
            If j“ú–@‰ü³_U‘Ö‹x“ú1 <= HolidayDate And HolidayDate < j“ú–@‰ü³_U‘Ö‹x“ú2 Then
                'j“ú‚ª“ú—j“ú‚Ìê‡‚Í‚»‚Ì—‚“ú‚ÌŒ—j“ú‚ğ‹x“ú‚Æ‚·‚é
                RET.Add DateAdd("d", 1, HolidayDate)
            ElseIf j“ú–@‰ü³_U‘Ö‹x“ú2 <= HolidayDate Then
                'u‘–¯‚Ìj“úv‚ª“ú—j“ú‚É“–‚½‚é‚Æ‚«A‚»‚Ì“ú‚ÌŒã‚ÌÅ‚à‹ß‚¢•½“ú‚ğ‹x“ú‚Æ‚·‚é
                '@Ignore UseMeaningfulName
                Dim ‹x“ú As Date
                ‹x“ú = DateAdd("d", 1, HolidayDate)
                Do
                    If Not HolidayList.Exists(‹x“ú) Then Exit Do
                    ‹x“ú = DateAdd("d", 1, ‹x“ú)
                Loop
                RET.Add ‹x“ú
            End If
        End If
    Next
    Set U‘Ö‹x“ú = RET
End Function
'@Ignore UseMeaningfulName
Private Function ‘–¯‚Ì‹x“ú(ByVal HolidayList As Dictionary) As Collection
    Dim RET As Collection
    Set RET = New Collection
    Dim HolidayDate As Variant
    For Each HolidayDate In HolidayList
        If j“ú–@‰ü³_‘–¯‚Ì‹x“ú <= HolidayDate Then
            '‘O“ú‚Æ—‚“ú‚Ì—¼•û‚ğu‘–¯‚Ìj“úv‚É‹²‚Ü‚ê‚½•½“ú‚Í‹x“ú‚Æ‚·‚é
            '@Ignore UseMeaningfulName
            Dim ‹x“ú As Date
            ‹x“ú = DateAdd("d", 1, HolidayDate)
            Do
                If Not HolidayList.Exists(‹x“ú) Then Exit Do
                ‹x“ú = DateAdd("d", 1, ‹x“ú)
            Loop
            If Weekday(‹x“ú) <> vbSunday Then
                '@Ignore UseMeaningfulName
                Dim ‹x“ú‚Ì—‚“ú As Date
                ‹x“ú‚Ì—‚“ú = DateAdd("d", 1, ‹x“ú)
                If HolidayList.Exists(‹x“ú‚Ì—‚“ú) Then
                    RET.Add ‹x“ú
                End If
            End If
        End If
    Next
    Set ‘–¯‚Ì‹x“ú = RET
End Function
'@Ignore UseMeaningfulName
Private Function IsValidYears(ByVal ”N As Long) As Boolean
    If ”N < Year(j“ú–@{s) Or LimitYear < ”N Then
        MsgBox "”N‚Ìw’è‚ÉŒë‚è‚ª‚ ‚è‚Ü‚·B", vbExclamation, "“ú•tw’èƒGƒ‰["
        Exit Function
    End If
    IsValidYears = True
End Function
'@Ignore UseMeaningfulName
Private Function IsValidMonths(ByVal Œ As Long) As Boolean
    If Œ < 1 Or 12 < Œ Then
        MsgBox "Œ‚Ìw’è‚ÉŒë‚è‚ª‚ ‚è‚Ü‚·B", vbExclamation, "“ú•tw’èƒGƒ‰["
        Exit Function
    End If
    IsValidMonths = True
End Function
'@Ignore UseMeaningfulName
Private Function IsValidDays(ByVal “ú As Long) As Boolean
    If “ú < 1 Or 31 < “ú Then
        MsgBox "“ú‚Ìw’è‚ÉŒë‚è‚ª‚ ‚è‚Ü‚·B", vbExclamation, "“ú•tw’èƒGƒ‰["
        Exit Function
    End If
    IsValidDays = True
End Function
'@Ignore UseMeaningfulName
Private Function IsValidWeeks(ByVal T As Long) As Boolean
    If T < 1 Or 5 < T Then
        MsgBox "T‚Ìw’è‚ÉŒë‚è‚ª‚ ‚è‚Ü‚·B", vbExclamation, "“ú•tw’èƒGƒ‰["
        Exit Function
    End If
    IsValidWeeks = True
End Function
'@Ignore UseMeaningfulName
Private Function IsValidWeekday(ByVal —j“ú As Long) As Boolean
    If —j“ú < 1 Or 7 < —j“ú Then
        MsgBox "—j“ú‚Ìw’è‚ÉŒë‚è‚ª‚ ‚è‚Ü‚·B", vbExclamation, "“ú•tw’èƒGƒ‰["
        Exit Function
    End If
    IsValidWeekday = True
End Function


