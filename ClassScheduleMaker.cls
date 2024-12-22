VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassScheduleMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Schedule")
Option Explicit
Private HeaderBody_()                As Variant
Private OfficialHolidayArray_()   As Variant
Private Const C_DATE As Long = 1
Private Const C_NAME As Long = 2
Private Sub Class_Initialize()
    Dim Directory As IDirectory
    Set Directory = New Directory
    HeaderBody_ = LIB.ReadAllLineAsArray2D(Directory.ScheduleHeaderFile)
    OfficialHolidayArray_ = LIB.ReadAllLineAsArray2D(Directory.HolidayFile)
End Sub
Public Function AddDataBody(ByVal WesternYear As Long) As Variant()
    Dim OfficialHolidayList As Dictionary
    Set OfficialHolidayList = AddOfficialHolidayList
    Dim HolidayList As Dictionary
    Set HolidayList = OfficialHolidayList
    Dim Years As Long
    For Years = 0 To 1
        With New ClassHoliday
            Dim UnOfficialHolidayList As Dictionary
            Set UnOfficialHolidayList = .AddHolidays(WesternYear + Years)
        End With
        Set HolidayList = LIB.MergeDictionary(HolidayList, UnOfficialHolidayList)
    Next
    Dim Dates As Dictionary
    Set Dates = AddDates(WesternYear)
    Dim RET() As Variant
    ReDim RET(1 To Dates.Count + 1, 1 To UBound(HeaderBody_, 2))
    Dim R As Long
    Dim C As Long
    R = 1
    For C = 1 To UBound(HeaderBody_, 2)
        RET(R, C) = HeaderBody_(R, C)
    Next
    R = 2
    '@Ignore UseMeaningfulName
    Dim C“ú•t As Long
    '@Ignore UseMeaningfulName
    Dim Cj“ú As Long
    '@Ignore UseMeaningfulName
    Dim C“ú’¼“ú As Long
    '@Ignore UseMeaningfulName
    Dim Cö‹Æ“ú As Long
    C“ú•t = Columns“ú•t
    Cj“ú = Columnsj“ú
    C“ú’¼“ú = Columns“ú’¼“ú
    Cö‹Æ“ú = Columnsö‹Æ“ú
    Dim Target As Variant
    For Each Target In Dates
        RET(R, C“ú•t) = Target
        If HolidayList.Exists(CDate(Target)) Then
            '@Ignore IIfSideEffect
            RET(R, Cj“ú) = IIf(HolidayList.Item(CDate(Target)) Like "‹x“ú", "U‘Ö‹x“ú", HolidayList.Item(CDate(Target)))
            RET(R, C“ú’¼“ú) = False
            RET(R, Cö‹Æ“ú) = False
        ElseIf Weekday(Target) = vbSunday Or Weekday(Target) = vbSaturday Then
            RET(R, C“ú’¼“ú) = False
            RET(R, Cö‹Æ“ú) = False
        Else
            RET(R, C“ú’¼“ú) = True
            RET(R, Cö‹Æ“ú) = True
        End If
        R = R + 1
    Next
    AddDataBody = RET
End Function
Private Function AddDates(ByVal WesternYear As Long) As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim FirstDate As Date
    Dim LastDate As Date
    FirstDate = DateSerial(WesternYear, 4, 1)
    LastDate = DateSerial(WesternYear + 1, 3, 31)
    Dim Target As Date
    For Target = FirstDate To LastDate
        RET.Add Target, vbNullString
    Next
    Set AddDates = RET
End Function
Private Function AddOfficialHolidayList() As Dictionary
    If LIB.IsNoneArray(OfficialHolidayArray_) Then
        MsgBox "j“úƒf[ƒ^‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñB", vbExclamation, "ƒtƒ@ƒCƒ‹ƒGƒ‰["
        Exit Function
    End If
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim i As Long
    For i = LBound(OfficialHolidayArray_) + 1 To UBound(OfficialHolidayArray_)
        If OfficialHolidayArray_(i, C_NAME) Like "‹x“ú" Then OfficialHolidayArray_(i, C_NAME) = "U‘Ö‹x“ú"
        RET.Add CDate(OfficialHolidayArray_(i, C_DATE)), OfficialHolidayArray_(i, C_NAME)
    Next
    Set AddOfficialHolidayList = RET
End Function
Private Function Columns“ú•t() As Long
    Columns“ú•t = ColumnIndex("“ú•t")
End Function
Private Function Columnsj“ú() As Long
    Columnsj“ú = ColumnIndex("j“ú")
End Function
Private Function Columns“ú’¼“ú() As Long
    Columns“ú’¼“ú = ColumnIndex("“ú’¼“ú")
End Function
Private Function Columnsö‹Æ“ú() As Long
    Columnsö‹Æ“ú = ColumnIndex("ö‹Æ“ú")
End Function
Private Function ColumnIndex(ByVal ItemName As String) As Long
    If LIB.IsNoneArray(HeaderBody_) Then Exit Function
    Dim RET As Long
    For RET = LBound(HeaderBody_, 2) To UBound(HeaderBody_, 2)
        If RET = UBound(HeaderBody_, 2) Then
            If Not HeaderBody_(1, RET) Like ItemName Then Exit Function
        Else
            If HeaderBody_(1, RET) Like ItemName Then Exit For
        End If
    Next
    ColumnIndex = RET
End Function
Public Function IsValidHeader() As Boolean
    If Columns“ú•t = 0 Or Columnsj“ú = 0 Or Columns“ú’¼“ú = 0 Or Columnsö‹Æ“ú = 0 Then Exit Function
    IsValidHeader = True
End Function

