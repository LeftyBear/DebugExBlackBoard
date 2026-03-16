VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTableHeaderInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Parser")
Option Explicit

Private Type Member
    ClassHourType   As String
    Period          As String
    Grade           As String
End Type

Private This As Member

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Interprete(ByVal Config As Dom_SchoolConfig, ByVal ColumnDTO As App_ColumnDTO, ByRef TimeTables() As String)
    ParseHeader ColumnDTO.HeaderName
    Dim ClassHourType As Dom_ClassHourType
    ClassHourType = ResolveClassHourType(This.ClassHourType)
    Dim Period As Dom_Period
    Set Period = ResolvePeriod(This.Period)
    If Config.PeriodCount < Period.Value Then Err.Raise DomErrInvalidRange, "Dom_TimeTableHeaderInterpreter", "校時は設定範囲内である必要があります。"
    Dim Grade As Dom_Grade
    Set Grade = ResolveGrade(This.Grade)
    If Config.MainStreamGradeCount < Grade.Value Then Err.Raise DomErrInvalidRange, "Dom_TimeTableHeaderInterpreter", "学年は設定範囲内である必要があります。"
    TimeTables(ClassHourType, Period.Value, Grade.Value) = ColumnDTO.RawValue
End Sub

Private Sub ParseHeader(ByVal HeaderName As String)
    Dim Parts() As String
    Parts = VBA.Split(HeaderName, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        If 0 < VBA.InStr(1, Parts(i), jpnKeikaku) Or 0 < VBA.InStr(1, Parts(i), jpnJisshi) Then
            This.ClassHourType = Parts(i)
        ElseIf 0 < VBA.InStr(1, Parts(i), jpnKouji) Then
            This.Period = Parts(i)
        ElseIf 0 < VBA.InStr(1, Parts(i), jpnNen) Then
            This.Grade = Parts(i)
        End If
    Next
End Sub

Private Function ResolveClassHourType(ByVal Part As String) As Dom_ClassHourType
    Select Case Part
    Case jpnKeikaku
        ResolveClassHourType = Dom_ClassHourType.Plan
    Case jpnJisshi
        ResolveClassHourType = Dom_ClassHourType.Execution
    End Select
End Function

Private Function ResolvePeriod(ByVal Part As String) As Dom_Period
    Dim Period As Dom_Period
    Set Period = New Dom_Period
    Dim Buffer As String
    Buffer = VBA.Replace(Part, jpnKouji, vbNullString)
    If VBA.IsNumeric(Buffer) Then
        Set ResolvePeriod = Period.Create(CLng(Buffer))
    Else
        Set ResolvePeriod = Period.Create(0)
    End If
End Function

Private Function ResolveGrade(ByVal Part As String) As Dom_Grade
    Dim Grade As Dom_Grade
    Set Grade = New Dom_Grade
    Dim Buffer As String
    Buffer = VBA.Replace(Part, jpnNen, vbNullString)
    If VBA.IsNumeric(Buffer) Then
        Set ResolveGrade = Grade.Create(CLng(Buffer))
    Else
        Set ResolveGrade = Grade.Create(0)
    End If
End Function
