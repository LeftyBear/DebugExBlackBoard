VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_GenerateConfigUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    LimitValue  As Dom_EntityAggregate
    Subject     As Dom_EntityAggregate
End Type

Private This As Member

Public Sub Initialize(ByVal LimitValue As Dom_EntityAggregate, ByVal Subject As Dom_EntityAggregate)
    Set This.LimitValue = LimitValue
    Set This.Subject = Subject
End Sub

Public Function Execute() As Dom_SchoolConfig
    Dim LimitValue As Dom_LimitValue
    Set LimitValue = CreateLimitValue
    Dim Subject As Dom_Subject
    Subject = CreateSubject
    Dim Result As Dom_SchoolConfig
    Set Result = New Dom_SchoolConfig
    Result.Initialize LimitValue, Subject
    Set Execute = Result
End Function

Private Function CreateLimitValue() As Dom_LimitValue
    Dim R As Long
    For R = 1 To This.LimitValue.Count
        Dim Context As Dom_ColumnContext
        Set Context = This.LimitValue.GetColumnContext(R)
        Dim C As Long
        For C = 1 To Context.Count
            Dim LongValueColumn As Dom_LongValueColumn
            Set LongValueColumn = Context.GetColumn(C)
            Dim MainStreamGradeCount As Long, MainStreamClassCount As Long, SpecialGradeCount As Long, SpecialClassCount As Long, PeriodCount As Long
            ParseLimitValueColumn LongValueColumn.HeaderName, LongValueColumn.Value, MainStreamGradeCount, MainStreamClassCount, SpecialGradeCount, SpecialClassCount, PeriodCount
        Next
    Next
    Dim Result As Dom_LimitValue
    Set Result = New Dom_LimitValue
    Result.Initialize MainStreamGradeCount, MainStreamClassCount, SpecialGradeCount, SpecialClassCount, PeriodCount
    Set CreateLimitValue = Result
End Function

Private Function CreateSubject() As Dom_Subject
    Dim Names As VBA.Collection
    Set Names = New VBA.Collection
    Dim TargetGrades As VBA.Collection
    Set TargetGrades = New VBA.Collection
    Dim R As Long
    For R = 1 To This.Subject.Count
        Dim Context As Dom_ColumnContext
        Set Context = This.Subject.GetColumnContext(R)
        Dim C As Long
        For C = 1 To Context.Count
            Dim StringValueColumn As Dom_StringValueColumn
            Set StringValueColumn = Context.GetColumn(C)
            Select Case StringValueColumn.HeaderName
            Case "項目名"
                Names.Add StringValueColumn.Value
            Case "対象学年"
                TargetGrades.Add StringValueColumn.Value
            End Select
        Next
    Next
    Dim Result As Dom_Subject
    Set Result = New Dom_Subject
    Result.Initialize Names, TargetGrades
    Set CreateSubject = Result
End Function

Private Sub ParseLimitValueColumn(ByVal HeaderName As String, ByVal Value As Long, ByRef MainStreamGradeCount As Long, ByRef MainStreamClassCount As Long, ByRef SpecialGradeCount As Long, ByRef SpecialClassCount As Long, ByRef PeriodCount As Long)
    Select Case HeaderName
    Case "普通学年数"
        MainStreamGradeCount = Value
    Case "普通学級数"
        MainStreamClassCount = Value
    Case "特支学年数"
        SpecialGradeCount = Value
    Case "特支学級数"
        SpecialClassCount = Value
    Case "時間割数"
        PeriodCount = Value
    End Select
End Sub
