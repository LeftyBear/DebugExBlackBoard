VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    DateID          As Dom_DateID
    PlanValueList   As Scripting.Dictionary
    PlanMarkList    As Scripting.Dictionary
    ExeValueList    As Scripting.Dictionary
    ExeMarkList     As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal DateID As Dom_DateID, ByVal PlanValueList As Scripting.Dictionary, ByVal PlanMarkList As Scripting.Dictionary, ByVal ExeValueList As Scripting.Dictionary, ByVal ExeMarkList As Scripting.Dictionary)
    Set This.DateID = DateID
    Set This.PlanValueList = PlanValueList
    Set This.PlanMarkList = PlanMarkList
    Set This.ExeValueList = ExeValueList
    Set This.ExeMarkList = ExeMarkList
End Sub

Public Sub UpdatePlanValue(ByVal ID As Dom_ClassHourID, ByVal NewValue As Dom_ClassHourValue)
    Dim Target As Dom_ClassHour
    Set Target = FindPlanValue(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "ClassHourオブジェクトが必要です。"
    Dim Value As Dom_ClassHourValue
    With New Dom_ClassHourValue
        Set Value = .Create(NewValue.Value)
    End With
    Target.ChangeValueInternal Value
    ValidateTotal ID
End Sub

Public Sub UpdatePlanMark(ByVal ID As Dom_TimeTableID, ByVal NewValue As Dom_TimeTableMark)
    Dim Target As Dom_TimeTable
    Set Target = FindPlanMark(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "ClassHourオブジェクトが必要です。"
    Dim Value As Dom_TimeTableMark
    With New Dom_TimeTableMark
        Set Value = .Create(NewValue.Value)
    End With
    Target.ChangeValueInternal Value
    ValidateTotal ToClassHourID(ID)
End Sub

Public Sub UpdateExeValue(ByVal ID As Dom_ClassHourID, ByVal NewValue As Dom_ClassHourValue)
    Dim Target As Dom_ClassHour
    Set Target = FindExeValue(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "ClassHourオブジェクトが必要です。"
    Dim Value As Dom_ClassHourValue
    With New Dom_ClassHourValue
        Set Value = .Create(NewValue.Value)
    End With
    Target.ChangeValueInternal Value
    ValidateTotal ID
End Sub

Public Sub UpdateExeMark(ByVal ID As Dom_TimeTableID, ByVal NewValue As Dom_TimeTableMark)
    Dim Target As Dom_TimeTable
    Set Target = FindExeMark(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "ClassHourオブジェクトが必要です。"
    Dim Value As Dom_TimeTableMark
    With New Dom_TimeTableMark
        Set Value = .Create(NewValue.Value)
    End With
    Target.ChangeValueInternal Value
    ValidateTotal ToClassHourID(ID)
End Sub

Public Function GetDateKey() As String
    GetDateKey = This.DateID.ToKey
End Function

Public Function GetPlanValueKeys() As Variant()
    GetPlanValueKeys = This.PlanValueList.Keys
End Function

Public Function GetPlanMarkKeys() As Variant()
    GetPlanMarkKeys = This.PlanMarkList.Keys
End Function

Public Function GetExeValueKeys() As Variant()
    GetExeValueKeys = This.ExeValueList.Keys
End Function

Public Function GetExeMarkKeys() As Variant()
    GetExeMarkKeys = This.ExeMarkList.Keys
End Function

Public Function GetPlanValues() As Variant()
    GetPlanValues = This.PlanValueList.Items
End Function

Public Function GetPlanMarks() As Variant()
    GetPlanMarks = This.PlanMarkList.Items
End Function

Public Function GetExeValues() As Variant()
    GetExeValues = This.ExeValueList.Items
End Function

Public Function GetExeMarks() As Variant()
    GetExeMarks = This.ExeMarkList.Items
End Function

Private Sub ValidateTotal(ByVal ID As Dom_ClassHourID)
    With New Dom_ClassHourFilter
        Dim ClassHourFilter As Dom_ClassHourFilter
        Set ClassHourFilter = .Create(ID.ClassHourType, Nothing, ID.Grade)
    End With
    Dim HourTotal As Long
    HourTotal = GetHourTotal(ClassHourFilter)
    With New Dom_TimeTableFilter
        Dim TimeTableFilter As Dom_TimeTableFilter
        Set TimeTableFilter = .Create(ID.ClassHourType, Nothing, ID.Grade)
    End With
    Dim MarkCount As Long
    MarkCount = GetMarkCount(TimeTableFilter)
    If HourTotal = MarkCount Then Exit Sub
    Err.Raise DomErrUnmatch, "Dom_ClassHourAggregate", ID.ClassHourType.ToKey & charHalfSpace & ID.Grade.ToKey & "の時数と時間割を一致させてください。"
End Sub

Private Function GetHourTotal(ByVal Filter As Dom_ClassHourFilter) As Long
    Select Case Filter.ClassHourType.Value
    Case Dom_ClassHourTypeEnum.Plan
        Dim Items() As Variant
        Items = This.PlanValueList.Items
    Case Dom_ClassHourTypeEnum.Execution
        Items = This.ExeValueList.Items
    End Select
    '@Ignore SelfAssignedDeclaration
    Dim Calculator As New Dom_ClassHourCalculator
    GetHourTotal = Calculator.Calculate(Items, Filter)
End Function

Private Function GetMarkCount(ByVal Filter As Dom_TimeTableFilter) As Long
    Select Case Filter.ClassHourType.Value
    Case Dom_ClassHourTypeEnum.Plan
        Dim Items() As Variant
        Items = This.PlanMarkList.Items
    Case Dom_ClassHourTypeEnum.Execution
        Items = This.ExeMarkList.Items
    End Select
    '@Ignore SelfAssignedDeclaration
    Dim Calculator As New Dom_TimeTableCalculator
    GetMarkCount = Calculator.Calculate(Items, Filter)
End Function

Private Function FindPlanValue(ByVal ID As Dom_ClassHourID) As Dom_ClassHour
    Dim Key As String
    Key = ID.ToKey
    If This.PlanValueList.Exists(Key) Then
        Set FindPlanValue = This.PlanValueList.Item(Key)
    End If
End Function

Private Function FindPlanMark(ByVal ID As Dom_TimeTableID) As Dom_TimeTable
    Dim Key As String
    Key = ID.ToKey
    If This.PlanMarkList.Exists(Key) Then
        Set FindPlanMark = This.PlanMarkList.Item(Key)
    End If
End Function

Private Function FindExeValue(ByVal ID As Dom_ClassHourID) As Dom_ClassHour
    Dim Key As String
    Key = ID.ToKey
    If This.ExeValueList.Exists(Key) Then
        Set FindExeValue = This.ExeValueList.Item(Key)
    End If
End Function

Private Function FindExeMark(ByVal ID As Dom_TimeTableID) As Dom_TimeTable
    Dim Key As String
    Key = ID.ToKey
    If This.ExeMarkList.Exists(Key) Then
        Set FindExeMark = This.ExeMarkList.Item(Key)
    End If
End Function

Private Function ToClassHourID(ByVal ID As Dom_TimeTableID) As Dom_ClassHourID
    With New Dom_ClassHourID
        Set ToClassHourID = .Create(ID.ClassHourType, Nothing, ID.Grade)
    End With
End Function
