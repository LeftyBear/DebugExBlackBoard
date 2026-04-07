VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
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

Public Sub UpdatePlanValue(ByVal ID As Dom_ClassHourID, ByVal NewValue As Long)
    Dim Target As Dom_ClassHour
    Set Target = GetPlanValueByID(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourColumn", "IDをキーに持つオブジェクトが必要です。"
    With New Dom_ClassHourValue
        Dim Value As Dom_ClassHourValue
        Set Value = .Create(NewValue)
    End With
    Target.ChangeValueInternal Value
End Sub

Public Sub UpdatePlanMark(ByVal ID As Dom_ClassHourID, ByVal NewMark As String)
    Dim Target As Dom_TimeTable
    Set Target = GetPlanMarkByID(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourColumn", "IDをキーに持つオブジェクトが必要です。"
    Target.ChangeValueInternal NewMark
End Sub

Public Sub UpdateExeValue(ByVal ID As Dom_ClassHourID, ByVal NewValue As Long)
    Dim Target As Dom_ClassHour
    Set Target = GetExeValueByID(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourColumn", "IDをキーに持つオブジェクトが必要です。"
    With New Dom_ClassHourValue
        Dim Value As Dom_ClassHourValue
        Set Value = .Create(NewValue)
    End With
    Target.ChangeValueInternal Value
End Sub

Public Sub UpdateExeMark(ByVal ID As Dom_ClassHourID, ByVal NewMark As String)
    Dim Target As Dom_TimeTable
    Set Target = GetExeMarkByID(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourColumn", "IDをキーに持つオブジェクトが必要です。"
    Target.ChangeValueInternal NewMark
End Sub

Public Function GetDateKey() As String
    GetDateKey = This.DateID.ToKey
End Function

Public Function GetHourTotal(ByVal Filter As Dom_ClassHourFilter) As Long
    '@Ignore SelfAssignedDeclaration
    Dim Calculator As New Dom_ClassHourCalculator
    GetHourTotal = Calculator.Calculate(ClassHourItems(Filter), Filter)
End Function

Public Function GetMarkCount(ByVal Filter As Dom_TimeTableFilter) As Long
    '@Ignore SelfAssignedDeclaration
    Dim Calculator As New Dom_TimeTableMarkCalculator
    GetMarkCount = Calculator.Calculate(TimeTableItems(Filter), Filter)
End Function

Private Function GetPlanValueByID(ByVal ID As Dom_ClassHourID) As Dom_ClassHour
    Dim Key As String
    Key = ID.ToKey
    If This.PlanValueList.Exists(Key) Then
        Set GetPlanValueByID = This.PlanValueList.Item(Key)
    End If
End Function

Private Function GetPlanMarkByID(ByVal ID As Dom_ClassHourID) As Dom_TimeTable
    Dim Key As String
    Key = ID.ToKey
    If This.PlanMarkList.Exists(Key) Then
        Set GetPlanMarkByID = This.PlanMarkList.Item(Key)
    End If
End Function

Private Function GetExeValueByID(ByVal ID As Dom_ClassHourID) As Dom_ClassHour
    Dim Key As String
    Key = ID.ToKey
    If This.ExeValueList.Exists(Key) Then
        Set GetExeValueByID = This.ExeValueList.Item(Key)
    End If
End Function

Private Function GetExeMarkByID(ByVal ID As Dom_ClassHourID) As Dom_TimeTable
    Dim Key As String
    Key = ID.ToKey
    If This.ExeMarkList.Exists(Key) Then
        Set GetExeMarkByID = This.ExeMarkList.Item(Key)
    End If
End Function

Private Function ClassHourItems(ByVal Filter As Dom_ClassHourFilter) As Variant()
    If Not Filter.HasType And Not Filter.HasSubject And Not Filter.HasGrade Then
        If Filter.ClassHourType.ToKey = jpnJisshi Then
            ClassHourItems = This.PlanValueList.Items
        Else
            ClassHourItems = This.ExeValueList.Items
        End If
    Else
        Err.Raise DomErrEmptyFilter, "Dom_ClassHourColumn", "時数フィルターが設定されている必要があります。"
    End If
End Function

Private Function TimeTableItems(ByVal Filter As Dom_TimeTableFilter) As Variant()
    If Not Filter.HasType And Not Filter.HasPeriod And Not Filter.HasGrade Then
        If Filter.ClassHourType.ToKey = jpnJisshi Then
            TimeTableItems = This.PlanMarkList.Items
        Else
            TimeTableItems = This.ExeMarkList.Items
        End If
    Else
        Err.Raise DomErrEmptyFilter, "Dom_ClassHourColumn", "時間割フィルターが設定されている必要があります。"
    End If
End Function
