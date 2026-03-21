VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    Items       As Scripting.Dictionary
    ValueIndex  As Scripting.Dictionary
    MarkIndex   As Scripting.Dictionary
    HourTotal() As Long
    MarkTotal() As Long
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Items = New Scripting.Dictionary
    Set This.ValueIndex = New Scripting.Dictionary
    Set This.MarkIndex = New Scripting.Dictionary
End Sub

Public Sub Add(ByVal Item As Dom_ClassHour)
    Dim IDKey As String
    IDKey = Item.GetID.ToKey
    If This.Items.Exists(IDKey) Then
        Err.Raise DomErrNotUnique, "Dom_ClassHourAggregate", "IDは一意である必要があります。"
    End If
    This.Items.Add IDKey, Item
    Dim ValueIndexKey As String
    ValueIndexKey = Item.GetID.ClassHourType & charUnderScore & Item.GetID.Subject.ToKey & Item.GetID.Grade.ToKey
End Sub

Public Sub UpdateValue(ByVal ID As Dom_ClassHourID, ByVal NewValue As Long)
    Dim Target As Dom_ClassHour
    Set Target = GetById(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "IDに紐づけられたClassHourが必要です。"
    With New Dom_ClassHourValue
        Dim Value As Dom_ClassHourValue
        Set Value = .Create(NewValue)
    End With
    Target.ChangeValueInternal Value
End Sub

Public Sub UpdateMark(ByVal ID As Dom_ClassHourID, ByVal NewMark As String)
    Dim Target As Dom_ClassHour
    Set Target = GetById(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "IDに紐づけられたClassHourが必要です。"
    With New Dom_ClassHourMark
        Dim Mark As Dom_ClassHourMark
        Set Mark = .Create(NewMark)
    End With
    Target.ChangeMarkInternal Mark
End Sub

Public Function Count() As Long
    Count = This.Items.Count
End Function

Public Function GetHours(ByVal ID As Dom_ClassHourID) As Dom_ClassHourValue
    Set GetHours = GetById(ID).GetHours
End Function

Public Function GetMark(ByVal ID As Dom_ClassHourID) As Dom_ClassHourMark
    Set GetMark = GetById(ID).GetMark
End Function

Public Function GetTotal(ByVal Filter As Dom_ClassHourFilter) As Long
    Dim Item As Dom_ClassHour
    For Each Item In This.Items
        If Filter.Match(Item.GetID) Then
            Dim Total As Long
            Total = Total + Item.GetHours.Value
        End If
    Next
    GetTotal = Total
End Function

Private Function Match(ByVal Item As Dom_ClassHour, ByVal Filter As Dom_ClassHourFilter) As Boolean
    If Not Filter.Match(Item.GetID) Then Exit Function
    Match = True
End Function

Private Sub ValidateTotal()
    Dim Item As Dom_ClassHour
    For Each Item In This.Items
        Dim Total As Long
        Total = Total + Item.GetHours.Value
    Next
    If Total > 100 Then
        Err.Raise 5, , "合計超過"
    End If
End Sub

Private Function GetById(ByVal ID As Dom_ClassHourID) As Dom_ClassHour
    Dim Key As String
    Key = ID.ToKey
    If This.Items.Exists(Key) Then
        Set GetById = This.Items(Key)
    End If
End Function
'
'Public Sub Initialize(ByVal MaxType As Long, ByVal MaxSubject As Long, ByVal MaxPeriod As Long, ByVal MaxGrade As Long)
'    ReDim This.Hours(MaxType, MaxSubject, MaxGrade)
'    ReDim This.Marks(MaxType, MaxPeriod, MaxGrade)
'End Sub
'
'Public Sub AddHours(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade, ByVal Hours As Long)
'    This.Hours(ClassHourType, Subject.Index, Grade.Value) = Hours
'End Sub
'
'Public Sub AddMark(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade, ByVal Mark As String)
'    This.Marks(ClassHourType, Period.Value, Grade.Value) = Mark
'End Sub
'
'Public Function GetHours(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Long
'    GetHours = This.Hours(ClassHourType, Subject.Index, Grade.Value)
'End Function
'
'Public Function GetMark(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As String
'    GetMark = This.Marks(ClassHourType, Period.Value, Grade.Value)
'End Function
'
'Public Function MarkCount(ByVal ClassHourType As Dom_ClassHourType, ByVal Grade As Dom_Grade) As Long
'    Dim Period  As Long
'    For Period = 1 To UBound(This.Marks, 2)
'        If This.Marks(ClassHourType, Period, Grade.Value) <> vbNullString Then
'            Dim Count As Long
'            Count = Count + 1
'        End If
'    Next
'    MarkCount = Count
'End Function
'
'Public Function SubjectCount() As Long
'    SubjectCount = UBound(This.Hours, 2)
'End Function
'
'Public Function GradeCount() As Long
'    GradeCount = UBound(This.Hours, 3)
'End Function
'
'Public Function UpdateHours(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade, ByVal NewValue As Long) As Long
'    Dim OldValue As Long
'    OldValue = This.Hours(ClassHourType, Subject.Index, Grade.Value)
'    This.Hours(ClassHourType, Subject.Index, Grade.Value) = NewValue
'    UpdateHours = NewValue - OldValue
'End Function
