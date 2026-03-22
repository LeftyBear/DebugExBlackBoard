VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    NormDate    As Date
    ItemList    As Scripting.Dictionary
    ValueIndex  As Scripting.Dictionary
    MarkIndex   As Scripting.Dictionary
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.ItemList = New Scripting.Dictionary
    Set This.ValueIndex = New Scripting.Dictionary
    Set This.MarkIndex = New Scripting.Dictionary
End Sub

Public Sub AddInitialClassHour(ByVal Item As Dom_ClassHour)
    This.NormDate = Item.GetDate
    Dim IDKey As String
    IDKey = Item.GetID.ToKey
    If This.ItemList.Exists(IDKey) Then Err.Raise DomErrNotUnique, "Dom_ClassHourAggregate", "キーは一意である必要があります。" & vbCrLf & "重複しているIDキー: " & IDKey
    This.ItemList.Add IDKey, Item
    Dim ValueIndexKey As String
    ValueIndexKey = Item.GetID.ClassHourType.ToKey & charUnderScore & Item.GetID.Subject.ToKey & Item.GetID.Grade.ToKey
    AddToIndex This.ValueIndex, ValueIndexKey, IDKey, Item
    Dim MarkIndexKey As String
    MarkIndexKey = Item.GetID.ClassHourType.ToKey & charUnderScore & Item.GetID.Period.ToKey & Item.GetID.Grade.ToKey
    AddToIndex This.MarkIndex, MarkIndexKey, IDKey, Item
End Sub

Public Sub UpdateHours(ByVal ID As Dom_ClassHourID, ByVal NewHours As Long)
    Dim Target As Dom_ClassHour
    Set Target = GetById(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "IDをキーに持つオブジェクトが必要です。"
    With New Dom_ClassHourValue
        Dim Value As Dom_ClassHourValue
        Set Value = .Create(NewHours)
    End With
    Target.ChangeValueInternal Value
    ValidateHourTotal ID
End Sub

Public Sub UpdateMark(ByVal ID As Dom_ClassHourID, ByVal NewMark As String)
    Dim Target As Dom_ClassHour
    Set Target = GetById(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "IDをキーに持つオブジェクトが必要です。"
    With New Dom_ClassHourMark
        Dim Mark As Dom_ClassHourMark
        Set Mark = .Create(NewMark)
    End With
    Target.ChangeMarkInternal Mark
    ValidateHourTotal ID
End Sub

Public Function Count() As Long
    Count = This.ItemList.Count
End Function

Public Function GetDate() As Date
    GetDate = This.NormDate
End Function

Public Function GetHourTotal(ByVal Filter As Dom_ClassHourFilter) As Long
    '@Ignore SelfAssignedDeclaration
    Dim Calculator As New Dom_ClassHourCalculator
    GetHourTotal = Calculator.SumSubjectHour(ClassHours(Filter), Filter)
End Function

Public Function GetMarkCount(ByVal Filter As Dom_ClassHourFilter) As Long
    '@Ignore SelfAssignedDeclaration
    Dim Calculator As New Dom_ClassHourCalculator
    GetMarkCount = Calculator.SumMark(ClassHours(Filter), Filter)
End Function

Private Sub ValidateHourTotal(ByVal ID As Dom_ClassHourID)
    With New Dom_ClassHourFilter
        Dim Filter As Dom_ClassHourFilter
        Set Filter = .Create(ID.ClassHourType, ID.Subject, ID.Period, ID.Grade)
    End With
    Dim HourTotal As Long
    HourTotal = Me.GetHourTotal(Filter)
    Dim MarkCount As Long
    MarkCount = Me.GetMarkCount(Filter)
    If HourTotal <> MarkCount Then
        Err.Raise DomErrUnmatch, "Dom_ClassHourAggregate", ID.ClassHourType.ToKey & charHalfSpace & ID.Grade.ToKey & "の時数と時間割を一致させてください。"
    End If
End Sub

Private Sub AddToIndex(ByVal Index As Scripting.Dictionary, ByVal Key As String, ByVal IDKey As String, ByVal Item As Dom_ClassHour)
    Dim Bucket As Scripting.Dictionary
    If Index.Exists(Key) Then
        Set Bucket = Index.Item(Key)
    Else
        Set Bucket = New Scripting.Dictionary
        Index.Add Key, Bucket
    End If
    Bucket.Add IDKey, Item
End Sub

Private Function GetById(ByVal ID As Dom_ClassHourID) As Dom_ClassHour
    Dim Key As String
    Key = ID.ToKey
    If This.ItemList.Exists(Key) Then
        Set GetById = This.ItemList.Item(Key)
    End If
End Function

Private Function ClassHours(ByVal Filter As Dom_ClassHourFilter) As Variant()
    If Not Filter.HasType And Not Filter.HasSubject And Not Filter.HasGrade Then
        Dim ValueIndexKey As String
        ValueIndexKey = Filter.ClassHourType.ToKey & charUnderScore & Filter.Subject.ToKey & Filter.Grade.ToKey
        If This.ValueIndex.Exists(ValueIndexKey) Then
            Dim Target As Scripting.Dictionary
            Set Target = This.ValueIndex.Item(ValueIndexKey)
            ClassHours = Target.Items
            Exit Function
        End If
    End If
    If Not Filter.HasType And Not Filter.HasPeriod And Not Filter.HasGrade Then
        Dim MarkIndexKey As String
        MarkIndexKey = Filter.ClassHourType.ToKey & charUnderScore & Filter.Period.ToKey & Filter.Grade.ToKey
        If This.MarkIndex.Exists(MarkIndexKey) Then
            Set Target = This.MarkIndex.Item(MarkIndexKey)
            ClassHours = Target.Items
            Exit Function
        End If
    End If
    ClassHours = This.ItemList.Items
End Function
