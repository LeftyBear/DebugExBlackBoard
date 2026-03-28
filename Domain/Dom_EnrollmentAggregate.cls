VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregate"
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
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.ItemList = New Scripting.Dictionary
    Set This.ValueIndex = New Scripting.Dictionary
End Sub

Public Sub AddInitialEnrollment(ByVal Item As Dom_Enrollment)
    This.NormDate = Item.GetDate
    Dim IDKey As String
    IDKey = Item.GetID.ToKey
    If This.ItemList.Exists(IDKey) Then Err.Raise DomErrNotUnique, "Dom_EnrollmentAggregate", "キーは一意である必要があります。" & vbCrLf & "重複しているIDキー: " & IDKey
    This.ItemList.Add IDKey, Item
    If Item.GetID.StreamType.Value = Dom_StreamTypeEnum.Main Then
        Dim ValueIndexKey As String
        ValueIndexKey = Item.GetID.Grade.Value & charUnderScore & Item.GetID.ClassNo.ToKey & charUnderScore & Item.GetID.Gender.ToKey
    Else
        ValueIndexKey = jpnTokushi & charUnderScore & Item.GetID.ClassNo.ToKey & Item.GetID.Gender.ToKey
    End If
    AddToIndex This.ValueIndex, ValueIndexKey, IDKey, Item
End Sub

Public Sub UpdateValue(ByVal ID As Dom_EnrollmentID, ByVal NewValue As Long)
    Dim Target As Dom_Enrollment
    Set Target = GetById(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_EnrollmentAggregate", "IDをキーに持つオブジェクトが必要です。"
    With New Dom_EnrollmentValue
        Dim Value As Dom_EnrollmentValue
        Set Value = .Create(NewValue)
    End With
    Target.ChangeValueInternal Value
    ValidateValueTotal ID
End Sub

Public Function Count() As Long
    Count = This.ItemList.Count
End Function

Public Function GetDate() As Date
    GetDate = This.NormDate
End Function

Public Function GetValueTotal(ByVal Filter As Dom_EnrollmentFilter) As Long
    '@Ignore SelfAssignedDeclaration
    Dim Calculator As New Dom_EnrollmentCalculator
    GetValueTotal = Calculator.SumValue(Enrollments(Filter), Filter)
End Function

Private Sub ValidateValueTotal(ByVal ID As Dom_EnrollmentID)
    With New Dom_EnrollmentFilter
        Dim Filter As Dom_EnrollmentFilter
        Set Filter = .Create(ID.StreamType, ID.Grade, ID.ClassNo, ID.Gender)
    End With
    Dim ValueTotal As Long
    ValueTotal = Me.GetValueTotal(Filter)
    If ValueTotal < 0 Then Err.Raise DomErrNegativeNumber, "Dom_EnrollmentAggregate", "合計は正の数になるはずです。"
End Sub

Private Sub AddToIndex(ByVal Index As Scripting.Dictionary, ByVal Key As String, ByVal IDKey As String, ByVal Item As Dom_Enrollment)
    Dim Bucket As Scripting.Dictionary
    If Index.Exists(Key) Then
        Set Bucket = Index.Item(Key)
    Else
        Set Bucket = New Scripting.Dictionary
        Index.Add Key, Bucket
    End If
    Bucket.Add IDKey, Item
End Sub

Private Function GetById(ByVal ID As Dom_EnrollmentID) As Dom_Enrollment
    Dim Key As String
    Key = ID.ToKey
    If This.ItemList.Exists(Key) Then
        Set GetById = This.ItemList.Item(Key)
    End If
End Function

Private Function Enrollments(ByVal Filter As Dom_EnrollmentFilter) As Variant()
    If Not Filter.HasType And Not Filter.HasClassNo And Not Filter.HasGender Then
        Dim ValueIndexKey As String
        ValueIndexKey = jpnTokushi & charUnderScore & Filter.ClassNo.ToKey & charUnderScore & Filter.Gender.ToKey
    ElseIf Not Filter.HasGrade And Not Filter.HasClassNo And Not Filter.HasGender Then
        ValueIndexKey = Filter.Grade.ToKey & charUnderScore & Filter.ClassNo.ToKey & charUnderScore & Filter.Gender.ToKey
    End If
    If This.ValueIndex.Exists(ValueIndexKey) Then
        Dim Target As Scripting.Dictionary
        Set Target = This.ValueIndex.Item(ValueIndexKey)
        Enrollments = Target.Items
        Exit Function
    End If
    Enrollments = This.ItemList.Items
End Function
