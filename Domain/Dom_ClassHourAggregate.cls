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
    Entities As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal Entities As VBA.Collection)
    If Entities Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ClassHourAggregate", "コレクションが生成されている必要があります。"
    If Entities.Count = 0 Then Err.Raise DomErrEmptyObject, "Dom_ClassHourAggregate", "コレクションのアイテムは１件以上ある必要があります。"
    ValidateRecord Entities
    Set This.Entities = Entities
End Sub

Public Property Get Count() As Long
    Count = This.Entities.Count
End Property

Public Function GetClassHour(ByVal Index As Long) As Dom_ClassHour
    Set GetClassHour = This.Entities.Item(Index)
End Function

Private Sub ValidateRecord(ByVal Entities As VBA.Collection)
    Dim i As Long
    For i = 1 To Entities.Count
        If Not TypeOf Entities.Item(i) Is Dom_ClassHour Then
            Err.Raise DomErrInvalidTypeOfObject, "Dom_ClassHourAggregate", "エンティティは ClassHour である必要があります。"
        End If
    Next
End Sub
