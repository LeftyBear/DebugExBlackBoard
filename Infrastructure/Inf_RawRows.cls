VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_RawRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.DTO")
Option Explicit
Private Type Member
    Fields  As VBA.Collection
    Records As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Fields = New VBA.Collection
    Set This.Records = New VBA.Collection
End Sub

Public Sub InitializeFileds()
    Set This.Fields = New VBA.Collection
End Sub

Public Sub AddFiled(ByVal RawText As String)
    This.Fields.Add RawText
End Sub

Public Sub AddRecord(ByVal Fields As VBA.Collection)
    This.Records.Add Fields
End Sub

Public Sub Remove(ByVal Index As Long)
    This.Records.Remove Index
End Sub

Public Property Get Fields() As VBA.Collection
    Set Fields = This.Fields
End Property

Public Property Get Records() As VBA.Collection
    Set Records = This.Records
End Property

Public Function FieldsCount() As Long
    If This.Records.Count = 0 Then Exit Function
    Dim Field As VBA.Collection
    Set Field = This.Records.Item(1)
    FieldsCount = Fields.Count
End Function

Public Function RecordsCount() As Long
    If This.Records.Count = 0 Then Exit Function
    RecordsCount = This.Records.Count
End Function

Public Function GetHeaderRows() As VBA.Collection
    If This.Records.Count = 0 Then Exit Function
    Set GetHeaderRows = This.Records.Item(1)
End Function

Public Function GetField(ByVal RecordIndex As Long, ByVal FieldIndex As Long) As String
    If This.Records.Count = 0 Then Exit Function
    If RecordIndex < 1 Or This.Records.Count < RecordIndex Then Exit Function
    Dim Field As VBA.Collection
    Set Field = This.Records.Item(RecordIndex)
    If FieldIndex < 1 Or Fields.Count < FieldIndex Then Exit Function
    GetField = Field.Item(FieldIndex)
End Function
