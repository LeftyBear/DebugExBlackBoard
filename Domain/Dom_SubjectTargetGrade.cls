VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SubjectTargetGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Values As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal Values As VBA.Collection)
    Set This.Values = Values
End Sub

Public Property Get Count() As Long
    Count = This.Values.Count
End Property

Public Property Get GetValue(ByVal Index As Long) As String
    GetValue = This.Values.Item(Index)
End Property
