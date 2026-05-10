VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentMetaMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.DTO")
Option Explicit
Private Type Member
    MetaList As Scripting.Dictionary
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.MetaList = New Scripting.Dictionary
End Sub

Public Sub Add(ByVal Key As String, ByVal Meta As Inf_EnrollmentColumnMeta)
    This.MetaList.Add Key, Meta
End Sub

Public Function Count() As Long
    Count = This.MetaList.Count
End Function

Public Function Exists(ByVal Key As String) As Boolean
    Exists = This.MetaList.Exists(Key)
End Function

Public Function Item(ByVal Key As String) As Inf_EnrollmentColumnMeta
    Set Item = This.MetaList.Item(Key)
End Function
