VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ISpecialStreamRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.RepositoryInterface")
Option Explicit

Public Sub Add(ByVal Entry As Dom_SpecialStream)
End Sub

Public Function FindByIndex(ByVal Index As Long) As Dom_SpecialStream
End Function

Public Function GetAll() As VBA.Collection
End Function

Public Function Replace(ByVal Name As String, ByVal NewEntry As Dom_SpecialStream) As VBA.Collection
End Function

Public Function ReadAll() As VBA.Collection
End Function

Public Function Count() As Long
End Function
