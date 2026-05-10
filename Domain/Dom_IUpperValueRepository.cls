VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_IUpperValueRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.RepositoryInterface")
Option Explicit

Public Sub Add(ByVal Entry As Dom_UpperValue)
End Sub

Public Function Item() As Dom_UpperValue
End Function

Public Function Replace(ByVal NewEntry As Dom_UpperValue) As VBA.Collection
End Function

Public Function ReadAll() As VBA.Collection
End Function
