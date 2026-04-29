VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_IClassHourRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Interface")
Option Explicit

Public Sub Add(ByVal Item As Dom_ClassHourAggregate)
End Sub

Public Function FindByID(ByVal ID As Dom_DateID) As Dom_ClassHourAggregate
End Function

Public Function GetAll() As VBA.Collection
End Function
