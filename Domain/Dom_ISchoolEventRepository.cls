VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ISchoolEventRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.RepositoryInterface")
Option Explicit

Public Sub Save()
End Sub

Public Sub Replace(ByVal Name As String, ByVal NewEntry As Dom_SchoolEvent)
End Sub

Public Function FindByName(ByVal Name As String) As Dom_SchoolEvent
End Function
