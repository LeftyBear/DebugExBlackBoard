VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SubjectFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal Row As App_SubjectRow) As Dom_Subject
    Dim Entity As Dom_Subject
    Set Entity = New Dom_Subject
    Entity.Initialize Row.Name, Row.TargetGrade, Row.Mark
    Set Create = Entity
End Function
