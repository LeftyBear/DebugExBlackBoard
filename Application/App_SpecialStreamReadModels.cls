VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SpecialStreamReadModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.DTOCollection")
Option Explicit
Private Type Member
    Models As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Models = New VBA.Collection
End Sub

Public Sub Add(ByVal Model As App_SpecialStreamReadModel)
    This.Models.Add Model
End Sub

Public Function GradeCount() As Long
    GradeCount = 1
End Function

Public Function ClassNoCount() As Long
    ClassNoCount = This.Models.Count
End Function

Public Function Item(ByVal Index As Long) As App_SpecialStreamReadModel
    Set Item = This.Models.Item(Index)
End Function
