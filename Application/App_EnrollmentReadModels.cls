VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentReadModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.DTOCollection")
Option Explicit
Private Type Member
    MainMaleModels      As VBA.Collection
    MainFemaleModels    As VBA.Collection
    SpecialMaleModels   As VBA.Collection
    SpecialFemaleModels As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.MainMaleModels = New VBA.Collection
    Set This.MainFemaleModels = New VBA.Collection
    Set This.SpecialMaleModels = New VBA.Collection
    Set This.SpecialFemaleModels = New VBA.Collection
End Sub

Public Sub AddMainMale(ByVal Model As App_EnrollmentReadModel)
    This.MainMaleModels.Add Model
End Sub

Public Sub AddMainFemale(ByVal Model As App_EnrollmentReadModel)
    This.MainFemaleModels.Add Model
End Sub

Public Sub AddSpecialMale(ByVal Model As App_EnrollmentReadModel)
    This.SpecialMaleModels.Add Model
End Sub

Public Sub AddSpecialFemale(ByVal Model As App_EnrollmentReadModel)
    This.SpecialFemaleModels.Add Model
End Sub

Public Function MainMaleCount() As Long
    MainMaleCount = This.MainMaleModels.Count
End Function

Public Function MainFemaleCount() As Long
    MainFemaleCount = This.MainFemaleModels.Count
End Function

Public Function SpecialMaleCount() As Long
    SpecialMaleCount = This.SpecialMaleModels.Count
End Function

Public Function SpecialFemaleCount() As Long
    SpecialFemaleCount = This.SpecialFemaleModels.Count
End Function

Public Function GetMainMale(ByVal INDEX As Long) As App_EnrollmentReadModel
    Set GetMainMale = This.MainMaleModels.Item(INDEX)
End Function

Public Function GetMainFemale(ByVal INDEX As Long) As App_EnrollmentReadModel
    Set GetMainFemale = This.MainFemaleModels.Item(INDEX)
End Function

Public Function GetSpecialMaleGrand(ByVal INDEX As Long) As App_EnrollmentReadModel
    Set GetSpecialMaleGrand = This.SpecialMaleModels.Item(INDEX)
End Function

Public Function GetSpecialFemaleGrand(ByVal INDEX As Long) As App_EnrollmentReadModel
    Set GetSpecialFemaleGrand = This.SpecialFemaleModels.Item(INDEX)
End Function
