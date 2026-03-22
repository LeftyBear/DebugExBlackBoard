VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_GenerateSchoolStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Repository As Dom_IConfigRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_IConfigRepository)
    Set This.Repository = Repository
End Sub

Public Function Execute() As Dom_SchoolStructure
    With New App_UpperValueBuilder
        Dim UpperValue As Dom_UpperValue
        Set UpperValue = .Build(This.Repository)
    End With
    With New App_SpecialStreamCatalogBuilder
        Dim SpecialStreamCatalog As Dom_SpecialStreamCatalog
        Set SpecialStreamCatalog = .Build(This.Repository)
    End With
    With New App_SubjectCatalogBuilder
        Dim SubjectCatalog As Dom_SubjectCatalog
        Set SubjectCatalog = .Build(This.Repository)
    End With
    Dim Result As Dom_SchoolStructure
    Set Result = New Dom_SchoolStructure
    Result.Initialize UpperValue, SpecialStreamCatalog, SubjectCatalog
    Set Execute = Result
End Function
