VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_StructureUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    SchoolEvent     As App_LoadSchoolEventUseCase
    SpecialStream   As App_LoadSpecialStreamUseCase
    Subject         As App_LoadSubjectUseCase
    UpperValue      As App_LoadUpperValueUseCase
End Type

Private This As Member

Public Sub Initialize(ByVal SchoolEvent As App_LoadSchoolEventUseCase, ByVal SpecialStream As App_LoadSpecialStreamUseCase, ByVal Subject As App_LoadSubjectUseCase, ByVal UpperValue As App_LoadUpperValueUseCase)
    Set This.SchoolEvent = SchoolEvent
    Set This.SpecialStream = SpecialStream
    Set This.Subject = Subject
    Set This.UpperValue = UpperValue
End Sub

Public Function Create() As App_LoadStructureUseCase
    Dim Result As App_LoadStructureUseCase
    Set Result = New App_LoadStructureUseCase
    Result.Initialize This.SchoolEvent, This.SpecialStream, This.Subject, This.UpperValue
    Set Create = Result
End Function
