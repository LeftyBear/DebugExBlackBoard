VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Service")
Option Explicit

Private Type Member
    SchoolEvent     As Dom_ISchoolEventRepository
    SpecialStream   As Dom_ISpecialStreamRepository
    Subject         As Dom_ISubjectRepository
    UpperValue      As Dom_IUpperValueRepository
End Type

Private This As Member

Public Sub Initialize(ByVal SchoolEvent As Dom_ISchoolEventRepository, ByVal SpecialStream As Dom_ISpecialStreamRepository, ByVal Subject As Dom_ISubjectRepository, ByVal UpperValue As Dom_IUpperValueRepository)
    Set This.SchoolEvent = SchoolEvent
    Set This.SpecialStream = SpecialStream
    Set This.Subject = Subject
    Set This.UpperValue = UpperValue
End Sub

Property Get SchoolEvent() As Dom_ISchoolEventRepository
    Set SchoolEvent = This.SchoolEvent
End Property

Property Get SpecialStream() As Dom_ISpecialStreamRepository
    Set SpecialStream = This.SpecialStream
End Property

Property Get Subject() As Dom_ISubjectRepository
    Set Subject = This.Subject
End Property

Property Get UpperValue() As Dom_IUpperValueRepository
    Set UpperValue = This.UpperValue
End Property

