VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Facade")
Option Explicit

Private Type Member
    UpperValue              As Dom_UpperValue
    SpecialStreamCatalog    As Dom_SpecialStreamCatalog
    SubjectCatalog          As Dom_SubjectCatalog
    SchoolEventCatalog      As Dom_SchoolEventCatalog
End Type

Private This As Member

Public Sub Initialize(ByVal UpperValue As Dom_UpperValue, ByVal SpecialStreamCatalog As Dom_SpecialStreamCatalog, ByVal SubjectCatalog As Dom_SubjectCatalog, ByVal SchoolEventCatalog As Dom_SchoolEventCatalog)
    Set This.UpperValue = UpperValue
    Set This.SpecialStreamCatalog = SpecialStreamCatalog
    Set This.SubjectCatalog = SubjectCatalog
    Set This.SchoolEventCatalog = SchoolEventCatalog
End Sub

Public Property Get MainStreamGradeCount() As Long
    MainStreamGradeCount = This.UpperValue.MainStreamGradeCount
End Property

Public Property Get MainStreamClassCount() As Long
    MainStreamClassCount = This.UpperValue.MainStreamClassCount
End Property

Public Function ResolveSpecialStream(ByVal Name As String) As Dom_SpecialStream
    Set ResolveSpecialStream = This.SpecialStreamCatalog.Item(Name)
End Function

Public Function SpecialStreamClassNameOf(ByVal Index As Long) As String
    SpecialStreamClassNameOf = This.SubjectCatalog.NameOf(Index)
End Function

Public Property Get SpecialStreamGradeCount() As Long
    SpecialStreamGradeCount = 1
End Property

Public Property Get SpecialStreamClassCount() As Long
    SpecialStreamClassCount = This.SpecialStreamCatalog.Count
End Property

Public Function ResolveSubject(ByVal Name As String) As Dom_Subject
    Set ResolveSubject = This.SubjectCatalog.Resolve(Name)
End Function

Public Function SubjectIndexOf(ByVal Index As Long) As Dom_Subject
    SubjectIndexOf = This.SubjectCatalog.IndexOf(Index)
End Function

Public Function SubjectNameOf(ByVal Index As Long) As String
    SubjectNameOf = This.SubjectCatalog.NameOf(Index)
End Function

Public Function SubjectCount() As Long
    SubjectCount = This.SubjectCatalog.Count
End Function

Public Function SubjectNames() As String()
    SubjectNames = This.SubjectCatalog.Names
End Function

Public Function IsSubject(ByVal Name As String) As Boolean
    IsSubject = This.SubjectCatalog.Exists(Name)
End Function

Public Function IsSchoolEvent(ByVal Name As String) As Boolean
    IsSchoolEvent = This.SchoolEventCatalog.Exists(Name)
End Function

Public Property Get PeriodCount() As Long
    PeriodCount = This.UpperValue.PeriodCount
End Property

Public Function SchoolEventNames() As String()
    SchoolEventNames = This.SchoolEventCatalog.Names
End Function
