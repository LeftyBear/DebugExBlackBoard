VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectNameMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Public Function ToRaw(ByVal SubjectName As Dom_SubjectName) As String
    ToRaw = SubjectName.Value
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_SubjectName
    Set ToDomain = Dom_ValueObjectFactory.CreateSubjectName(Raw)
End Function
