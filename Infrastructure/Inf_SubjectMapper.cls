VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Public Function ToRaw(ByVal Subject As Dom_SubjectNew) As String
    ToRaw = Subject.Value
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_SubjectNew
    Dim Factory As New Dom_ValueObjectFactory
    Set ToDomain = Factory.CreateSubject(Raw)
End Function
