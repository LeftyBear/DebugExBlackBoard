VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourTypeMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Public Function ToRaw(ByVal TypeCode As Dom_ClassHourType) As String
    If TypeCode.IsPlan Then
        ToRaw = jpnKeikaku
    ElseIf TypeCode.IsExecution Then
        ToRaw = jpnJisshi
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassHourType
    Dim TypeCode As Long
    If 0 < VBA.InStr(1, Raw, jpnKeikaku) Then
        TypeCode = Dom_ClassHourTypePolicy.Plan
    ElseIf 0 < VBA.InStr(1, Raw, jpnJisshi) Then
        TypeCode = Dom_ClassHourTypePolicy.Execution
    End If
    Dim Factory As New Dom_ValueObjectFactory
    Set ToDomain = Factory.CreateClassHourType(TypeCode)
End Function
