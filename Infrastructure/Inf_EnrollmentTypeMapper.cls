VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentTypeMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Public Function ToRaw(ByVal TypeCode As Dom_StreamType) As String
    If TypeCode.IsMain Then
        ToRaw = vbNullString
    ElseIf TypeCode.IsSpecial Then
        ToRaw = jpnTokushi
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_StreamType
    Dim TypeCode As Long
    If 0 < VBA.InStr(1, Raw, jpnNen) Then
        TypeCode = Dom_StreamTypePolicy.Main
    ElseIf 0 < VBA.InStr(1, Raw, jpnTokushi) Then
        TypeCode = Dom_StreamTypePolicy.Special
    End If
    Dim Factory As Dom_ValueObjectFactory
    Set Factory = New Dom_ValueObjectFactory
    Set ToDomain = Factory.CreateStreamType(TypeCode)
End Function
