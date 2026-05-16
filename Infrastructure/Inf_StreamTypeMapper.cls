VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_StreamTypeMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Private Type Member
    Factory As Dom_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function ToRaw(ByVal TypeCode As Dom_StreamType) As String
    If TypeCode.IsMain Then
        ToRaw = vbNullString
    ElseIf TypeCode.IsSpecial Then
        ToRaw = jpnTokushi
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_StreamType
    If 0 < VBA.InStr(1, Raw, jpnNen) Then
        Dim TypeCode As Long
        TypeCode = Dom_StreamTypePolicy.Main
    ElseIf 0 < VBA.InStr(1, Raw, jpnTokushi) Then
        TypeCode = Dom_StreamTypePolicy.Special
    End If
    Set ToDomain = This.Factory.CreateStreamType(TypeCode)
End Function

Public Function ToModel(ByVal Raw As String) As String
    If 0 < VBA.InStr(1, Raw, jpnNen) Then
        ToModel = "Main"
    ElseIf 0 < VBA.InStr(1, Raw, jpnTokushi) Then
        ToModel = "Special"
    End If
End Function
