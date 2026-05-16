VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassNoMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Mapper")
Option Explicit

Public Function ToRaw(ByVal ClassNo As Dom_ClassNo) As String
    ToRaw = CStr(ClassNo.Value) & jpnKumi
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassNo
    Dim ClassNo As Long
    If 0 < VBA.InStr(1, Raw, jpnKumi) Then
        Dim Buffer As String
        Buffer = VBA.Replace(Raw, jpnKumi, vbNullString)
        ClassNo = Inf_NumericUtility.ExtractNumber(Buffer)
    End If
    Dim Factory As Dom_ValueObjectFactory
    Set Factory = New Dom_ValueObjectFactory
    Set ToDomain = Factory.CreateClassNo(ClassNo)
End Function
