VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassNoMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Const COLUMN_CLASSNO As String = "‘g"

Public Function ToRaw(ByVal ClassNo As Dom_ClassNo) As String
    ToRaw = CStr(ClassNo.Value) & COLUMN_CLASSNO
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassNo
    If VBA.InStr(1, Raw, COLUMN_CLASSNO) = 0 Then Exit Function
    Dim ClassNo As Long
    ClassNo = ParseClassNo(Raw)
    Set ToDomain = Dom_ValueObjectFactory.CreateClassNo(ClassNo)
End Function

Private Function ParseClassNo(ByVal Raw As String) As Long
    Dim ClassNo As String
    ClassNo = VBA.Replace(Raw, COLUMN_CLASSNO, vbNullString)
    ParseClassNo = Inf_NumericUtility.ExtractNumber(ClassNo)
End Function
