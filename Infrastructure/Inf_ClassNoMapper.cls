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
    Dim ClassNoString As String
    ClassNoString = VBA.Replace(Raw, COLUMN_CLASSNO, vbNullString)
    Dim ClassNo As Long
    ClassNo = Inf_NumericUtility.ExtractNumber(ClassNoString)
    Set ToDomain = Dom_ValueObjectFactory.CreateClassNo(ClassNo)
End Function
