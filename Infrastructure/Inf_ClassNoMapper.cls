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

Private Const CLASS_NO As String = "‘g"

Public Property Get ColumnClassNo() As String
    ColumnClassNo = CLASS_NO
End Property

Public Function ToRaw(ByVal ClassNo As Dom_ClassNo) As String
    ToRaw = CStr(ClassNo.Value) & CLASS_NO
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassNo
    If VBA.InStr(1, Raw, CLASS_NO) = 0 Then Exit Function
    Dim ClassNo As Long
    ClassNo = GetClassNo(Raw)
    Set ToDomain = Dom_ValueObjectFactory.CreateClassNo(ClassNo)
End Function

Public Function ToModel(ByVal Raw As String) As Long
    ToModel = GetClassNo(Raw)
End Function

Private Function GetClassNo(ByVal Raw As String) As Long
    Dim Buffer As String
    Buffer = VBA.Replace(Raw, CLASS_NO, vbNullString)
    GetClassNo = Inf_NumericUtility.ExtractNumber(Buffer)
End Function
