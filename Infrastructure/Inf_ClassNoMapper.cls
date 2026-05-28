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

Private Type Member
    Factory As Dom_ValueObjectFactory
End Type

Private This As Member

Private Const CLASS_NUMBER As String = "‘g"

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function ToRaw(ByVal ClassNo As Dom_ClassNo) As String
    ToRaw = CStr(ClassNo.Value) & CLASS_NUMBER
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassNo
    If 0 < VBA.InStr(1, Raw, CLASS_NUMBER) Then
        Dim ClassNo As String
        ClassNo = GetClassNo(Raw)
    Else
        Exit Function
    End If
    Set ToDomain = This.Factory.CreateClassNo(ClassNo)
End Function

Public Function ToModel(ByVal Raw As String) As Long
    ToModel = GetClassNo(Raw)
End Function

Private Function GetClassNo(ByVal Raw As String) As Long
    Dim Buffer As String
    Buffer = VBA.Replace(Raw, CLASS_NUMBER, vbNullString)
    GetClassNo = Inf_NumericUtility.ExtractNumber(Buffer)
End Function

Public Property Get ColumnClassNo() As String
    ColumnClassNo = CLASS_NUMBER
End Function
