VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourTypeMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Const COLUMN_PLAN       As String = "Œv‰æ"
Private Const COLUMN_EXECUTION  As String = "ŽÀŽ{"

Public Function ToRaw(ByVal ClassHourType As Dom_ClassHourType) As String
    If ClassHourType.IsPlan Then
        ToRaw = COLUMN_PLAN
    ElseIf ClassHourType.IsExecution Then
        ToRaw = COLUMN_EXECUTION
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassHourType
    Dim TypeCode As Dom_ClassHourTypeCode
    If 0 < VBA.InStr(1, Raw, COLUMN_PLAN) Then
        TypeCode = PlanCode
    ElseIf 0 < VBA.InStr(1, Raw, COLUMN_EXECUTION) Then
        TypeCode = ExeCode
    Else
        Exit Function
    End If
    Set ToDomain = Dom_ValueObjectFactory.CreateClassHourType(TypeCode)
End Function
