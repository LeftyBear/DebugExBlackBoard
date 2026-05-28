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

Private Type Member
    Factory As Dom_ValueObjectFactory
End Type

Private This As Member

Private Const COLUMN_PLAN       As String = "Œv‰æ"
Private Const COLUMN_EXECUTION  As String = "ŽÀŽ{"

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function ToRaw(ByVal TypeCode As Dom_ClassHourType) As String
    If TypeCode.IsPlan Then
        ToRaw = COLUMN_PLAN
    ElseIf TypeCode.IsExecution Then
        ToRaw = COLUMN_EXECUTION
    End If
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_ClassHourType
    If 0 < VBA.InStr(1, Raw, COLUMN_PLAN) Then
        Dim TypeCode As Long
        TypeCode = Dom_ClassHourTypePolicy.Plan
    ElseIf 0 < VBA.InStr(1, Raw, COLUMN_EXECUTION) Then
        TypeCode = Dom_ClassHourTypePolicy.Execution
    Else
        Exit Function
    End If
    Set ToDomain = This.Factory.CreateClassHourType(TypeCode)
End Function

Public Function ToModel(ByVal Raw As String) As String
    If 0 < VBA.InStr(1, Raw, COLUMN_PLAN) Then
        ToModel = "Plan"
    ElseIf 0 < VBA.InStr(1, Raw, COLUMN_EXECUTION) Then
        ToModel = "Exe"
    End If
End Function
