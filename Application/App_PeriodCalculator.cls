VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_PeriodCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Calculator")
Option Explicit

Public Function CalculateTotal(ByVal Models As App_ClassHourReadModels, ByVal TypeCode As String, ByVal Grade As Long) As Long
    Dim i As Long
    For i = 1 To Models.Count
        If IsTarget(Models.Item(i), TypeCode, Grade) Then
            Dim Result As Long
            Result = Result + 1
        End If
    Next
    CalculateTotal = Result
End Function

Private Function IsTarget(ByVal Model As App_ClassHourReadModel, ByVal TypeCode As String, ByVal Grade As Long) As Boolean
    If Model.TypeCode <> TypeCode Then Exit Function
    If Model.Period = 0 Then Exit Function
    If Model.Grade <> Grade Then Exit Function
    If Model.Mark = vbNullString Then Exit Function
    IsTarget = True
End Function
