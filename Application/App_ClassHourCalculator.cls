VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Calculator")
Option Explicit

Public Function CalculateTotal(ByVal Models As App_ClassHourReadModels, ByVal Key As String, ByVal Grade As Long) As Long
    Dim i As Long
    For i = 1 To Models.ClassHourCount
        If IsTarget(Models.ClassHourItem(i), Key, Grade) Then
            Dim Result As Long
            Result = Result + Models.ClassHourItem(i).Value
        End If
    Next
    CalculateTotal = Result
End Function

Private Function IsTarget(ByVal Model As App_ClassHourReadModel, ByVal Key As String, ByVal Grade As Long) As Boolean
    If Model.Key <> Key Then Exit Function
    If Model.Subject = vbNullString Then Exit Function
    If Model.Grade <> Grade Then Exit Function
    IsTarget = True
End Function
