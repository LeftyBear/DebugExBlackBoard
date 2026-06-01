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

Public Function CalculateTotal(ByVal Models As App_ClassHourReadModels, ByVal TypeCode As String, ByVal Grade As Long) As Long
    Dim i As Long
    For i = 1 To Models.Count
        If Models.Item(i).TypeCode = TypeCode Then
            If Models.Item(i).Grade = Grade Then
                Dim Result As Long
                Result = Result + Models.Item(i).Value
            End If
        End If
    Next
    CalculateTotal = Result
End Function
