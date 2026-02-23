VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolYearCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Service")
Option Explicit

Public Function Calculate(ByVal TargetDate As Date) As Long
    If VBA.Month(TargetDate) < 4 Then
        Calculate = VBA.Year(TargetDate) - 1
    Else
        Calculate = VBA.Year(TargetDate)
    End If
End Function
