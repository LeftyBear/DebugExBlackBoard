VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_DateIndexCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Service")
Option Explicit

Private Type Member
    SchoolYearCalculator As Dom_SchoolYearCalculator
End Type
Private This As Member

Private Sub Class_Initialize()
    Set This.SchoolYearCalculator = New Dom_SchoolYearCalculator
End Sub

Public Function Calculate(ByVal TargetDate As Date) As Long
    Dim SchoolYear As Long
    SchoolYear = This.SchoolYearCalculator.Calculate(TargetDate)
    Calculate = TargetDate - CDate(VBA.DateSerial(CInt(SchoolYear), 4, 1)) + 1
End Function
