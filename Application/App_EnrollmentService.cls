VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function GetDailyTotal(ByVal Yearly As Dom_EnrollmentYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Calculator As Dom_EnrollmentCalculator, ByVal Filter As Dom_EnrollmentFilter) As Long
    Dim Daily As Dom_EnrollmentAggregate
    Set Daily = Yearly.GetByDate(TargetDate)
    Select Case Filter.StreamType.Value
    Case Dom_StreamTypeEnum.Main
        Dim Items() As Variant
        Items = Daily.GetMainStreams
    Case Dom_StreamTypeEnum.Special
        Items = Daily.GetSpecialStreams
    End Select
    GetDailyTotal = Calculator.Calculate(Items, Filter)
End Function
