VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_PeriodMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Const COLUMN_PERIOD As String = "çZéû"

Public Function ToRaw(ByVal Period As Dom_Period) As String
    ToRaw = CStr(Period.Value) & COLUMN_PERIOD
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Period
    If VBA.InStr(1, Raw, COLUMN_PERIOD) = 0 Then Exit Function
    Dim EndPosition As Long
    EndPosition = VBA.InStr(1, Raw, COLUMN_PERIOD)
    Dim BeginPosition As Long
    BeginPosition = EndPosition - 1
    Dim PeriodString As String
    PeriodString = VBA.Mid$(Raw, BeginPosition, EndPosition)
    Dim Period As Long
    Period = Inf_NumericUtility.ExtractNumber(PeriodString)
    Set ToDomain = Dom_ValueObjectFactory.CreatePeriod(Period)
End Function
