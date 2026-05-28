VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_PeriodMapper"
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

Private Const COLUMN_PERIOD As String = "çZéû"

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function ToRaw(ByVal Period As Dom_Period) As String
    ToRaw = CStr(Period.Value) & COLUMN_PERIOD
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Period
    If 0 < VBA.InStr(1, Raw, COLUMN_PERIOD) Then
        Dim Period As Long
        Period = GetPeriod(Raw)
    Else
        Exit Function
    End If
    Set ToDomain = This.Factory.CreatePeriod(Period)
End Function

Public Function ToModel(ByVal Raw As String) As Long
    ToModel = GetPeriod(Raw)
End Function

Private Function GetPeriod(ByVal Raw As String) As Long
    Dim Buffer As String
    Buffer = VBA.Replace(Raw, COLUMN_PERIOD, vbNullString)
    GetPeriod = Inf_NumericUtility.ExtractNumber(Buffer)
End Function
