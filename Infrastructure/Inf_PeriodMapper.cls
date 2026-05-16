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

Public Function ToRaw(ByVal Period As Dom_Period) As String
    ToRaw = CStr(Period.Value) & jpnKouji
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_Period
    Dim Period As Long
    If 0 < VBA.InStr(1, Raw, jpnKouji) Then
        Dim Buffer As String
        Buffer = VBA.Replace(Raw, jpnKouji, vbNullString)
        Period = Inf_NumericUtility.ExtractNumber(Buffer)
    End If
    Dim Factory As Dom_ValueObjectFactory
    Set Factory = New Dom_ValueObjectFactory
    Set ToDomain = Factory.CreatePeriod(Period)
End Function
