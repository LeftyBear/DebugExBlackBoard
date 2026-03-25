VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_UpperValueFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Public Function Create(ByRef RawCSVRows() As String) As Dom_UpperValue
    Dim R As Long
    For R = LBound(RawCSVRows) + 1 To UBound(RawCSVRows)
        Dim C As Long
        For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
            Select Case RawCSVRows(1, C)
            Case jpnFutsuGakunenSu
                Dim MainStreamGrade As Long
                MainStreamGrade = Dom_Util_Long.NormalizeToLong(RawCSVRows(R, C))
            Case jpnFutsuGakkyuSu
                Dim MainStreamClass As Long
                MainStreamClass = Dom_Util_Long.NormalizeToLong(RawCSVRows(R, C))
            Case jpnJikanwariSu
                Dim Period As Long
                Period = Dom_Util_Long.NormalizeToLong(RawCSVRows(R, C))
            End Select
        Next
    Next
    With New Dom_UpperValue
        Set Create = .Create(MainStreamGrade, MainStreamClass, Period)
    End With
End Function
