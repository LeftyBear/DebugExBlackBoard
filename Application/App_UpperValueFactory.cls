VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_UpperValueFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Factory As App_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As App_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function Create(ByRef RawRows() As String) As Dom_UpperValue
    Dim R As Long
    For R = LBound(RawRows) + 1 To UBound(RawRows)
        Dim C As Long
        For C = LBound(RawRows, 2) To UBound(RawRows, 2)
            Select Case RawRows(1, C)
            Case jpnFutsuGakunenSu
                Dim MainStreamGrade As Long
                MainStreamGrade = App_NumericUtility.ExtractNumber(RawRows(R, C))
            Case jpnFutsuGakkyuSu
                Dim MainStreamClass As Long
                MainStreamClass = App_NumericUtility.ExtractNumber(RawRows(R, C))
            Case jpnJikanwariSu
                Dim Period As Long
                Period = App_NumericUtility.ExtractNumber(RawRows(R, C))
            End Select
        Next
    Next
    Set Create = This.Factory.CreateUpperValue(MainStreamGrade, MainStreamClass, Period)
End Function
