VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentColumnResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function Resolve(ByRef HeaderRow() As String) As Scripting.Dictionary
    Dim Map As Scripting.Dictionary
    Set Map = New Scripting.Dictionary
    Dim C As Long
    For C = LBound(HeaderRow) To UBound(HeaderRow)
        Dim ColumnName As String
        ColumnName = HeaderRow(C)
        Dim Meta As App_EnrollmentColumnMeta
        Meta = ParseColumnMeta(ColumnName)
        Map.Add C, Meta
    Next
    Set Resolve = Map
End Function

Private Function ParseColumnMeta(ByVal ColumnName As String) As App_EnrollmentColumnMeta
    Dim Meta As App_EnrollmentColumnMeta
    Set Meta = New App_EnrollmentColumnMeta
    Dim Parts() As String
    Parts = VBA.Split(ColumnName, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnHizuke) Then
            Meta.ColDate = Part
        End If
        If 0 < VBA.InStr(1, Part, jpnTokushi) Then
            Meta.TypeCode = Dom_StreamTypePolicy.Special
            Meta.Grade = 1
        ElseIf 0 < VBA.InStr(1, Part, jpnNen) Then
            Meta.TypeCode = Dom_StreamTypePolicy.Main
            Dim GradeString As String
            GradeString = VBA.Replace(Part, jpnNen, vbNullString)
            Meta.Grade = App_NumericUtility.ExtractNumber(GradeString)
        ElseIf 0 < VBA.InStr(1, Part, jpnKumi) Then
            Dim ClassNoString As String
            ClassNoString = VBA.Replace(Part, jpnKumi, vbNullString)
            Meta.ClassNo = App_NumericUtility.ExtractNumber(ClassNoString)
        ElseIf 0 < VBA.InStr(1, Part, jpnOtoko) Then
            Meta.Gender = Dom_GenderTypePolicy.Male
        ElseIf 0 < VBA.InStr(1, Part, jpnOnna) Then
            Meta.Gender = Dom_GenderTypePolicy.Female
        End If
        If 0 < VBA.InStr(1, Part, jpnIdou) Then
            Meta.ColTransfer = Part
        End If
        If 0 < VBA.InStr(1, Part, jpnBikou) Then
            Meta.ColRemarks = Part
        End If
    Next
    ParseColumnMeta = Meta
End Function
