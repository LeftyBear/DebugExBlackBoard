VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourColumnResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Resolver")
Option Explicit

Private Type Member
    Structure As Dom_SchoolStructure
End Type

Private This As Member

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
End Sub

Public Function Resolve(ByRef HeaderRow() As String) As Scripting.Dictionary
    Dim Map As Scripting.Dictionary
    Set Map = New Scripting.Dictionary
    Dim C As Long
    For C = LBound(HeaderRow) To UBound(HeaderRow)
        Dim ColumnName As String
        ColumnName = HeaderRow(C)
        If 0 < VBA.InStr(1, ColumnName, charUnderScore) Then
            Dim Meta As App_ClassHourColumnMeta
            Meta = ParseColumnMeta(ColumnName)
            Map.Add C, Meta
        End If
    Next
    Set Resolve = Map
End Function

Private Function ParseColumnMeta(ByVal ColumnName As String) As App_ClassHourColumnMeta
    Dim Result As App_ClassHourColumnMeta
    Set Result = New App_ClassHourColumnMeta
    Dim Parts() As String
    Parts = VBA.Split(ColumnName, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnKeikaku) Then
            Result.TypeCode = Dom_ClassHourTypeEnum.Plan
        ElseIf 0 < VBA.InStr(1, Part, jpnJisshi) Then
            Result.TypeCode = Dom_ClassHourTypeEnum.Execution
        End If
        If This.Structure.IsSubject(Part) Then
            Result.Subject = Part
        End If
        If 0 < VBA.InStr(1, Part, jpnNen) Then
            Dim GradeString As String
            GradeString = VBA.Replace(Part, jpnNen, vbNullString)
            Result.Grade = App_NumericUtility.ExtractNumber(GradeString)
        End If
    Next
    ParseColumnMeta = Result
End Function
