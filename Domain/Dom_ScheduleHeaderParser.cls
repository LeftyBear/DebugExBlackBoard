VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleHeaderParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Parser")
Option Explicit

Private Type Member
    Structure As Dom_SchoolStructure
End Type

Private This As Member

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
End Sub

Public Function Parse(ByVal HeaderName As String) As Dom_ScheduleHeader
    If 0 < VBA.InStr(1, HeaderName, jpnHizuke) Then
        Dim RawDate As String
        RawDate = HeaderName
    ElseIf This.Structure.IsSchoolEvent(HeaderName) Then
        Dim RawText As String
        RawText = HeaderName
    End If
    Dim Result As Dom_ScheduleHeader
    Set Result = New Dom_ScheduleHeader
    Result.Initialize RawDate, RawText
    Set Parse = Result
End Function
