VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit
Implements Dom_IEnrollmentRepository
Private Type Member
    Calculator  As Dom_SchoolYearCalculator
    Resolver    As Inf_FilePathResolver
End Type
Private This As Member

Public Sub Initialize(ByVal Calculator As Dom_SchoolYearCalculator, ByVal Resolver As Inf_FilePathResolver)
    Set This.Calculator = Calculator
    Set This.Resolver = Resolver
End Sub

Private Function Dom_IEnrollmentRepository_LoadLines(ByVal TargetDate As Date) As String()
    Dim SchoolYear As Long
    SchoolYear = This.Calculator.Calculate(TargetDate)
    Dim FilePath As String
    This.Resolver.Initialize New Inf_WorkbookPathProvider
    FilePath = This.Resolver.Resolve(EnrollmentFile, SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Err.Raise Inf_ErrNum.NotFoundFile, , "File not found: " & FilePath
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    Stream.Type = adTypeText
    Stream.Charset = "Shift_JIS"
    Stream.Open
    Stream.LoadFromFile FilePath
    Dim TextData As String
    TextData = Stream.ReadText(ADODB.StreamReadEnum.adReadAll)
    Stream.Close
    Set Stream = Nothing
    Dim Lines() As String
    If VBA.Len(TextData) = 0 Then
        ReDim Lines(0): Lines(0) = vbNullString
        Dom_IEnrollmentRepository_LoadLines = Lines
        Exit Function
    End If
    TextData = VBA.Replace(TextData, vbCrLf, vbLf)
    TextData = VBA.Replace(TextData, vbCr, vbLf)
    Lines = VBA.Split(TextData, vbLf)
    If 0 <= UBound(Lines) Then
        If VBA.Len(Lines(UBound(Lines))) = 0 Then ReDim Preserve Lines(UBound(Lines) - 1)
    End If
    Dom_IEnrollmentRepository_LoadLines = Lines
End Function
