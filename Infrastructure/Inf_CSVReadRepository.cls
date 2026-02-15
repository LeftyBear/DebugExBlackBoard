VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_CSVReadRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Repository"
Option Explicit
Implements Inf_IReadRepository
Private Type Member
    FilePath As String
End Type

Private This As Member

Public Sub Initialize(ByVal FilePath As String)
    This.FilePath = FilePath
End Sub

Private Function ReadAllLines(ByVal FilePath As String) As String()
    Dim Result() As String
    '@Ignore IntegerDataType
    Dim FileNo As Integer
    FileNo = VBA.FreeFile
    Open FilePath For Input As #FileNo
    Do Until VBA.EOF(FileNo)
        Dim Line As String
        Line Input #FileNo, Line
        Dim i As Long
        ReDim Preserve Result(i)
        Result(i) = Line
        i = i + 1
    Loop
    Close #FileNo
    ReadAllLines = Result
End Function

Private Function SplitCSV(ByVal Line As String) As String()
    Dim Result() As String
    Dim i As Long
    Dim Position As Long
    For Position = 1 To VBA.Len(Line)
        Dim Char As String
        Char = VBA.Mid$(Line, Position, 1)
        Select Case Char
        Case Util_Character.DoubleQuotation
            Dim InQuotes As Boolean
            InQuotes = Not InQuotes
        Case Util_Character.Comma
            If InQuotes Then
                Dim Buffer As String
                Buffer = Buffer & Char
            Else
                ReDim Preserve Result(i)
                Result(i) = Buffer
                i = i + 1
                Buffer = vbNullString
            End If
        Case Else
            Buffer = Buffer & Char
        End Select
    Next
    ReDim Preserve Result(i)
    Result(i) = Buffer
    SplitCSV = Result
End Function

Private Function Inf_IReadRepository_ReadAll(ByVal FilePath As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Lines() As String
    Lines = ReadAllLines(FilePath)
    Dim Headers() As String
    Headers = SplitCSV(Lines(LBound(Lines)))
    Dim R As Long
    For R = LBound(Lines) + 1 To UBound(Lines)
        Dim Values() As String
        Values = SplitCSV(Lines(R))
        If UBound(Headers) <> UBound(Values) Then
            Err.Raise Inf_ErrNum.MismatchColumnCount, , "Mismatch column count of header and value"
        End If
        Dim RawRow As Scripting.Dictionary
        Set RawRow = New Scripting.Dictionary
        Dim C As Long
        For C = LBound(Headers) To UBound(Headers)
            RawRow.Add Headers(C), Values(C)
        Next
        Result.Add RawRow
    Next
    Set Inf_IReadRepository_ReadAll = Result
End Function

