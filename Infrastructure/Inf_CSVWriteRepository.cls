VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_CSVWriteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Repository"
Option Explicit
Implements Inf_IWriteRepository
Private Type Member
    FilePath As String
End Type

Private This As Member

Public Sub Initialize(ByVal FilePath As String)
    This.FilePath = FilePath
End Sub

Private Function ReadAllLines(ByVal FilePath As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    '@Ignore IntegerDataType
    Dim FileNo As Integer
    FileNo = VBA.FreeFile
    Open FilePath For Input As #FileNo
    Do Until VBA.EOF(FileNo)
        Dim Line As String
        Line Input #FileNo, Line
        Result.Add Line
    Loop
    Close #FileNo
    Set ReadAllLines = Result
End Function

Private Sub WriteAllLines(ByVal FilePath As String, ByVal Lines As VBA.Collection)
    '@Ignore IntegerDataType
    Dim FileNo As Integer
    FileNo = VBA.FreeFile
    Open FilePath For Output As #FileNo
    Dim i As Long
    For i = 1 To Lines.Count
        Print #FileNo, Lines.Item(i)
    Next
    Close #FileNo
End Sub

Private Function BuildCSVLine(ByVal Row As Scripting.Dictionary) As String
    Dim Parts() As String
    ReDim Parts(0 To Row.Count - 1)
    Dim i As Long
    For i = 0 To Row.Count - 1
        Parts(i) = EscapeCSV(Row.Item(i))
    Next
    BuildCSVLine = VBA.Join(Parts, Util_Character.Comma)
End Function

Private Function EscapeCSV(ByVal Text As String) As String
    Dim Result As String
    If 0 < VBA.InStr(1, Text, Util_Character.DoubleQuotation) Then
        Result = VBA.Replace(Text, Util_Character.DoubleQuotation, """""")
    End If
    If 0 < VBA.InStr(1, Text, Util_Character.Comma) Or 0 < VBA.InStr(1, Text, Util_Character.DoubleQuotation) Or 0 < VBA.InStr(1, Text, vbCrLf) Then
        Result = Util_Character.DoubleQuotation & Text & Util_Character.DoubleQuotation
    End If
    EscapeCSV = Result
End Function

Private Function ExtractKey(ByVal Line As String) As Variant
    Dim Values As Variant
    Values = Split(Line, ",")
    ExtractKey = Values(0)
End Function

Private Sub Inf_IWriteRepository_Delete(ByVal Key As Variant)
    Dim Lines As VBA.Collection
    Set Lines = ReadAllLines(This.FilePath)
    Dim i As Long
    For i = Lines.Count To 1 Step -1
        If ExtractKey(Lines.Item(i)) Like Key Then
            Lines.Remove i
            Exit For
        End If
    Next
    WriteAllLines This.FilePath, Lines
End Sub

Private Sub Inf_IWriteRepository_WriteAll(ByVal RawRows As VBA.Collection)
    If RawRows.Count = 0 Then Exit Sub
    Dim Lines As VBA.Collection
    Set Lines = New VBA.Collection
    Dim FirstRow As Scripting.Dictionary
    Set FirstRow = RawRows.Item(1)
    Dim HeaderLine As String
    HeaderLine = BuildCSVLine(FirstRow)
    Lines.Add HeaderLine
    Dim Row As Scripting.Dictionary
    For Each Row In RawRows
        Dim Line As String
        Line = BuildCSVLine(Row)
        Lines.Add Line
    Next
    WriteAllLines This.FilePath, Lines
End Sub

