VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_CSVPersistence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Persistence")
Option Explicit

Private Enum ParserState
    OutsideField
    InsideField
    InsideQuotedField
End Enum

Public Function ReadAll(ByVal FilePath As String) As Inf_RawRows
    Dim Reader As Inf_TextStreamReader
    Set Reader = New Inf_TextStreamReader
    Dim TextData As String
    TextData = Reader.ReadAll(FilePath)
    Set ReadAll = ParsedCSV(TextData)
End Function

Public Sub WriteAll(ByVal FilePath As String, ByVal Rows As VBA.Collection)
    Dim TempPath As String
    TempPath = FilePath & ".tmp"
    Dim Lines As VBA.Collection
    Set Lines = BuildLine(Rows)
    WriteAllLines TempPath, Lines
    If 0 < VBA.Len(VBA.Dir(FilePath)) Then VBA.Kill FilePath
    'TODO
    Name TempPath As FilePath
End Sub

Public Sub Delete(ByVal FilePath As String, ByVal DateKey As String)
    Dim Rows As VBA.Collection
    Set Rows = Me.ReadAll(FilePath)
    Dim i As Long
    For i = Rows.Count To 1 Step -1
        If ExtractKey(Rows.Item(i)) = DateKey Then
            Rows.Remove i
            Exit For
        End If
    Next
    Me.WriteAll FilePath, Rows
End Sub

Private Function BuildLine(ByVal Rows As VBA.Collection) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Line As String
        Line = EscapedCSV(Rows.Item(i))
        Result.Add Line
    Next
    Set BuildLine = Result
End Function

Private Function EscapedCSV(ByVal Part As String) As String
    If 0 < VBA.InStr(1, Part, charDoubleQuotation) Then
        EscapedCSV = VBA.Replace(Part, charDoubleQuotation, charEscapeDoubleQuotation)
    End If
    If 0 < VBA.InStr(1, Part, charComma) Or 0 < VBA.InStr(1, Part, charDoubleQuotation) Or 0 < VBA.InStr(1, Part, vbCrLf) Then
        EscapedCSV = charDoubleQuotation & Part & charDoubleQuotation
    End If
End Function

Private Function ExtractKey(ByVal Line As String) As Variant
    Const DateColumns = 1
    Dim Parts() As String
    Parts = VBA.Split(Line, charComma)
    ExtractKey = Parts(DateColumns)
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

Private Function ParsedCSV(ByVal TextData As String) As Inf_RawRows
    Dim RawRows As Inf_RawRows
    Set RawRows = New Inf_RawRows
    Dim State As ParserState
    State = OutsideField
    Dim Buffer As String
    Buffer = vbNullString
    Dim i As Long
    For i = 1 To VBA.Len(TextData)
        Dim Char As String
        Char = VBA.Mid$(TextData, i, 1)
        If i < VBA.Len(TextData) Then
            Dim NextChar As String
            NextChar = VBA.Mid$(TextData, i + 1, 1)
        Else
            NextChar = vbNullString
        End If
        Select Case State
        Case OutsideField:      TryAddRecord RawRows, State, Buffer, i, Char, NextChar
        Case InsideField:       AddRecord RawRows, State, Buffer, i, Char, NextChar
        Case InsideQuotedField: EscapeDoubleQuotation State, Buffer, i, Char, NextChar
        End Select
    Next
    If VBA.Trim$(Buffer) <> vbNullString Then
        RawRows.AddFiled Buffer
        RawRows.AddRecord RawRows.Fields
    End If
    Set ParsedCSV = RawRows
End Function

Private Sub TryAddRecord(ByRef RawRows As Inf_RawRows, ByRef State As ParserState, ByRef Buffer As String, ByRef i As Long, ByRef Char As String, ByRef NextChar As String)
    If Char = charDoubleQuotation Then
        State = InsideQuotedField
    ElseIf Char = charComma Then
        RawRows.AddFiled Buffer
        Buffer = vbNullString
    ElseIf Char = vbCr Then
        If NextChar = vbLf Then
            With RawRows
                .AddFiled Buffer
                .AddRecord .Fields
                .InitializeFileds
            End With
            Buffer = vbNullString
            i = i + 1
        End If
    Else
        Buffer = Buffer & Char
        State = InsideField
    End If
End Sub

Private Sub AddRecord(ByRef RawRows As Inf_RawRows, ByRef State As ParserState, ByRef Buffer As String, ByRef i As Long, ByRef Char As String, ByRef NextChar As String)
    If Char = charComma Then
        RawRows.AddFiled Buffer
        Buffer = vbNullString
        State = OutsideField
    ElseIf Char = vbCr Then
        If NextChar = vbLf Then
            With RawRows
                .AddFiled Buffer
                .AddRecord .Fields
                .InitializeFileds
            End With
            Buffer = vbNullString
            State = OutsideField
            i = i + 1
        End If
    Else
        Buffer = Buffer & Char
    End If
End Sub

Private Sub EscapeDoubleQuotation(ByRef State As ParserState, ByRef Buffer As String, ByRef i As Long, ByRef Char As String, ByRef NextChar As String)
    If Char = charDoubleQuotation Then
        If NextChar = charDoubleQuotation Then
            Buffer = Buffer & charDoubleQuotation
            i = i + 1
        Else
            State = OutsideField
        End If
    Else
        Buffer = Buffer & Char
    End If
End Sub
