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

Private Const COMMA             As String = ","
Private Const DOUBLE_QUOTATION  As String = """"

Public Function ReadAll(ByVal FilePath As String) As Inf_RawRows
    Dim Reader As Inf_TextStreamReader
    Set Reader = New Inf_TextStreamReader
    Dim TextData As String
    TextData = Reader.ReadAll(FilePath)
    Set ReadAll = ParseCSV(TextData)
End Function

Public Sub WriteAll(ByVal FilePath As String, ByVal RawRows As Inf_RawRows)
    Dim TempPath As String
    TempPath = FilePath & ".tmp"
    Dim Lines As VBA.Collection
    Set Lines = BuildLine(RawRows)
    WriteAllLines TempPath, Lines
    If 0 < VBA.Len(VBA.Dir(FilePath)) Then VBA.Kill FilePath
    Name TempPath As FilePath
End Sub

Public Sub Delete(ByVal FilePath As String, ByVal DateKey As String)
    Dim Rows As Inf_RawRows
    Set Rows = Me.ReadAll(FilePath)
    Dim i As Long
    For i = Rows.RowsCount To 1 Step -1
        If Rows.GetRow(i, 1) = DateKey Then
            Rows.Remove i
            Exit For
        End If
    Next
    Me.WriteAll FilePath, Rows
End Sub

Private Function BuildLine(ByVal RawRows As Inf_RawRows) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim R As Long
    For R = 1 To RawRows.RowsCount
        Dim Fields() As Variant
        ReDim Fields(RawRows.ColumnsCount(R) - 1)
        Dim C As Long
        For C = 1 To RawRows.ColumnsCount(R)
            Dim Field As String
            Field = EscapeCSV(RawRows.GetRow(R, C))
            Fields(C - 1) = Field
        Next
        Dim Line As String
        Line = VBA.Join(Fields, COMMA)
        Result.Add Line
    Next
    Set BuildLine = Result
End Function

Private Function EscapeCSV(ByVal Part As String) As String
    Const ESCAPE_DOUBLE_QUOTATION As String = """"""
    If 0 < VBA.InStr(1, Part, DOUBLE_QUOTATION) Then
        EscapeCSV = VBA.Replace(Part, DOUBLE_QUOTATION, ESCAPE_DOUBLE_QUOTATION)
    End If
    If 0 < VBA.InStr(1, Part, COMMA) Or 0 < VBA.InStr(1, Part, DOUBLE_QUOTATION) Or 0 < VBA.InStr(1, Part, vbCrLf) Then
        EscapeCSV = DOUBLE_QUOTATION & Part & DOUBLE_QUOTATION
    End If
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

Private Function ParseCSV(ByVal TextData As String) As Inf_RawRows
    Dim Result As Inf_RawRows
    Set Result = New Inf_RawRows
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
        Case OutsideField:      TryAddRecord Result, State, Buffer, i, Char, NextChar
        Case InsideField:       AddRows Result, State, Buffer, i, Char, NextChar
        Case InsideQuotedField: EscapeDoubleQuotation State, Buffer, i, Char, NextChar
        End Select
    Next
    If VBA.Trim$(Buffer) <> vbNullString Then
        Result.AddRow Buffer
        Result.AddRowsInternal
    End If
    Set ParseCSV = Result
End Function

Private Sub TryAddRecord(ByRef RawRows As Inf_RawRows, ByRef State As ParserState, ByRef Buffer As String, ByRef i As Long, ByRef Char As String, ByRef NextChar As String)
    If Char = DOUBLE_QUOTATION Then
        State = InsideQuotedField
    ElseIf Char = COMMA Then
        RawRows.AddRow Buffer
        Buffer = vbNullString
    ElseIf Char = vbCr Then
        If NextChar = vbLf Then
            With RawRows
                .AddRow Buffer
                .AddRowsInternal
            End With
            Buffer = vbNullString
            i = i + 1
        End If
    Else
        Buffer = Buffer & Char
        State = InsideField
    End If
End Sub

Private Sub AddRows(ByRef RawRows As Inf_RawRows, ByRef State As ParserState, ByRef Buffer As String, ByRef i As Long, ByRef Char As String, ByRef NextChar As String)
    If Char = COMMA Then
        RawRows.AddRow Buffer
        Buffer = vbNullString
        State = OutsideField
    ElseIf Char = vbCr Then
        If NextChar = vbLf Then
            With RawRows
                .AddRow Buffer
                .AddRowsInternal
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
    If Char = DOUBLE_QUOTATION Then
        If NextChar = DOUBLE_QUOTATION Then
            Buffer = Buffer & DOUBLE_QUOTATION
            i = i + 1
        Else
            State = OutsideField
        End If
    Else
        Buffer = Buffer & Char
    End If
End Sub
