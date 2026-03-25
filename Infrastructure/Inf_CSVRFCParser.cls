VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_CSVRFCParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Parser")
Option Explicit

Private Enum ParserState
    OutsideField
    InsideField
    InsideQuotedField
End Enum

Public Function Parse(ByVal RawText As String) As String()
    Dim RawRows As VBA.Collection
    Set RawRows = New VBA.Collection
    Dim CurrentRow As VBA.Collection
    Set CurrentRow = New VBA.Collection
    Dim State As ParserState
    State = OutsideField
    Dim Buffer As String
    Buffer = vbNullString
    Dim i As Long
    For i = 1 To VBA.Len(RawText)
        Dim Char As String
        Char = VBA.Mid$(RawText, i, 1)
        If i < VBA.Len(RawText) Then
            Dim NextChar As String
            NextChar = VBA.Mid$(RawText, i + 1, 1)
        Else
            NextChar = vbNullString
        End If
        Select Case State
        Case OutsideField
            If Char = charDoubleQuotation Then
                State = InsideQuotedField
            ElseIf Char = charComma Then
                CurrentRow.Add Buffer
                Buffer = vbNullString
            ElseIf Char = vbCr Then
                If NextChar = vbLf Then
                    CurrentRow.Add Buffer
                    RawRows.Add CurrentRow
                    Set CurrentRow = New VBA.Collection
                    Buffer = vbNullString
                    i = i + 1
                End If
            Else
                Buffer = Buffer & Char
                State = InsideField
            End If
        Case InsideField
            If Char = charComma Then
                CurrentRow.Add Buffer
                Buffer = vbNullString
                State = OutsideField
            ElseIf Char = vbCr Then
                If NextChar = vbLf Then
                    CurrentRow.Add Buffer
                    RawRows.Add CurrentRow
                    Set CurrentRow = New VBA.Collection
                    Buffer = vbNullString
                    State = OutsideField
                    i = i + 1
                End If
            Else
                Buffer = Buffer & Char
            End If
        Case InsideQuotedField
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
        End Select
    Next
    If VBA.Trim$(Buffer) <> vbNullString Then
        CurrentRow.Add Buffer
        RawRows.Add CurrentRow
    End If
    Parse = ToArray2D(RawRows)
End Function

Private Function ToArray2D(ByVal RawRows As VBA.Collection) As String()
    Dim Result() As String
    Dim RowCount As Long
    RowCount = RawRows.Count
    If RowCount = 0 Then
        ToArray2D = Result
        Exit Function
    End If
    Dim CurrentRow As VBA.Collection
    Set CurrentRow = RawRows.Item(1)
    Dim ColumnCount As Long
    ColumnCount = CurrentRow.Count
    ReDim Result(1 To RowCount, 1 To ColumnCount)
    Dim R As Long
    For R = 1 To RowCount
        Set CurrentRow = RawRows.Item(R)
        Dim C As Long
        For C = 1 To ColumnCount
            Result(R, C) = CurrentRow.Item(C)
        Next
    Next
    ToArray2D = Result
End Function
