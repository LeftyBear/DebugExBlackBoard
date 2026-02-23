VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_CSVRFCParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Parser")
Option Explicit

Public Function ParseAll(ByRef Lines() As String) As VBA.Collection
    Dim Records As VBA.Collection
    Set Records = New VBA.Collection
    Dim i As Long
    For i = LBound(Lines) To UBound(Lines)
        Dim Buffer As String
        If 0 < VBA.Len(Buffer) Then Buffer = Buffer & vbLf & Lines(i) Else Buffer = Lines(i)
        Dim InQuote As Boolean
        InQuote = HasUnclosedQuote(Buffer)
        If Not InQuote Then Records.Add Me.ParseRecord(Buffer): Buffer = vbNullString
    Next
    If 0 < VBA.Len(Buffer) Then Err.Raise errUnclosedQuotedField, , "Unclosed quoted field"
    Set ParseAll = Records
End Function

Public Function ParseRecord(ByVal Line As String) As String()
    Dim Result() As String
    ReDim Result(0)
    Dim i As Long
    For i = 1 To VBA.Len(Line)
        Dim Char As String
        Char = VBA.Mid$(Line, i, 1)
        If Char = charDoubleQuotation Then
            Dim InQuote As Boolean
            If InQuote And i < VBA.Len(Line) And VBA.Mid$(Line, i + 1, 1) = charDoubleQuotation Then
                Dim Buffer As String
                Buffer = Buffer & charDoubleQuotation
                i = i + 1
            Else
                InQuote = Not InQuote
            End If
        ElseIf Char = charComma And Not InQuote Then
            Dim FieldIndex As Long
            Result(FieldIndex) = Buffer
            Buffer = vbNullString
            FieldIndex = FieldIndex + 1
            ReDim Preserve Result(FieldIndex)
        Else
            Buffer = Buffer & Char
        End If
    Next
    If InQuote Then Err.Raise errUnmatchedQuote, , "Unmatched quote"
    Result(FieldIndex) = Buffer
    ParseRecord = Result
End Function

Private Function HasUnclosedQuote(ByVal Line As String) As Boolean
    Dim i As Long
    For i = 1 To VBA.Len(Line)
        Dim Count As Long
        If VBA.Mid$(Line, i, 1) = charDoubleQuotation Then Count = Count + 1
    Next
    HasUnclosedQuote = (Count Mod 2 <> 0)
End Function
