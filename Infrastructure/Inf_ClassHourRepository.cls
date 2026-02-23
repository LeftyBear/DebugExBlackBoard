VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit

Public Function LoadLines(ByVal FilePath As String) As String()
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Err.Raise Inf_ErrNum.NotFoundFile, , "File not found: " & FilePath
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    Stream.Type = adTypeText
    Stream.Charset = "Shift_JIS"
    Stream.Open
    Stream.LoadFromFile FilePath
    Dim TextData As String
    TextData = Stream.ReadText
    Stream.Close
    Set Stream = Nothing
    Dim Lines() As String
    If VBA.Len(TextData) = 0 Then ReDim Lines(0): Lines(0) = vbNullString: LoadLines = Lines: Exit Function
    TextData = VBA.Replace(TextData, vbCrLf, vbLf)
    TextData = VBA.Replace(TextData, vbCr, vbLf)
    Lines = VBA.Split(TextData, vbLf)
    If 0 <= UBound(Lines) Then If VBA.Len(Lines(UBound(Lines))) = 0 Then ReDim Preserve Lines(UBound(Lines) - 1)
    LoadLines = Lines
End Function

