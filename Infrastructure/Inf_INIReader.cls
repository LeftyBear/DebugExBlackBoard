VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_INIReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Loader"
Option Explicit

Public Function ReadAll(ByVal FilePath As String) As String()
    Dim Result() As String
    '@Ignore IntegerDataType
    Dim FileNo As Integer
    FileNo = VBA.FreeFile
    Open FilePath For Input As #FileNo
    Do Until VBA.EOF(FileNo)
        Dim Buffer As String
        Line Input #FileNo, Buffer
        Dim Index As Long
        ReDim Preserve Result(Index)
        Result(Index) = Buffer
        Index = Index + 1
    Loop
    ReadAll = Result
End Function
