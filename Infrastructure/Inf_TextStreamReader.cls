VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_TextStreamReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Persistence")
Option Explicit

Public Function ReadAll(ByVal FilePath As String) As String
    Dim Stream As Object
    Set Stream = CreateObject("ADODB.Stream")
    With Stream
        .Type = 2 ':=adTypeText
        .Charset = "Shift_JIS"
        .Open
        .LoadFromFile FilePath
        ReadAll = .ReadText(-1) ':=adReadAll
        .Close
    End With
    Set Stream = Nothing
End Function
