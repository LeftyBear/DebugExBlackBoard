VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_TextStreamReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit

Public Function ReadAll(ByVal FilePath As String) As String
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    With Stream
        .Type = adTypeText
        .Charset = "Shift_JIS"
        .Open
        .LoadFromFile FilePath
        ReadAll = .ReadText(adReadAll)
        .Close
    End With
    Set Stream = Nothing
End Function
