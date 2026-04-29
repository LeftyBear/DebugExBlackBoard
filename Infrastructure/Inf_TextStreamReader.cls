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

Private Type Member
    Stream As ADODB.Stream
End Type

Private This As Member

Public Sub Initialize(ByVal Stream As ADODB.Stream)
    Set This.Stream = Stream
End Sub

Public Function ReadAll(ByVal FilePath As String) As String
    With This.Stream
        .Type = adTypeText
        .Charset = "Shift_JIS"
        .Open
        .LoadFromFile FilePath
        ReadAll = .ReadText(adReadAll)
        .Close
    End With
    Set This.Stream = Nothing
End Function
