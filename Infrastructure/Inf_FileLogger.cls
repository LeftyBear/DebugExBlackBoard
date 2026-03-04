VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Logger")
Option Explicit
Implements Inf_ILogger
Private Type Member
    FilePath As String
End Type

Private This As Member

Public Sub Initialize(ByVal FilePath As String)
    This.FilePath = FilePath
End Sub

Private Sub Inf_ILogger_WriteLog(ByVal Message As String)
    '@Ignore IntegerDataType
    Dim FileNo As Integer
    FileNo = VBA.FreeFile
    Open This.FilePath For Append As #FileNo
    Print #FileNo, Now & charHalfSpace & Message
    Close #FileNo
End Sub

