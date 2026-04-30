VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit
Implements App_ILogger

Private Sub App_ILogger_WriteLog(ByVal Message As String)
    Dim Provider As Inf_WorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim FilePath As String
    FilePath = Provider.GetBaseFolderPath
    '@Ignore IntegerDataType
    Dim FileNo As Integer
    FileNo = VBA.FreeFile
    Open FilePath For Append As #FileNo
    Print #FileNo, Now & charHalfSpace & Message
    Close #FileNo
End Sub
