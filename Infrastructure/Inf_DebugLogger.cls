VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DebugLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Utility.Logger"
Option Explicit
Implements Inf_ILogger

Private Sub Inf_ILogger_WriteLog(ByVal Message As String)
    Debug.Print Now & charHalfSpace & Message
End Sub
