VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DebugLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Persistence")
Option Explicit
Implements App_ILogPersistence

Private Sub App_ILogPersistence_Log(ByVal Message As String)
    Debug.Print Now & charHalfSpace & Message
End Sub
