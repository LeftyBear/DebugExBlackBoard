VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LogFilePathBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Provider As App_IWorkbookPathProvider
End Type

Private This As Member

Public Sub Initialize(ByVal Provider As App_IWorkbookPathProvider)
    Set This.Provider = Provider
End Sub

Public Function BuiledPath() As String
    BuiledPath = This.Provider.GetBaseFolderPath & "root\data\errorlog\error.log"
End Function
