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

Public Function Builed() As String
    Dim Provider As Inf_WorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Builed = Inf_StringUtility.BuildToPath(Provider.GetBaseFolderPath, "root", "data", "errorlog", "error.log")
End Function
