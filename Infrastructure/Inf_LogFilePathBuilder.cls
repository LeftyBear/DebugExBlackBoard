VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_LogFilePathBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Builder")
Option Explicit

Private Type Member
    Provider As Inf_WorkbookPathProvider
End Type

Private This As Member

Public Sub Initialize(ByVal Provider As Inf_WorkbookPathProvider)
    Set This.Provider = Provider
End Sub

Public Function Builed() As String
    Builed = This.Provider.GetBaseFolderPath & "root\data\errorlog\error.log"
End Function
