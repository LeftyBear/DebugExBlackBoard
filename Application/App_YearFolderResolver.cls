VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_YearFolderResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Application.Resolver

Option Explicit

Public Function Resolve(ByVal BasePath As String, ByVal SchoolYearKey As App_SchoolYearKey) As String
    Resolve = BasePath & "\" & SchoolYearKey.Value
End Function
