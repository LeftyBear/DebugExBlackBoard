VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_RemarksKeywordMatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function MatchAll(ByVal Remarks As String, ByVal Keywords As VBA.Collection) As Boolean
    Dim Keyword As Variant
    For Each Keyword In Keywords
        If VBA.InStr(1, Remarks, CStr(Keyword)) = 0 Then MatchAll = False: Exit Function
    Next
    MatchAll = True
End Function
