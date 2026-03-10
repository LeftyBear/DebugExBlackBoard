VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ConfigFileNameResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Resolver")
Option Explicit

Public Function Resolve(ByVal ConfigName As Dom_SchoolStructure) As String
    Select Case ConfigName
    Case Dom_SchoolStructure.LimitValue
        Resolve = "LimitValue.csv"
    Case Dom_SchoolStructure.Subject
        Resolve = "Subject.csv"
    End Select
End Function

