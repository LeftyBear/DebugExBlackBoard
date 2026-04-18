VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_DateIDFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal NormDate As Date) As Dom_DateID
    With New Dom_DateID
        Set Create = .Create(NormDate)
    End With
End Function
