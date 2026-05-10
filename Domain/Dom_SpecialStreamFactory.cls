VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SpecialStreamFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal Row As App_SchoolEventRow) As Dom_SpecialStream
    Dim Entity As Dom_SpecialStream
    Set Entity = New Dom_SpecialStream
    Entity.Initialize Row.Name
    Set Create = Entity
End Function
