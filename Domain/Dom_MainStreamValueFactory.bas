Attribute VB_Name = "Dom_MainStreamValueFactory"
'@Folder("Domain.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal UpperGrade As Long, ByVal UpperClassNo As Long) As Dom_MainStreamValue
    Dim Value As Dom_MainStreamValue
    Set Value = New Dom_MainStreamValue
    Value.Initialize UpperGrade, UpperClassNo
    Set Create = Value
End Function
