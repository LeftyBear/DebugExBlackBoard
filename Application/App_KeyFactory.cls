VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_KeyFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.Factory"
Option Explicit

Public Function Create(ByVal RawKey As String, ByVal RawValue As Variant) As App_KeyValue
    Dim KeyValue As App_KeyValue
    Set KeyValue = New App_KeyValue
    KeyValue.Initialize RawKey, RawValue
    Set Create = KeyValue
End Function

