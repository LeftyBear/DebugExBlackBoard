VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_KeyFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.Factory"
Option Explicit

Public Function Create(ByVal RawKey As String, ByVal RawValue As Variant) As Dom_KeyValue
    Dim KeyValue As Dom_KeyValue
    Set KeyValue = New Dom_KeyValue
    KeyValue.Initialize RawKey, RawValue
    Set Create = KeyValue
End Function
