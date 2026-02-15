VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollKeyFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.Factory"
Option Explicit

Public Function Create(ByVal RawKey As String, ByVal RawValue As Variant) As Dom_EnrollKeyValue
    Dim KeyValue As Dom_EnrollKeyValue
    Set KeyValue = New Dom_EnrollKeyValue
    KeyValue.Initialize RawKey, RawValue
    Set Create = KeyValue
End Function
