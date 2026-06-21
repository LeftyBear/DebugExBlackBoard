VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_StreamTypeMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service.Mapper")
Option Explicit

Private Const COLUMN_SPECIAL As String = "“ÁŽx"

Public Function ToRaw(ByVal StreamType As Dom_StreamType) As String
    If StreamType.IsSpecial Then ToRaw = COLUMN_SPECIAL
End Function

Public Function ToDomain(ByVal Raw As String) As Dom_StreamType
    Dim TypeCode As Dom_StreamTypeCode
    If 0 < VBA.InStr(1, Raw, COLUMN_SPECIAL) Then
        TypeCode = SpecialCode
    Else
        TypeCode = MainCode
    End If
    Set ToDomain = Dom_ValueObjectFactory.CreateStreamType(TypeCode)
End Function
