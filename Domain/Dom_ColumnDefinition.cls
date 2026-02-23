VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ColumnDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Metadata")
Option Explicit

Private Type Member
    StreamType  As Dom_StreamType
    Required    As Boolean
End Type

Private This As Member

Public Sub Initialize(ByVal StreamType As Dom_StreamType, ByVal Required As Boolean)
    This.StreamType = StreamType
    This.Required = Required
End Sub

Public Property Get StreamType() As Dom_StreamType
    StreamType = This.StreamType
End Property

Public Property Get Required() As Boolean
    Required = This.Required
End Property
