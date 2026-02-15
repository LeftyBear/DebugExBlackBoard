VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentConfigFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Factory"
Option Explicit

Public Function Create(ByVal INISettings As Scripting.Dictionary) As App_EnrollmentConfig
    Dim KeyValues As VBA.Collection
    Set KeyValues = New VBA.Collection
    Dim Factory As Dom_KeyFactory
    Set Factory = New Dom_KeyFactory
    Dim SectionKey As Variant
    For Each SectionKey In INISettings.Keys
        If SectionKey Like "Common" Or SectionKey Like "Enrollment*" Then
            Dim SectionItems As Scripting.Dictionary
            Set SectionItems = INISettings.Item(SectionKey)
            Dim RawKey As Variant
            For Each RawKey In SectionItems.Keys
                Dim KeyValue As Dom_KeyValue
                Set KeyValue = Factory.Create(RawKey, SectionItems.Item(RawKey))
                KeyValues.Add KeyValue
            Next
        End If
    Next
    Dim Config As App_EnrollmentConfig
    Set Config = New App_EnrollmentConfig
    Config.Initialize KeyValues
    Set Create = Config
End Function

