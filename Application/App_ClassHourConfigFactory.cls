VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourConfigFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Factory"
Option Explicit

Public Function Create(ByVal INISettings As Scripting.Dictionary) As App_ClassHourConfig
    Dim KeyValues As VBA.Collection
    Set KeyValues = New VBA.Collection
    Dim Factory As App_KeyFactory
    Set Factory = New App_KeyFactory
    Dim SectionKey As Variant
    For Each SectionKey In INISettings.Keys
        If SectionKey Like "Common" Or SectionKey Like "ClassHour*" Then
            Dim SectionItems As Scripting.Dictionary
            Set SectionItems = INISettings.Item(SectionKey)
            Dim RawKey As Variant
            For Each RawKey In SectionItems.Keys
                Dim KeyValue As App_KeyValue
                Set KeyValue = Factory.Create(RawKey, SectionItems.Item(RawKey))
                KeyValues.Add KeyValue
            Next
        End If
    Next
    Dim Config As App_ClassHourConfig
    Set Config = New App_ClassHourConfig
    Config.Initialize KeyValues
    Set Create = Config
End Function

