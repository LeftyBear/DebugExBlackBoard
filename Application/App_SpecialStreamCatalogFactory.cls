VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SpecialStreamCatalogFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Factory As Dom_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function Create(ByRef RawRows() As String) As Dom_SpecialStreamCatalog
    Dim Catalog As Dom_SpecialStreamCatalog
    Set Catalog = New Dom_SpecialStreamCatalog
    Dim R As Long
    For R = LBound(RawRows) To UBound(RawRows)
        Dim SpecialStream As Dom_SpecialStream
        Set SpecialStream = This.Factory.CreateSpecialStream(R, RawRows(R, 1))
        Catalog.Add SpecialStream
    Next
    Set Create = Catalog
End Function
