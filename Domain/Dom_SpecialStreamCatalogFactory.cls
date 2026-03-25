VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SpecialStreamCatalogFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Public Function Create(ByRef RawCSVRows() As String) As Dom_SpecialStreamCatalog
    Dim Catalog As Dom_SpecialStreamCatalog
    Set Catalog = New Dom_SpecialStreamCatalog
    Dim R As Long
    For R = LBound(RawCSVRows) To UBound(RawCSVRows)
        With New Dom_SpecialStream
            Dim SpecialStream As Dom_SpecialStream
            Set SpecialStream = .Create(R, RawCSVRows(R, 1))
        End With
        Catalog.Add SpecialStream
    Next
    Set Create = Catalog
End Function
