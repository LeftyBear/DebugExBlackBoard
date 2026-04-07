VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolEventCatalogFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByRef RawCSVRows() As String) As Dom_SchoolEventCatalog
    Dim Result As Dom_SchoolEventCatalog
    Set Result = New Dom_SchoolEventCatalog
    Dim R As Long
    For R = LBound(RawCSVRows) + 1 To UBound(RawCSVRows)
        Dim C As Long
        For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
            If RawCSVRows(1, C) = jpnSukejuruMei Then
                With New Dom_SchoolEvent
                    Dim SchoolEvent As Dom_SchoolEvent
                    Set SchoolEvent = .Create(R, RawCSVRows(R, C))
                End With
                Result.Add SchoolEvent
            End If
        Next
    Next
    Set Create = Result
End Function
