VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SchoolEventCatalogFactory"
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

Public Function Create(ByRef RawRows() As String) As Dom_SchoolEventCatalog
    Dim Result As Dom_SchoolEventCatalog
    Set Result = New Dom_SchoolEventCatalog
    Dim R As Long
    For R = LBound(RawRows) + 1 To UBound(RawRows)
        Dim C As Long
        For C = LBound(RawRows, 2) To UBound(RawRows, 2)
            If RawRows(1, C) = jpnSukejuruMei Then
                Dim SchoolEvent As Dom_SchoolEvent
                Set SchoolEvent = This.Factory.CreateSchoolEvent(R, RawRows(R, C))
                Result.Add SchoolEvent
            End If
        Next
    Next
    Set Create = Result
End Function
