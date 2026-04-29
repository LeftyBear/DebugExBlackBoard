VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.DTO")
Option Explicit

Public NormDate     As Date
Public TypeCode     As Long
Public Grade        As Long
Public ClassNo      As Long
Public Gender       As Long
Public Value        As Long
Public RawTransfer  As String
Public RawRemarks   As String
