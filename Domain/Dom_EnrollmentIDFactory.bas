Attribute VB_Name = "Dom_EnrollmentIDFactory"
'@Folder("Domain.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal ClassName As Dom_ClassName, ByVal Gender As Dom_Gender) As Dom_EnrollmentID
    Dim ID As Dom_EnrollmentID
    Set ID = New Dom_EnrollmentID
    ID.Initialize StreamType, Grade, ClassNo, ClassName, Gender
    Set Create = ID
End Function
