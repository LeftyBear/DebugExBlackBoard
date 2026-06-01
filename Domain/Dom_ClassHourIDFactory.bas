Attribute VB_Name = "Dom_ClassHourIDFactory"
'@Folder("Domain.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal ClassHourType As Dom_ClassHourType, ByVal SubjectName As Dom_SubjectName, ByVal Grade As Dom_Grade) As Dom_ClassHourID
    Dim ID As Dom_ClassHourID
    Set ID = New Dom_ClassHourID
    ID.Initialize ClassHourType, SubjectName, Grade
    Set Create = ID
End Function
