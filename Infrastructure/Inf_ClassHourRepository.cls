VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Repository")
Option Explicit
Implements Dom_IClassHourRepository
Private Type Member
    AggregateList As Scripting.Dictionary
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.AggregateList = New Scripting.Dictionary
End Sub

Private Function BuiledFilePath(ByVal SchoolYear As Long) As String
    Dim Provider As Inf_WorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim BaseFolderPath As String
    BaseFolderPath = Provider.GetBaseFolderPath
    Dim FileName As String
    FileName = "ClassHour" & CStr(SchoolYear) & ".csv"
    BuiledFilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "data", "schoolyears", CStr(SchoolYear), FileName)
End Function

Private Sub Dom_IClassHourRepository_Add(ByVal Aggregate As Dom_DailyClassHour)
    This.AggregateList.Add Aggregate.DateID.ToKey, Aggregate
End Sub

Private Function Dom_IClassHourRepository_FindByDate(ByVal TargetDate As Date) As Dom_DailyClassHour
    If This.AggregateList.Exists(CStr(TargetDate)) Then
        Set Dom_IClassHourRepository_FindByDate = This.AggregateList.Aggregate(CStr(TargetDate))
    End If
End Function

Private Function Dom_IClassHourRepository_GetAll() As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Aggregate As Dom_DailyClassHour
    For Each Aggregate In This.AggregateList.Items
        Result.Add Aggregate
    Next
    Set Dom_IClassHourRepository_GetAll = Result
End Function

Private Function Dom_IClassHourRepository_ReadAll(ByVal SchoolYear As Long) As VBA.Collection
    Dim FilePath As String
    FilePath = BuiledFilePath(SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Exit Function
    Dim Persistence As Inf_CSVPersistence
    Set Persistence = New Inf_CSVPersistence
    Set Dom_IClassHourRepository_ReadAll = Persistence.ReadAll(FilePath)
End Function

Private Function Dom_IClassHourRepository_Replace(ByVal NewAggregate As Dom_DailyClassHour) As VBA.Collection
    Dim Rows As VBA.Collection
    Set Rows = Dom_IClassHourRepository_GetAll
    Dim i As Long
    For i = 1 To Rows.Count
        Dim OldAggregate As Dom_DailyClassHour
        Set OldAggregate = Rows.Item(i)
        If OldAggregate.DateID.Equals(NewAggregate.DateID) Then
            Set Rows.Item(i) = NewAggregate
        End If
    Next
    Set Dom_IClassHourRepository_Replace = Rows
End Function
