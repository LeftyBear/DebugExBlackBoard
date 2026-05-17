Attribute VB_Name = "Inf_VBComponentUtility"
'@Folder("Infrastructure.Service")
Option Explicit

Private Const CtStdModule   As Long = 1
Private Const CtClassModule As Long = 2
Private Const CtMsForm      As Long = 3
Private Const RootPath      As String = "C:\Users\biz\Documents\GitHub\DebugExBlackBoard\"

Public Sub BuildAddin()
    Dim FilePath As String
    FilePath = ThisWorkbook.Path & charBackSlash & "test_" & VBA.Format$(Now, "yyyymmddhhnnss") & ".xlam"
    Dim AddinWorkbook As Excel.Workbook
    Set AddinWorkbook = CreateAddinWorkbook(FilePath)
    ImportAllComponents AddinWorkbook.VBProject, RootPath
    Application.DisplayAlerts = False
    AddinWorkbook.SaveAs FileName:=FilePath, FileFormat:=xlOpenXMLAddIn
    Application.DisplayAlerts = True
    AddinWorkbook.Close SaveChanges:=False
End Sub

Public Sub ExportAllModules()
    If Inf_Environment.GetEnvironmentTypeCode = ReleaseMode Then Exit Sub
    Dim Component As Object
    For Each Component In ThisWorkbook.VBProject.VBComponents
        If IsExportTarget(Component) Then ExportComponent Component
    Next
    CleanupUnusedComponents RootPath
End Sub

Private Sub RemoveExtraSheets(ByVal Workbook As Excel.Workbook)
    Application.DisplayAlerts = False
    Do While 1 < Workbook.Worksheets.Count
        Workbook.Worksheets(Workbook.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
End Sub

Private Sub ImportAllComponents(ByVal Project As Object, ByVal FolderPath As String)
    TraverseFolders Project, FolderPath
End Sub

Private Sub TraverseFolders(ByVal Project As Object, ByVal FolderPath As String)
    Dim FileName As String
    FileName = VBA.Dir(FolderPath & "*.*")
    Do While FileName <> vbNullString
        If FileName <> "." And FileName <> ".." Then ImportComponent Project, FolderPath, FileName
        FileName = VBA.Dir()
    Loop
    Dim FolderName As String
    FolderName = VBA.Dir(FolderPath, vbDirectory)
    Do While FolderName <> vbNullString
        If FolderName <> "." And FolderName <> ".." Then
            If (VBA.GetAttr(FolderPath & FolderName) And vbDirectory) = vbDirectory Then
                Dim i As Long
                i = i + 1
                Dim SubFolders() As String
                ReDim Preserve SubFolders(i)
                SubFolders(i) = FolderPath & FolderName & charBackSlash
            End If
        End If
        FolderName = VBA.Dir()
    Loop
    If 0 < i Then
        For i = 1 To i
            TraverseFolders Project, SubFolders(i)
        Next
    End If
End Sub

Private Sub ImportComponent(ByVal Project As Object, ByVal FolderPath As String, ByVal FileName As String)
    If Not IsImportTarget(FileName) Then Exit Sub
    Project.VBComponents.Import FolderPath & FileName
End Sub

Private Sub ExportComponent(ByVal Component As Object)
    Dim FilePath As String
    FilePath = ResolveFilePath(Component)
    If VBA.Dir(FilePath) <> vbNullString Then VBA.Kill FilePath
    Component.Export FilePath
End Sub

Private Sub CleanupUnusedComponents(ByVal RootPath As String)
    CleanupFolder RootPath & "CompositionRoot\"
    CleanupFolder RootPath & "Policy\"
    CleanupFolder RootPath & "Domain\"
    CleanupFolder RootPath & "Application\"
    CleanupFolder RootPath & "Presentation\"
    CleanupFolder RootPath & "Infrastructure\"
End Sub

Private Sub CleanupFolder(ByVal FolderPath As String)
    If VBA.Dir(FolderPath, vbDirectory) <> vbNullString Then Exit Sub
    Dim FileName As String
    FileName = VBA.Dir(FolderPath & "*.*")
    Do While 0 < VBA.Len(FileName)
        If FileName <> ".gitkeep" Then
            If Not IsModuleStillExists(RemoveExtension(FileName)) Then VBA.Kill FolderPath & FileName
        End If
        FileName = VBA.Dir
    Loop
End Sub

Private Function CreateAddinWorkbook(ByVal FilePath As String) As Excel.Workbook
    Dim Workbook As Excel.Workbook
    Set Workbook = Application.Workbooks.Add(xlWBATWorksheet)
    RemoveExtraSheets Workbook
    Application.DisplayAlerts = False
    Workbook.SaveAs FileName:=FilePath, FileFormat:=xlOpenXMLAddIn
    Application.DisplayAlerts = True
    Set CreateAddinWorkbook = Workbook
End Function

Private Function RemoveExtension(ByVal FileName As String) As String
    RemoveExtension = VBA.Left$(FileName, VBA.InStrRev(FileName, charPeriod) - 1)
End Function

Private Function IsImportTarget(ByVal FileName As String) As Boolean
    Dim Extension As String
    Extension = VBA.Mid$(FileName, VBA.InStrRev(FileName, charPeriod))
    Select Case VBA.LCase$(Extension)
    Case ".bas", ".cls", ".frm": IsImportTarget = True
    End Select
End Function

Private Function IsExportTarget(ByVal Component As Object) As Boolean
    If Component.Type <> CtStdModule And Component.Type <> CtClassModule And Component.Type <> CtMsForm Then Exit Function
    If HasLayerPrefix(Component.Name) Then IsExportTarget = True
End Function

Private Function IsModuleStillExists(ByVal ModuleName As String) As Boolean
    Dim Component As Object
    For Each Component In ThisWorkbook.VBProject.VBComponents
        If Component.Name = ModuleName Then
            If IsExportTarget(Component) Then
                IsModuleStillExists = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function HasLayerPrefix(ByVal ModuleName As String) As Boolean
    If VBA.Left$(ModuleName, 4) = "Dom_" Then HasLayerPrefix = True
    If VBA.Left$(ModuleName, 4) = "App_" Then HasLayerPrefix = True
    If VBA.Left$(ModuleName, 4) = "Pre_" Then HasLayerPrefix = True
    If VBA.Left$(ModuleName, 4) = "Inf_" Then HasLayerPrefix = True
    If VBA.Left$(ModuleName, 5) = "Compo" Then HasLayerPrefix = True
    If VBA.Right$(ModuleName, 6) = "Policy" Then HasLayerPrefix = True
End Function

Private Function ResolveLayerFolder(ByVal ModuleName As String) As String
    If ModuleName Like "Compo*" Then
        ResolveLayerFolder = RootPath & "CompositionRoot\"
    ElseIf ModuleName Like "Dom_*" Then
        ResolveLayerFolder = RootPath & "Domain\"
    ElseIf ModuleName Like "App_*" Then
        ResolveLayerFolder = RootPath & "Application\"
    ElseIf ModuleName Like "Pre_*" Then
        ResolveLayerFolder = RootPath & "Presentation\"
    ElseIf ModuleName Like "Inf_*" Then
        ResolveLayerFolder = RootPath & "Infrastructure\"
    ElseIf ModuleName Like "*Policy" Then
        ResolveLayerFolder = RootPath & "Policy\"
    Else
        Err.Raise InfErrNotFoundLayerPrefix, "Util_VBComponent", "Layer prefix not found: " & ModuleName
    End If
End Function

Private Function ResolveFilePath(ByVal Component As Object) As String
    Dim LayerFolder As String
    LayerFolder = ResolveLayerFolder(Component.Name)
    Select Case Component.Type
    Case CtStdModule
        ResolveFilePath = LayerFolder & Component.Name & ".bas"
    Case CtClassModule
        ResolveFilePath = LayerFolder & Component.Name & ".cls"
    Case CtMsForm
        ResolveFilePath = LayerFolder & Component.Name & ".frm"
    Case Else
        Err.Raise InfErrUnsupportedComponentType, "Util_VBComponent", "Unsupported component type."
    End Select
End Function
