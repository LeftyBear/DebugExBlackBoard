Attribute VB_Name = "Inf_VBComponentUtility"
'@Folder("Infrastructure.Service")
Option Explicit

Private Const CtStdModule   As Long = 1
Private Const CtClassModule As Long = 2
Private Const CtMsForm      As Long = 3
Private Const ROOT_PATH     As String = "C:\Users\biz\Documents\GitHub\DebugExBlackBoard\"
Private Const MODULE_FILE   As String = ".bas"
Private Const CLASS_FILE    As String = ".cls"
Private Const FORM_FILE     As String = ".frm"

Public Sub ExportAllModules()
    If Inf_Environment.GetEnvironmentTypeCode = ReleaseMode Then Exit Sub
    Dim Component As Object
    For Each Component In ThisWorkbook.VBProject.VBComponents
        If IsExportTarget(Component) Then ExportComponent Component
    Next
    CleanupUnusedComponents ROOT_PATH
End Sub

Private Sub ExportComponent(ByVal Component As Object)
    Dim FilePath As String
    FilePath = ResolveFilePath(Component)
    If VBA.Dir(FilePath) <> vbNullString Then VBA.Kill FilePath
    Component.Export FilePath
End Sub

Private Sub CleanupUnusedComponents(ByVal ROOT_PATH As String)
    CleanupFolder ROOT_PATH & "CompositionRoot\"
    CleanupFolder ROOT_PATH & "Policy\"
    CleanupFolder ROOT_PATH & "Domain\"
    CleanupFolder ROOT_PATH & "Application\"
    CleanupFolder ROOT_PATH & "Presentation\"
    CleanupFolder ROOT_PATH & "Infrastructure\"
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

Private Function RemoveExtension(ByVal FileName As String) As String
    RemoveExtension = VBA.Left$(FileName, VBA.InStrRev(FileName, ".") - 1)
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

Private Function ResolveFilePath(ByVal Component As Object) As String
    Dim LayerFolder As String
    LayerFolder = ResolveLayerFolder(Component.Name)
    Select Case Component.Type
    Case CtStdModule
        ResolveFilePath = LayerFolder & Component.Name & MODULE_FILE
    Case CtClassModule
        ResolveFilePath = LayerFolder & Component.Name & CLASS_FILE
    Case CtMsForm
        ResolveFilePath = LayerFolder & Component.Name & FORM_FILE
    Case Else
        Err.Raise InfErrUnsupportedComponentType, "Util_VBComponent", "Unsupported component type."
    End Select
End Function

Private Function ResolveLayerFolder(ByVal ModuleName As String) As String
    If ModuleName Like "Compo*" Then
        ResolveLayerFolder = ROOT_PATH & "CompositionRoot\"
    ElseIf ModuleName Like "Dom_*" Then
        ResolveLayerFolder = ROOT_PATH & "Domain\"
    ElseIf ModuleName Like "App_*" Then
        ResolveLayerFolder = ROOT_PATH & "Application\"
    ElseIf ModuleName Like "Pre_*" Then
        ResolveLayerFolder = ROOT_PATH & "Presentation\"
    ElseIf ModuleName Like "Inf_*" Then
        ResolveLayerFolder = ROOT_PATH & "Infrastructure\"
    ElseIf ModuleName Like "*Policy" Then
        ResolveLayerFolder = ROOT_PATH & "Policy\"
    Else
        Err.Raise InfErrNotFoundLayerPrefix, "Util_VBComponent", "Layer prefix not found: " & ModuleName
    End If
End Function

