Attribute VB_Name = "modVbaSync"
Option Explicit

' Full VBA sync helper:
' - Export all VBA components to text files.
' - Import them back from files.
' - Includes standard modules, class modules, userforms, and document modules
'   (ThisWorkbook + worksheet code modules).
'
' Notes:
' 1) Enable "Trust access to the VBA project object model" in Excel settings.
' 2) If VBA project is password-protected, import/export will fail.
' 3) The module named SYNC_MODULE_NAME is protected from removal during import.

Private Const DEFAULT_SYNC_DIR As String = "_vba_sync"
Private Const MANIFEST_FILE As String = "manifest.tsv"

Private Const FOLDER_MODULES As String = "modules"
Private Const FOLDER_CLASSES As String = "classes"
Private Const FOLDER_FORMS As String = "forms"
Private Const FOLDER_DOCUMENTS As String = "documents"

Private Const KIND_STD As String = "STD"
Private Const KIND_CLS As String = "CLS"
Private Const KIND_FRM As String = "FRM"
Private Const KIND_DOC As String = "DOC"

Private Const VBEXT_CT_STD_MODULE As Long = 1
Private Const VBEXT_CT_CLASS_MODULE As Long = 2
Private Const VBEXT_CT_MS_FORM As Long = 3
Private Const VBEXT_CT_DOCUMENT As Long = 100

Private Const FOR_READING As Long = 1
Private Const FOR_WRITING As Long = 2
Private Const TRISTATE_TRUE As Long = -1

Private Const SYNC_MODULE_NAME As String = "modVbaSync"

Public Sub ExportAllModules()
    ExportAllModulesTo vbNullString
End Sub

Public Sub ImportAllModules()
    ImportAllModulesFrom vbNullString
End Sub

Public Sub ExportAllModules_UI()
    On Error GoTo EH

    Dim rootPath As String
    rootPath = ResolveSyncRoot(vbNullString)

    Dim chosenFolder As String
    chosenFolder = PickFolderPath("Select folder for VBA export", rootPath)
    If Len(chosenFolder) = 0 Then
        Exit Sub
    End If

    ExportAllModulesTo chosenFolder
    Exit Sub

EH:
    MsgBox "Export UI failed: " & Err.Description, vbCritical
End Sub

Public Sub ImportAllModules_UI()
    On Error GoTo EH

    Dim rootPath As String
    rootPath = ResolveSyncRoot(vbNullString)

    Dim chosenFolder As String
    chosenFolder = PickFolderPath("Select folder with VBA files to import", rootPath)
    If Len(chosenFolder) = 0 Then
        Exit Sub
    End If

    Dim answer As VbMsgBoxResult
    answer = MsgBox("Import VBA components from folder?" & vbCrLf & chosenFolder & vbCrLf & vbCrLf & _
                    "This will replace current project modules (except modVbaSync).", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Import")
    If answer <> vbYes Then
        Exit Sub
    End If

    ImportAllModulesFrom chosenFolder
    Exit Sub

EH:
    MsgBox "Import UI failed: " & Err.Description, vbCritical
End Sub

Public Sub ExportAllModulesTo(ByVal syncRoot As String)
    On Error GoTo EH

    Dim rootPath As String
    rootPath = ResolveSyncRoot(syncRoot)

    Dim modulesPath As String
    Dim classesPath As String
    Dim formsPath As String
    Dim docsPath As String

    modulesPath = CombinePath(rootPath, FOLDER_MODULES)
    classesPath = CombinePath(rootPath, FOLDER_CLASSES)
    formsPath = CombinePath(rootPath, FOLDER_FORMS)
    docsPath = CombinePath(rootPath, FOLDER_DOCUMENTS)

    EnsureFolder rootPath
    EnsureCleanFolder modulesPath
    EnsureCleanFolder classesPath
    EnsureCleanFolder formsPath
    EnsureCleanFolder docsPath
    DeleteFileIfExists CombinePath(rootPath, MANIFEST_FILE)

    Dim vbProj As Object
    Set vbProj = GetVbProjectOrThrow()

    Dim manifestRows As Collection
    Set manifestRows = New Collection

    Dim vbComp As Object
    Dim kind As String
    Dim relPath As String
    Dim fullPath As String
    Dim codeText As String
    Dim exportCount As Long

    For Each vbComp In vbProj.VBComponents
        kind = KindFromComponentType(CLng(vbComp.Type))
        If Len(kind) = 0 Then
            GoTo NextComponent
        End If

        relPath = RelativePathForComponent(kind, CStr(vbComp.Name))
        fullPath = CombinePath(rootPath, relPath)

        Select Case kind
            Case KIND_STD, KIND_CLS, KIND_FRM
                vbComp.Export fullPath
            Case KIND_DOC
                codeText = GetComponentCodeText(vbComp)
                WriteUnicodeTextFile fullPath, codeText
            Case Else
                GoTo NextComponent
        End Select

        manifestRows.Add BuildManifestRow(kind, CStr(vbComp.Name), relPath)
        exportCount = exportCount + 1
NextComponent:
    Next vbComp

    WriteManifestFile CombinePath(rootPath, MANIFEST_FILE), manifestRows

    MsgBox "Export completed." & vbCrLf & _
           "Components exported: " & CStr(exportCount) & vbCrLf & _
           "Folder: " & rootPath, vbInformation
    Exit Sub

EH:
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

Public Sub ImportAllModulesFrom(ByVal syncRoot As String)
    On Error GoTo EH

    Dim rootPath As String
    rootPath = ResolveSyncRoot(syncRoot)

    Dim manifestPath As String
    manifestPath = CombinePath(rootPath, MANIFEST_FILE)

    If Not FileExists(manifestPath) Then
        Err.Raise vbObjectError + 1100, "modVbaSync.ImportAllModulesFrom", _
                  "Manifest file not found: " & manifestPath
    End If

    Dim manifestRows As Collection
    Set manifestRows = ReadManifestFile(manifestPath)
    ValidateManifestRows manifestRows, rootPath

    Dim vbProj As Object
    Set vbProj = GetVbProjectOrThrow()

    Dim removedCount As Long
    removedCount = RemoveAllNonDocumentComponents(vbProj, SYNC_MODULE_NAME)

    Dim rowText As Variant
    Dim kind As String
    Dim compName As String
    Dim relPath As String
    Dim fullPath As String
    Dim importedComp As Object

    Dim importedCount As Long
    Dim updatedDocsCount As Long
    Dim skippedSelfCount As Long

    For Each rowText In manifestRows
        ParseManifestRow CStr(rowText), kind, compName, relPath
        fullPath = CombinePath(rootPath, relPath)

        Select Case kind
            Case KIND_STD, KIND_CLS, KIND_FRM
                If StrComp(compName, SYNC_MODULE_NAME, vbTextCompare) = 0 Then
                    skippedSelfCount = skippedSelfCount + 1
                    GoTo NextImportRow
                End If

                Set importedComp = vbProj.VBComponents.Import(fullPath)
                On Error Resume Next
                If StrComp(CStr(importedComp.Name), compName, vbTextCompare) <> 0 Then
                    importedComp.Name = compName
                End If
                On Error GoTo EH

                importedCount = importedCount + 1

            Case KIND_DOC
                Dim targetComp As Object
                Set targetComp = FindComponentByName(vbProj, compName)
                If targetComp Is Nothing Then
                    Err.Raise vbObjectError + 1101, "modVbaSync.ImportAllModulesFrom", _
                              "Document component not found in workbook: " & compName
                End If
                If CLng(targetComp.Type) <> VBEXT_CT_DOCUMENT Then
                    Err.Raise vbObjectError + 1102, "modVbaSync.ImportAllModulesFrom", _
                              "Component is not a document module: " & compName
                End If

                ReplaceComponentCode targetComp, ReadUnicodeTextFile(fullPath)
                updatedDocsCount = updatedDocsCount + 1
        End Select
NextImportRow:
    Next rowText

    MsgBox "Import completed." & vbCrLf & _
           "Removed non-document components: " & CStr(removedCount) & vbCrLf & _
           "Imported modules/classes/forms: " & CStr(importedCount) & vbCrLf & _
           "Updated document modules: " & CStr(updatedDocsCount) & vbCrLf & _
           "Skipped protected module: " & CStr(skippedSelfCount), vbInformation
    Exit Sub

EH:
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub

Private Function ResolveSyncRoot(ByVal syncRoot As String) As String
    Dim value As String
    value = Trim$(syncRoot)

    If Len(value) = 0 Then
        If Len(ThisWorkbook.Path) = 0 Then
            Err.Raise vbObjectError + 1000, "modVbaSync.ResolveSyncRoot", _
                      "Workbook must be saved before using default sync folder."
        End If
        ResolveSyncRoot = CombinePath(ThisWorkbook.Path, DEFAULT_SYNC_DIR)
    Else
        ResolveSyncRoot = value
    End If
End Function

Private Function PickFolderPath(ByVal dialogTitle As String, ByVal initialFolder As String) As String
    On Error GoTo Fallback

    Const MsoFileDialogFolderPicker As Long = 4

    Dim dlg As Object
    Set dlg = Application.FileDialog(MsoFileDialogFolderPicker)

    With dlg
        .Title = dialogTitle
        .AllowMultiSelect = False
        If Len(initialFolder) > 0 Then
            .InitialFileName = EnsureTrailingBackslash(initialFolder)
        End If
        If .Show = -1 Then
            If .SelectedItems.Count > 0 Then
                PickFolderPath = CStr(.SelectedItems(1))
            End If
        End If
    End With
    Exit Function

Fallback:
    PickFolderPath = InputBox(dialogTitle, "Folder Path", initialFolder)
End Function

Private Function KindFromComponentType(ByVal componentType As Long) As String
    Select Case componentType
        Case VBEXT_CT_STD_MODULE
            KindFromComponentType = KIND_STD
        Case VBEXT_CT_CLASS_MODULE
            KindFromComponentType = KIND_CLS
        Case VBEXT_CT_MS_FORM
            KindFromComponentType = KIND_FRM
        Case VBEXT_CT_DOCUMENT
            KindFromComponentType = KIND_DOC
        Case Else
            KindFromComponentType = vbNullString
    End Select
End Function

Private Function RelativePathForComponent(ByVal kind As String, ByVal componentName As String) As String
    Dim safeName As String
    safeName = SafeFileName(componentName)

    Select Case kind
        Case KIND_STD
            RelativePathForComponent = FOLDER_MODULES & "\" & safeName & ".bas"
        Case KIND_CLS
            RelativePathForComponent = FOLDER_CLASSES & "\" & safeName & ".cls"
        Case KIND_FRM
            RelativePathForComponent = FOLDER_FORMS & "\" & safeName & ".frm"
        Case KIND_DOC
            RelativePathForComponent = FOLDER_DOCUMENTS & "\" & safeName & ".bas"
        Case Else
            Err.Raise vbObjectError + 1001, "modVbaSync.RelativePathForComponent", _
                      "Unknown component kind: " & kind
    End Select
End Function

Private Function SafeFileName(ByVal value As String) As String
    Dim result As String
    result = value

    Dim badChars As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    Dim i As Long
    For i = LBound(badChars) To UBound(badChars)
        result = Replace(result, CStr(badChars(i)), "_")
    Next i

    If Len(result) = 0 Then
        result = "_"
    End If

    SafeFileName = result
End Function

Private Function BuildManifestRow(ByVal kind As String, ByVal name As String, ByVal relativePath As String) As String
    BuildManifestRow = kind & vbTab & name & vbTab & relativePath
End Function

Private Sub ParseManifestRow(ByVal rowText As String, ByRef kind As String, ByRef name As String, ByRef relativePath As String)
    Dim parts() As String
    parts = Split(rowText, vbTab)

    If UBound(parts) < 2 Then
        Err.Raise vbObjectError + 1002, "modVbaSync.ParseManifestRow", _
                  "Invalid manifest row: " & rowText
    End If

    kind = Trim$(parts(0))
    name = Trim$(parts(1))
    relativePath = Trim$(parts(2))
End Sub

Private Sub WriteManifestFile(ByVal manifestPath As String, ByVal rows As Collection)
    Dim outputText As String
    outputText = "Kind" & vbTab & "Name" & vbTab & "RelativePath" & vbCrLf

    Dim rowText As Variant
    For Each rowText In rows
        outputText = outputText & CStr(rowText) & vbCrLf
    Next rowText

    WriteUnicodeTextFile manifestPath, outputText
End Sub

Private Function ReadManifestFile(ByVal manifestPath As String) As Collection
    Dim rows As Collection
    Set rows = New Collection

    Dim content As String
    content = ReadUnicodeTextFile(manifestPath)

    Dim lines As Variant
    lines = SplitToLines(content)

    Dim i As Long
    Dim lineText As String
    For i = LBound(lines) To UBound(lines)
        lineText = CStr(lines(i))
        If Len(Trim$(lineText)) = 0 Then
            GoTo NextLine
        End If
        If i = LBound(lines) Then
            GoTo NextLine
        End If
        rows.Add lineText
NextLine:
    Next i

    Set ReadManifestFile = rows
End Function

Private Sub ValidateManifestRows(ByVal rows As Collection, ByVal rootPath As String)
    If rows Is Nothing Or rows.Count = 0 Then
        Err.Raise vbObjectError + 1003, "modVbaSync.ValidateManifestRows", _
                  "Manifest has no component rows."
    End If

    Dim rowText As Variant
    Dim kind As String
    Dim compName As String
    Dim relPath As String
    Dim fullPath As String

    For Each rowText In rows
        ParseManifestRow CStr(rowText), kind, compName, relPath
        If Len(kind) = 0 Or Len(compName) = 0 Or Len(relPath) = 0 Then
            Err.Raise vbObjectError + 1004, "modVbaSync.ValidateManifestRows", _
                      "Manifest row has empty fields: " & CStr(rowText)
        End If
        fullPath = CombinePath(rootPath, relPath)
        If Not FileExists(fullPath) Then
            Err.Raise vbObjectError + 1005, "modVbaSync.ValidateManifestRows", _
                      "Component file listed in manifest not found: " & fullPath
        End If
    Next rowText
End Sub

Private Function RemoveAllNonDocumentComponents(ByVal vbProj As Object, ByVal protectedComponentName As String) As Long
    Dim namesToRemove As Collection
    Set namesToRemove = New Collection

    Dim vbComp As Object
    For Each vbComp In vbProj.VBComponents
        If CLng(vbComp.Type) <> VBEXT_CT_DOCUMENT Then
            If StrComp(CStr(vbComp.Name), protectedComponentName, vbTextCompare) <> 0 Then
                namesToRemove.Add CStr(vbComp.Name)
            End If
        End If
    Next vbComp

    Dim nameItem As Variant
    For Each nameItem In namesToRemove
        vbProj.VBComponents.Remove vbProj.VBComponents(CStr(nameItem))
        RemoveAllNonDocumentComponents = RemoveAllNonDocumentComponents + 1
    Next nameItem
End Function

Private Function FindComponentByName(ByVal vbProj As Object, ByVal componentName As String) As Object
    On Error Resume Next
    Set FindComponentByName = vbProj.VBComponents(componentName)
    On Error GoTo 0
End Function

Private Function GetComponentCodeText(ByVal vbComp As Object) As String
    Dim codeModule As Object
    Set codeModule = vbComp.CodeModule

    If codeModule.CountOfLines <= 0 Then
        GetComponentCodeText = vbNullString
    Else
        GetComponentCodeText = codeModule.Lines(1, codeModule.CountOfLines)
    End If
End Function

Private Sub ReplaceComponentCode(ByVal vbComp As Object, ByVal codeText As String)
    Dim codeModule As Object
    Set codeModule = vbComp.CodeModule

    If codeModule.CountOfLines > 0 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines
    End If

    codeText = NormalizeLineEndings(codeText)
    If Len(codeText) > 0 Then
        codeModule.AddFromString codeText
    End If
End Sub

Private Function NormalizeLineEndings(ByVal text As String) As String
    Dim result As String
    result = Replace(text, vbCrLf, vbLf)
    result = Replace(result, vbCr, vbLf)
    NormalizeLineEndings = Replace(result, vbLf, vbCrLf)
End Function

Private Function SplitToLines(ByVal text As String) As Variant
    Dim normalized As String
    normalized = Replace(text, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    SplitToLines = Split(normalized, vbLf)
End Function

Private Function GetVbProjectOrThrow() As Object
    On Error GoTo EH

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim componentCount As Long
    componentCount = vbProj.VBComponents.Count

    Set GetVbProjectOrThrow = vbProj
    Exit Function

EH:
    Err.Raise vbObjectError + 1006, "modVbaSync.GetVbProjectOrThrow", _
              "Cannot access VBA project. Enable Trust Center option " & _
              "'Trust access to the VBA project object model'."
End Function

Private Function CombinePath(ByVal leftPart As String, ByVal rightPart As String) As String
    If Len(leftPart) = 0 Then
        CombinePath = rightPart
    ElseIf Len(rightPart) = 0 Then
        CombinePath = leftPart
    ElseIf Right$(leftPart, 1) = "\" Then
        CombinePath = leftPart & rightPart
    Else
        CombinePath = leftPart & "\" & rightPart
    End If
End Function

Private Function EnsureTrailingBackslash(ByVal value As String) As String
    If Len(value) = 0 Then
        EnsureTrailingBackslash = value
    ElseIf Right$(value, 1) = "\" Then
        EnsureTrailingBackslash = value
    Else
        EnsureTrailingBackslash = value & "\"
    End If
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function

Private Sub DeleteFileIfExists(ByVal filePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath, True
    End If
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(folderPath) = 0 Then
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        Exit Sub
    End If

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 Then
        EnsureFolder parentPath
    End If

    fso.CreateFolder folderPath
End Sub

Private Sub EnsureCleanFolder(ByVal folderPath As String)
    EnsureFolder folderPath

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderObj As Object
    Set folderObj = fso.GetFolder(folderPath)

    Dim fileObj As Object
    For Each fileObj In folderObj.Files
        fileObj.Delete True
    Next fileObj

    Dim subFolderObj As Object
    For Each subFolderObj In folderObj.SubFolders
        subFolderObj.Delete True
    Next subFolderObj
End Sub

Private Sub WriteUnicodeTextFile(ByVal filePath As String, ByVal content As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(filePath)
    If Len(parentPath) > 0 Then
        EnsureFolder parentPath
    End If

    Dim stream As Object
    Set stream = fso.OpenTextFile(filePath, FOR_WRITING, True, TRISTATE_TRUE)
    stream.Write content
    stream.Close
End Sub

Private Function ReadUnicodeTextFile(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim stream As Object
    Set stream = fso.OpenTextFile(filePath, FOR_READING, False, TRISTATE_TRUE)
    ReadUnicodeTextFile = stream.ReadAll
    stream.Close
End Function
