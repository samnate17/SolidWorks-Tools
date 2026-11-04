Attribute VB_Name = "FeatureWorksStepAssemblyImporter"
Option Explicit

' ============================================================================
' FeatureWorksStepAssemblyImporter
'
' Toolbar-button macro: pick one STEP file (a part OR an assembly) and an
' output folder. The macro:
'
'   1. Imports the STEP file.
'   2. If it's a single part: runs FeatureWorks on it directly.
'   3. If it's an assembly: for every TOP-LEVEL component, saves it out as
'      its own native .SLDPRT into the output folder (this also externalizes
'      a component that came in "virtual"/embedded, since SaveAs on a
'      virtual component's document is how SOLIDWORKS gives it its own
'      file), then runs FeatureWorks on that same component document in
'      place. Because the component keeps referencing the exact file we just
'      edited (we never swap it for a differently-named file), there is no
'      separate "replace component" step - the live assembly already points
'      at the now-recognized geometry once it's rebuilt.
'   4. Rebuilds the assembly, and - on SOLIDWORKS 2025 SP01 or newer only -
'      calls AssemblyDoc::AutoMateRepair to auto-heal mates broken by the
'      swap from dumb imported geometry to recognized features. On older
'      versions this step is skipped and reported as such; broken mates show
'      up greyed-out/red in the Mates folder and need a manual look either
'      way (see README - this cannot be guaranteed for arbitrary geometry).
'   5. Saves the assembly into the output folder.
'
' NOTE: sub-assemblies nested under the top-level assembly are NOT recursed
' into - their components are skipped and reported. Extend
' ProcessOneComponent if you need that.
'
' Same FeatureWorks settings as FeatureWorksStepImporter.bas:
'   Recognition mode = Automatic, Feature type = Standard features,
'   everything available in that mode except Volume (Draft has no
'   automatic-mode option at all, so it's never included either way).
'
' Entry point for a toolbar button: RunFeatureWorksOnStepAssembly
' ============================================================================

' --- fwAutomaticRecognitionOptions_e (FeatureWorks type library) ---
Private Const fwExtrudeOption      As Long = &H1
Private Const fwVolume             As Long = &H2    ' EXCLUDED by request
Private Const fwRevolve            As Long = &H4
Private Const fwHoles              As Long = &H8
Private Const fwChamfils           As Long = &H10
Private Const fwRibs               As Long = &H20
Private Const STANDARD_AUTO_OPTIONS As Long = fwExtrudeOption Or fwRevolve Or fwHoles Or fwChamfils Or fwRibs

' --- fwFeatureCreationOptions_e ---
Private Const fwAllowFailFeatureCreation As Integer = &H2

' ----------------------------------------------------------------------------
' Assign this Sub to a toolbar button (Tools > Customize).
' ----------------------------------------------------------------------------
Sub RunFeatureWorksOnStepAssembly()

    Dim swApp As SldWorks.SldWorks
    Dim fwApp As Object
    Dim stepPath As String
    Dim outputFolder As String
    Dim fileOptions As Long
    Dim fileConfig As String, fileDispName As String
    Dim filterStr As String
    Dim swImportStepData As SldWorks.ImportStepData
    Dim loadErrors As Long
    Dim swModel As SldWorks.ModelDoc2
    Dim report As String

    Set swApp = Application.SldWorks

    Set fwApp = swApp.GetAddInObject("FeatureWorks.FeatureWorksApp")
    If fwApp Is Nothing Then
        MsgBox "The FeatureWorks add-in is not active." & vbCrLf & _
               "Enable it under Tools > Add-ins and try again.", vbExclamation, "FeatureWorks"
        Exit Sub
    End If

    filterStr = "STEP Files (*.step;*.stp)|*.step;*.stp|All Files (*.*)|*.*"
    stepPath = swApp.GetOpenFileName("Select a STEP file (part or assembly)", "", filterStr, fileOptions, fileConfig, fileDispName)
    If Len(stepPath) = 0 Then Exit Sub

    outputFolder = PickOutputFolder()
    If Len(outputFolder) = 0 Then Exit Sub

    Set swImportStepData = swApp.GetImportFileData(stepPath)
    swImportStepData.MapConfigurationData = True

    Set swModel = swApp.LoadFile4(stepPath, "r", swImportStepData, loadErrors)
    If swModel Is Nothing Then
        MsgBox "Could not open " & stepPath & " (error code " & loadErrors & ")", vbExclamation, "FeatureWorks"
        Exit Sub
    End If

    Select Case swModel.GetType
        Case swDocumentTypes_e.swDocASSEMBLY
            report = ProcessStepAssembly(swApp, fwApp, swModel, outputFolder)
        Case swDocumentTypes_e.swDocPART
            report = ProcessStepPart(swApp, fwApp, swModel, outputFolder)
        Case Else
            swApp.CloseDoc swModel.GetTitle
            report = "The STEP file didn't open as a part or an assembly - nothing to do."
    End Select

    MsgBox report, vbInformation, "FeatureWorks"

End Sub

' Single-part STEP: run FeatureWorks directly, save into the output folder.
Private Function ProcessStepPart(swApp As SldWorks.SldWorks, fwApp As Object, swModel As SldWorks.ModelDoc2, outputFolder As String) As String

    Dim recognizedCount As Long
    Dim savePath As String
    Dim saveErrors As Long, saveWarnings As Long
    Dim baseName As String
    Dim dotPos As Long

    baseName = swModel.GetTitle
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then baseName = Left$(baseName, dotPos - 1)

    recognizedCount = RecognizeActiveDocFeatures(fwApp, swModel)

    savePath = outputFolder & "\" & baseName & ".SLDPRT"
    swModel.Extension.SaveAs savePath, 0, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, saveErrors, saveWarnings
    swApp.CloseDoc swModel.GetTitle

    ProcessStepPart = "Single part. Recognized " & recognizedCount & " feature(s)." & vbCrLf & _
                       "Saved to: " & savePath
End Function

' STEP assembly: externalize + recognize every top-level component in place,
' then rebuild/repair/save the assembly itself.
Private Function ProcessStepAssembly(swApp As SldWorks.SldWorks, fwApp As Object, swModel As SldWorks.ModelDoc2, outputFolder As String) As String

    Dim swAssy As SldWorks.AssemblyDoc
    Dim comps As Variant
    Dim singleComp As Variant
    Dim comp As SldWorks.Component2
    Dim usedNames As Collection
    Dim assyTitle As String
    Dim processedParts As Long, skippedParts As Long, totalRecognized As Long
    Dim thisCount As Long
    Dim swAssyActive As Object
    Dim actErrors As Long
    Dim mateReport As String
    Dim assySavePath As String
    Dim saveErrors As Long, saveWarnings As Long
    Dim baseName As String
    Dim dotPos As Long

    Set swAssy = swModel
    assyTitle = swModel.GetTitle

    comps = swAssy.GetComponents(True) ' top-level components only

    Set usedNames = New Collection

    If Not IsEmpty(comps) Then
        For Each singleComp In comps
            Set comp = singleComp
            If ProcessOneComponent(swApp, fwApp, comp, outputFolder, usedNames, thisCount) Then
                processedParts = processedParts + 1
                totalRecognized = totalRecognized + thisCount
            Else
                skippedParts = skippedParts + 1
            End If
        Next singleComp
    End If

    ' Back to the assembly itself.
    Set swAssyActive = swApp.ActivateDoc3(assyTitle, False, swRebuildOnActivation_e.swRebuildActiveDoc, actErrors)
    If swAssyActive Is Nothing Then Set swAssyActive = swModel

    swAssyActive.ForceRebuild3 False

    mateReport = TryAutoMateRepair(swAssyActive)

    baseName = assyTitle
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then baseName = Left$(baseName, dotPos - 1)
    assySavePath = outputFolder & "\" & baseName & ".SLDASM"

    swAssyActive.Extension.SaveAs assySavePath, 0, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, saveErrors, saveWarnings

    ProcessStepAssembly = "Assembly processed." & vbCrLf & _
        "Components recognized: " & processedParts & " (total features: " & totalRecognized & ")" & vbCrLf & _
        "Components skipped (sub-assemblies or unresolvable): " & skippedParts & vbCrLf & _
        mateReport & vbCrLf & _
        "Saved assembly to: " & assySavePath & vbCrLf & _
        "Saved parts to: " & outputFolder
End Function

' Externalizes one top-level component to its own SLDPRT in outputFolder and
' runs FeatureWorks on it in place. Returns False (and leaves the component
' untouched) for sub-assemblies or anything that can't be resolved.
Private Function ProcessOneComponent(swApp As SldWorks.SldWorks, fwApp As Object, comp As SldWorks.Component2, _
                                      outputFolder As String, usedNames As Collection, ByRef recognizedCount As Long) As Boolean

    Dim compDoc As SldWorks.ModelDoc2
    Dim activeDoc As Object
    Dim outPath As String
    Dim saveErrors As Long, saveWarnings As Long
    Dim actErrors As Long

    recognizedCount = 0
    ProcessOneComponent = False

    On Error Resume Next
    Set compDoc = comp.GetModelDoc2()
    On Error GoTo 0
    If compDoc Is Nothing Then Exit Function
    If compDoc.GetType <> swDocumentTypes_e.swDocPART Then Exit Function ' nested sub-assembly - skip

    outPath = outputFolder & "\" & UniqueFileBaseName(comp.Name2, usedNames) & ".SLDPRT"

    ' Saving the component's own document directly (not via ActiveDoc) both
    ' writes the file AND, for a virtual/embedded component, is what gives
    ' it its own external file in the live assembly.
    compDoc.Extension.SaveAs outPath, 0, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, saveErrors, saveWarnings
    If saveErrors <> 0 Then Exit Function

    Set activeDoc = swApp.ActivateDoc3(compDoc.GetTitle, False, swRebuildOnActivation_e.swDontRebuildActiveDoc, actErrors)
    If activeDoc Is Nothing Then Exit Function

    recognizedCount = RecognizeActiveDocFeatures(fwApp, activeDoc)

    activeDoc.Save3 swSaveAsOptions_e.swSaveAsOptions_Silent, saveErrors, saveWarnings

    ProcessOneComponent = True
End Function

' Shared recognition step: runs on whatever SOLIDWORKS currently considers
' the active document (FeatureWorks's automation methods take no document
' argument at all, so the right doc must already be active before calling).
Private Function RecognizeActiveDocFeatures(fwApp As Object, activeDoc As Object) As Long

    Dim recognizedCount As Long

    recognizedCount = fwApp.RecognizeFeatureAutomatic(STANDARD_AUTO_OPTIONS)
    If recognizedCount > 0 Then
        fwApp.CreateFeatures fwAllowFailFeatureCreation
    End If
    activeDoc.EditRebuild3

    RecognizeActiveDocFeatures = recognizedCount
End Function

' Best-effort mate repair. AssemblyDoc::AutoMateRepair only exists from
' SOLIDWORKS 2025 SP01 onward, so this is late-bound and error-guarded to
' keep the macro compiling/running fine on older versions too.
Private Function TryAutoMateRepair(assyDoc As Object) As String

    Dim processedMates As Variant
    Dim failedMates As Variant
    Dim processedCount As Long, failedCount As Long

    On Error GoTo NotAvailable

    assyDoc.Extension.SelectByID2 "", "MATEGROUP", 0, 0, 0, False, 0, Nothing, 0
    assyDoc.AutoMateRepair processedMates, failedMates
    assyDoc.ClearSelection2 True

    If Not IsEmpty(processedMates) Then processedCount = UBound(processedMates) - LBound(processedMates) + 1
    If Not IsEmpty(failedMates) Then failedCount = UBound(failedMates) - LBound(failedMates) + 1

    TryAutoMateRepair = "Mate auto-repair: " & processedCount & " repaired, " & failedCount & _
                         " still broken (check the Mates folder in the FeatureManager tree)."
    Exit Function

NotAvailable:
    TryAutoMateRepair = "Mate auto-repair needs SOLIDWORKS 2025 SP01 or newer - skipped on this version. " & _
                         "Check the Mates folder manually for anything greyed-out or red."
End Function

' Lets the user pick a folder without any extra VBA references.
Private Function PickOutputFolder() As String

    Dim shellApp As Object
    Dim folderItem As Object

    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    Set folderItem = shellApp.BrowseForFolder(0, "Select an output folder for the recognized parts/assembly", 0, "")
    On Error GoTo 0

    If folderItem Is Nothing Then
        PickOutputFolder = ""
    Else
        PickOutputFolder = folderItem.Self.Path
    End If
End Function

' Builds a filesystem-safe, collision-free base filename from a component's
' display name (strips the "-N" instance suffix and "^AssemblyName" virtual-
' component suffix SOLIDWORKS adds).
Private Function UniqueFileBaseName(componentName As String, usedNames As Collection) As String

    Dim baseName As String
    Dim candidateName As String
    Dim caretPos As Long, dashPos As Long
    Dim n As Long
    Dim nm As Variant
    Dim isUsed As Boolean

    baseName = componentName

    caretPos = InStr(baseName, "^")
    If caretPos > 0 Then baseName = Left$(baseName, caretPos - 1)

    dashPos = InStrRev(baseName, "-")
    If dashPos > 0 Then
        If IsNumeric(Mid$(baseName, dashPos + 1)) Then baseName = Left$(baseName, dashPos - 1)
    End If

    baseName = CleanFileName(baseName)

    candidateName = baseName
    n = 1
    Do
        isUsed = False
        For Each nm In usedNames
            If StrComp(CStr(nm), candidateName, vbTextCompare) = 0 Then
                isUsed = True
                Exit For
            End If
        Next nm
        If isUsed Then
            n = n + 1
            candidateName = baseName & "_" & n
        End If
    Loop While isUsed

    usedNames.Add candidateName
    UniqueFileBaseName = candidateName
End Function

Private Function CleanFileName(ByVal s As String) As String

    Dim badChars As String
    Dim i As Long

    badChars = "\/:*?""<>|"
    For i = 1 To Len(badChars)
        s = Replace(s, Mid$(badChars, i, 1), "_")
    Next i

    CleanFileName = Trim$(s)
    If Len(CleanFileName) = 0 Then CleanFileName = "Component"
End Function
