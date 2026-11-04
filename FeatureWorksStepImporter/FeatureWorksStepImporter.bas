Attribute VB_Name = "FeatureWorksStepImporter"
Option Explicit

' ============================================================================
' FeatureWorksStepImporter
'
' Toolbar-button macro: pick one or more STEP files, import each one, and run
' FeatureWorks on it with:
'   Recognition mode = Automatic
'   Feature type     = Standard features (not sheet metal)
'   Auto-recognize    = everything available in that mode EXCEPT Volume
'                        (Draft has no automatic-mode option at all, so it is
'                        never included)
'
' Entry point for a toolbar button: RunFeatureWorksOnStepFiles
' ============================================================================

' --- fwAutomaticRecognitionOptions_e (FeatureWorks type library) ---
' https://help.solidworks.com/.../fwAutomaticRecognitionOptions_e.html
Private Const fwExtrudeOption      As Long = &H1    ' extrude/boss/cut - Standard
Private Const fwVolume             As Long = &H2    ' volume features  - EXCLUDED by request
Private Const fwRevolve            As Long = &H4    ' Standard
Private Const fwHoles              As Long = &H8    ' Standard
Private Const fwChamfils           As Long = &H10   ' fillets & chamfers - Standard
Private Const fwRibs               As Long = &H20   ' Standard
Private Const fwBaseFlange         As Long = &H40   ' Sheet Metal - not "Standard features"
Private Const fwSketchedBend       As Long = &H80   ' Sheet Metal
Private Const fwAutoEdgeFlange     As Long = &H100  ' Sheet Metal
Private Const fwAutoHemFlange      As Long = &H200  ' Sheet Metal

' All "Standard features" automatic-recognition bits except Volume.
' (There is no automatic-mode bit for Draft - FeatureWorks only recognizes
' draft features interactively - so "except ... drafts" is satisfied for free.)
Private Const STANDARD_AUTO_OPTIONS As Long = fwExtrudeOption Or fwRevolve Or fwHoles Or fwChamfils Or fwRibs

' --- fwFeatureCreationOptions_e ---
' https://help.solidworks.com/.../fwFeatureCreationOptions_e.html
Private Const fwAllowFailFeatureCreation As Integer = &H2

' --- swDocumentTypes_e.swDocPART (standard SOLIDWORKS constant) ---
Private Const SW_DOC_PART As Long = 1

' ----------------------------------------------------------------------------
' Assign this Sub to a toolbar button:
'   Tools > Customize > Macros tab (or Commands tab, Macro category)
'   > drag a button onto a toolbar > Browse to this macro's .swp file
'   > pick "RunFeatureWorksOnStepFiles" > choose an icon/tooltip > OK
' ----------------------------------------------------------------------------
Sub RunFeatureWorksOnStepFiles()

    Dim swApp As SldWorks.SldWorks
    Dim fwApp As Object
    Dim stepPath As String
    Dim fileOptions As Long
    Dim fileConfig As String
    Dim fileDispName As String
    Dim filterStr As String
    Dim keepGoing As VbMsgBoxResult
    Dim okCount As Long, failCount As Long
    Dim resultMsg As String

    Set swApp = Application.SldWorks

    Set fwApp = swApp.GetAddInObject("FeatureWorks.FeatureWorksApp")
    If fwApp Is Nothing Then
        MsgBox "The FeatureWorks add-in is not active." & vbCrLf & _
               "Enable it under Tools > Add-ins (tick both ""Active"" and " & _
               """Start Up"" for FeatureWorks) and try again.", vbExclamation, "FeatureWorks"
        Exit Sub
    End If

    filterStr = "STEP Files (*.step;*.stp)|*.step;*.stp|All Files (*.*)|*.*"

    Do
        stepPath = swApp.GetOpenFileName("Select a STEP file for FeatureWorks recognition", _
                                          "", filterStr, fileOptions, fileConfig, fileDispName)

        If Len(stepPath) = 0 Then Exit Do ' user cancelled the picker

        resultMsg = ImportAndRecognizeStep(swApp, fwApp, stepPath)
        If Len(resultMsg) = 0 Then
            okCount = okCount + 1
        Else
            failCount = failCount + 1
            MsgBox resultMsg, vbExclamation, "FeatureWorks"
        End If

        keepGoing = MsgBox("Select another STEP file?", vbYesNo + vbQuestion, "FeatureWorks")
    Loop While keepGoing = vbYes

    If okCount + failCount > 0 Then
        MsgBox "FeatureWorks batch complete." & vbCrLf & _
               "Succeeded: " & okCount & vbCrLf & _
               "Failed/skipped: " & failCount, vbInformation, "FeatureWorks"
    End If

End Sub

' Imports one STEP file and runs FeatureWorks on it.
' Returns "" on success, or an error description on failure.
Private Function ImportAndRecognizeStep(swApp As SldWorks.SldWorks, fwApp As Object, stepPath As String) As String

    Dim swImportStepData As SldWorks.ImportStepData
    Dim swModel As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim loadErrors As Long
    Dim recognizedCount As Long
    Dim created As Boolean
    Dim docTitle As String
    Dim savePath As String
    Dim saveErrors As Long, saveWarnings As Long
    Dim dotPos As Long

    On Error GoTo HandleError

    Set swImportStepData = swApp.GetImportFileData(stepPath)
    swImportStepData.MapConfigurationData = True

    Set swModel = swApp.LoadFile4(stepPath, "r", swImportStepData, loadErrors)
    If swModel Is Nothing Then
        ImportAndRecognizeStep = "Could not open: " & stepPath & " (error code " & loadErrors & ")"
        Exit Function
    End If

    docTitle = swModel.GetTitle

    If swModel.GetType <> SW_DOC_PART Then
        swApp.CloseDoc docTitle
        ImportAndRecognizeStep = "Skipped (not a single-body part): " & stepPath
        Exit Function
    End If

    Set swModelDocExt = swModel.Extension

    recognizedCount = fwApp.RecognizeFeatureAutomatic(STANDARD_AUTO_OPTIONS)

    If recognizedCount > 0 Then
        created = fwApp.CreateFeatures(fwAllowFailFeatureCreation)
    End If

    swModel.EditRebuild3

    dotPos = InStrRev(stepPath, ".")
    If dotPos > 0 Then
        savePath = Left$(stepPath, dotPos - 1) & ".SLDPRT"
    Else
        savePath = stepPath & ".SLDPRT"
    End If

    swModelDocExt.SaveAs savePath, 0, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, saveErrors, saveWarnings

    swApp.CloseDoc docTitle

    If saveErrors <> 0 Then
        ImportAndRecognizeStep = "Recognized " & recognizedCount & " feature(s) but failed to save: " & savePath
    Else
        ImportAndRecognizeStep = "" ' success
    End If

    Exit Function

HandleError:
    ImportAndRecognizeStep = "Error on " & stepPath & ": " & Err.Description

End Function
