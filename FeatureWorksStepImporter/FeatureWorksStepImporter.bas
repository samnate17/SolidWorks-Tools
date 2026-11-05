Attribute VB_Name = "FeatureWorksStepImporter"
Option Explicit

' ============================================================================
' FeatureWorksStepImporter
'
' Toolbar-button macro: pick one or more STEP files, import each one, and run
' FeatureWorks on it so the saved .SLDPRT shows real recognized features
' (Extrude1, Hole1, Fillet1, ...) instead of a single dumb "Imported1" body.
'
' IMPORTANT - how this actually drives FeatureWorks:
' The documented automation API (IFeatureWorksApp::RecognizeFeatureAutomatic
' + CreateFeatures) is well known to silently do nothing for a lot of people
' - it returns 0 recognized features even run exactly per SOLIDWORKS's own
' sample. This macro instead uses the sequence SOLIDWORKS's own macro
' recorder captures when you manually run Insert > FeatureWorks > Recognize
' Features with "Automatic" selected in the PropertyManager - an
' undocumented but verified-working route via ISldWorks::CallBack into
' "fworks@FWPlaybackManager". See RunFeatureWorksRecognition below.
'
' WHICH feature types get recognized (Standard features vs sheet metal,
' which specific types, Volume on/off, Draft never) is controlled by
' FeatureWorks's OWN remembered PropertyManager settings - not by this
' macro, because there's no verified way to set that from code. Do a
' one-time manual run first to lock in the settings you want. See README.
'
' Entry point for a toolbar button: RunFeatureWorksOnStepFiles
' ============================================================================

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

    ' Sanity check only - actual recognition below doesn't go through fwApp.
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

        resultMsg = ImportAndRecognizeStep(swApp, stepPath)
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
' Returns "" on success, or an error/warning description on failure.
Private Function ImportAndRecognizeStep(swApp As SldWorks.SldWorks, stepPath As String) As String

    Dim swImportStepData As SldWorks.ImportStepData
    Dim swModel As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim loadErrors As Long
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

    If swModel.GetType <> swDocumentTypes_e.swDocPART Then
        swApp.CloseDoc docTitle
        ImportAndRecognizeStep = "Skipped (not a single-body part): " & stepPath
        Exit Function
    End If

    Set swModelDocExt = swModel.Extension

    If Not RunFeatureWorksRecognition(swApp, swModel) Then
        swApp.CloseDoc docTitle
        ImportAndRecognizeStep = "Skipped (no imported body feature found to recognize): " & stepPath
        Exit Function
    End If

    dotPos = InStrRev(stepPath, ".")
    If dotPos > 0 Then
        savePath = Left$(stepPath, dotPos - 1) & ".SLDPRT"
    Else
        savePath = stepPath & ".SLDPRT"
    End If

    swModelDocExt.SaveAs savePath, 0, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, saveErrors, saveWarnings

    If saveErrors <> 0 Then
        swApp.CloseDoc docTitle
        ImportAndRecognizeStep = "Recognition ran but failed to save: " & savePath
        Exit Function
    End If

    If HasImportedFeature(swModel) Then
        swApp.CloseDoc docTitle
        ImportAndRecognizeStep = "Saved, but an Imported* feature still remains in the tree " & _
                                  "(partial or no recognition) - check manually: " & savePath
        Exit Function
    End If

    swApp.CloseDoc docTitle
    ImportAndRecognizeStep = "" ' success - no dumb imported body left in the tree

    Exit Function

HandleError:
    ImportAndRecognizeStep = "Error on " & stepPath & ": " & Err.Description

End Function

' Runs FeatureWorks on whatever SOLIDWORKS currently considers the active
' document, using the verified callback sequence described at the top of
' this file. Returns False only if there was no imported body feature to
' select in the first place (nothing to recognize).
Private Function RunFeatureWorksRecognition(swApp As Object, activeDoc As Object) As Boolean

    Dim feat As Object
    Dim retval As Long

    Set feat = activeDoc.FirstFeature
    If feat Is Nothing Then
        RunFeatureWorksRecognition = False
        Exit Function
    End If
    feat.Select2 False, -1 ' select the imported body - required, regardless of mode

    retval = swApp.CallBack("fworks@FWPlaybackManager", 0, "SetAutomaticFeatureRecognition 0")
    retval = swApp.CallBack("fworks@FWPlaybackManager", 0, "SetInteractiveFeatureRecognition 1")
    retval = swApp.CallBack("fworks@FWPlaybackManager", 0, "RecognizeFeatures")
    retval = swApp.CallBack("fworks@FWPlaybackManager", 0, "SetAutoDimOptions 0 0 0 0")
    retval = swApp.CallBack("fworks@FWPlaybackManager", 0, "SetAutoRelOptions 0")

    activeDoc.ClearSelection2 True
    activeDoc.ForceRebuild3 True

    RunFeatureWorksRecognition = True
End Function

' True if any feature in the tree is still a dumb imported body
' ("Imported1", "Imported2", ...) - i.e. recognition left something behind.
Private Function HasImportedFeature(doc As Object) As Boolean

    Dim feat As Object

    Set feat = doc.FirstFeature
    Do While Not feat Is Nothing
        If Left$(feat.Name, 8) = "Imported" Then
            HasImportedFeature = True
            Exit Function
        End If
        Set feat = feat.GetNextFeature
    Loop

    HasImportedFeature = False
End Function
