'==============================================================
' Module: Main
' Purpose: Workflow orchestration and user entry points
' Version: 1.0
' Last Updated: June 2026
'==============================================================
Option Explicit

Public Enum WorkflowState
    STATE_IDLE = 0
    STATE_CONFIG_LOADED = 1
    STATE_TAGS_READY = 2
    STATE_EXTRACTION_SUBMITTED = 3
    STATE_VALIDATION_COMPLETE = 4
    STATE_QA_IN_PROGRESS = 5
    STATE_QA_REVIEWED = 6
    STATE_COMPLETE = 7
    STATE_ARCHIVED = 8
End Enum

Private gCurrentState As WorkflowState

Public Sub InitializeFramework()
    gCurrentState = STATE_IDLE
End Sub

Public Function GetCurrentState() As WorkflowState
    GetCurrentState = gCurrentState
End Function

Public Sub LoadEngagement()
    On Error GoTo ErrorHandler
    If gCurrentState = 0 Then InitializeFramework

    Dim configLoaded As Boolean
    configLoaded = ConfigManager.LoadConfig()

    If configLoaded Then
        gCurrentState = STATE_CONFIG_LOADED
        MsgBox "Engagement configuration loaded successfully.", vbInformation, "Framework"
    Else
        MsgBox "Failed to load engagement configuration. Check the CONFIG sheet.", vbExclamation, "Framework"
    End If
    Exit Sub

ErrorHandler:
    HandleError "Main", "LoadEngagement", Err.Number, Err.Description
End Sub

Public Sub BuildTags()
    On Error GoTo ErrorHandler
    If Not CanProceed(STATE_CONFIG_LOADED) Then Exit Sub

    Dim result As Boolean
    result = TagBuilder.BuildTags()

    If result Then
        gCurrentState = STATE_TAGS_READY
        MsgBox "Tags generated successfully. Review TAG_ENGINE.SourceTag.", vbInformation, "Framework"
    Else
        MsgBox "Tag generation failed. Review the TAG_ENGINE sheet for errors.", vbExclamation, "Framework"
    End If
    Exit Sub

ErrorHandler:
    HandleError "Main", "BuildTags", Err.Number, Err.Description
End Sub

Public Sub ValidateInput()
    On Error GoTo ErrorHandler
    If Not CanProceed(STATE_TAGS_READY) Then Exit Sub

    Dim validationPassed As Boolean
    validationPassed = Validator.ValidateInput()

    If validationPassed Then
        gCurrentState = STATE_VALIDATION_COMPLETE
        MsgBox "Input validation passed. Ready to process extraction.", vbInformation, "Framework"
    Else
        MsgBox "Input validation completed with issues. Review VALIDATION sheet.", vbExclamation, "Framework"
    End If
    Exit Sub

ErrorHandler:
    HandleError "Main", "ValidateInput", Err.Number, Err.Description
End Sub

Public Sub ProcessExtraction()
    On Error GoTo ErrorHandler
    If Not CanProceed(STATE_VALIDATION_COMPLETE) Then Exit Sub

    Dim processResult As Boolean
    processResult = DataMapper.MapRawToSchema()

    If processResult Then
        Dim qaPassed As Boolean
        qaPassed = QAEngine.ApplyQARules()
        gCurrentState = STATE_QA_IN_PROGRESS

        If qaPassed Then
            MsgBox "Extraction processed successfully and QA checks completed.", vbInformation, "Framework"
        Else
            MsgBox "Extraction processed, but QA found issues. Review QA sheet.", vbExclamation, "Framework"
        End If
    Else
        MsgBox "Extraction processing found issues. Review the QA sheet.", vbExclamation, "Framework"
    End If
    Exit Sub

ErrorHandler:
    HandleError "Main", "ProcessExtraction", Err.Number, Err.Description
End Sub

Public Sub ReviewExceptions()
    On Error GoTo ErrorHandler
    If Not CanProceed(STATE_QA_IN_PROGRESS) Then Exit Sub

    MsgBox "Please review QA exceptions in the QA sheet, then mark decisions.", vbInformation, "Framework"
    Exit Sub

ErrorHandler:
    HandleError "Main", "ReviewExceptions", Err.Number, Err.Description
End Sub

Public Sub ApproveOutput()
    On Error GoTo ErrorHandler
    If Not CanProceed(STATE_QA_IN_PROGRESS) Then Exit Sub

    Dim approvalOkay As Boolean
    approvalOkay = AuditLog.RecordSignOff("System", "Approval", "Approved")

    If approvalOkay Then
        OutputManager.FinalizeOutput "System"
        gCurrentState = STATE_COMPLETE
        MsgBox "Output approved and workflow completed.", vbInformation, "Framework"
    Else
        MsgBox "Approval failed. Review the audit log and retry.", vbExclamation, "Framework"
    End If
    Exit Sub

ErrorHandler:
    HandleError "Main", "ApproveOutput", Err.Number, Err.Description
End Sub

Public Sub ExportResults()
    On Error GoTo ErrorHandler
    If gCurrentState <> STATE_COMPLETE Then
        MsgBox "Final approval is required before exporting results.", vbExclamation, "Framework"
        Exit Sub
    End If

    MsgBox "Exporting results is not implemented in this initial version.", vbInformation, "Framework"
    Exit Sub

ErrorHandler:
    HandleError "Main", "ExportResults", Err.Number, Err.Description
End Sub

Public Sub GenerateReport()
    On Error GoTo ErrorHandler
    If gCurrentState < STATE_COMPLETE Then
        MsgBox "Please complete the workflow before generating the audit report.", vbExclamation, "Framework"
        Exit Sub
    End If

    MsgBox "Audit report generation is not implemented in this initial version.", vbInformation, "Framework"
    Exit Sub

ErrorHandler:
    HandleError "Main", "GenerateReport", Err.Number, Err.Description
End Sub

Private Function CanProceed(requiredState As WorkflowState) As Boolean
    CanProceed = (gCurrentState >= requiredState)
    If Not CanProceed Then
        MsgBox "Workflow is not in the correct state to perform this action.", vbExclamation, "Framework"
    End If
End Function

Private Sub HandleError(moduleName As String, functionName As String, errorCode As Long, errorMessage As String)
    Dim auditMessage As String
    auditMessage = moduleName & "." & functionName & " error " & CStr(errorCode) & ": " & errorMessage
    AuditLog.RecordError moduleName, functionName, errorCode, errorMessage
    MsgBox auditMessage, vbCritical, "Framework Error"
End Sub
