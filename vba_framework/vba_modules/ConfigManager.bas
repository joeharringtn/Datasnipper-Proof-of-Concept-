'==============================================================
' Module: ConfigManager
' Purpose: Load and validate engagement configuration
' Version: 1.0
' Last Updated: June 2026
'==============================================================
Option Explicit

Public Type DocumentSpec
    doc_id As String
    file_path As String
    page_count As Long
    doc_type As String
    extraction_method As String
End Type

Public Type TagDefinition
    tag_id As String
    tag_type As String
    extraction_method As String
    field_name As String
    source_document As String
    field_type As String
    required As Boolean
    search_keywords As String
    start_anchor As String
    end_anchor As String
    coord_page As Long
    coord_x As Long
    coord_y As Long
    coord_width As Long
    coord_height As Long
    tolerance As Long
    fallback_keywords As String
    notes As String
End Type

Public Type QARule
    rule_id As String
    rule_name As String
    field_name As String
    rule_type As String
    rule_definition As String
    fail_action As String
    severity As String
    priority As Long
End Type

Public Type SchemaField
    field_name As String
    data_type As String
    required As Boolean
    format_spec As String
End Type

Public Type ApprovalWorkflow
    workflow_type As String
    level1_approver As String
    level1_email As String
    level2_approver As String
    level2_email As String
End Type

Public Type ConfigObject
    engagement_id As String
    engagement_type As String
    period_start_date As Date
    period_end_date As Date
    lead_auditor As String
    client_name As String
    client_contact As String
    framework_version As String
    documents() As DocumentSpec
    tags() As TagDefinition
    qa_rules() As QARule
    output_schema() As SchemaField
    approval_workflow As ApprovalWorkflow
End Type

Private gConfig As ConfigObject
Private gConfigLoaded As Boolean

Public Function LoadConfig() As Boolean
    On Error GoTo ErrorHandler

    If Not WorksheetExists("CONFIG") Then
        MsgBox "The CONFIG sheet is missing. Create a sheet named CONFIG and retry.", vbExclamation, "ConfigManager"
        LoadConfig = False
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CONFIG")

    gConfig.engagement_id = Trim$(CStr(ReadConfigValue(ws, "EngagementID")))
    gConfig.engagement_type = Trim$(CStr(ReadConfigValue(ws, "EngagementType")))
    gConfig.period_start_date = CDate(ReadConfigValue(ws, "PeriodStartDate"))
    gConfig.period_end_date = CDate(ReadConfigValue(ws, "PeriodEndDate"))
    gConfig.lead_auditor = Trim$(CStr(ReadConfigValue(ws, "LeadAuditor")))
    gConfig.client_name = Trim$(CStr(ReadConfigValue(ws, "ClientName")))
    gConfig.client_contact = Trim$(CStr(ReadConfigValue(ws, "ClientContact")))
    gConfig.framework_version = Trim$(CStr(ReadConfigValue(ws, "FrameworkVersion")))

    If Not LoadTagDefinitions() Then
        LoadConfig = False
        Exit Function
    End If

    If Not LoadQARules() Then
        LoadConfig = False
        Exit Function
    End If

    LoadOutputSchema
    LoadApprovalWorkflow

    gConfigLoaded = True
    LoadConfig = True
    Exit Function

ErrorHandler:
    HandleError "ConfigManager", "LoadConfig", Err.Number, Err.Description
    LoadConfig = False
End Function

Public Function ValidateConfig() As Boolean
    On Error GoTo ErrorHandler

    Dim valid As Boolean
    valid = True

    If Not gConfigLoaded Then
        MsgBox "Configuration must be loaded before validation.", vbExclamation, "ConfigManager"
        ValidateConfig = False
        Exit Function
    End If

    If Len(gConfig.engagement_id) = 0 Then
        valid = False
        MsgBox "CONFIG: EngagementID is required.", vbExclamation, "ConfigManager"
    End If

    If Len(gConfig.engagement_type) = 0 Then
        valid = False
        MsgBox "CONFIG: EngagementType is required.", vbExclamation, "ConfigManager"
    End If

    If gConfig.period_start_date <= 0 Or gConfig.period_end_date <= 0 Then
        valid = False
        MsgBox "CONFIG: PeriodStartDate and PeriodEndDate must be populated.", vbExclamation, "ConfigManager"
    ElseIf gConfig.period_start_date > gConfig.period_end_date Then
        valid = False
        MsgBox "CONFIG: PeriodStartDate must be before PeriodEndDate.", vbExclamation, "ConfigManager"
    End If

    If Len(gConfig.lead_auditor) = 0 Then
        valid = False
        MsgBox "CONFIG: LeadAuditor is recommended for audit traceability.", vbInformation, "ConfigManager"
    End If

    If Len(gConfig.client_name) = 0 Then
        valid = False
        MsgBox "CONFIG: ClientName is required.", vbExclamation, "ConfigManager"
    End If

    If Len(gConfig.client_contact) = 0 Then
        valid = False
        MsgBox "CONFIG: ClientContact is recommended.", vbInformation, "ConfigManager"
    End If

    ValidateConfig = valid
    Exit Function

ErrorHandler:
    HandleError "ConfigManager", "ValidateConfig", Err.Number, Err.Description
    ValidateConfig = False
End Function

Public Function IsConfigLoaded() As Boolean
    IsConfigLoaded = gConfigLoaded
End Function

Public Function GetEngagementType() As String
    If gConfigLoaded Then
        GetEngagementType = gConfig.engagement_type
    Else
        GetEngagementType = ""
    End If
End Function

Public Function GetDocSpecs() As DocumentSpec()
    GetDocSpecs = gConfig.documents
End Function

Public Function GetTagDefinitions() As TagDefinition()
    GetTagDefinitions = gConfig.tags
End Function

Public Function GetQARules() As QARule()
    GetQARules = gConfig.qa_rules
End Function

Public Function GetOutputSchema() As SchemaField()
    GetOutputSchema = gConfig.output_schema
End Function

Public Function GetApprovalWorkflow() As ApprovalWorkflow
    GetApprovalWorkflow = gConfig.approval_workflow
End Function

Private Function ReadConfigValue(ws As Worksheet, key As String) As Variant
    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For rowIndex = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, 1).Value)), key, vbTextCompare) = 0 Then
            ReadConfigValue = ws.Cells(rowIndex, 2).Value
            Exit Function
        End If
    Next rowIndex

    ReadConfigValue = vbNullString
End Function

Private Function LoadTagDefinitions() As Boolean
    Dim ws As Worksheet
    Dim headerMap As Object
    Dim rowIndex As Long
    Dim currentTag As TagDefinition
    Dim sourceValue As Variant

    If Not WorksheetExists("TAG_ENGINE") Then
        MsgBox "The TAG_ENGINE sheet is missing. Create a sheet named TAG_ENGINE and retry.", vbExclamation, "ConfigManager"
        LoadTagDefinitions = False
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets("TAG_ENGINE")
    Set headerMap = BuildHeaderMap(ws)

    If Not headerMap.Exists("TagID") Then
        MsgBox "TAG_ENGINE sheet must include a TagID column.", vbExclamation, "ConfigManager"
        LoadTagDefinitions = False
        Exit Function
    End If

    ReDim gConfig.tags(1 To 0)
    For rowIndex = 2 To ws.Cells(ws.Rows.Count, headerMap("TagID")).End(xlUp).Row
        sourceValue = Trim$(CStr(ws.Cells(rowIndex, headerMap("TagID")).Value))
        If Len(sourceValue) = 0 Then GoTo NextTagRow

        currentTag.tag_id = sourceValue
        currentTag.tag_type = Trim$(CStr(GetCellValue(ws, headerMap, "TagType", rowIndex)))
        currentTag.extraction_method = Trim$(CStr(GetCellValue(ws, headerMap, "ExtractionMethod", rowIndex)))
        currentTag.field_name = Trim$(CStr(GetCellValue(ws, headerMap, "FieldName", rowIndex)))
        currentTag.source_document = Trim$(CStr(GetCellValue(ws, headerMap, "SourceDocument", rowIndex)))
        currentTag.field_type = Trim$(CStr(GetCellValue(ws, headerMap, "FieldType", rowIndex)))
        currentTag.required = IsYes(Trim$(CStr(GetCellValue(ws, headerMap, "Required", rowIndex))))
        currentTag.search_keywords = Trim$(CStr(GetCellValue(ws, headerMap, "SearchKeywords", rowIndex)))
        currentTag.start_anchor = Trim$(CStr(GetCellValue(ws, headerMap, "StartAnchor", rowIndex)))
        currentTag.end_anchor = Trim$(CStr(GetCellValue(ws, headerMap, "EndAnchor", rowIndex)))
        currentTag.coord_page = CLng(Val(GetCellValue(ws, headerMap, "CoordPage", rowIndex)))
        currentTag.coord_x = CLng(Val(GetCellValue(ws, headerMap, "CoordX", rowIndex)))
        currentTag.coord_y = CLng(Val(GetCellValue(ws, headerMap, "CoordY", rowIndex)))
        currentTag.coord_width = CLng(Val(GetCellValue(ws, headerMap, "CoordWidth", rowIndex)))
        currentTag.coord_height = CLng(Val(GetCellValue(ws, headerMap, "CoordHeight", rowIndex)))
        currentTag.tolerance = CLng(Val(GetCellValue(ws, headerMap, "Tolerance", rowIndex)))
        currentTag.fallback_keywords = Trim$(CStr(GetCellValue(ws, headerMap, "FallbackKeywords", rowIndex)))
        currentTag.notes = Trim$(CStr(GetCellValue(ws, headerMap, "Notes", rowIndex)))

        ReDim Preserve gConfig.tags(1 To UBound(gConfig.tags) + 1)
        gConfig.tags(UBound(gConfig.tags)) = currentTag
NextTagRow:
    Next rowIndex

    LoadTagDefinitions = True
End Function

Private Function LoadQARules() As Boolean
    Dim ws As Worksheet
    Dim headerMap As Object
    Dim rowIndex As Long
    Dim currentRule As QARule
    Dim sourceValue As Variant

    If Not WorksheetExists("QA") Then
        ReDim gConfig.qa_rules(1 To 0)
        LoadQARules = True
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets("QA")
    Set headerMap = BuildHeaderMap(ws)

    If Not headerMap.Exists("RuleID") Then
        ReDim gConfig.qa_rules(1 To 0)
        LoadQARules = True
        Exit Function
    End If

    ReDim gConfig.qa_rules(1 To 0)
    For rowIndex = 2 To ws.Cells(ws.Rows.Count, headerMap("RuleID")).End(xlUp).Row
        sourceValue = Trim$(CStr(ws.Cells(rowIndex, headerMap("RuleID")).Value))
        If Len(sourceValue) = 0 Then GoTo NextRuleRow

        currentRule.rule_id = sourceValue
        currentRule.rule_name = Trim$(CStr(GetCellValue(ws, headerMap, "RuleName", rowIndex)))
        currentRule.field_name = Trim$(CStr(GetCellValue(ws, headerMap, "FieldName", rowIndex)))
        currentRule.rule_type = Trim$(CStr(GetCellValue(ws, headerMap, "RuleType", rowIndex)))
        currentRule.rule_definition = Trim$(CStr(GetCellValue(ws, headerMap, "RuleDefinition", rowIndex)))
        currentRule.fail_action = Trim$(CStr(GetCellValue(ws, headerMap, "FailAction", rowIndex)))
        currentRule.severity = Trim$(CStr(GetCellValue(ws, headerMap, "Severity", rowIndex)))
        currentRule.priority = CLng(Val(GetCellValue(ws, headerMap, "Priority", rowIndex)))

        ReDim Preserve gConfig.qa_rules(1 To UBound(gConfig.qa_rules) + 1)
        gConfig.qa_rules(UBound(gConfig.qa_rules)) = currentRule
NextRuleRow:
    Next rowIndex

    LoadQARules = True
End Function

Private Sub LoadApprovalWorkflow()
    Dim ws As Worksheet
    If Not WorksheetExists("CONFIG") Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("CONFIG")
    gConfig.approval_workflow.workflow_type = Trim$(CStr(ReadConfigValue(ws, "WorkflowType")))
    gConfig.approval_workflow.level1_approver = Trim$(CStr(ReadConfigValue(ws, "Level1Approver")))
    gConfig.approval_workflow.level1_email = Trim$(CStr(ReadConfigValue(ws, "Level1Email")))
    gConfig.approval_workflow.level2_approver = Trim$(CStr(ReadConfigValue(ws, "Level2Approver")))
    gConfig.approval_workflow.level2_email = Trim$(CStr(ReadConfigValue(ws, "Level2Email")))
End Sub

Private Sub LoadOutputSchema()
    Dim ws As Worksheet
    Dim headerMap As Object
    Dim key As Variant
    Dim currentSchema As SchemaField

    If Not WorksheetExists("OUTPUT") Then
        ReDim gConfig.output_schema(1 To 0)
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("OUTPUT")
    Set headerMap = BuildHeaderMap(ws)

    ReDim gConfig.output_schema(1 To 0)
    For Each key In headerMap.Keys
        currentSchema.field_name = CStr(key)
        currentSchema.data_type = "String"
        currentSchema.required = False
        currentSchema.format_spec = ""
        ReDim Preserve gConfig.output_schema(1 To UBound(gConfig.output_schema) + 1)
        gConfig.output_schema(UBound(gConfig.output_schema)) = currentSchema
    Next key
End Sub

Private Function BuildHeaderMap(ws As Worksheet) As Object
    Dim headerMap As Object
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerName As String

    Set headerMap = CreateObject("Scripting.Dictionary")
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For columnIndex = 1 To lastColumn
        headerName = Trim$(CStr(ws.Cells(1, columnIndex).Value))
        If Len(headerName) > 0 Then
            headerMap(headerName) = columnIndex
        End If
    Next columnIndex

    Set BuildHeaderMap = headerMap
End Function

Private Function GetCellValue(ws As Worksheet, headerMap As Object, headerName As String, rowIndex As Long) As Variant
    If headerMap.Exists(headerName) Then
        GetCellValue = ws.Cells(rowIndex, headerMap(headerName)).Value
    Else
        GetCellValue = vbNullString
    End If
End Function

Private Function IsYes(value As String) As Boolean
    Dim normalized As String
    normalized = Trim$(UCase$(value))
    IsYes = (normalized = "YES" Or normalized = "TRUE" Or normalized = "1")
End Function

Private Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function IsValidEmail(emailAddress As String) As Boolean
    IsValidEmail = (InStr(1, emailAddress, "@", vbTextCompare) > 1 And InStrRev(emailAddress, ".") > InStr(1, emailAddress, "@", vbTextCompare))
End Function

Private Sub HandleError(moduleName As String, functionName As String, errorCode As Long, errorMessage As String)
    Dim auditMessage As String
    auditMessage = moduleName & "." & functionName & " error " & CStr(errorCode) & ": " & errorMessage
    AuditLog.RecordError moduleName, functionName, errorCode, errorMessage
    MsgBox auditMessage, vbCritical, "ConfigManager Error"
End Sub
