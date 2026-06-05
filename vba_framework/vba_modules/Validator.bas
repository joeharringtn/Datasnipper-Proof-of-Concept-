'==============================================================
' Module: Validator
' Purpose: Validate workbook structure, configuration, and tag definitions
' Version: 1.0
' Last Updated: June 2026
'==============================================================
Option Explicit

Public Function ValidateInput() As Boolean
    On Error GoTo ErrorHandler

    Dim validationPassed As Boolean
    validationPassed = True

    Dim wsValidation As Worksheet
    Set wsValidation = GetOrCreateWorksheet("VALIDATION")
    wsValidation.Cells.Clear
    WriteValidationHeader wsValidation

    If Not ConfigManager.IsConfigLoaded() Then
        If Not ConfigManager.LoadConfig() Then
            validationPassed = False
            AppendValidation wsValidation, "CONFIG", "Config", "FAIL", "Failed to load configuration."
            ValidateInput = False
            Exit Function
        End If
    End If

    If Not ConfigManager.ValidateConfig() Then
        validationPassed = False
        AppendValidation wsValidation, "CONFIG", "Config", "FAIL", "Configuration validation failed."
    Else
        AppendValidation wsValidation, "CONFIG", "Config", "PASS", "Configuration validation succeeded."
    End If

    If Not ControlFramework.RequireWorksheet("TAG_ENGINE") Then
        validationPassed = False
        AppendValidation wsValidation, "Tag Engine", "TAG_ENGINE", "FAIL", "TAG_ENGINE worksheet is missing."
    ElseIf Not ValidateTagEngine(wsValidation) Then
        validationPassed = False
    Else
        AppendValidation wsValidation, "Tag Engine", "TAG_ENGINE", "PASS", "TAG_ENGINE definitions are valid."
    End If

    If Not ControlFramework.RequireWorksheet("EXTRACTION_INPUT") Then
        validationPassed = False
        AppendValidation wsValidation, "Extraction Input", "EXTRACTION_INPUT", "FAIL", "EXTRACTION_INPUT worksheet is missing."
    Else
        AppendValidation wsValidation, "Extraction Input", "EXTRACTION_INPUT", "PASS", "EXTRACTION_INPUT worksheet is present."
    End If

    If Not ControlFramework.RequireWorksheet("OUTPUT") Then
        AppendValidation wsValidation, "Output", "OUTPUT", "WARN", "OUTPUT sheet is not present; it will be created during extraction mapping."
    Else
        AppendValidation wsValidation, "Output", "OUTPUT", "PASS", "OUTPUT worksheet is present."
    End If

    ValidateInput = validationPassed
    Exit Function

ErrorHandler:
    HandleError "Validator", "ValidateInput", Err.Number, Err.Description
    ValidateInput = False
End Function

Private Function ValidateTagEngine(wsValidation As Worksheet) As Boolean
    Dim tags() As TagDefinition
    Dim rowIndex As Long
    Dim valid As Boolean
    Dim currentTag As TagDefinition

    valid = True
    tags = ConfigManager.GetTagDefinitions()

    If Not IsArrayAllocated(tags) Then
        AppendValidation wsValidation, "Tag Engine", "TAG_ENGINE", "FAIL", "No tag definitions were loaded."
        ValidateTagEngine = False
        Exit Function
    End If

    For rowIndex = LBound(tags) To UBound(tags)
        currentTag = tags(rowIndex)
        If Len(Trim$(currentTag.tag_id)) = 0 Then
            valid = False
            AppendValidation wsValidation, "Tag Engine", "Row " & CStr(rowIndex + 1), "FAIL", "TagID is required."
            GoTo NextValidationRow
        End If

        If Len(Trim$(currentTag.extraction_method)) = 0 And Len(Trim$(currentTag.tag_type)) = 0 Then
            valid = False
            AppendValidation wsValidation, "Tag Engine", currentTag.tag_id, "FAIL", "ExtractionMethod or TagType is required."
            GoTo NextValidationRow
        End If

        Select Case UCase$(Trim$(currentTag.extraction_method))
            Case "DS_SEARCH"
                If Len(Trim$(currentTag.start_anchor)) = 0 Or Len(Trim$(currentTag.end_anchor)) = 0 Then
                    valid = False
                    AppendValidation wsValidation, "Tag Engine", currentTag.tag_id, "FAIL", "DS_SEARCH tags require StartAnchor and EndAnchor."
                End If
            Case "DS_COORDS"
                If currentTag.coord_page <= 0 Or currentTag.coord_width <= 0 Or currentTag.coord_height <= 0 Then
                    valid = False
                    AppendValidation wsValidation, "Tag Engine", currentTag.tag_id, "FAIL", "DS_COORDS tags require page, width, and height."
                End If
            Case Else
                valid = False
                AppendValidation wsValidation, "Tag Engine", currentTag.tag_id, "FAIL", "ExtractionMethod must be DS_SEARCH or DS_COORDS."
        End Select
NextValidationRow:
    Next rowIndex

    ValidateTagEngine = valid
End Function

Private Sub WriteValidationHeader(ws As Worksheet)
    ws.Cells(1, 1).Value = "CheckID"
    ws.Cells(1, 2).Value = "Category"
    ws.Cells(1, 3).Value = "Result"
    ws.Cells(1, 4).Value = "Message"
    ws.Cells(1, 5).Value = "Sheet"
End Sub

Private Sub AppendValidation(ws As Worksheet, category As String, sheetName As String, result As String, message As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = "CHK-" & Format(lastRow - 1, "000")
    ws.Cells(lastRow, 2).Value = category
    ws.Cells(lastRow, 3).Value = result
    ws.Cells(lastRow, 4).Value = message
    ws.Cells(lastRow, 5).Value = sheetName
End Sub

Private Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateWorksheet = ws
End Function

Private Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function IsArrayAllocated(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = (Not IsEmpty(LBound(arr)))
    On Error GoTo 0
End Function

Private Sub HandleError(moduleName As String, functionName As String, errorCode As Long, errorMessage As String)
    AuditLog.RecordError moduleName, functionName, errorCode, errorMessage
    MsgBox moduleName & "." & functionName & " error " & CStr(errorCode) & ": " & errorMessage, vbCritical, "Validator Error"
End Sub