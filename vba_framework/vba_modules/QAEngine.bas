'==============================================================
' Module: QAEngine
' Purpose: Execute QA rules against output data
' Version: 1.0
' Last Updated: June 2026
'==============================================================
Option Explicit

Public Function ApplyQARules() As Boolean
    On Error GoTo ErrorHandler

    Dim rules() As QARule
    Dim wsQA As Worksheet
    Dim wsOutput As Worksheet
    Dim qaHeader As Object
    Dim outputHeader As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim failCount As Long
    Dim ruleResult As String
    Dim resultValue As String

    If Not ConfigManager.IsConfigLoaded() Then
        If Not ConfigManager.LoadConfig() Then
            ApplyQARules = False
            Exit Function
        End If
    End If

    rules = ConfigManager.GetQARules()
    If Not IsArrayAllocated(rules) Then
        MsgBox "No QA rules are defined. Skipping QA evaluation.", vbInformation, "QAEngine"
        ApplyQARules = True
        Exit Function
    End If

    If Not WorksheetExists("OUTPUT") Then
        MsgBox "OUTPUT sheet is required for QA evaluation.", vbExclamation, "QAEngine"
        ApplyQARules = False
        Exit Function
    End If

    Set wsQA = ThisWorkbook.Worksheets("QA")
    Set wsOutput = ThisWorkbook.Worksheets("OUTPUT")
    Set qaHeader = BuildHeaderMap(wsQA)
    Set outputHeader = BuildHeaderMap(wsOutput)

    EnsureQAHeaders wsQA, qaHeader
    qaHeader = BuildHeaderMap(wsQA)

    lastRow = wsQA.Cells(wsQA.Rows.Count, qaHeader("RuleID")).End(xlUp).Row

    failCount = 0
    For rowIndex = 2 To lastRow
        ruleResult = EvaluateRule(wsOutput, outputHeader, wsQA, qaHeader, rowIndex)
        resultValue = ruleResult

        wsQA.Cells(rowIndex, qaHeader("Result")).Value = IIf(Left$(resultValue, 4) = "PASS", "PASS", "FAIL")
        wsQA.Cells(rowIndex, qaHeader("Comment")).Value = resultValue

        If Left$(resultValue, 4) <> "PASS" Then
            failCount = failCount + 1
        End If
    Next rowIndex

    ApplyQARules = (failCount = 0)
    Exit Function

ErrorHandler:
    HandleError "QAEngine", "ApplyQARules", Err.Number, Err.Description
    ApplyQARules = False
End Function

Private Function EvaluateRule(wsOutput As Worksheet, outputHeader As Object, wsQA As Worksheet, qaHeader As Object, ruleRow As Long) As String
    Dim ruleId As String
    Dim ruleType As String
    Dim fieldName As String
    Dim ruleDefinition As String
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim valueToCheck As Variant
    Dim valueString As String
    Dim minValue As Double
    Dim maxValue As Double
    Dim pattern As String
    Dim foundBlank As Boolean
    Dim duplicateCheck As Object

    ruleId = Trim$(CStr(wsQA.Cells(ruleRow, qaHeader("RuleID")).Value))
    fieldName = Trim$(CStr(wsQA.Cells(ruleRow, qaHeader("FieldName")).Value))
    ruleType = UCase$(Trim$(CStr(wsQA.Cells(ruleRow, qaHeader("RuleType")).Value)))
    ruleDefinition = Trim$(CStr(wsQA.Cells(ruleRow, qaHeader("RuleDefinition")).Value))

    If Len(ruleType) = 0 Then
        EvaluateRule = "FAIL: RuleType is missing."
        Exit Function
    End If

    If Not outputHeader.Exists(fieldName) Then
        EvaluateRule = "FAIL: Output field '" & fieldName & "' not found."
        Exit Function
    End If

    columnIndex = outputHeader(fieldName)

    Select Case ruleType
        Case "REQUIRED"
            foundBlank = False
            For rowIndex = 2 To wsOutput.Cells(wsOutput.Rows.Count, columnIndex).End(xlUp).Row
                If Len(Trim$(CStr(wsOutput.Cells(rowIndex, columnIndex).Value))) = 0 Then
                    foundBlank = True
                    Exit For
                End If
            Next rowIndex
            If foundBlank Then
                EvaluateRule = "FAIL: Required values are missing in " & fieldName
            Else
                EvaluateRule = "PASS: Required values present"
            End If

        Case "RANGE"
            If InStr(ruleDefinition, "|") = 0 Then
                EvaluateRule = "FAIL: RANGE definition must use min|max format."
                Exit Function
            End If
            minValue = Val(Trim$(Split(ruleDefinition, "|")(0)))
            maxValue = Val(Trim$(Split(ruleDefinition, "|")(1)))
            foundBlank = False
            For rowIndex = 2 To wsOutput.Cells(wsOutput.Rows.Count, columnIndex).End(xlUp).Row
                valueToCheck = wsOutput.Cells(rowIndex, columnIndex).Value
                If Len(Trim$(CStr(valueToCheck))) > 0 Then
                    If Not IsNumeric(valueToCheck) Then
                        EvaluateRule = "FAIL: Non-numeric value found in " & fieldName
                        Exit Function
                    End If
                    If CDbl(valueToCheck) < minValue Or CDbl(valueToCheck) > maxValue Then
                        EvaluateRule = "FAIL: " & fieldName & " out of range"
                        Exit Function
                    End If
                End If
            Next rowIndex
            EvaluateRule = "PASS: Range validation succeeded"

        Case "UNIQUE"
            Set duplicateCheck = CreateObject("Scripting.Dictionary")
            For rowIndex = 2 To wsOutput.Cells(wsOutput.Rows.Count, columnIndex).End(xlUp).Row
                valueString = Trim$(CStr(wsOutput.Cells(rowIndex, columnIndex).Value))
                If Len(valueString) > 0 Then
                    If duplicateCheck.Exists(valueString) Then
                        EvaluateRule = "FAIL: Duplicate value '" & valueString & "' found in " & fieldName
                        Exit Function
                    End If
                    duplicateCheck(valueString) = True
                End If
            Next rowIndex
            EvaluateRule = "PASS: Unique values confirmed"

        Case "FORMAT"
            pattern = ruleDefinition
            If Len(pattern) = 0 Then
                EvaluateRule = "FAIL: FORMAT rule requires a pattern."
                Exit Function
            End If
            For rowIndex = 2 To wsOutput.Cells(wsOutput.Rows.Count, columnIndex).End(xlUp).Row
                valueString = Trim$(CStr(wsOutput.Cells(rowIndex, columnIndex).Value))
                If Len(valueString) > 0 Then
                    If Not valueString Like pattern Then
                        EvaluateRule = "FAIL: Value '" & valueString & "' in " & fieldName & " does not match format"
                        Exit Function
                    End If
                End If
            Next rowIndex
            EvaluateRule = "PASS: Format validated"

        Case "CUSTOM"
            EvaluateRule = "REVIEW: Custom rule needs manual assessment"

        Case Else
            EvaluateRule = "FAIL: Unsupported rule type '" & ruleType & "'"
    End Select
End Function

Private Sub EnsureQAHeaders(ws As Worksheet, headerMap As Object)
    Dim requiredHeaders As Variant
    Dim headerName As Variant
    Dim nextColumn As Long

    requiredHeaders = Array("RuleID", "RuleName", "FieldName", "RuleType", "RuleDefinition", "FailAction", "Severity", "Priority", "Result", "Comment")
    nextColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1

    For Each headerName In requiredHeaders
        If Not headerMap.Exists(headerName) Then
            ws.Cells(1, nextColumn).Value = headerName
            headerMap(headerName) = nextColumn
            nextColumn = nextColumn + 1
        End If
    Next headerName
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
    MsgBox moduleName & "." & functionName & " error " & CStr(errorCode) & ": " & errorMessage, vbCritical, "QAEngine Error"
End Sub