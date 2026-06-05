'==============================================================
' Module: OutputManager
' Purpose: Normalize extraction results into output deliverables
' Version: 1.0
' Last Updated: June 2026
'==============================================================
Option Explicit

Public Function PrepareOutput() As Boolean
    On Error GoTo ErrorHandler

    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim inputHeader As Object
    Dim outputHeader As Object
    Dim lastInputRow As Long
    Dim lastOutputRow As Long
    Dim rowIndex As Long
    Dim outputRow As Long
    Dim tagId As String
    Dim extractedValue As Variant
    Dim extractedCell As String

    If Not WorksheetExists("EXTRACTION_INPUT") Then
        MsgBox "EXTRACTION_INPUT sheet is missing. Populate raw extraction values before output generation.", vbExclamation, "OutputManager"
        PrepareOutput = False
        Exit Function
    End If

    Set wsInput = ThisWorkbook.Worksheets("EXTRACTION_INPUT")
    Set wsOutput = GetOrCreateWorksheet("OUTPUT")
    Set inputHeader = BuildHeaderMap(wsInput)
    Set outputHeader = BuildHeaderMap(wsOutput)

    EnsureOutputHeaders wsOutput
    lastInputRow = wsInput.Cells(wsInput.Rows.Count, inputHeader("TagID")).End(xlUp).Row
    wsOutput.Rows("2:" & wsOutput.Rows.Count).ClearContents

    outputRow = 2
    For rowIndex = 2 To lastInputRow
        tagId = Trim$(CStr(wsInput.Cells(rowIndex, inputHeader("TagID")).Value))
        If Len(tagId) = 0 Then GoTo NextInputRow

        extractedValue = wsInput.Cells(rowIndex, inputHeader("ExtractedValue")).Value
        extractedCell = Trim$(CStr(wsInput.Cells(rowIndex, inputHeader("ExtractedCell")).Value))

        wsOutput.Cells(outputRow, outputHeader("OutputID")).Value = "OUT-" & Format(outputRow - 1, "000")
        wsOutput.Cells(outputRow, outputHeader("SourceTagID")).Value = tagId
        wsOutput.Cells(outputRow, outputHeader("ExtractedValue")).Value = extractedValue
        wsOutput.Cells(outputRow, outputHeader("ExtractedCell")).Value = extractedCell
        wsOutput.Cells(outputRow, outputHeader("ValidationStatus")).Value = "Pending"
        wsOutput.Cells(outputRow, outputHeader("Reviewer")).Value = ""
        wsOutput.Cells(outputRow, outputHeader("FinalComment")).Value = ""
        wsOutput.Cells(outputRow, outputHeader("Timestamp")).Value = Now

        outputRow = outputRow + 1
NextInputRow:
    Next rowIndex

    PrepareOutput = True
    Exit Function

ErrorHandler:
    HandleError "OutputManager", "PrepareOutput", Err.Number, Err.Description
    PrepareOutput = False
End Function

Public Function FinalizeOutput(approvedBy As String) As Boolean
    On Error GoTo ErrorHandler

    Dim wsOutput As Worksheet
    Dim outputHeader As Object
    Dim lastRow As Long
    Dim rowIndex As Long

    If Not WorksheetExists("OUTPUT") Then
        MsgBox "OUTPUT sheet is missing. Cannot finalize.", vbExclamation, "OutputManager"
        FinalizeOutput = False
        Exit Function
    End If

    Set wsOutput = ThisWorkbook.Worksheets("OUTPUT")
    Set outputHeader = BuildHeaderMap(wsOutput)

    If Not outputHeader.Exists("ValidationStatus") Then
        MsgBox "OUTPUT worksheet does not contain ValidationStatus column.", vbExclamation, "OutputManager"
        FinalizeOutput = False
        Exit Function
    End If

    lastRow = wsOutput.Cells(wsOutput.Rows.Count, outputHeader("OutputID")).End(xlUp).Row
    For rowIndex = 2 To lastRow
        wsOutput.Cells(rowIndex, outputHeader("ValidationStatus")).Value = "Approved"
        If outputHeader.Exists("Reviewer") Then
            wsOutput.Cells(rowIndex, outputHeader("Reviewer")).Value = approvedBy
        End If
    Next rowIndex

    AuditLog.RecordSignOff approvedBy, "FinalizeOutput", "Approved"
    FinalizeOutput = True
    Exit Function

ErrorHandler:
    HandleError "OutputManager", "FinalizeOutput", Err.Number, Err.Description
    FinalizeOutput = False
End Function

Private Sub EnsureOutputHeaders(ws As Worksheet)
    Dim requiredHeaders As Variant
    Dim i As Long
    Dim headerName As String
    Dim headerMap As Object

    requiredHeaders = Array("OutputID", "SourceTagID", "ExtractedValue", "ExtractedCell", "ValidationStatus", "Reviewer", "FinalComment", "Timestamp")
    Set headerMap = BuildHeaderMap(ws)

    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        headerName = requiredHeaders(i)
        If Not headerMap.Exists(headerName) Then
            ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1).Value = headerName
            headerMap(headerName) = ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column).Column
        End If
    Next i
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

Private Sub HandleError(moduleName As String, functionName As String, errorCode As Long, errorMessage As String)
    AuditLog.RecordError moduleName, functionName, errorCode, errorMessage
    MsgBox moduleName & "." & functionName & " error " & CStr(errorCode) & ": " & errorMessage, vbCritical, "OutputManager Error"
End Sub