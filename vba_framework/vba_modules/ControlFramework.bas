'==============================================================
' Module: ControlFramework
' Purpose: Provide control checks and audit checkpoint logging
' Version: 1.0
' Last Updated: June 2026
'==============================================================
Option Explicit

Public Function RequireWorksheet(sheetName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim exists As Boolean
    exists = WorksheetExists(sheetName)

    If exists Then
        LogCheckpoint "RequireWorksheet", "PASS", "Worksheet '" & sheetName & "' is present."
        RequireWorksheet = True
    Else
        LogCheckpoint "RequireWorksheet", "FAIL", "Worksheet '" & sheetName & "' is missing."
        RequireWorksheet = False
    End If

    Exit Function

ErrorHandler:
    AuditLog.RecordError "ControlFramework", "RequireWorksheet", Err.Number, Err.Description
    RequireWorksheet = False
End Function

Public Function LogCheckpoint(checkpointName As String, status As String, description As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetAuditLogSheet()

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "CHECKPOINT"
    ws.Cells(nextRow, 3).Value = checkpointName
    ws.Cells(nextRow, 4).Value = status
    ws.Cells(nextRow, 5).Value = description
    ws.Cells(nextRow, 6).Value = Application.UserName

    LogCheckpoint = True
    Exit Function

ErrorHandler:
    LogCheckpoint = False
End Function

Private Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function GetAuditLogSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("AUDIT_LOG")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "AUDIT_LOG"
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "EntryType"
        ws.Cells(1, 3).Value = "Checkpoint"
        ws.Cells(1, 4).Value = "Status"
        ws.Cells(1, 5).Value = "Description"
        ws.Cells(1, 6).Value = "User"
    End If

    Set GetAuditLogSheet = ws
End Function