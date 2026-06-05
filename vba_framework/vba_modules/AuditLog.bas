'==============================================================
' Module: AuditLog
' Purpose: Persist framework errors, approvals, and checkpoints
' Version: 1.0
' Last Updated: June 2026
'==============================================================
Option Explicit

Public Sub RecordError(moduleName As String, functionName As String, errorCode As Long, errorMessage As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetAuditLogSheet()

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "ERROR"
    ws.Cells(nextRow, 3).Value = moduleName
    ws.Cells(nextRow, 4).Value = functionName
    ws.Cells(nextRow, 5).Value = errorCode
    ws.Cells(nextRow, 6).Value = errorMessage
    ws.Cells(nextRow, 7).Value = Application.UserName

    Exit Sub

ErrorHandler:
    Resume Next
End Sub

Public Function RecordSignOff(requestor As String, action As String, status As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetAuditLogSheet()

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "SIGNOFF"
    ws.Cells(nextRow, 3).Value = requestor
    ws.Cells(nextRow, 4).Value = action
    ws.Cells(nextRow, 5).Value = status
    ws.Cells(nextRow, 6).Value = ""
    ws.Cells(nextRow, 7).Value = Application.UserName

    RecordSignOff = True
    Exit Function

ErrorHandler:
    RecordSignOff = False
End Function

Public Function GenerateAuditTrail() As Boolean
    On Error GoTo ErrorHandler
    GenerateAuditTrail = True
    Exit Function

ErrorHandler:
    GenerateAuditTrail = False
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
        ws.Cells(1, 3).Value = "Module"
        ws.Cells(1, 4).Value = "Action"
        ws.Cells(1, 5).Value = "CodeOrStatus"
        ws.Cells(1, 6).Value = "Details"
        ws.Cells(1, 7).Value = "User"
    End If

    Set GetAuditLogSheet = ws
End Function