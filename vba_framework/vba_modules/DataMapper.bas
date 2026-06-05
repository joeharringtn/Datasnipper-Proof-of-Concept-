'==============================================================
' Module: DataMapper
' Purpose: Map raw extraction input into the OUTPUT sheet schema
' Version: 1.0
' Last Updated: June 2026
'==============================================================
Option Explicit

Public Function MapRawToSchema() As Boolean
    On Error GoTo ErrorHandler

    If Not OutputManager.PrepareOutput() Then
        MapRawToSchema = False
        Exit Function
    End If

    MapRawToSchema = True
    Exit Function

ErrorHandler:
    AuditLog.RecordError "DataMapper", "MapRawToSchema", Err.Number, Err.Description
    MapRawToSchema = False
End Function