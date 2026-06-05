'==============================================================
' Module: TagBuilder
' Purpose: Generate DataSnipper DS_SEARCH and DS_COORDS tags
' Version: 1.0
' Last Updated: June 2026
'==============================================================
Option Explicit

Public Function BuildTags() As Boolean
    On Error GoTo ErrorHandler

    Dim tags() As TagDefinition
    Dim tagString As String
    Dim outputCol As Long
    Dim ws As Worksheet
    Dim headerMap As Object
    Dim rowIndex As Long
    Dim buildSuccess As Boolean

    buildSuccess = True
    If Not ConfigManager.IsConfigLoaded() Then
        MsgBox "Configuration must be loaded before building tags.", vbExclamation, "TagBuilder"
        BuildTags = False
        Exit Function
    End If

    tags = ConfigManager.GetTagDefinitions()
    If Not IsArrayAllocated(tags) Then
        MsgBox "No tag definitions were found in the configuration.", vbExclamation, "TagBuilder"
        BuildTags = False
        Exit Function
    End If

    If Not WorksheetExists("TAG_ENGINE") Then
        MsgBox "TAG_ENGINE sheet does not exist.", vbExclamation, "TagBuilder"
        BuildTags = False
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets("TAG_ENGINE")
    Set headerMap = BuildHeaderMap(ws)
    outputCol = EnsureColumn(ws, headerMap, "SourceTag")

    For rowIndex = LBound(tags) To UBound(tags)
        If Len(Trim$(tags(rowIndex).tag_id)) = 0 Then
            buildSuccess = False
            GoTo NextTagRow
        End If

        Select Case UCase$(Trim$(tags(rowIndex).extraction_method))
            Case "DS_SEARCH"
                tagString = BuildSearchTag(tags(rowIndex))
            Case "DS_COORDS"
                tagString = BuildCoordTag(tags(rowIndex))
            Case Else
                tagString = BuildSearchTag(tags(rowIndex))
        End Select

        If Not ValidateTagSyntax(tagString) Then
            buildSuccess = False
            ws.Cells(rowIndex + 1, outputCol).Value = "INVALID TAG"
        Else
            ws.Cells(rowIndex + 1, outputCol).Value = tagString
        End If
NextTagRow:
    Next rowIndex

    BuildTags = buildSuccess
    Exit Function

ErrorHandler:
    HandleError "TagBuilder", "BuildTags", Err.Number, Err.Description
    BuildTags = False
End Function

Public Function BuildSearchTag(tagDef As TagDefinition) As String
    On Error GoTo ErrorHandler

    Dim parts As String
    parts = "start=" & tagDef.start_anchor & "|end=" & tagDef.end_anchor & "|type=" & tagDef.field_type

    If Len(Trim$(tagDef.search_keywords)) > 0 Then
        parts = parts & "|keywords=" & tagDef.search_keywords
    End If

    BuildSearchTag = "DS_SEARCH:" & tagDef.tag_id & ":" & tagDef.field_name & ":(" & parts & ")"
    Exit Function

ErrorHandler:
    HandleError "TagBuilder", "BuildSearchTag", Err.Number, Err.Description
    BuildSearchTag = ""
End Function

Public Function BuildCoordTag(tagDef As TagDefinition) As String
    On Error GoTo ErrorHandler

    Dim parts As String
    parts = "page=" & CStr(tagDef.coord_page) & "|x=" & CStr(tagDef.coord_x) & "|y=" & CStr(tagDef.coord_y)
    parts = parts & "|width=" & CStr(tagDef.coord_width) & "|height=" & CStr(tagDef.coord_height)
    parts = parts & "|type=" & tagDef.field_type

    If Len(Trim$(tagDef.search_keywords)) > 0 Then
        parts = parts & "|keywords=" & tagDef.search_keywords
    End If

    BuildCoordTag = "DS_COORDS:" & tagDef.tag_id & ":" & tagDef.field_name & ":(" & parts & ")"
    Exit Function

ErrorHandler:
    HandleError "TagBuilder", "BuildCoordTag", Err.Number, Err.Description
    BuildCoordTag = ""
End Function

Public Function ValidateTagSyntax(tagString As String) As Boolean
    On Error GoTo ErrorHandler

    If Len(Trim$(tagString)) = 0 Then
        ValidateTagSyntax = False
        Exit Function
    End If

    If Left$(tagString, 9) = "DS_SEARCH:" Or Left$(tagString, 10) = "DS_COORDS:" Then
        ValidateTagSyntax = True
    Else
        ValidateTagSyntax = False
    End If
    Exit Function

ErrorHandler:
    HandleError "TagBuilder", "ValidateTagSyntax", Err.Number, Err.Description
    ValidateTagSyntax = False
End Function

Private Function EnsureColumn(ws As Worksheet, headerMap As Object, headerName As String) As Long
    If headerMap.Exists(headerName) Then
        EnsureColumn = headerMap(headerName)
    Else
        EnsureColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, EnsureColumn).Value = headerName
    End If
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

Private Function IsArrayAllocated(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = (Not IsEmpty(LBound(arr)))
    On Error GoTo 0
End Function

Private Sub HandleError(moduleName As String, functionName As String, errorCode As Long, errorMessage As String)
    Dim auditMessage As String
    auditMessage = moduleName & "." & functionName & " error " & CStr(errorCode) & ": " & errorMessage
    AuditLog.RecordError moduleName, functionName, errorCode, errorMessage
    MsgBox auditMessage, vbCritical, "TagBuilder Error"
End Sub
