$ErrorActionPreference = 'Stop'

$basePath = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
Set-Location $basePath\..\

$templates = @(
    @{Path = 'templates\config_template.xlsx'; Sheet = 'CONFIG'; Headers = @('Key','Value','DataType','Scope','Description','Required')},
    @{Path = 'templates\tag_engine_template.xlsx'; Sheet = 'TAG_ENGINE'; Headers = @('TagID','TagType','ExtractionMethod','FieldName','SourceDocument','FieldType','Required','SearchKeywords','StartAnchor','EndAnchor','CoordPage','CoordX','CoordY','CoordWidth','CoordHeight','Tolerance','FallbackKeywords','Notes')},
    @{Path = 'templates\coord_reference_template.xlsx'; Sheet = 'COORD_REFERENCE'; Headers = @('CoordID','WorksheetName','StartCell','EndCell','LookupKey','CoordType','Usage','Notes')},
    @{Path = 'templates\qa_template.xlsx'; Sheet = 'QA'; Headers = @('RuleID','RuleName','FieldName','RuleType','RuleDefinition','FailAction','Severity','Priority','Result','Comment')},
    @{Path = 'templates\output_template.xlsx'; Sheet = 'OUTPUT'; Headers = @('OutputID','SourceTagID','ExtractedValue','ExtractedCell','ValidationStatus','Reviewer','FinalComment','Timestamp')}
)

if (-not (Test-Path 'templates')) {
    New-Item -ItemType Directory -Path 'templates' | Out-Null
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

try {
    foreach ($entry in $templates) {
        $fullPath = Join-Path (Get-Location) $entry.Path
        if (Test-Path $fullPath) {
            Remove-Item $fullPath -Force
        }

        $workbook = $excel.Workbooks.Add()
        while ($workbook.Sheets.Count -gt 1) {
            $workbook.Sheets.Item(1).Delete()
        }

        $sheet = $workbook.Sheets.Item(1)
        $sheet.Name = $entry.Sheet
        for ($i = 0; $i -lt $entry.Headers.Count; $i++) {
            $sheet.Cells.Item(1, $i + 1).Value2 = $entry.Headers[$i]
        }

        $workbook.SaveAs($fullPath, 51)
        $workbook.Close($false)
    }
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Output 'Excel templates created successfully.'
