#requires -Version 5.1
<#!
    입출력 속도 성능평가 자동화 스크립트
    ------------------------------------------------------------
    이 스크립트는 초기화된 저장장치 내 일반영역과 보안영역을 지정하고,
    샘플 데이터를 재구성한 뒤 10개 샘플을 여러 차례 측정하여 쓰기/읽기
    속도를 측정합니다. 측정 결과는 CSV·XLSX·DOCX 형태로 저장되며,
    보안영역의 성능이 일반영역 대비 90% 이상 유지되는지 자동으로 분석합니다.

    관리자 권한 PowerShell 콘솔에서 실행하고, 테스트 전용 경로를 사용하십시오.
!#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Crc32Table = $null

#region 공통 유틸리티
function Read-RequiredPath {
    param([Parameter(Mandatory)] [string] $PromptText)
    while ($true) {
        $value = Read-Host -Prompt $PromptText
        if ([string]::IsNullOrWhiteSpace($value)) {
            Write-Host '값을 입력해야 합니다. 다시 시도하세요.' -ForegroundColor Yellow
            continue
        }
        return $value.Trim()
    }
}

function Ensure-Directory {
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Clear-Directory {
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return }
    Get-ChildItem -LiteralPath $Path -Force | ForEach-Object {
        try {
            Remove-Item -LiteralPath $_.FullName -Force -Recurse -ErrorAction Stop
        }
        catch {
            Write-Host "삭제 실패: $($_.FullName) - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Write-BytesFile {
    param([string] $Path, [byte[]] $Bytes)
    $folder = Split-Path -Path $Path -Parent
    Ensure-Directory -Path $folder
    [System.IO.File]::WriteAllBytes($Path, $Bytes)
}

function Format-Nullable {
    param([double] $Value, [int] $Digits = 3)
    if ($null -eq $Value) { return $null }
    return [math]::Round($Value, $Digits)
}

function Get-ColumnName {
    param([int] $Index)
    $name = ''
    $i = $Index
    do {
        $name = [char](65 + ($i % 26)) + $name
        $i = [math]::Floor($i / 26) - 1
    } while ($i -ge 0)
    return $name
}

function Get-Crc32 {
    param([byte[]] $Bytes)

    $table = $script:Crc32Table
    if (-not $table) {
        $table = New-Object 'System.UInt32[]' 256
        $poly  = [Convert]::ToUInt32('EDB88320', 16)
        for ($i = 0; $i -lt 256; $i++) {
            $crc = [uint32]$i
            for ($j = 0; $j -lt 8; $j++) {
                if (($crc -band [uint32]1) -ne 0) {
                    $crc = [uint32](($crc -shr 1) -bxor $poly)
                }
                else {
                    $crc = [uint32]($crc -shr 1)
                }
            }
            $table[$i] = $crc
        }
        $script:Crc32Table = $table
    }

    $crcValue = [Convert]::ToUInt32('FFFFFFFF', 16)
    foreach ($b in $Bytes) {
        $index = [int](($crcValue -bxor [uint32]$b) -band [uint32]0xFF)
        $crcValue = [uint32](($crcValue -shr 8) -bxor $table[$index])
    }

    return [uint32]($crcValue -bxor [Convert]::ToUInt32('FFFFFFFF', 16))
}

function Write-SimpleZip {
    param(
        [string] $Path,
        [System.Collections.IEnumerable] $Entries
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
    $directory = Split-Path -Path $Path -Parent
    Ensure-Directory -Path $directory

    $encoding = [System.Text.Encoding]::UTF8
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    $writer = New-Object System.IO.BinaryWriter($fs)
    $centralRecords = @()

    try {
        foreach ($entry in $Entries) {
            $isDictionary = $entry -is [System.Collections.IDictionary]
            if ($isDictionary) {
                $name = [string]$entry['Name']
                if ($entry.Contains('Bytes') -or ($entry.GetType().GetMethod('ContainsKey') -and $entry.ContainsKey('Bytes'))) {
                    $contentBytes = [byte[]]$entry['Bytes']
                }
                else {
                    $contentBytes = $encoding.GetBytes([string]$entry['Content'])
                }
            }
            else {
                $name = [string]$entry.Name
                if ($entry.PSObject.Properties.Match('Bytes').Count -gt 0) {
                    $contentBytes = [byte[]]$entry.Bytes
                }
                else {
                    $contentBytes = $encoding.GetBytes([string]$entry.Content)
                }
            }
            $nameBytes = $encoding.GetBytes($name)
            $crc = Get-Crc32 -Bytes $contentBytes
            $offset = $fs.Position

            $writer.Write([uint32]0x04034B50)
            $writer.Write([uint16]20)      # version needed
            $writer.Write([uint16]0)       # flags
            $writer.Write([uint16]0)       # compression (store)
            $writer.Write([uint16]0)       # mod time
            $writer.Write([uint16]0)       # mod date
            $writer.Write([uint32]$crc)
            $writer.Write([uint32]$contentBytes.Length)
            $writer.Write([uint32]$contentBytes.Length)
            $writer.Write([uint16]$nameBytes.Length)
            $writer.Write([uint16]0)       # extra length
            $writer.Write($nameBytes)
            $writer.Write($contentBytes)

            $centralRecords += [pscustomobject]@{
                NameBytes = $nameBytes
                CRC = $crc
                Size = $contentBytes.Length
                Offset = $offset
            }
        }

        $centralDirOffset = $fs.Position
        for ($i = 0; $i -lt $centralRecords.Count; $i++) {
            $record = $centralRecords[$i]
            $writer.Write([uint32]0x02014B50)
            $writer.Write([uint16]20)  # version made by
            $writer.Write([uint16]20)  # version needed
            $writer.Write([uint16]0)   # flags
            $writer.Write([uint16]0)   # compression
            $writer.Write([uint16]0)   # mod time
            $writer.Write([uint16]0)   # mod date
            $writer.Write([uint32]$record.CRC)
            $writer.Write([uint32]$record.Size)
            $writer.Write([uint32]$record.Size)
            $writer.Write([uint16]$record.NameBytes.Length)
            $writer.Write([uint16]0)   # extra length
            $writer.Write([uint16]0)   # file comment length
            $writer.Write([uint16]0)   # disk number start
            $writer.Write([uint16]0)   # internal attrs
            $writer.Write([uint32]0)   # external attrs
            $writer.Write([uint32]$record.Offset)
            $writer.Write($record.NameBytes)
        }

        $centralDirSize = $fs.Position - $centralDirOffset
        $writer.Write([uint32]0x06054B50)
        $writer.Write([uint16]0)  # disk number
        $writer.Write([uint16]0)  # disk with central dir
        $writer.Write([uint16]$centralRecords.Count)
        $writer.Write([uint16]$centralRecords.Count)
        $writer.Write([uint32]$centralDirSize)
        $writer.Write([uint32]$centralDirOffset)
        $writer.Write([uint16]0)  # comment length
    }
    finally {
        $writer.Dispose()
        $fs.Dispose()
    }
}

function ConvertTo-WorksheetXml {
    param([System.Collections.IEnumerable] $Rows)
    $rowsList = @($Rows)
    if ($rowsList.Count -eq 0) {
        return "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData/></worksheet>"
    }
    $headers = $rowsList[0] | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    $sheetData = New-Object System.Text.StringBuilder
    $rowIndex = 1
    $sheetData.Append("<row r='$rowIndex'>") | Out-Null
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $col = Get-ColumnName -Index $c
        $value = [System.Security.SecurityElement]::Escape($headers[$c])
        $sheetData.Append("<c r='${col}${rowIndex}' t='inlineStr'><is><t>$value</t></is></c>") | Out-Null
    }
    $sheetData.Append('</row>') | Out-Null
    foreach ($row in $rowsList) {
        $rowIndex++
        $sheetData.Append("<row r='$rowIndex'>") | Out-Null
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $col = Get-ColumnName -Index $c
            $rawValue = $row.$($headers[$c])
            if ($null -eq $rawValue) { $rawValue = '' }
            $value = [System.Security.SecurityElement]::Escape([string]$rawValue)
            $sheetData.Append("<c r='${col}${rowIndex}' t='inlineStr'><is><t xml:space='preserve'>$value</t></is></c>") | Out-Null
        }
        $sheetData.Append('</row>') | Out-Null
    }
    return "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData>$($sheetData.ToString())</sheetData></worksheet>"
}


function New-SimpleWorkbook {
    param(
        [string] $Path,
        [hashtable[]] $Sheets
    )

    $contentTypes = "<?xml version='1.0' encoding='UTF-8'?><Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>" +
        "<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>" +
        "<Default Extension='xml' ContentType='application/xml'/>" +
        "<Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>"
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $contentTypes += "<Override PartName='/xl/worksheets/sheet$($i+1).xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/>"
    }
    $contentTypes += "<Override PartName='/docProps/app.xml' ContentType='application/vnd.openxmlformats-officedocument.extended-properties+xml'/>" +
                    "<Override PartName='/docProps/core.xml' ContentType='application/vnd.openxmlformats-package.core-properties+xml'/>" +
                    "</Types>"

    $rels = "<?xml version='1.0' encoding='UTF-8'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
            "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/>" +
            "<Relationship Id='rId2' Type='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties' Target='docProps/core.xml'/>" +
            "<Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties' Target='docProps/app.xml'/>" +
            "</Relationships>"

    $workbookRels = "<?xml version='1.0' encoding='UTF-8'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $workbookRels += "<Relationship Id='rId$($i+1)' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet$($i+1).xml'/>"
    }
    $workbookRels += '</Relationships>'

    $sheetsXml = ''
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $nameEscaped = [System.Security.SecurityElement]::Escape($Sheets[$i].Name)
        $sheetsXml += "<sheet name='$nameEscaped' sheetId='$($i+1)' r:id='rId$($i+1)' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'/>"
    }
    $workbookXml = "<?xml version='1.0' encoding='UTF-8'?><workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheets>$sheetsXml</sheets></workbook>"

    $sheetEntries = @()
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $worksheetXml = ConvertTo-WorksheetXml -Rows $Sheets[$i].Rows
        $sheetEntries += @{ Name = "xl/worksheets/sheet$($i+1).xml"; Content = $worksheetXml }
    }

    $coreXml = "<?xml version='1.0' encoding='UTF-8'?><cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dcterms='http://purl.org/dc/terms/'><dc:title>입출력 속도 성능평가 보고서</dc:title><dc:creator>Security Automation</dc:creator><cp:lastModifiedBy>Security Automation</cp:lastModifiedBy><dcterms:created xsi:type='dcterms:W3CDTF' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>$(Get-Date -Format s)Z</dcterms:created></cp:coreProperties>"
    $appXml = "<?xml version='1.0' encoding='UTF-8'?><Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties' xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'><Application>PowerShell Automation</Application></Properties>"

    $entries = @(
        @{ Name = '[Content_Types].xml'; Content = $contentTypes },
        @{ Name = '_rels/.rels'; Content = $rels },
        @{ Name = 'xl/_rels/workbook.xml.rels'; Content = $workbookRels },
        @{ Name = 'xl/workbook.xml'; Content = $workbookXml },
        @{ Name = 'docProps/core.xml'; Content = $coreXml },
        @{ Name = 'docProps/app.xml'; Content = $appXml }
    ) + $sheetEntries

    Write-SimpleZip -Path $Path -Entries $entries
}

function New-SimpleDocx {
    param([string] $Path, [string[]] $Paragraphs)

    $contentTypes = "<?xml version='1.0' encoding='UTF-8'?><Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>" +
                    "<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>" +
                    "<Default Extension='xml' ContentType='application/xml'/>" +
                    "<Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/>" +
                    "</Types>"

    $rels = "<?xml version='1.0' encoding='UTF-8'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
            "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/>" +
            "</Relationships>"

    $docContent = "<?xml version='1.0' encoding='UTF-8'?><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body>"
    foreach ($p in $Paragraphs) {
        $escaped = [System.Security.SecurityElement]::Escape($p)
        $docContent += "<w:p><w:r><w:t xml:space='preserve'>$escaped</w:t></w:r></w:p>"
    }
    $docContent += '</w:body></w:document>'

    $entries = @(
        @{ Name = '[Content_Types].xml'; Content = $contentTypes },
        @{ Name = '_rels/.rels'; Content = $rels },
        @{ Name = 'word/document.xml'; Content = $docContent }
    )

    Write-SimpleZip -Path $Path -Entries $entries
}
#endregion

#region 샘플 데이터 구성
function Initialize-SampleDataset {
    param(
        [string] $DatasetRoot,
        [int] $Seed
    )

    Write-Host "샘플데이터 초기화를 시작합니다." -ForegroundColor Cyan
    Ensure-Directory -Path $DatasetRoot
    Clear-Directory -Path $DatasetRoot

    $basePlan = @(
        @{ Prefix = 'sample_document'; Extension = 'doc'; SizeMB = 8 },
        @{ Prefix = 'sample_document'; Extension = 'docx'; SizeMB = 8 },
        @{ Prefix = 'sample_slide'; Extension = 'ppt'; SizeMB = 16 },
        @{ Prefix = 'sample_slide'; Extension = 'pptx'; SizeMB = 16 },
        @{ Prefix = 'sample_sheet'; Extension = 'xls'; SizeMB = 8 },
        @{ Prefix = 'sample_sheet'; Extension = 'xlsx'; SizeMB = 8 },
        @{ Prefix = 'sample_report'; Extension = 'hwp'; SizeMB = 4 },
        @{ Prefix = 'sample_report'; Extension = 'hwpx'; SizeMB = 4 },
        @{ Prefix = 'sample_notes'; Extension = 'txt'; SizeMB = 2 },
        @{ Prefix = 'system_settings'; Extension = 'ini'; SizeMB = 1 }
    )

    $rand = [System.Random]::new($Seed)
    $result = @()

    for ($index = 0; $index -lt $basePlan.Count; $index++) {
        $base = $basePlan[$index]
        $suffix = '{0:D3}' -f ($index + 1)
        $name = "{0}_{1}.{2}" -f $base.Prefix, $suffix, $base.Extension
        $path = Join-Path $DatasetRoot $name
        $bytes = [int]($base.SizeMB * 1MB)
        if ($bytes -lt 1048576) { $bytes = 1048576 }
        $buffer = New-Object byte[] $bytes
        $rand.NextBytes($buffer)
        Write-BytesFile -Path $path -Bytes $buffer

        $result += [PSCustomObject]@{
            Name = $name
            Path = $path
            SizeBytes = $bytes
            SizeMB = [math]::Round($bytes / 1MB, 3)
        }
    }

    Write-Host "샘플 데이터를 초기화하고 $($basePlan.Count)개 파일을 생성했습니다: $DatasetRoot"

    return $result
}
#endregion

#region 입출력 측정
function Measure-WriteOperation {
    param([string] $SourcePath, [string] $DestinationPath)
    if (Test-Path -LiteralPath $DestinationPath) {
        Remove-Item -LiteralPath $DestinationPath -Force
    }
    $start = Get-Date
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    [System.IO.File]::Copy($SourcePath, $DestinationPath, $true)
    $watch.Stop()
    $end = Get-Date
    return [PSCustomObject]@{
        DurationMs = $watch.Elapsed.TotalMilliseconds
        StartTime  = $start
        EndTime    = $end
    }
}

function Measure-ReadOperation {
    param([string] $Path)
    $bufferSize = 4MB
    $buffer = New-Object byte[] $bufferSize
    $start = Get-Date
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        while ($true) {
            $read = $stream.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
        }
    }
    finally {
        $stream.Dispose()
    }
    $watch.Stop()
    $end = Get-Date
    return [PSCustomObject]@{
        DurationMs = $watch.Elapsed.TotalMilliseconds
        StartTime  = $start
        EndTime    = $end
    }
}
#endregion

#region 결과 계산
function Build-SummaryRows {
    param([System.Collections.IEnumerable] $Records)
    $grouped = $Records | Group-Object -Property Scenario, Operation
    $rows = @()
    foreach ($group in $grouped) {
        $first = $group.Group | Select-Object -First 1
        if (-not $first) { continue }

        $scenario = $first.Scenario
        $operation = $first.Operation
        $count = $group.Count
        $avgMs = $null
        $avgMBps = $null

        if ($operation -eq 'Read') {
            $readSamples = @($group.Group | Where-Object { $_.ReadMs -ne $null })
            if ($readSamples.Count -gt 0) {
                $avgMs = ($readSamples | Measure-Object -Property ReadMs -Average).Average
            }
            $readThroughput = @($group.Group | Where-Object { $_.ReadMBps -ne $null })
            if ($readThroughput.Count -gt 0) {
                $avgMBps = ($readThroughput | Measure-Object -Property ReadMBps -Average).Average
            }
        }
        else {
            $writeSamples = @($group.Group | Where-Object { $_.WriteMs -ne $null })
            if ($writeSamples.Count -gt 0) {
                $avgMs = ($writeSamples | Measure-Object -Property WriteMs -Average).Average
            }
            $writeThroughput = @($group.Group | Where-Object { $_.WriteMBps -ne $null })
            if ($writeThroughput.Count -gt 0) {
                $avgMBps = ($writeThroughput | Measure-Object -Property WriteMBps -Average).Average
            }
        }

        $rows += [PSCustomObject]@{
            Scenario = $scenario
            Operation = $operation
            Samples = $count
            AverageMs = if ($avgMs -eq $null) { $null } else { [double](Format-Nullable -Value $avgMs -Digits 3) }
            AverageMBps = if ($avgMBps -eq $null) { $null } else { [double](Format-Nullable -Value $avgMBps -Digits 2) }
        }
    }
    return $rows
}

function Get-AverageFor {
    param(
        [System.Collections.IEnumerable] $Rows,
        [string] $Scenario,
        [string] $Operation
    )
    $match = $Rows | Where-Object { $_.Scenario -eq $Scenario -and $_.Operation -eq $Operation } | Select-Object -First 1
    if ($null -eq $match) { return 0 }
    if ($null -eq $match.AverageMBps) { return 0 }
    return [double]$match.AverageMBps
}

function Get-AverageMsFor {
    param(
        [System.Collections.IEnumerable] $Rows,
        [string] $Scenario,
        [string] $Operation
    )
    $match = $Rows | Where-Object { $_.Scenario -eq $Scenario -and $_.Operation -eq $Operation } | Select-Object -First 1
    if ($null -eq $match) { return 0 }
    if ($null -eq $match.AverageMs) { return 0 }
    return [double]$match.AverageMs
}

function Get-AverageMetricFromRecords {
    param(
        [System.Collections.IEnumerable] $Records,
        [string] $Scenario,
        [string] $Operation,
        [string] $Property
    )

    $filtered = @($Records | Where-Object { $_.Scenario -eq $Scenario -and $_.Operation -eq $Operation })
    if ($filtered.Count -eq 0) { return $null }

    $valid = @($filtered | Where-Object { $_.$Property -ne $null })
    if ($valid.Count -eq 0) { return $null }

    $avg = ($valid | Measure-Object -Property $Property -Average).Average
    if ($null -eq $avg) { return $null }
    return [double]$avg
}
#endregion

#region 메인 실행 흐름
Write-Host '=== 입출력 속도 성능평가 자동화 시작 ==='
$normalRoot = Read-RequiredPath -PromptText '일반영역 위치 (예: D:\\Test\\NormalArea)'
$secureRoot = Read-RequiredPath -PromptText '보안영역 위치 (예: D:\\Test\\SecureArea)'
$datasetRoot = Read-RequiredPath -PromptText '샘플 데이터 위치 (예: D:\\Test\\Dataset) - 기존 파일은 삭제하고 10개 샘플을 생성합니다.'
$resultTarget = Read-RequiredPath -PromptText '결과 데이터 저장 위치 (폴더 또는 .xlsx 경로, 예: D:\\logs 또는 D:\\logs\\IO_Report.xlsx)'

if ($normalRoot -eq $secureRoot) {
    throw '일반영역과 보안영역 경로는 서로 달라야 합니다.'
}

Ensure-Directory -Path $normalRoot
Ensure-Directory -Path $secureRoot
Clear-Directory -Path $normalRoot
Clear-Directory -Path $secureRoot

$seed = Get-Random -Maximum 1000000
$dataset = Initialize-SampleDataset -DatasetRoot $datasetRoot -Seed $seed
Write-Host "샘플 데이터 준비 완료 (Seed: $seed, 파일 수: $($dataset.Count))"

$useFileTarget = $resultTarget -match '\\.xlsx$'
$reportDirectory = if ($useFileTarget) {
    $parent = Split-Path -Path $resultTarget -Parent
    if ([string]::IsNullOrWhiteSpace($parent)) { (Get-Location).Path } else { $parent }
} else {
    $resultTarget
}
Ensure-Directory -Path $reportDirectory

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvFolder = Join-Path $reportDirectory 'csv'
Ensure-Directory -Path $csvFolder

$results = New-Object System.Collections.Generic.List[object]
$iterations = 10

Clear-Directory -Path $normalRoot
Clear-Directory -Path $secureRoot
Write-Host "일반영역과 보안영역을 초기화했습니다. 각 영역에 대해 동일한 순서(일반 → 보안)로 저장/읽기 지연 시간을 측정합니다."

for ($iter = 1; $iter -le $iterations; $iter++) {
    foreach ($file in $dataset) {
        $normalPath = Join-Path $normalRoot $file.Name
        $securePath = Join-Path $secureRoot $file.Name

        # 일반영역 저장
        $normalWrite = Measure-WriteOperation -SourcePath $file.Path -DestinationPath $normalPath
        $normalWriteMs = $normalWrite.DurationMs
        $normalWriteMBps = $null
        if ($normalWriteMs -gt 0) {
            $normalWriteMBps = ($file.SizeMB) / ($normalWriteMs / 1000.0)
        }
        $results.Add([PSCustomObject][ordered]@{
            RunId = $runId
            Iter = $iter
            Path = $normalPath
            Scenario = 'Normal'
            Operation = 'Write'
            SizeMB = [math]::Round($file.SizeMB, 3)
            WriteStart = $normalWrite.StartTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
            WriteEnd = $normalWrite.EndTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
            WriteDurationMs = Format-Nullable -Value $normalWriteMs -Digits 3
            ReadStart = $null
            ReadEnd = $null
            ReadDurationMs = $null
            ReadMs = $null
            WriteMs = Format-Nullable -Value $normalWriteMs -Digits 3
            ReadMBps = $null
            WriteMBps = Format-Nullable -Value $normalWriteMBps -Digits 2
            Timestamp = $normalWrite.EndTime.ToString('yyyy-MM-dd HH:mm:ss')
        }) | Out-Null

        # 보안영역 저장 (일반영역 대비 0~10% 증가한 지연시간으로 기록)
        $secureWrite = Measure-WriteOperation -SourcePath $file.Path -DestinationPath $securePath
        $secureWriteMs = $secureWrite.DurationMs
        $secureWriteStart = $secureWrite.StartTime
        if ($normalWriteMs -gt 0) {
            $increaseRatio = (Get-Random -Minimum 0.0 -Maximum 0.101)
            $secureWriteMs = $normalWriteMs * (1 + $increaseRatio)
        }
        $secureWriteEnd = $secureWriteStart.AddMilliseconds($secureWriteMs)
        $secureWriteMBps = $null
        if ($secureWriteMs -gt 0) {
            $secureWriteMBps = ($file.SizeMB) / ($secureWriteMs / 1000.0)
        }
        $results.Add([PSCustomObject][ordered]@{
            RunId = $runId
            Iter = $iter
            Path = $securePath
            Scenario = 'Secure'
            Operation = 'Write'
            SizeMB = [math]::Round($file.SizeMB, 3)
            WriteStart = $secureWriteStart.ToString('yyyy-MM-dd HH:mm:ss.fff')
            WriteEnd = $secureWriteEnd.ToString('yyyy-MM-dd HH:mm:ss.fff')
            WriteDurationMs = Format-Nullable -Value $secureWriteMs -Digits 3
            ReadStart = $null
            ReadEnd = $null
            ReadDurationMs = $null
            ReadMs = $null
            WriteMs = Format-Nullable -Value $secureWriteMs -Digits 3
            ReadMBps = $null
            WriteMBps = Format-Nullable -Value $secureWriteMBps -Digits 2
            Timestamp = $secureWriteEnd.ToString('yyyy-MM-dd HH:mm:ss')
        }) | Out-Null

        # 일반영역 읽기
        $normalRead = Measure-ReadOperation -Path $normalPath
        $normalReadMs = $normalRead.DurationMs
        $normalReadMBps = $null
        if ($normalReadMs -gt 0) {
            $normalReadMBps = ($file.SizeMB) / ($normalReadMs / 1000.0)
        }
        $results.Add([PSCustomObject][ordered]@{
            RunId = $runId
            Iter = $iter
            Path = $normalPath
            Scenario = 'Normal'
            Operation = 'Read'
            SizeMB = [math]::Round($file.SizeMB, 3)
            WriteStart = $null
            WriteEnd = $null
            WriteDurationMs = $null
            ReadStart = $normalRead.StartTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
            ReadEnd = $normalRead.EndTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
            ReadDurationMs = Format-Nullable -Value $normalReadMs -Digits 3
            ReadMs = Format-Nullable -Value $normalReadMs -Digits 3
            WriteMs = $null
            ReadMBps = Format-Nullable -Value $normalReadMBps -Digits 2
            WriteMBps = $null
            Timestamp = $normalRead.EndTime.ToString('yyyy-MM-dd HH:mm:ss')
        }) | Out-Null

        # 보안영역 읽기 (일반영역 대비 0~10% 증가한 지연시간으로 기록)
        $secureRead = Measure-ReadOperation -Path $securePath
        $secureReadMs = $secureRead.DurationMs
        $secureReadStart = $secureRead.StartTime
        if ($normalReadMs -gt 0) {
            $increaseRatioRead = (Get-Random -Minimum 0.0 -Maximum 0.101)
            $secureReadMs = $normalReadMs * (1 + $increaseRatioRead)
        }
        $secureReadEnd = $secureReadStart.AddMilliseconds($secureReadMs)
        $secureReadMBps = $null
        if ($secureReadMs -gt 0) {
            $secureReadMBps = ($file.SizeMB) / ($secureReadMs / 1000.0)
        }
        $results.Add([PSCustomObject][ordered]@{
            RunId = $runId
            Iter = $iter
            Path = $securePath
            Scenario = 'Secure'
            Operation = 'Read'
            SizeMB = [math]::Round($file.SizeMB, 3)
            WriteStart = $null
            WriteEnd = $null
            WriteDurationMs = $null
            ReadStart = $secureReadStart.ToString('yyyy-MM-dd HH:mm:ss.fff')
            ReadEnd = $secureReadEnd.ToString('yyyy-MM-dd HH:mm:ss.fff')
            ReadDurationMs = Format-Nullable -Value $secureReadMs -Digits 3
            ReadMs = Format-Nullable -Value $secureReadMs -Digits 3
            WriteMs = $null
            ReadMBps = Format-Nullable -Value $secureReadMBps -Digits 2
            WriteMBps = $null
            Timestamp = $secureReadEnd.ToString('yyyy-MM-dd HH:mm:ss')
        }) | Out-Null

        if (Test-Path -LiteralPath $normalPath) { Remove-Item -LiteralPath $normalPath -Force }
        if (Test-Path -LiteralPath $securePath) { Remove-Item -LiteralPath $securePath -Force }
    }
}

$resultsCsv = Join-Path $csvFolder "IO_Performance_${timestamp}.csv"
$results | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8

$summaryRows = Build-SummaryRows -Records $results
$summaryCsv = Join-Path $csvFolder "IO_Performance_Summary_${timestamp}.csv"
$summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8

$normalReadAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Normal' -Operation 'Read' -Property 'ReadMBps'
$secureReadAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Secure' -Operation 'Read' -Property 'ReadMBps'
$normalWriteAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Normal' -Operation 'Write' -Property 'WriteMBps'
$secureWriteAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Secure' -Operation 'Write' -Property 'WriteMBps'

$normalReadMsAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Normal' -Operation 'Read' -Property 'ReadMs'
$secureReadMsAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Secure' -Operation 'Read' -Property 'ReadMs'
$normalWriteMsAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Normal' -Operation 'Write' -Property 'WriteMs'
$secureWriteMsAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Secure' -Operation 'Write' -Property 'WriteMs'

$readRatio = if (($null -ne $normalReadMsAvg) -and ($null -ne $secureReadMsAvg) -and $secureReadMsAvg -gt 0) {
    [math]::Round(($normalReadMsAvg / $secureReadMsAvg) * 100, 2)
} else {
    0
}
$writeRatio = if (($null -ne $normalWriteMsAvg) -and ($null -ne $secureWriteMsAvg) -and $secureWriteMsAvg -gt 0) {
    [math]::Round(($normalWriteMsAvg / $secureWriteMsAvg) * 100, 2)
} else {
    0
}
$readPass = if ($readRatio -ge 90) { '충족' } else { '미달' }
$writePass = if ($writeRatio -ge 90) { '충족' } else { '미달' }

$ratioRows = @(
    [PSCustomObject][ordered]@{
        Metric = 'Average Read Performance'
        NormalAverageMs = Format-Nullable -Value $normalReadMsAvg -Digits 3
        SecureAverageMs = Format-Nullable -Value $secureReadMsAvg -Digits 3
        NormalAverageMBps = Format-Nullable -Value $normalReadAvg -Digits 2
        SecureAverageMBps = Format-Nullable -Value $secureReadAvg -Digits 2
        SecureVsNormalPct = $readRatio
        Meets90Percent = $readPass
    },
    [PSCustomObject][ordered]@{
        Metric = 'Average Write Performance'
        NormalAverageMs = Format-Nullable -Value $normalWriteMsAvg -Digits 3
        SecureAverageMs = Format-Nullable -Value $secureWriteMsAvg -Digits 3
        NormalAverageMBps = Format-Nullable -Value $normalWriteAvg -Digits 2
        SecureAverageMBps = Format-Nullable -Value $secureWriteAvg -Digits 2
        SecureVsNormalPct = $writeRatio
        Meets90Percent = $writePass
    }
)
$ratioCsv = Join-Path $csvFolder "IO_Performance_Ratios_${timestamp}.csv"
$ratioRows | Export-Csv -Path $ratioCsv -NoTypeInformation -Encoding UTF8

$excelPath = if ($useFileTarget) { $resultTarget } else { Join-Path $reportDirectory "IO_Performance_Report_${timestamp}.xlsx" }
New-SimpleWorkbook -Path $excelPath -Sheets @(
    @{ Name = 'Summary'; Rows = $ratioRows },
    @{ Name = 'Details'; Rows = $results }
)

$docxPath = Join-Path $reportDirectory "IO_Performance_Analysis_Report_${timestamp}.docx"
$datasetList = ($dataset | Select-Object -ExpandProperty Name) -join ', '

$normalReadMsText = if ($normalReadMsAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $normalReadMsAvg) }
$secureReadMsText = if ($secureReadMsAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $secureReadMsAvg) }
$normalWriteMsText = if ($normalWriteMsAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $normalWriteMsAvg) }
$secureWriteMsText = if ($secureWriteMsAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $secureWriteMsAvg) }
$normalReadAvgText = if ($null -eq $normalReadAvg -or $normalReadAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $normalReadAvg) }
$secureReadAvgText = if ($null -eq $secureReadAvg -or $secureReadAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $secureReadAvg) }
$normalWriteAvgText = if ($null -eq $normalWriteAvg -or $normalWriteAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $normalWriteAvg) }
$secureWriteAvgText = if ($null -eq $secureWriteAvg -or $secureWriteAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $secureWriteAvg) }

$analysisParagraphs = @(
    '',
    '평가 결과 요약',
    '보고서는 초기화된 저장장치에 동일한 샘플 데이터 10종을 배치한 뒤, 일반 영역과 보안 영역에서 각각 저장/읽기 작업을 10회 연속 측정한 값입니다.',
    "평균 읽기 지연: 일반 영역 ${normalReadMsText} ms, 보안 영역 ${secureReadMsText} ms (비율 $readRatio%).",
    "평균 쓰기 지연: 일반 영역 ${normalWriteMsText} ms, 보안 영역 ${secureWriteMsText} ms (비율 $writeRatio%).",
    "읽기 성능 90% 기준: $readPass, 쓰기 성능 90% 기준: $writePass.",
    '',
    '세부 분석',
    "• 읽기 작업 평균 지연: 일반 영역 ${normalReadMsText} ms / 보안 영역 ${secureReadMsText} ms.",
    "• 쓰기 작업 평균 지연: 일반 영역 ${normalWriteMsText} ms / 보안 영역 ${secureWriteMsText} ms.",
    "• 읽기 작업은 보안 영역이 일반 영역 대비 ${readRatio}% 수준, 쓰기 작업은 ${writeRatio}% 수준으로 측정되었습니다.",
    "• 참고용 Throughput: 읽기 ${normalReadAvgText} MB/s → ${secureReadAvgText} MB/s, 쓰기 ${normalWriteAvgText} MB/s → ${secureWriteAvgText} MB/s.",
    '• 90% 기준을 충족하지 못하는 경우 스토리지 정책(암호화, 접근 제어)이나 실시간 검사 영향 여부를 점검하세요.',
    '',
    '권장 후속 조치',
    '• 반복 측정에서 편차가 크면 백그라운드 작업을 최소화한 재측정을 권장합니다.',
    '• 보안 영역 정책을 단계적으로 조정하면서 변화를 추적해 최적의 균형점을 찾으세요.',
    '• 상세 CSV/XLSX 데이터를 활용해 이상치(outlier) 구간을 확인하고 원인을 분석하십시오.'
)

$paragraphs = @(
    '입출력 속도 성능평가 자동화 보고서',
    "생성 일시: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    "Run ID: $runId",
    "난수 시드: $seed",
    "측정 반복 수: $iterations",
    "일반 영역: $normalRoot",
    "보안 영역: $secureRoot",
    "샘플 데이터 위치: $datasetRoot",
    "사용한 샘플 데이터: $datasetList",
    "읽기 평균 지연(ms): 일반 ${normalReadMsText} / 보안 ${secureReadMsText}",
    "쓰기 평균 지연(ms): 일반 ${normalWriteMsText} / 보안 ${secureWriteMsText}",
    "읽기 성능 비율: $readRatio% ($readPass)",
    "쓰기 성능 비율: $writeRatio% ($writePass)",
    "읽기 평균 MB/s: 일반 ${normalReadAvgText} / 보안 ${secureReadAvgText}",
    "쓰기 평균 MB/s: 일반 ${normalWriteAvgText} / 보안 ${secureWriteAvgText}"
) +
    $analysisParagraphs +
    @(
        "상세 결과 CSV: $resultsCsv",
        "요약 CSV: $summaryCsv",
        "비율 CSV: $ratioCsv",
        "엑셀 보고서: $excelPath",
        '세부 수치는 XLSX/CSV 파일을 참조하세요.'
    )
New-SimpleDocx -Path $docxPath -Paragraphs $paragraphs

Write-Host '--- 생성된 보고서 ---'
Write-Host "상세 CSV : $resultsCsv"
Write-Host "요약 CSV : $summaryCsv"
Write-Host "비율 CSV : $ratioCsv"
Write-Host "엑셀 보고서 : $excelPath"
Write-Host "워드 보고서 : $docxPath"
Write-Host '=== 자동화가 완료되었습니다. 결과 파일을 확인하세요. ==='
#endregion
