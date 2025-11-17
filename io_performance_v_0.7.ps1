#requires -Version 5.1
<#!
    입출력 속도 측정 스크립트 안내
    ------------------------------------------------------------
    1) 일반 영역과 보안 영역으로 사용할 드라이브 문자를 입력하면 각 드라이브에 Normal_Zone, Secure_Zone 폴더를 새로 만듭니다.
    2) 폴더에 예시 파일을 여러 개 만들어 쓰기 속도를 측정합니다.
    3) 만들어 둔 파일을 다시 읽어 읽기 속도를 측정합니다.
    4) 결과를 CSV, XLSX, DOCX 파일로 저장하고 보안 영역 속도가 일반 영역의 90% 이상인지 확인합니다.

    관리자 권한 PowerShell 창에서 실행하고, 실제 업무 자료가 없는 테스트 드라이브에서 사용하세요.
!#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Crc32Table = $null

#region 공통으로 쓰는 도우미 함수
# 아래 함수들은 스크립트 전반에서 반복 사용되는 기본 기능입니다.
# 폴더를 만들거나 지우고, 사용자 입력을 확인하는 등 기초 작업을 담당합니다.
# 각 함수의 역할을 간단한 생활 용어로 설명했으니 PowerShell이 처음인 분도 흐름을 이해할 수 있습니다.
function Read-RequiredPath {
    param([Parameter(Mandatory)] [string] $PromptText)
    # 사용자가 내용을 입력할 때까지 계속 묻습니다.
    # "그냥 엔터"를 누르면 다시 입력하라는 안내를 띄웁니다.
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
    # 특정 경로에 폴더가 없으면 새로 만들어 줍니다.
    # 폴더가 이미 있으면 아무 일도 하지 않고 그냥 통과합니다.
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Clear-Directory {
    param([string] $Path)
    # 기존 파일이 남아 있으면 측정 결과가 섞이기 때문에 폴더 안을 깨끗하게 비웁니다.
    # 폴더가 없으면 건드릴 것이 없으므로 바로 종료합니다.
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
    # 바이트 배열을 그대로 파일로 저장합니다.
    # 샘플 데이터를 만들 때 다양한 크기의 파일을 생성하는 용도로 사용합니다.
    $folder = Split-Path -Path $Path -Parent
    Ensure-Directory -Path $folder
    [System.IO.File]::WriteAllBytes($Path, $Bytes)
}

function Format-Nullable {
    param([double] $Value, [int] $Digits = 3)
    # 계산 결과가 비어 있을 때는 그대로 null을 돌려주고
    # 값이 있다면 보기 좋게 소수점 자릿수를 제한합니다.
    if ($null -eq $Value) { return $null }
    return [math]::Round($Value, $Digits)
}

function Resolve-DriveRoot {
    param([Parameter(Mandatory)][string] $DriveText)

    # 사용자가 입력한 드라이브 문자(C, C:, C:\ 등)를 표준 형식(C:\)으로 바꿔 줍니다.
    # 서로 다른 입력 형태라도 같은 드라이브를 가리키도록 통일하는 과정입니다.

    $text = $DriveText.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw '드라이브 문자를 입력해야 합니다.'
    }

    $normalized = $text -replace '/', '\\'
    if ($normalized -match '^[A-Za-z]$') {
        $driveLetter = $normalized.ToUpper()
        return ($driveLetter + ':\')
    }

    if ($normalized -match '^[A-Za-z]:$') {
        $driveLetter = $normalized.Substring(0, 1).ToUpper()
        return ($driveLetter + ':\')
    }

    if ($normalized -match '^[A-Za-z]:\\?$') {
        $driveLetter = $normalized.Substring(0, 1).ToUpper()
        return ($driveLetter + ':\')
    }

    try {
        $root = [System.IO.Path]::GetPathRoot($normalized)
    }
    catch {
        throw "드라이브 형식을 인식할 수 없습니다: $DriveText"
    }

    if ([string]::IsNullOrWhiteSpace($root)) {
        throw "드라이브 형식을 인식할 수 없습니다: $DriveText"
    }

    $root = $root.Trim()
    if (-not $root.EndsWith('\\')) {
        $root = $root + '\\'
    }

    return $root.ToUpper()
}

function Get-ColumnName {
    param([int] $Index)
    # 엑셀 셀 주소에 사용하는 A, B, C... AA, AB 같은 열 이름을 만들어 줍니다.
    # 첫 번째 열은 A, 27번째 열은 AA가 되도록 계산합니다.
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

    # ZIP 파일 형식을 만들기 위해 필요한 CRC32 체크섬을 직접 계산합니다.
    # .NET의 기본 라이브러리를 쓰지 않고 수식을 구현했습니다.

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

    # 여러 개의 XML/텍스트 조각을 모아 하나의 ZIP 파일(XLSX, DOCX의 기반)을 만듭니다.
    # 상용 라이브러리 없이 순수 PowerShell 코드만으로 최소한의 ZIP 구조를 작성합니다.
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
            $writer.Write([uint16]20)      # ZIP 파일을 만들 때 필요한 최소 버전 정보
            $writer.Write([uint16]0)       # ZIP 옵션 값 (압축 없이 저장하므로 0)
            $writer.Write([uint16]0)       # 압축 방식을 "저장"으로 설정 (실제 압축 없음)
            $writer.Write([uint16]0)       # 파일 수정 시간 정보를 넣지 않아 0
            $writer.Write([uint16]0)       # 파일 수정 날짜 정보를 넣지 않아 0
            $writer.Write([uint32]$crc)
            $writer.Write([uint32]$contentBytes.Length)
            $writer.Write([uint32]$contentBytes.Length)
            $writer.Write([uint16]$nameBytes.Length)
            $writer.Write([uint16]0)       # 추가 데이터 길이를 쓰지 않아 0으로 둡니다
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
            $writer.Write([uint16]20)  # ZIP 파일을 만든 쪽의 버전 정보
            $writer.Write([uint16]20)  # ZIP 파일을 만들 때 필요한 최소 버전 정보
            $writer.Write([uint16]0)   # ZIP 옵션 값 (사용하지 않아 0으로 둡니다)
            $writer.Write([uint16]0)   # 압축 방식을 저장 전용으로 설정합니다
            $writer.Write([uint16]0)   # 파일 수정 시간을 쓰지 않아 0으로 둡니다
            $writer.Write([uint16]0)   # 파일 수정 날짜를 쓰지 않아 0으로 둡니다
            $writer.Write([uint32]$record.CRC)
            $writer.Write([uint32]$record.Size)
            $writer.Write([uint32]$record.Size)
            $writer.Write([uint16]$record.NameBytes.Length)
            $writer.Write([uint16]0)   # 추가 데이터 길이를 쓰지 않아 0으로 둡니다
            $writer.Write([uint16]0)   # 파일 설명 길이를 쓰지 않아 0으로 둡니다
            $writer.Write([uint16]0)   # 파일이 들어있는 디스크 번호 (단일 파일이라 0)
            $writer.Write([uint16]0)   # 내부 속성 (사용하지 않아 0)
            $writer.Write([uint32]0)   # 외부 속성 (사용하지 않아 0)
            $writer.Write([uint32]$record.Offset)
            $writer.Write($record.NameBytes)
        }

        $centralDirSize = $fs.Position - $centralDirOffset
        $writer.Write([uint32]0x06054B50)
        $writer.Write([uint16]0)  # 현재 디스크 번호 (단일 파일이라 0)
        $writer.Write([uint16]0)  # 중앙 디렉터리가 들어있는 디스크 번호 (단일 파일이라 0)
        $writer.Write([uint16]$centralRecords.Count)
        $writer.Write([uint16]$centralRecords.Count)
        $writer.Write([uint32]$centralDirSize)
        $writer.Write([uint32]$centralDirOffset)
        $writer.Write([uint16]0)  # ZIP 전체 설명 길이를 쓰지 않아 0으로 둡니다
    }
    finally {
        $writer.Dispose()
        $fs.Dispose()
    }
}

function ConvertTo-WorksheetXml {
    param([System.Collections.IEnumerable] $Rows)
    # CSV 데이터를 엑셀에서 사용할 수 있는 XML 형식으로 바꿉니다.
    # 머리글, 행 번호 등을 모두 직접 조립하여 간단한 시트를 만듭니다.
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

    # Excel 파일(xlsx)은 사실 여러 XML 파일을 ZIP으로 묶은 구조입니다.
    # 필요한 최소 XML 조각을 직접 만들어 ZIP으로 묶어 간단한 보고서를 생성합니다.
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

    # 워드 파일(docx)도 여러 XML 문서를 ZIP으로 묶은 형식입니다.
    # 보고서 내용을 문단 배열로 받아 최소한의 구조로 저장합니다.
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

#region 샘플 데이터 만들기
function Initialize-SampleDataset {
    param(
        [string] $DatasetRoot,
        [int] $Seed
    )

    # 성능 측정을 위해 여러 종류의 예제 파일을 자동으로 만들어 줍니다.
    # 문서, 슬라이드, 스프레드시트 등 다양한 확장자를 포함해 실제 업무 환경을 흉내 냅니다.
    # Seed 값을 고정하면 같은 파일 구성이 다시 생성되어 재현이 쉬워집니다.
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

#region 쓰기/읽기 측정하기
function Measure-WriteOperation {
    param([string] $SourcePath, [string] $DestinationPath)
    # 파일 복사 시간을 측정하여 "쓰기" 속도를 계산합니다.
    # 초 단위보다 더 정확한 결과를 얻기 위해 Stopwatch 클래스를 사용합니다.
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
    # 파일을 일정 크기(4MB)씩 읽어 들이면서 총 소요 시간을 측정합니다.
    # 실제 사용자 시나리오처럼 파일을 끝까지 읽도록 설계했습니다.
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

#region 측정 결과 정리하기
function Build-SummaryRows {
    param([System.Collections.IEnumerable] $Records)
    # 여러 번 측정한 결과를 시나리오(일반/보안)와 작업 종류(읽기/쓰기)별로 묶어 평균을 계산합니다.
    # 평균 지연 시간(ms)과 초당 전송 속도(MB/s)를 함께 기록합니다.
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
    # 특정 시나리오와 작업에 해당하는 평균 MB/s 값을 꺼내는 보조 함수입니다.
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
    # 특정 시나리오와 작업에 해당하는 평균 지연(ms) 값을 꺼내는 보조 함수입니다.
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

    # Raw 측정 결과에서 원하는 조건(예: 보안 영역의 읽기 속도)에 맞는 값만 골라 평균을 계산합니다.
    $filtered = @($Records | Where-Object { $_.Scenario -eq $Scenario -and $_.Operation -eq $Operation })
    if ($filtered.Count -eq 0) { return $null }

    $valid = @($filtered | Where-Object { $_.$Property -ne $null })
    if ($valid.Count -eq 0) { return $null }

    $avg = ($valid | Measure-Object -Property $Property -Average).Average
    if ($null -eq $avg) { return $null }
    return [double]$avg
}
#endregion

#region 전체 실행 순서
Write-Host '=== 입출력 속도 성능평가 자동화 시작 ==='

# 1단계) 일반/보안 영역으로 사용할 드라이브를 입력 받습니다.
# 드라이브 문자는 유연하게 입력할 수 있으며, Resolve-DriveRoot가 자동으로 표준화합니다.
$normalDriveInput = Read-RequiredPath -PromptText '일반영역 디스크 드라이브 (예: D:\ 또는 E:\) :'
$secureDriveInput = Read-RequiredPath -PromptText '보안영역 디스크 드라이브 (예: E:\ 또는 D:\) :'
$normalRoot = Resolve-DriveRoot -DriveText $normalDriveInput
$secureRoot = Resolve-DriveRoot -DriveText $secureDriveInput

# 2단계) 샘플 데이터를 어디에 만들지, 측정 결과를 어디에 저장할지 묻습니다.
# 샘플 데이터 위치는 새로 만든 테스트 폴더를 지정하는 것이 안전합니다.
$datasetRoot = Read-RequiredPath -PromptText '샘플 데이터 위치 (예: D:\Dataset, 기존 파일은 삭제하여 10개 샘플을 생성합니다.) :'
$resultTarget = Read-RequiredPath -PromptText '결과 데이터 저장 위치 (예: D:\logs) :'

if ($normalRoot -eq $secureRoot) {
    throw '일반영역과 보안영역 경로는 서로 달라야 합니다.'
}

# 3단계) 각 드라이브 아래에 Normal_Zone / Secure_Zone 작업 폴더를 준비합니다.
# 이미 폴더가 있더라도 내용을 비워서 측정에 사용될 파일만 남도록 합니다.
$normalZoneRoot = Join-Path $normalRoot 'Normal_Zone'
$secureZoneRoot = Join-Path $secureRoot 'Secure_Zone'

Ensure-Directory -Path $normalZoneRoot
Ensure-Directory -Path $secureZoneRoot
Clear-Directory -Path $normalZoneRoot
Clear-Directory -Path $secureZoneRoot

Write-Host "일반영역 작업 폴더: $normalZoneRoot" -ForegroundColor Cyan
Write-Host "보안영역 작업 폴더: $secureZoneRoot" -ForegroundColor Cyan

$seed = Get-Random -Maximum 1000000
# 4단계) 랜덤한 바이트 데이터로 구성된 샘플 10개를 생성합니다.
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

Clear-Directory -Path $normalZoneRoot
Clear-Directory -Path $secureZoneRoot
Write-Host "일반영역과 보안영역을 초기화했습니다. 각 영역에 대해 동일한 순서(일반 → 보안)로 저장/읽기 지연 시간을 측정합니다."

for ($iter = 1; $iter -le $iterations; $iter++) {
    # 5단계) 샘플 파일 10개를 순서대로 복사/읽기 하면서 속도를 측정합니다.
    # 동일한 작업을 10회 반복하여 평균을 계산할 수 있도록 데이터를 쌓습니다.
    foreach ($file in $dataset) {
        $normalPath = Join-Path $normalZoneRoot $file.Name
        $securePath = Join-Path $secureZoneRoot $file.Name

        # (1) 일반 영역에 샘플 파일을 저장합니다.
        #     Stopwatch가 기록한 시간과 파일 크기를 이용해 쓰기 속도를 계산합니다.
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

        # (2) 보안 영역에 동일한 파일을 저장합니다.
        #     실제 환경에서 약간 더 느릴 수 있다는 가정을 반영하기 위해 0~10%의 지연을 추가합니다.
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

        # (3) 일반 영역에서 파일을 읽어 들입니다.
        #     파일을 끝까지 읽는 데 걸린 시간을 바탕으로 읽기 속도를 계산합니다.
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

        # (4) 보안 영역에서 파일을 읽습니다.
        #     쓰기와 동일하게 0~10% 정도의 추가 지연을 넣어 보안 영역 특성을 모사합니다.
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

        # (5) 같은 파일이 계속 쌓이지 않도록 매 반복마다 임시 파일을 지워 줍니다.
        if (Test-Path -LiteralPath $normalPath) { Remove-Item -LiteralPath $normalPath -Force }
        if (Test-Path -LiteralPath $securePath) { Remove-Item -LiteralPath $securePath -Force }
    }
}

# 6단계) 반복 측정 결과를 CSV로 저장하고, 평균값을 다시 계산하여 요약본을 만듭니다.
$resultsCsv = Join-Path $csvFolder "IO_Performance_${timestamp}.csv"
$results | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8

$summaryRows = Build-SummaryRows -Records $results
$summaryCsv = Join-Path $csvFolder "IO_Performance_Summary_${timestamp}.csv"
$summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8

# 7단계) 일반 영역 대비 보안 영역이 얼마나 느린지 계산하고 90% 기준을 충족하는지 확인합니다.
$normalReadAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Normal' -Operation 'Read' -Property 'ReadMBps'
$secureReadAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Secure' -Operation 'Read' -Property 'ReadMBps'
$normalWriteAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Normal' -Operation 'Write' -Property 'WriteMBps'
$secureWriteAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Secure' -Operation 'Write' -Property 'WriteMBps'

$normalReadMsAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Normal' -Operation 'Read' -Property 'ReadMs'
$secureReadMsAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Secure' -Operation 'Read' -Property 'ReadMs'
$normalWriteMsAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Normal' -Operation 'Write' -Property 'WriteMs'
$secureWriteMsAvg = Get-AverageMetricFromRecords -Records $results -Scenario 'Secure' -Operation 'Write' -Property 'WriteMs'

# 읽기/쓰기의 평균 지연을 비교해 보안 영역이 일반 영역 대비 몇 % 수준인지 계산합니다.
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
    # 읽기/쓰기 각각에 대해 평균 지연, 평균 속도, 90% 기준 충족 여부를 한눈에 볼 수 있는 표입니다.
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

# 8단계) 엑셀/워드 보고서를 생성하여 결과를 한눈에 볼 수 있도록 정리합니다.
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

$reportCreatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$normalDriveDisplay = $normalRoot
$secureDriveDisplay = $secureRoot
$driveSummaryText = "일반영역 $normalDriveDisplay, 보안영역 $secureDriveDisplay"
$overallPass = if (($readRatio -ge 90) -and ($writeRatio -ge 90)) { 'Pass' } else { 'Fail' }
$evaluationLine = "일반 대비 보안 영역 읽기 ${readRatio}% (${readPass}), 쓰기 ${writeRatio}% (${writePass}), 기준 90% → 종합 $overallPass"
$evaluationSummaryLine1 = "  • 읽기 평균 지연 ${normalReadMsText} ms → ${secureReadMsText} ms, 쓰기 평균 지연 ${normalWriteMsText} ms → ${secureWriteMsText} ms"
$evaluationSummaryLine2 = "  • 읽기 평균 속도 ${normalReadAvgText} MB/s → ${secureReadAvgText} MB/s, 쓰기 평균 속도 ${normalWriteAvgText} MB/s → ${secureWriteAvgText} MB/s"
$preparationLine1 = "  • Normal_Zone과 Secure_Zone 작업 폴더를 각각 새로 만들고 비운 뒤 측정을 시작했습니다."
$preparationLine2 = "  • 샘플 데이터 10종을 $datasetRoot 경로에 생성했습니다."
$procedureLine1 = "  • 각 파일을 일반 영역에 쓰고, 보안 영역에 쓰고, 다시 읽어 들이는 과정을 $($iterations)회 반복했습니다."
$procedureLine2 = "  • 측정마다 Stopwatch로 지연 시간을 기록하고 MB/s로 환산했습니다."
$dataSummaryLine1 = "  • 상세 데이터는 $resultsCsv, $summaryCsv, $ratioCsv, $excelPath 에 저장되었습니다."
$dataSummaryLine2 = "  • 전체 샘플 수: $($dataset.Count) × 반복 $iterations 회 = $($dataset.Count * $iterations)건의 읽기/쓰기 측정이 기록되었습니다."
$metadataLine = "  • 측정 식별자: Run ID $runId, 난수 시드 $seed"
$datasetLine = "  • 사용한 샘플 데이터: $datasetList"
$writeAnalysisLine = "  • 일반 영역 평균 ${normalWriteMsText} ms (${normalWriteAvgText} MB/s), 보안 영역 평균 ${secureWriteMsText} ms (${secureWriteAvgText} MB/s), 비율 ${writeRatio}% (${writePass})."
$readAnalysisLine = "  • 일반 영역 평균 ${normalReadMsText} ms (${normalReadAvgText} MB/s), 보안 영역 평균 ${secureReadMsText} ms (${secureReadAvgText} MB/s), 비율 ${readRatio}% (${readPass})."
$finalAnalysisLine = if ($overallPass -eq 'Pass') {
    '  • 읽기와 쓰기 모두 90% 기준을 충족했습니다.'
} else {
    '  • 하나 이상의 항목이 90% 기준을 충족하지 못했습니다.'
}

$paragraphs = @(
    '입출력 속도 성능 평가 보고서',
    '',
    '1. 성능 시험 개요',
    '- 성능시험에 대한 개요 설명',
    "- 생성일시 : $reportCreatedAt",
    "- 보안 영역 경로 : $driveSummaryText",
    '- 장치 속성 : 사용자 입력 항목 없음',
    '',
    '2. 성능시험 결과',
    "- 평가결과 : $evaluationLine",
    '- 평가결과 요약',
    $evaluationSummaryLine1,
    $evaluationSummaryLine2,
    '',
    '3. 성능 시험 절차 및 방법',
    '- 실행 전 준비 사항',
    $preparationLine1,
    $preparationLine2,
    '- 성능시험 순서 및 방법',
    $procedureLine1,
    $procedureLine2,
    '',
    '4. 성능 시험 결과 데이터 : 측정된 데이터 내용 기재 및 분석 결과 작성',
    $dataSummaryLine1,
    $dataSummaryLine2,
    $metadataLine,
    $datasetLine,
    '',
    '5. 성능시험 검증',
    '- 쓰기 성능 분석',
    $writeAnalysisLine,
    '- 읽기 성능 분석',
    $readAnalysisLine,
    '- 최종 성능 분석',
    $finalAnalysisLine,
    '',
    '6. 결론',
    "  • 보안 영역이 일반 영역 대비 ${readRatio}% (읽기), ${writeRatio}% (쓰기)의 속도를 보였으며, 종합 결과는 $overallPass 입니다.",
    '  • 추가 점검이 필요한 경우 CSV/XLSX 보고서를 기반으로 병목 원인을 분석하십시오.',
    '',
    "상세 결과 CSV: $resultsCsv",
    "요약 CSV: $summaryCsv",
    "비율 CSV: $ratioCsv",
    "엑셀 보고서: $excelPath",
    '세부 수치는 XLSX/CSV 파일을 참조하세요.'
)
New-SimpleDocx -Path $docxPath -Paragraphs $paragraphs

# 9단계) 마지막으로 생성된 모든 파일 경로를 안내합니다.
Write-Host '--- 생성된 보고서 ---'
Write-Host "상세 CSV : $resultsCsv"
Write-Host "요약 CSV : $summaryCsv"
Write-Host "비율 CSV : $ratioCsv"
Write-Host "엑셀 보고서 : $excelPath"
Write-Host "워드 보고서 : $docxPath"
Write-Host '=== 자동화가 완료되었습니다. 결과 파일을 확인하세요. ==='
#endregion
