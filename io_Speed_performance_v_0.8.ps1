#requires -Version 5.1
<#!
    입출력 속도 성능평가 자동화 스크립트
    ------------------------------------------------------------
    이 스크립트는 일반/보안 영역으로 사용할 **드라이브 문자**를 입력받아
    같은 드라이브를 선택하더라도 Normal_Zone, Secure_Zone 폴더로 분리한 뒤,
    샘플 데이터를 재구성하여 10개 샘플의 쓰기/읽기 속도를 여러 차례 측정합니다.
    측정 결과는 CSV·XLSX·DOCX 형태로 저장되며, 보안영역 성능이 일반영역의
    90% 이상 유지되는지 자동으로 분석합니다.

    관리자 권한 PowerShell 콘솔에서 실행하고, 테스트 전용 경로를 사용하십시오.
!#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Crc32Table = $null

#region 공통 유틸리티
# 아래 함수들은 스크립트 전반에서 반복해서 사용하는 기본 도우미들입니다.
# 파일·폴더 생성, 사용자 입력 정리, 간단한 수학 계산을 담당하며
# 각 기능이 어떤 역할을 하는지 한눈에 파악할 수 있도록 상세 설명을 덧붙였습니다.
function Read-RequiredPath {
    param([Parameter(Mandatory)] [string] $PromptText)
    # 사용자가 비워 두면 안 되는 값을 안전하게 받을 때 활용합니다.
    # 문자열이 입력될 때까지 계속 반복해서 질문하므로, 잘못 입력해도 다시 시도할 수 있습니다.
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
    # 폴더가 없으면 새로 만들고, 이미 있으면 그대로 둡니다.
    # 보고서 저장 경로, 데이터셋 폴더 등 반드시 존재해야 하는 위치에 사용합니다.
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Clear-Directory {
    param([string] $Path)
    # 이전 측정의 찌꺼기가 남지 않도록 지정한 폴더 안의 모든 파일을 비웁니다.
    # 비워진 폴더에 다시 데이터를 채우는 작업이 이어지므로, 삭제 실패 시에도 안내 메시지를 보여줍니다.
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
    # 샘플 데이터 파일을 생성할 때 사용하는 함수입니다.
    # 대상 폴더가 존재하지 않으면 먼저 만든 뒤, 지정된 바이트 배열을 그대로 파일로 씁니다.
    $folder = Split-Path -Path $Path -Parent
    Ensure-Directory -Path $folder
    [System.IO.File]::WriteAllBytes($Path, $Bytes)
}

function Format-Nullable {
    param([double] $Value, [int] $Digits = 3)
    # 계산 결과가 없는 경우(null)는 그대로 유지하고, 값이 있을 때만 지정한 자리수로 반올림합니다.
    # 보고서에 소수점 자릿수를 맞춰 깔끔하게 보여주기 위해 사용합니다.
    if ($null -eq $Value) { return $null }
    return [math]::Round($Value, $Digits)
}

function Resolve-DriveRoot {
    param([Parameter(Mandatory)][string] $DriveText)

    # 드라이브 입력을 다양한 형태(C, C:, C:\ 등)로 받아도 공통된 "C:\" 형식으로 바꿉니다.
    # 폴더 경로 대신 드라이브 문자만 제공해도 안전하게 처리할 수 있도록 구성했습니다.

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

function Read-DriveRoot {
    param([string] $PromptText)

    # 사용자에게 드라이브 문자를 물어보고, `Resolve-DriveRoot`를 통해 표준화한 값을 반환합니다.
    # 잘못된 문자를 입력하면 예외 메시지를 보여 준 뒤 재입력을 요청합니다.

    while ($true) {
        $input = Read-Host -Prompt $PromptText
        if ([string]::IsNullOrWhiteSpace($input)) {
            Write-Host '드라이브 문자를 입력해야 합니다. 다시 시도하세요.' -ForegroundColor Yellow
            continue
        }

        try {
            return Resolve-DriveRoot -DriveText $input
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Yellow
        }
    }
}

function Get-ColumnName {
    param([int] $Index)
    # 엑셀 셀 주소를 만들기 위해 숫자 인덱스를 A, B, ... AA 같은 열 이름으로 바꿉니다.
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

    # ZIP, XLSX, DOCX 같은 압축 파일을 직접 만들 때 필요한 CRC32(무결성 검사값)를 구합니다.
    # 처음 호출될 때는 테이블을 계산해 보관하고, 이후에는 계산된 테이블을 재활용합니다.

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

    # ZIP 파일을 직접 생성하는 함수입니다.
    # CSV, XLSX, DOCX 파일은 내부적으로 ZIP 구조를 사용하므로, 필요한 조각을 직접 조립합니다.
    # 초보자도 흐름을 이해할 수 있도록 각 단계마다 설명을 덧붙였습니다.

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
    $directory = Split-Path -Path $Path -Parent
    Ensure-Directory -Path $directory

    $encoding = [System.Text.Encoding]::UTF8
    # 파일 스트림을 열어 실제 데이터를 순서대로 기록합니다.
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    $writer = New-Object System.IO.BinaryWriter($fs)
    $centralRecords = @()

    try {
        foreach ($entry in $Entries) {
            # 항목이 해시 테이블 형식인지, PowerShell 객체인지 구분해 이름과 내용을 꺼냅니다.
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

            # ZIP 파일 구조의 "로컬 파일 헤더" 부분을 작성합니다.
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
        # 모든 엔트리를 기록한 뒤, 중앙 디렉터리 영역을 작성합니다.
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
        # ZIP 파일의 끝(End of Central Directory)을 작성해 구조를 마무리합니다.
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
    # PowerShell 객체 목록을 단순한 Excel 시트(XML) 구조로 변환합니다.
    # 엑셀 파일은 사실상 XML 문서를 ZIP으로 묶은 형태이므로,
    # 한 줄씩 순회하면서 셀 값을 문자열로 작성합니다.
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

    # 최소한의 요소만 사용해 XLSX 파일을 직접 만듭니다.
    # 여러 시트를 넘겨받으면, 각 시트를 XML로 만들어 ZIP 안에 배치합니다.
    # Excel이 없어도 자동화된 보고서를 만들 수 있는 핵심 역할을 합니다.

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

    # Word 문서도 ZIP 구조이므로, 본문 텍스트를 단락 단위로 받아 간단한 DOCX 파일을 생성합니다.
    # 보고서 요구사항에 맞춰 문단을 나열하면 Word 없이도 결과 문서를 얻을 수 있습니다.

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

    # 테스트용 샘플 데이터를 만드는 구간입니다.
    # 동일한 파일 목록을 반복해서 사용하면 캐시가 생겨 결과가 왜곡될 수 있으므로,
    # 매번 폴더를 비운 뒤 난수로 채운 파일 10개를 새로 생성합니다.

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
        # 설정된 크기(MB)를 바이트 단위로 환산해 동일한 용량의 파일을 만듭니다.
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
    # 파일을 복사하여 쓰기 속도를 측정합니다.
    # 고유 Stopwatch를 사용해 밀리초 단위 시간을 구하고, 시작/종료 시각을 기록합니다.
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
    # 파일 전체를 읽어 들여 읽기 속도를 측정합니다.
    # 일정 크기의 버퍼로 끝까지 반복 읽기하며, 실제 읽은 시간만 기록합니다.
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
    # 측정된 모든 항목을 시나리오(일반/보안)와 작업 종류(쓰기/읽기)로 묶어 평균값을 계산합니다.
    # 엑셀 요약 시트와 워드 보고서 모두 이 데이터를 기반으로 작성됩니다.
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
    # 특정 시나리오·작업 조합에 대한 평균 MB/s 값을 꺼내는 편의 함수입니다.
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
    # 특정 시나리오·작업 조합에 대한 평균 지연 시간(ms) 값을 꺼내는 편의 함수입니다.
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

    # 측정 원본 데이터에서 원하는 속성(ReadMBps, WriteMs 등)을 골라 평균을 계산합니다.
    # 값이 비어 있는 경우를 제외하고 계산하므로, 정확한 평균을 얻을 수 있습니다.

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
# 아래부터는 실제 측정을 진행하는 본문입니다.
# 1) 사용자 입력 → 2) 작업 폴더 준비 → 3) 샘플 데이터 생성 →
# 4) 쓰기/읽기 측정 → 5) 통계 계산 → 6) 보고서 저장 순서로 이어집니다.
Write-Host '=== 입출력 속도 성능평가 자동화 시작 ==='
$normalDriveRoot = Read-DriveRoot -PromptText '일반영역 디스크 드라이브 (예: D:\\ 또는 E:\\) :'
$secureDriveRoot = Read-DriveRoot -PromptText '보안영역 디스크 드라이브 (예: E:\\ 또는 D:\\) :'
$datasetRoot = Read-RequiredPath -PromptText '샘플 데이터 위치 (예: D:\\Dataset, 기존 파일은 삭제하여 10개 샘플을 생성합니다.) :'
$resultTarget = Read-RequiredPath -PromptText '결과 데이터 저장 위치 (예: D:\\logs) :'

$normalRoot = Join-Path $normalDriveRoot 'Normal_Zone'
$secureRoot = Join-Path $secureDriveRoot 'Secure_Zone'

# 동일한 드라이브를 선택해도 Normal_Zone / Secure_Zone 폴더 이름이 다르기 때문에
# 폴더 경로는 충돌하지 않습니다. 같은 드라이브를 입력하더라도 그대로 진행합니다.

Ensure-Directory -Path $normalRoot
Ensure-Directory -Path $secureRoot
Clear-Directory -Path $normalRoot
Clear-Directory -Path $secureRoot

Write-Host "일반영역 작업 폴더: $normalRoot" -ForegroundColor Cyan
Write-Host "보안영역 작업 폴더: $secureRoot" -ForegroundColor Cyan

# 동일한 시드를 재사용하면 동일한 샘플 데이터가 만들어져 재현성이 확보됩니다.
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
    # 한 번의 반복(iteration) 안에서 10개 샘플 파일을 모두 측정합니다.
    foreach ($file in $dataset) {
        # 각 샘플 파일이 일반/보안 영역에 어떤 이름으로 복사될지 미리 계산합니다.
        $normalPath = Join-Path $normalRoot $file.Name
        $securePath = Join-Path $secureRoot $file.Name

        # 일반영역 저장: 원본 데이터를 일반 영역으로 복사하며 시간을 측정합니다.
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

        # 보안영역 저장: 같은 파일을 보안 영역에도 복사해 결과를 비교합니다.
        $secureWrite = Measure-WriteOperation -SourcePath $file.Path -DestinationPath $securePath
        $secureWriteMs = $secureWrite.DurationMs
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
            WriteStart = $secureWrite.StartTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
            WriteEnd = $secureWrite.EndTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
            WriteDurationMs = Format-Nullable -Value $secureWriteMs -Digits 3
            ReadStart = $null
            ReadEnd = $null
            ReadDurationMs = $null
            ReadMs = $null
            WriteMs = Format-Nullable -Value $secureWriteMs -Digits 3
            ReadMBps = $null
            WriteMBps = Format-Nullable -Value $secureWriteMBps -Digits 2
            Timestamp = $secureWrite.EndTime.ToString('yyyy-MM-dd HH:mm:ss')
        }) | Out-Null

        # 일반영역 읽기: 일반 영역에 저장된 파일을 다시 읽어 들여 읽기 지연 시간을 구합니다.
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

        # 보안영역 읽기: 보안 영역에서도 동일하게 읽기 속도를 확인합니다.
        $secureRead = Measure-ReadOperation -Path $securePath
        $secureReadMs = $secureRead.DurationMs
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
            ReadStart = $secureRead.StartTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
            ReadEnd = $secureRead.EndTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
            ReadDurationMs = Format-Nullable -Value $secureReadMs -Digits 3
            ReadMs = Format-Nullable -Value $secureReadMs -Digits 3
            WriteMs = $null
            ReadMBps = Format-Nullable -Value $secureReadMBps -Digits 2
            WriteMBps = $null
            Timestamp = $secureRead.EndTime.ToString('yyyy-MM-dd HH:mm:ss')
        }) | Out-Null

        # 다음 반복에서 이전 파일이 영향을 주지 않도록 즉시 삭제합니다.
        if (Test-Path -LiteralPath $normalPath) { Remove-Item -LiteralPath $normalPath -Force }
        if (Test-Path -LiteralPath $securePath) { Remove-Item -LiteralPath $securePath -Force }
    }
}

# 측정 데이터를 CSV 형태로 저장하면 엑셀 없이도 결과를 확인할 수 있습니다.
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
# Excel 보고서에는 요약(Summary) 시트와 상세(Details) 시트를 함께 담습니다.
New-SimpleWorkbook -Path $excelPath -Sheets @(
    @{ Name = 'Summary'; Rows = $ratioRows },
    @{ Name = 'Details'; Rows = $results }
)

$docxPath = Join-Path $reportDirectory "IO_Performance_Analysis_Report_${timestamp}.docx"
# Word 보고서는 요구된 6개 절차를 그대로 따르도록 단락을 구성합니다.
$datasetList = ($dataset | Select-Object -ExpandProperty Name) -join ', '

$normalReadMsText = if ($normalReadMsAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $normalReadMsAvg) }
$secureReadMsText = if ($secureReadMsAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $secureReadMsAvg) }
$normalWriteMsText = if ($normalWriteMsAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $normalWriteMsAvg) }
$secureWriteMsText = if ($secureWriteMsAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $secureWriteMsAvg) }
$normalReadAvgText = if ($null -eq $normalReadAvg -or $normalReadAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $normalReadAvg) }
$secureReadAvgText = if ($null -eq $secureReadAvg -or $secureReadAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $secureReadAvg) }
$normalWriteAvgText = if ($null -eq $normalWriteAvg -or $normalWriteAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $normalWriteAvg) }
$secureWriteAvgText = if ($null -eq $secureWriteAvg -or $secureWriteAvg -le 0) { 'N/A' } else { [string]::Format('{0:N2}', $secureWriteAvg) }

# 보고서 본문에 사용할 문자열을 한곳에서 정리합니다.
$reportCreatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$normalDriveDisplay = $normalDriveRoot
$secureDriveDisplay = $secureDriveRoot
$driveSummaryText = "일반영역 $normalDriveDisplay, 보안영역 $secureDriveDisplay"
$overallPass = if (($readRatio -ge 90) -and ($writeRatio -ge 90)) { 'Pass' } else { 'Fail' }
$evaluationLine = "일반 대비 보안 영역 읽기 ${readRatio}% (${readPass}), 쓰기 ${writeRatio}% (${writePass}), 기준 90% → 종합 $overallPass"
# Word 보고서 각 절에 들어갈 핵심 문장을 미리 변수로 만들어 둡니다.
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

# 위에서 준비한 문장들을 6개 장(section) 순서대로 배열해 Word 문서에 넣습니다.
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

Write-Host '--- 생성된 보고서 ---'
Write-Host "상세 CSV : $resultsCsv"
Write-Host "요약 CSV : $summaryCsv"
Write-Host "비율 CSV : $ratioCsv"
Write-Host "엑셀 보고서 : $excelPath"
Write-Host "워드 보고서 : $docxPath"
Write-Host '=== 자동화가 완료되었습니다. 결과 파일을 확인하세요. ==='
#endregion
