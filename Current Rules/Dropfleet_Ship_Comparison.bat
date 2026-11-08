@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem ---------------------------------------------------------------
rem Dropfleet Ship Image / Old Equivalent Comparison
rem Double-click this BAT from the root containing the OLD faction
rem folders. The NEW tree is configured below.
rem Requires Windows PowerShell 5.1 (included with Windows).
rem No Python, Excel, or PowerShell modules are required.
rem ---------------------------------------------------------------

set "DFC_NEW_ROOT=D:\DropFleetTTS_Repo\DropfleetTTS\RemasterShips"
set "DFC_OLD_ROOT=%~dp0"
set "DFC_BAT_PATH=%~f0"
set "DFC_PS_PATH=%TEMP%\Dropfleet_Comparison_%RANDOM%_%RANDOM%.ps1"

echo ============================================================
echo   DROPFLEET SHIP IMAGE COMPARISON
echo ============================================================
echo New folder: %DFC_NEW_ROOT%
echo Old folder: %DFC_OLD_ROOT%
echo.

rem Extract the embedded PowerShell section into a temporary script.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:DFC_BAT_PATH); $m=[regex]::Match($t,'(?m)^# POWERSHELL SECTION\r?\n'); if(-not $m.Success){throw 'Embedded PowerShell section not found.'}; [IO.File]::WriteAllText($env:DFC_PS_PATH,$t.Substring($m.Index+$m.Length),([Text.UTF8Encoding]::new($false)))"
if errorlevel 1 (
    echo ERROR: Could not extract the embedded script.
    goto :Failed
)

rem Read the root paths from environment variables in PowerShell. This avoids
rem Windows argument-quoting problems when %~dp0 ends in a backslash.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%DFC_PS_PATH%"
if errorlevel 1 goto :Failed

set "DFC_EXIT=0"
goto :Done

:Failed
set "DFC_EXIT=1"
echo.
echo Comparison failed. See the error message above.

:Done
if exist "%DFC_PS_PATH%" del /f /q "%DFC_PS_PATH%" >nul 2>&1
echo.
pause
exit /b %DFC_EXIT%

# POWERSHELL SECTION
# Root paths arrive through the inherited environment, not CLI arguments.
# PowerShell's -File argument parsing can corrupt a quoted path that ends in \.
$NewRoot = $env:DFC_NEW_ROOT
$OldRoot = $env:DFC_OLD_ROOT

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($NewRoot)) { throw 'DFC_NEW_ROOT is empty.' }
if ([string]::IsNullOrWhiteSpace($OldRoot)) { throw 'DFC_OLD_ROOT is empty.' }

function Say([string]$Message) {
    Write-Host $Message
}

function XmlEscape([object]$InputValue) {
    if ($null -eq $InputValue) { return '' }
    $s = [string]$InputValue
    # XML 1.0 only permits tabs, line feeds and carriage returns among controls.
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return [System.Security.SecurityElement]::Escape($s)
}

function ExcelColumn([int]$Index) {
    $out = ''
    while ($Index -gt 0) {
        $Index--
        $out = ([string][char](65 + ($Index % 26))) + $out
        $Index = [int][math]::Floor($Index / 26)
    }
    return $out
}

function MakeCell([int]$Column, [int]$Row, [object]$Value, [int]$Style = 0) {
    $ref = "$(ExcelColumn $Column)$Row"
    $styleAttr = if ($Style -gt 0) { ' s="' + $Style + '"' } else { '' }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) {
        return '<c r="' + $ref + '"' + $styleAttr + '><v>' + ([string]$Value) + '</v></c>'
    }
    $escaped = XmlEscape $Value
    return '<c r="' + $ref + '" t="inlineStr"' + $styleAttr + '><is><t xml:space="preserve">' + $escaped + '</t></is></c>'
}

function AddRow([System.Text.StringBuilder]$Builder, [int]$RowNumber, [object[]]$Values, [int]$Style = 0, [int]$Height = 0) {
    $rowAttr = if ($Height -gt 0) { ' ht="' + $Height + '" customHeight="1"' } else { '' }
    [void]$Builder.Append('<row r="' + $RowNumber + '"' + $rowAttr + '>')
    for ($i=0; $i -lt $Values.Count; $i++) {
        [void]$Builder.Append((MakeCell ($i+1) $RowNumber $Values[$i] $Style))
    }
    [void]$Builder.Append('</row>')
}

function AddMergedRow([System.Text.StringBuilder]$Builder, [int]$RowNumber, [string]$Value, [int]$Style) {
    [void]$Builder.Append('<row r="' + $RowNumber + '" ht="29" customHeight="1">')
    [void]$Builder.Append((MakeCell 1 $RowNumber $Value $Style))
    [void]$Builder.Append('</row>')
}

function NewWorksheetXml(
    [string]$Title,
    [string]$Subtitle,
    [object[]]$Columns,
    [object[]]$Rows,
    [int[]]$Widths,
    [bool]$ColorResults
) {
    $lastCol = ExcelColumn $Columns.Count
    $endRow = [math]::Max(4, 4 + $Rows.Count)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<dimension ref="A1:' + $lastCol + $endRow + '"/>')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/></sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="17"/>')
    [void]$sb.Append('<cols>')
    for ($i = 0; $i -lt $Widths.Count; $i++) {
        $n = $i + 1
        [void]$sb.Append('<col min="' + $n + '" max="' + $n + '" width="' + $Widths[$i] + '" customWidth="1"/>')
    }
    [void]$sb.Append('</cols><sheetData>')
    AddMergedRow $sb 1 $Title 1
    [void]$sb.Append('<row r="2">')
    [void]$sb.Append((MakeCell 1 2 $Subtitle 6))
    [void]$sb.Append('</row>')
    AddRow $sb 4 $Columns 2 25
    $r = 5
    foreach ($entry in $Rows) {
        $style = 0
        if ($ColorResults) {
            if (-not $entry.OldFound) { $style = 3 }
            elseif (-not ($entry.HasCard -and $entry.HasModel)) { $style = 4 }
            elseif (($r % 2) -eq 0) { $style = 5 }
        }
        $vals = @(
            $entry.Faction,
            $entry.Ship,
            $(if ($entry.HasCard) { 'YES' } else { 'NO' }),
            $(if ($entry.HasModel) { 'YES' } else { 'NO' }),
            $(if ($entry.OldFound) { 'YES' } else { 'NO' }),
            $entry.ImageStatus,
            $entry.MatchStatus,
            $entry.NewFiles,
            $entry.OldMatch
        )
        AddRow $sb $r $vals $style 18
        $r++
    }
    [void]$sb.Append('</sheetData>')
    [void]$sb.Append('<autoFilter ref="A4:' + $lastCol + $endRow + '"/>')
    [void]$sb.Append('<mergeCells count="2"><mergeCell ref="A1:' + $lastCol + '1"/><mergeCell ref="A2:' + $lastCol + '2"/></mergeCells>')
    [void]$sb.Append('<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function NewOverviewXml([object[]]$Items, [object[]]$Factions, [string]$NewPath, [string]$OldPath) {
    $total = $Items.Count
    $missing = @($Items | Where-Object { -not $_.OldFound }).Count
    $complete = @($Items | Where-Object { $_.HasCard -and $_.HasModel }).Count
    $incomplete = $total - $complete
    $matched = $total - $missing
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<dimension ref="A1:F' + ([math]::Max(16,12+$Factions.Count)) + '"/>')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="12" topLeftCell="A13" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A13" sqref="A13"/></sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="19"/>')
    [void]$sb.Append('<cols><col min="1" max="1" width="25" customWidth="1"/><col min="2" max="2" width="33" customWidth="1"/><col min="3" max="6" width="21" customWidth="1"/></cols>')
    [void]$sb.Append('<sheetData>')
    AddMergedRow $sb 1 'DROPFLEET - SHIP IMAGE AUDIT' 1
    AddRow $sb 3 @('New ships detected', $total) 0
    AddRow $sb 4 @('Old equivalents found', $matched) 0
    AddRow $sb 5 @('MISSING IN OLD', $missing) 3
    AddRow $sb 6 @('Both new images present', $complete) 0
    AddRow $sb 7 @('Incomplete image pairs', $incomplete) 4
    AddRow $sb 9 @('New image source', $NewPath) 0
    AddRow $sb 10 @('Old folders (BAT location)', $OldPath) 0
    AddRow $sb 11 @('Matching rule', 'Same faction + ship name, ignoring case') 0
    AddRow $sb 12 @('Faction', 'New ships', 'Both images', 'Missing image', 'In old', 'Missing in old') 2 25
    $r=13
    foreach ($f in $Factions) {
        AddRow $sb $r @($f.Name, $f.Total, $f.Both, $f.Incomplete, $f.Found, $f.Missing) $(if($f.Missing -gt 0){3}else{0}) 19
        $r++
    }
    [void]$sb.Append('</sheetData><mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells>')
    [void]$sb.Append('<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function AddZipEntry([System.IO.Compression.ZipArchive]$Zip, [string]$Name, [string]$Content) {
    $entry = $Zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
    try { $writer.Write($Content) }
    finally { $writer.Dispose() }
}

function SaveXlsx([string]$Path, [string]$Overview, [string]$AllShips, [string]$MissingShips) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
            '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
            '</Types>'
        $rootRel = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>'
        $workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
            '<sheet name="Overview" sheetId="1" r:id="rId1"/>' +
            '<sheet name="All New Ships" sheetId="2" r:id="rId2"/>' +
            '<sheet name="Missing in Old" sheetId="3" r:id="rId3"/>' +
            '</sheets></workbook>'
        $workbookRel = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>' +
            '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
            '</Relationships>'
        # Style IDs: 0 normal, 1 title, 2 header, 3 missing (red),
        # 4 incomplete (amber), 5 alternating row, 6 subtitle.
        $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<fonts count="4">' +
            '<font><sz val="11"/><color rgb="FF203047"/><name val="Calibri"/></font>' +
            '<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
            '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
            '<font><b/><sz val="11"/><color rgb="FF8A1B1B"/><name val="Calibri"/></font>' +
            '</fonts><fills count="6">' +
            '<fill><patternFill patternType="none"/></fill>' +
            '<fill><patternFill patternType="gray125"/></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FF17233D"/><bgColor indexed="64"/></patternFill></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FFFDE1E3"/><bgColor indexed="64"/></patternFill></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF0CC"/><bgColor indexed="64"/></patternFill></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FFF3F6FA"/><bgColor indexed="64"/></patternFill></fill>' +
            '</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="7">' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>' +
            '<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>' +
            '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
            '</styleSheet>'
        AddZipEntry $zip '[Content_Types].xml' $types
        AddZipEntry $zip '_rels/.rels' $rootRel
        AddZipEntry $zip 'xl/workbook.xml' $workbook
        AddZipEntry $zip 'xl/_rels/workbook.xml.rels' $workbookRel
        AddZipEntry $zip 'xl/styles.xml' $styles
        AddZipEntry $zip 'xl/worksheets/sheet1.xml' $Overview
        AddZipEntry $zip 'xl/worksheets/sheet2.xml' $AllShips
        AddZipEntry $zip 'xl/worksheets/sheet3.xml' $MissingShips
    }
    finally { $zip.Dispose() }
}

# New inventory: paired or unpaired image filenames are both represented.
if (-not (Test-Path -LiteralPath $NewRoot -PathType Container)) {
    throw "NEW source folder does not exist: $NewRoot"
}
if (-not (Test-Path -LiteralPath $OldRoot -PathType Container)) {
    throw "OLD source folder does not exist: $OldRoot"
}

function NormalizeRoot([string]$Path) {
    $normalized = (Get-Item -LiteralPath $Path).FullName
    if ($normalized -match '^[A-Za-z]:\$') { return $normalized }
    return $normalized.TrimEnd('\')
}
$NewRoot = NormalizeRoot $NewRoot
$OldRoot = NormalizeRoot $OldRoot
if ([string]::Equals($NewRoot, $OldRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The BAT is in RemasterShips itself; the new and old folder must be different.'
}

$newFactionDirs = @(Get-ChildItem -LiteralPath $NewRoot -Directory | Sort-Object Name)
if ($newFactionDirs.Count -eq 0) { throw 'No faction subfolders were found under RemasterShips.' }

$items = New-Object 'System.Collections.Generic.List[object]'
$stats = New-Object 'System.Collections.Generic.List[object]'
$columnNames = @('Faction', 'Ship', 'CardFrontImage', 'ModelImage', 'In old folder', 'New image status', 'Old match status', 'New image files (relative)', 'Old match path (relative)')

foreach ($faction in $newFactionDirs) {
    $ships = @{}
    $newFiles = @(Get-ChildItem -LiteralPath $faction.FullName -File -Recurse -ErrorAction Stop)
    foreach ($file in $newFiles) {
        if ($file.Name -notmatch '^(?<ship>.+)_(?<kind>CardFrontImage|ModelImage)\.png$') { continue }
        $ship = $Matches['ship']
        $kind = $Matches['kind']
        if (-not $ships.ContainsKey($ship)) {
            $ships[$ship] = [pscustomobject]@{
                Ship = $ship
                HasCard = $false
                HasModel = $false
                NewPaths = (New-Object 'System.Collections.Generic.List[string]')
            }
        }
        if ($kind -ieq 'CardFrontImage') { $ships[$ship].HasCard = $true }
        if ($kind -ieq 'ModelImage') { $ships[$ship].HasModel = $true }
        $relative = $file.FullName.Substring($NewRoot.Length).TrimStart('\')
        [void]$ships[$ship].NewPaths.Add($relative)
    }

    # Same-named faction in the OLD root. Recognition works with:
    #   <Ship>_CardFrontImage.png / <Ship>_ModelImage.png,
    #   a file whose full basename is <Ship> (e.g. <Ship>.obj), or
    #   a directory whose name is exactly <Ship>.
    $oldIndex = @{}
    $oldFaction = Join-Path $OldRoot $faction.Name
    if (Test-Path -LiteralPath $oldFaction -PathType Container) {
        foreach ($file in (Get-ChildItem -LiteralPath $oldFaction -File -Recurse -ErrorAction Stop)) {
            $name = $file.BaseName
            if ($file.Name -match '^(?<oldShip>.+)_(?:CardFrontImage|ModelImage)\.png$') {
                $name = $Matches['oldShip']
            }
            if (-not $oldIndex.ContainsKey($name)) {
                $oldIndex[$name] = $file.FullName.Substring($OldRoot.Length).TrimStart('\')
            }
        }
        foreach ($folder in (Get-ChildItem -LiteralPath $oldFaction -Directory -Recurse -ErrorAction Stop)) {
            if (-not $oldIndex.ContainsKey($folder.Name)) {
                $oldIndex[$folder.Name] = $folder.FullName.Substring($OldRoot.Length).TrimStart('\')
            }
        }
    }

    $localCount = 0
    $localBoth = 0
    $localFound = 0
    foreach ($shipName in @($ships.Keys | Sort-Object)) {
        $entry = $ships[$shipName]
        $found = $oldIndex.ContainsKey($shipName)
        $imageStatus = if ($entry.HasCard -and $entry.HasModel) { 'BOTH PRESENT' }
                       elseif (-not $entry.HasCard) { 'CARD MISSING' }
                       else { 'MODEL MISSING' }
        $items.Add([pscustomobject]@{
            Faction = $faction.Name
            Ship = $shipName
            HasCard = $entry.HasCard
            HasModel = $entry.HasModel
            OldFound = $found
            ImageStatus = $imageStatus
            MatchStatus = $(if ($found) { 'MATCHED' } else { 'MISSING IN OLD' })
            NewFiles = (($entry.NewPaths | Sort-Object) -join ' ; ')
            OldMatch = $(if ($found) { $oldIndex[$shipName] } else { '' })
        })
        $localCount++
        if ($entry.HasCard -and $entry.HasModel) { $localBoth++ }
        if ($found) { $localFound++ }
    }
    $stats.Add([pscustomobject]@{
        Name = $faction.Name
        Total = $localCount
        Both = $localBoth
        Incomplete = $localCount - $localBoth
        Found = $localFound
        Missing = $localCount - $localFound
    })
    Say ("{0}: {1} new ships, {2} missing in old, {3} missing image(s)" -f $faction.Name, $localCount, ($localCount-$localFound), ($localCount-$localBoth))
}

$all = @($items.ToArray() | Sort-Object Faction, Ship)
$missing = @($all | Where-Object { -not $_.OldFound })
$stamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$output = Join-Path $OldRoot ("Dropfleet_Ship_Comparison_{0}.xlsx" -f $stamp)
$increment = 1
while (Test-Path -LiteralPath $output) {
    $output = Join-Path $OldRoot ("Dropfleet_Ship_Comparison_{0}_{1}.xlsx" -f $stamp, $increment)
    $increment++
}

$overviewXml = NewOverviewXml $all @($stats.ToArray()) $NewRoot $OldRoot
$allXml = NewWorksheetXml 'ALL NEW SHIPS' ("{0} ships across {1} faction folders; red = no old equivalent; amber = an image is missing" -f $all.Count, $newFactionDirs.Count) $columnNames $all @(23,36,19,17,19,22,23,76,65) $true
$missingXml = NewWorksheetXml 'NEW SHIPS MISSING FROM OLD' ("{0} ships need an old-folder equivalent" -f $missing.Count) $columnNames $missing @(23,36,19,17,19,22,23,76,65) $true
SaveXlsx $output $overviewXml $allXml $missingXml

Say ''
Say '============================================================'
Say ("TOTAL NEW SHIPS:           {0}" -f $all.Count)
Say ("OLD EQUIVALENTS FOUND:     {0}" -f ($all.Count - $missing.Count))
Say ("NEW SHIPS MISSING IN OLD: {0}" -f $missing.Count)
Say ("INCOMPLETE NEW IMAGE PAIRS: {0}" -f @($all | Where-Object { -not ($_.HasCard -and $_.HasModel) }).Count)
Say '============================================================'
if ($missing.Count -gt 0) {
    Say 'MISSING OLD EQUIVALENTS (first 30):'
    foreach ($ship in @($missing | Select-Object -First 30)) {
        Say ('  [{0}] {1}' -f $ship.Faction, $ship.Ship)
    }
    if ($missing.Count -gt 30) { Say ('  ... and {0} more; see Excel.' -f ($missing.Count - 30)) }
    Say ''
}
Say "Excel report saved to: $output"
Say 'Open the "Missing in Old" tab for the highlighted action list.'
