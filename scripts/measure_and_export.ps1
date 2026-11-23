param(
    [Parameter(Mandatory=$true)][string]$DocxPath,
    [Parameter(Mandatory=$true)][string]$PdfPath,
    [Parameter(Mandatory=$true)][string]$JsonOutPath,
    [Parameter(Mandatory=$false)][string]$TargetDocxPath = "",
    [Parameter(Mandatory=$false)][double]$MaxDeltaPt = 5.25,
    [Parameter(Mandatory=$false)][double]$MinBottomPt = 745
)

$lockDir = Join-Path $env:USERPROFILE ".gemini\tmp"
if (-not (Test-Path $lockDir)) { New-Item -ItemType Directory -Path $lockDir -Force | Out-Null }
$lockFile = Join-Path $lockDir "word_com.lock"
$lockStream = $null
$deadline = (Get-Date).AddSeconds(120)
while ((Get-Date) -lt $deadline) {
    try {
        $lockStream = [System.IO.File]::Open($lockFile, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        break
    } catch {
        Start-Sleep -Seconds 1
    }
}
if ($null -eq $lockStream) {
    Write-Error "Timed out waiting for word_com.lock"
    exit 1
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$cols = @{}
$wordPages = 0
try {
    $doc = $word.Documents.Open($DocxPath, $false, $false)
    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.Repaginate()
    $wordPages = [int]$doc.ComputeStatistics(2)

    $idx = 0
    $inRefs = $false
    $orphans = @()
    $lineStarts = @()
    foreach ($p in $doc.Paragraphs) {
        $idx++
        $inTbl = $p.Range.Information(12)
        $r1 = $p.Range
        $r1.Collapse(1)
        $pg1 = [int]$r1.Information(3)
        $x1 = [math]::Round([double]$r1.Information(5), 2)
        $y1 = [math]::Round([double]$r1.Information(6), 2)

        $r2 = $p.Range
        $r2.Collapse(0)
        $null = $r2.Move(1, -1)
        $pg2 = [int]$r2.Information(3)
        $x2 = [math]::Round([double]$r2.Information(5), 2)
        $y2 = [math]::Round([double]$r2.Information(6), 2)

        $plain = ($p.Range.Text -replace "[\r\a]", "").Trim()
        if ($plain -match "^参考文献\s*$") { $inRefs = $true }
        elseif ($plain -match "^(附录|作者简介|致\s*谢)") { $inRefs = $false }
        if ((-not $inTbl) -and (-not $inRefs) -and ($plain.Length -gt 1) -and ($plain -notmatch "^(图|表|Fig\.|Tab\.)")) {
            $bodyRange = $p.Range.Duplicate
            $null = $bodyRange.MoveEnd(1, -1)
            if ($bodyRange.End -gt $bodyRange.Start) {
                $firstLine = $bodyRange.Duplicate
                $firstLine.Collapse(1)
                $null = $firstLine.Expand(5)
                $lastLine = $bodyRange.Duplicate
                $lastLine.Collapse(0)
                $null = $lastLine.Expand(5)
                if ($firstLine.Start -ne $lastLine.Start) {
                    $lastText = ($lastLine.Text -replace "\s", "" -replace "[。．.，,；;：:、）)]", "")
                    if (($lastText.Length -eq 1) -and ($lastText -match "[\u4e00-\u9fff]")) {
                        $orphans += [ordered]@{ page = $pg1; text = $plain.Substring(0, [Math]::Min(40, $plain.Length)) }
                    }
                }
                $cursor = $bodyRange.Duplicate
                $cursor.Collapse(1)
                $guard = 0
                while (($cursor.Start -lt $bodyRange.End) -and ($guard -lt 40)) {
                    $guard++
                    $ln = $cursor.Duplicate
                    $null = $ln.Expand(5)
                    $ltxt = $ln.Text.TrimStart()
                    $isList = $ltxt -match "^\d+）"
                    $badRef = $ltxt -match "^\[\d+\]"
                    $badUnit = $ltxt -match "^(MW|kW|GW|Hz|kHz|mHz|kV|mm|cm|km|pu|%)\b"
                    if ((-not $isList) -and ($badRef -or $badUnit)) {
                        $lineStarts += [ordered]@{ page = $pg1; text = $ltxt.Substring(0, [Math]::Min(30, $ltxt.Length)) }
                    }
                    if ($ln.End -ge $bodyRange.End) { break }
                    $cursor.SetRange($ln.End, $ln.End)
                    $null = $cursor.Move(1, 1)
                }
            }
        }

        if ($idx -le 8) { continue }
        $col1 = if ($x1 -lt 290) { "L" } else { "R" }
        $col2 = if ($x2 -lt 290) { "L" } else { "R" }
        if ((-not $inTbl) -and (($pg1 -ne $pg2) -or ($col1 -ne $col2)) -and ($p.Range.Start -lt ($p.Range.End - 2))) {
            $lo = $p.Range.Start
            $hi = $p.Range.End - 1
            while (($hi - $lo) -gt 1) {
                $mid = [int](($lo + $hi) / 2)
                $rm = $doc.Range($mid, $mid + 1)
                $pgm = [int]$rm.Information(3)
                $xm = [double]$rm.Information(5)
                $colm = if ($xm -lt 290) { "L" } else { "R" }
                if (($pgm -eq $pg1) -and ($colm -eq $col1)) { $lo = $mid } else { $hi = $mid }
            }
            $rLast1 = $doc.Range($lo, $lo + 1)
            $rFirst2 = $doc.Range($hi, $hi + 1)
            $yLast1 = [math]::Round([double]$rLast1.Information(6), 2)
            $yFirst2 = [math]::Round([double]$rFirst2.Information(6), 2)
            $k1 = "$pg1-$col1"
            $k2 = "$pg2-$col2"
            if (-not $cols.ContainsKey($k1)) { $cols[$k1] = @() }
            if (-not $cols.ContainsKey($k2)) { $cols[$k2] = @() }
            $cols[$k1] += $y1, $yLast1
            $cols[$k2] += $yFirst2, $y2
        } else {
            $k1 = "$pg1-$col1"
            if (-not $cols.ContainsKey($k1)) { $cols[$k1] = @() }
            $cols[$k1] += $y1, $y2
        }
    }
    $doc.ExportAsFixedFormat($PdfPath, 17)
    $doc.Close($false)
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    $lockStream.Close()
    Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
}

if ($TargetDocxPath -ne "") {
    try {
        $activeWord = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
        foreach ($d in @($activeWord.Documents)) {
            if ($d.FullName -eq $TargetDocxPath) { $d.Close($false) }
        }
    } catch {}
    Copy-Item -Path $DocxPath -Destination $TargetDocxPath -Force
}

$pdfBytes = [System.IO.File]::ReadAllBytes($PdfPath)
$pdfStr = [System.Text.Encoding]::ASCII.GetString($pdfBytes)
$pdfPages = ([regex]::Matches($pdfStr, "/Type\s*/Page\b")).Count

$pageList = @()
$allAligned = $true
$twoColPages = 0
for ($pg = 1; $pg -le $wordPages; $pg++) {
    $L = $cols["$pg-L"]
    $R = $cols["$pg-R"]
    $lt = if ($L) { ($L | Measure-Object -Minimum).Minimum } else { 0.0 }
    $lb = if ($L) { ($L | Measure-Object -Maximum).Maximum } else { 0.0 }
    $rt = if ($R) { ($R | Measure-Object -Minimum).Minimum } else { 0.0 }
    $rb = if ($R) { ($R | Measure-Object -Maximum).Maximum } else { 0.0 }
    $dT = [math]::Round($lt - $rt, 2)
    $dB = [math]::Round($lb - $rb, 2)
    if (($lb -gt 0) -and ($rb -gt 0)) {
        $twoColPages++
        if (([math]::Abs($dT) -gt $MaxDeltaPt) -or ([math]::Abs($dB) -gt $MaxDeltaPt) -or ($lb -lt $MinBottomPt) -or ($rb -lt $MinBottomPt)) {
            $allAligned = $false
        }
    }
    $pageList += [ordered]@{
        page = $pg
        L_top = $lt
        R_top = $rt
        dT = $dT
        L_bot = $lb
        R_bot = $rb
        dB = $dB
    }
}

if ($twoColPages -eq 0) { $allAligned = $false }
$outObj = [ordered]@{
    word_pages = $wordPages
    pdf_pages = $pdfPages
    pages_match = ($wordPages -eq $pdfPages)
    max_delta_pt = $MaxDeltaPt
    min_bottom_pt = $MinBottomPt
    all_columns_aligned_within_5_25pt = $allAligned
    pages = $pageList
    line_breaks = [ordered]@{
        orphans = @($orphans)
        line_starts = @($lineStarts)
    }
}
$outObj | ConvertTo-Json -Depth 5 | Set-Content -Path $JsonOutPath -Encoding UTF8
Write-Host "Success! Metrics written to: $JsonOutPath (word_pages=$wordPages, pdf_pages=$pdfPages, all_aligned=$allAligned)"
