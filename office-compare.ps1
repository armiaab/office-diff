param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Old,
    [Parameter(Mandatory = $true, Position = 1)][string]$New,
    [Parameter(Position = 2)][string]$Name,
    [switch]$ByDate
)
$ErrorActionPreference = 'Stop'

function Test-Missing([string]$file) {
    ($file -in '/dev/null', 'nul') -or -not (Test-Path -LiteralPath $file -PathType Leaf)
}

function Get-DocumentKind([string]$extension) {
    switch -Regex ($extension) {
        '^\.(doc|docx|docm)$' { 'word' }
        '^\.(ppt|pptx|pptm)$' { 'powerpoint' }
        '^\.(xls|xlsx|xlsm)$' { 'excel' }
    }
}

function Show-TextDiff([string]$oldFile, [string]$newFile) {
    git -c core.autocrlf=false -c core.safecrlf=false --no-pager diff --no-index --no-ext-diff --no-textconv --color -- $oldFile $newFile
}

function Get-Sha256([string]$file) {
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try { [BitConverter]::ToString($sha256.ComputeHash([IO.File]::ReadAllBytes($file))) -replace '-' }
    finally { $sha256.Dispose() }
}

function New-OfficeApp([string]$progId, [string]$product) {
    try { New-Object -ComObject $progId }
    catch { throw "$product (desktop app) is needed for this comparison but could not be started: $($_.Exception.Message)" }
}

function New-WorkFolder {
    $root = Join-Path $env:TEMP 'office-diff'
    Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    $folder = Join-Path $root (Get-Date -Format 'yyyyMMdd-HHmmss-fff')
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $folder
}

function Copy-Version([string]$source, [string]$destination) {
    Copy-Item -LiteralPath $source -Destination $destination
    Set-ItemProperty -LiteralPath $destination -Name IsReadOnly -Value $false
    Unblock-File -LiteralPath $destination -ErrorAction SilentlyContinue
    $destination
}

function Compare-WordDocuments([string]$oldCopy, [string]$newCopy) {
    $wdCompareDestinationNew = 2
    $wdGranularityWordLevel = 1
    $wdDoNotSaveChanges = 0
    $wdAlertsNone = 0
    $wdAlertsAll = -1
    $wdRevisionsMarkupAll = 2
    $wdRevisionsViewFinal = 0
    $confirmConversions = $false
    $readOnly = $true
    $addToRecentFiles = $false
    $compareFormatting = $compareCaseChanges = $compareWhitespace = $compareTables = $compareHeaders = $true
    $compareFootnotes = $compareTextboxes = $compareFields = $compareComments = $detectMoves = $true

    $word = New-OfficeApp 'Word.Application' 'Microsoft Word'
    $word.Visible = $true
    $word.DisplayAlerts = $wdAlertsNone
    $original = $revised = $null
    try {
        $original = $word.Documents.Open($oldCopy, $confirmConversions, $readOnly, $addToRecentFiles)
        $revised = $word.Documents.Open($newCopy, $confirmConversions, $readOnly, $addToRecentFiles)
        foreach ($document in $original, $revised) {
            if ($document.Revisions.Count) { $document.Revisions.AcceptAll() }
        }
        $result = $word.CompareDocuments($original, $revised, $wdCompareDestinationNew, $wdGranularityWordLevel,
            $compareFormatting, $compareCaseChanges, $compareWhitespace, $compareTables, $compareHeaders,
            $compareFootnotes, $compareTextboxes, $compareFields, $compareComments, $detectMoves, 'New version')
    } finally {
        foreach ($document in $original, $revised) {
            if ($document) { $document.Close($wdDoNotSaveChanges) }
        }
        $word.DisplayAlerts = $wdAlertsAll
    }
    $result.ActiveWindow.View.RevisionsFilter.Markup = $wdRevisionsMarkupAll
    $result.ActiveWindow.View.RevisionsFilter.View = $wdRevisionsViewFinal
    $result.Saved = $true
    $result.Activate()
    $word.Activate()
}

function Get-SlideTexts([string]$deckFile, [string]$textFile) {
    & (Join-Path $PSScriptRoot 'office-textconv.ps1') -Path $deckFile -OutFile $textFile
    $linesBySlide = @{}
    $current = $null
    foreach ($line in [IO.File]::ReadAllLines($textFile)) {
        if ($line -match '^=== Slide (\d+)') {
            $current = [Collections.Generic.List[string]]::new()
            $linesBySlide[[int]$Matches[1]] = $current
        } elseif ($null -ne $current) {
            $current.Add($line)
        }
    }
    $texts = @{}
    foreach ($number in $linesBySlide.Keys) { $texts[$number] = ($linesBySlide[$number] -join "`n").Trim() }
    $texts
}

function Export-Slides($powerPoint, [string]$deckFile, [string]$folder) {
    $msoTrue = -1
    $msoFalse = 0
    $ppSaveAsOpenXMLPresentation = 24
    $imageWidth = 1280
    $readOnly = $msoTrue
    $untitled = $msoFalse
    $withWindow = $msoFalse

    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $folderName = Split-Path -Leaf $folder
    $deck = $powerPoint.Presentations.Open($deckFile, $readOnly, $untitled, $withWindow)
    try {
        $imageHeight = [int][Math]::Round($imageWidth * $deck.PageSetup.SlideHeight / $deck.PageSetup.SlideWidth)
        $textSource = $deckFile
        if ([IO.Path]::GetExtension($deckFile) -eq '.ppt') {
            $textSource = Join-Path $folder 'converted.pptx'
            $deck.SaveCopyAs($textSource, $ppSaveAsOpenXMLPresentation)
        }
        $activity = "Rendering $(Split-Path -Leaf $deckFile)"
        $total = $deck.Slides.Count
        $slides = foreach ($slide in $deck.Slides) {
            $number = $slide.SlideIndex
            Write-Progress -Activity $activity -Status "Slide $number of $total" -PercentComplete (100 * $number / $total)
            $image = Join-Path $folder "slide$number.png"
            $slide.Export($image, 'PNG', $imageWidth, $imageHeight)
            [pscustomobject]@{
                Id          = $slide.SlideID
                Number      = $number
                Hidden      = $slide.SlideShowTransition.Hidden -eq $msoTrue
                Image       = "$folderName/slide$number.png"
                Fingerprint = Get-Sha256 $image
                Text        = ''
            }
        }
        Write-Progress -Activity $activity -Completed
    } finally {
        $deck.Close()
    }
    $texts = Get-SlideTexts $textSource (Join-Path $folder 'slides.txt')
    foreach ($slide in $slides) { $slide.Text = [string]$texts[$slide.Number] }
    , @($slides)
}

function Get-SlidesInPlace([int[]]$oldNumbers) {
    $count = $oldNumbers.Count
    $runLength = [int[]]::new($count)
    $previous = [int[]]::new($count)
    $best = -1
    for ($i = 0; $i -lt $count; $i++) {
        $runLength[$i] = 1
        $previous[$i] = -1
        for ($j = 0; $j -lt $i; $j++) {
            if ($oldNumbers[$j] -lt $oldNumbers[$i] -and $runLength[$j] + 1 -gt $runLength[$i]) {
                $runLength[$i] = $runLength[$j] + 1
                $previous[$i] = $j
            }
        }
        if ($best -lt 0 -or $runLength[$i] -gt $runLength[$best]) { $best = $i }
    }
    $inPlace = @{}
    for ($i = $best; $i -ge 0; $i = $previous[$i]) { $inPlace[$oldNumbers[$i]] = $true }
    $inPlace
}

function New-SlidePair($oldSlide, $newSlide, [bool]$moved) {
    $status = if (-not $oldSlide) { 'added' }
        elseif (-not $newSlide) { 'removed' }
        elseif ($oldSlide.Fingerprint -ne $newSlide.Fingerprint -or $oldSlide.Text -cne $newSlide.Text -or $oldSlide.Hidden -ne $newSlide.Hidden) { 'changed' }
        elseif ($moved) { 'moved' }
        else { 'unchanged' }
    [pscustomobject]@{
        status    = $status
        moved     = $moved
        oldNumber = $oldSlide.Number
        newNumber = $newSlide.Number
        oldImage  = $oldSlide.Image
        newImage  = $newSlide.Image
        oldText   = $oldSlide.Text
        newText   = $newSlide.Text
        oldHidden = [bool]$oldSlide.Hidden
        newHidden = [bool]$newSlide.Hidden
    }
}

function Join-Slides($oldSlides, $newSlides) {
    $oldIds = @{}
    foreach ($slide in $oldSlides) { $oldIds[$slide.Id] = $true }
    $sharesSlideIds = @($newSlides | Where-Object { $oldIds.ContainsKey($_.Id) }).Count -gt 0
    $key = if ($sharesSlideIds) { 'Id' } else { 'Number' }

    $oldByKey = @{}
    foreach ($slide in $oldSlides) { $oldByKey[$slide.$key] = $slide }
    $newKeys = @{}
    foreach ($slide in $newSlides) { $newKeys[$slide.$key] = $true }
    $removed = @($oldSlides | Where-Object { -not $newKeys.ContainsKey($_.$key) })
    $keptInNewOrder = @(foreach ($slide in $newSlides) { $oldByKey[$slide.$key] }) | Where-Object { $_ }
    $inPlace = Get-SlidesInPlace @($keptInNewOrder | ForEach-Object Number)

    $nextRemoved = 0
    foreach ($newSlide in $newSlides) {
        $oldSlide = $oldByKey[$newSlide.$key]
        if ($oldSlide) {
            while ($nextRemoved -lt $removed.Count -and $removed[$nextRemoved].Number -lt $oldSlide.Number) {
                New-SlidePair $removed[$nextRemoved] $null $false
                $nextRemoved++
            }
        }
        New-SlidePair $oldSlide $newSlide ($oldSlide -and -not $inPlace.ContainsKey($oldSlide.Number))
    }
    for (; $nextRemoved -lt $removed.Count; $nextRemoved++) {
        New-SlidePair $removed[$nextRemoved] $null $false
    }
}

function Write-SlideReport([string]$folder, [string]$title, [string]$oldLabel, [string]$newLabel, [object[]]$pairs) {
    $data = [pscustomobject]@{
        title    = $title
        oldLabel = $oldLabel
        newLabel = $newLabel
        slides   = $pairs
    }
    $json = (ConvertTo-Json -InputObject $data -Depth 4 -Compress) -replace '<', '<'
    $dataBlock = '<script type="application/json" id="report-data">{0}</script>'
    $template = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'slide-report.html'))
    $report = Join-Path $folder 'report.html'
    $html = $template.Replace(($dataBlock -f 'null'), ($dataBlock -f $json))
    [IO.File]::WriteAllText($report, $html, [Text.UTF8Encoding]::new($false))
    $report
}

function Compare-Presentations([string]$oldCopy, [string]$newCopy, [string]$work, [string]$title, [string]$oldLabel, [string]$newLabel) {
    $powerPointWasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
    $powerPoint = New-OfficeApp 'PowerPoint.Application' 'Microsoft PowerPoint'
    try {
        $oldSlides = Export-Slides $powerPoint $oldCopy (Join-Path $work 'old')
        $newSlides = Export-Slides $powerPoint $newCopy (Join-Path $work 'new')
    } finally {
        if (-not $powerPointWasRunning -and $powerPoint.Presentations.Count -eq 0) { $powerPoint.Quit() }
    }

    $pairs = @(Join-Slides $oldSlides $newSlides)
    $report = Write-SlideReport $work $title $oldLabel $newLabel $pairs
    $summary = foreach ($status in 'changed', 'added', 'removed', 'moved', 'unchanged') {
        $count = @($pairs | Where-Object status -eq $status).Count
        if ($count) { "$count $status" }
    }
    Write-Host "${title}: $(if ($summary) { $summary -join ', ' } else { 'no slides' })"
    Write-Host "Report: $report"
    Start-Process -FilePath $report
}

function Find-SpreadsheetCompare {
    $exe = 'Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE'
    foreach ($programFiles in @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ }) {
        $appVLauncher = Join-Path $programFiles 'Microsoft Office\root\Client\AppVLP.exe'
        $clickToRunExe = Join-Path $programFiles "Microsoft Office\root\vfs\ProgramFilesX86\$exe"
        if ((Test-Path -LiteralPath $appVLauncher) -and (Test-Path -LiteralPath $clickToRunExe)) {
            return @{ FilePath = $appVLauncher; Arguments = @("`"$(Join-Path ${env:ProgramFiles(x86)} $exe)`"") }
        }
        $msiExe = Join-Path $programFiles $exe
        if (Test-Path -LiteralPath $msiExe) {
            return @{ FilePath = $msiExe; Arguments = @() }
        }
    }
}

function Compare-Workbooks([string]$oldCopy, [string]$newCopy, [string]$work) {
    $tool = Find-SpreadsheetCompare
    if (-not $tool) {
        throw 'Spreadsheet Compare is not installed. It comes with Office Professional Plus and Microsoft 365 Apps for enterprise.'
    }
    $fileList = Join-Path $work 'files.txt'
    Set-Content -LiteralPath $fileList -Value $oldCopy, $newCopy -Encoding UTF8
    Start-Process -FilePath $tool.FilePath -ArgumentList ($tool.Arguments + "`"$fileList`"")
}

if ($ByDate -and -not (Test-Missing $Old) -and -not (Test-Missing $New) -and
    (Get-Item -LiteralPath $Old).LastWriteTime -gt (Get-Item -LiteralPath $New).LastWriteTime) {
    $Old, $New = $New, $Old
}

$separateFiles = (-not $Name) -or ($Name -eq $Old)
if ($separateFiles) { $Name = $New }
$oldExtension = [IO.Path]::GetExtension($(if ($separateFiles) { $Old } else { $Name }))
$newExtension = [IO.Path]::GetExtension($Name)
$kind = Get-DocumentKind $newExtension

if (-not $kind) {
    Show-TextDiff $Old $New
    exit 0
}
if ($separateFiles -and (Get-DocumentKind $oldExtension) -ne $kind) {
    throw "Cannot compare '$Old' with '$New': they are different kinds of documents."
}
$oldMissing = Test-Missing $Old
if ($oldMissing -or (Test-Missing $New)) {
    Write-Host "$Name was $(if ($oldMissing) { 'added' } else { 'deleted' }) - nothing to compare."
    exit 0
}

$work = New-WorkFolder
$oldBase = [IO.Path]::GetFileNameWithoutExtension($(if ($separateFiles) { $Old } else { $Name }))
$newBase = [IO.Path]::GetFileNameWithoutExtension($Name)
$oldCopy = Copy-Version $Old (Join-Path $work "$oldBase (OLD)$oldExtension")
$newCopy = Copy-Version $New (Join-Path $work "$newBase (NEW)$newExtension")

switch ($kind) {
    'word' {
        Compare-WordDocuments $oldCopy $newCopy
    }
    'powerpoint' {
        if ($separateFiles) {
            $oldLabel = Split-Path -Leaf $Old
            $newLabel = Split-Path -Leaf $New
            $title = "$oldLabel vs $newLabel"
        } else {
            $oldLabel = 'Old version'
            $newLabel = 'New version'
            $title = $Name
        }
        Compare-Presentations $oldCopy $newCopy $work $title $oldLabel $newLabel
    }
    'excel' {
        Compare-Workbooks $oldCopy $newCopy $work
    }
}
