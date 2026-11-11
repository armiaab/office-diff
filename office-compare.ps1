<#
.SYNOPSIS
    Compare two versions of an Office file with Microsoft Office's own tools.

.DESCRIPTION
    .doc / .docx / .docm   Word's built-in Compare (Review > Compare), shown as tracked changes.
    .ppt / .pptx / .pptm   Both versions side by side in PowerPoint (old left, new right) plus a
                           slide-by-slide text diff in the console. PowerPoint's own Compare was
                           retired in Microsoft 365 version 2502, so it cannot be used anymore.
    .xls / .xlsx / .xlsm   Microsoft Spreadsheet Compare (Office Professional Plus / Microsoft 365 Apps).
    anything else          Plain `git diff --no-index` in the console.

    Works as a git difftool (see install.ps1) or on its own:
        office-compare.ps1 old.docx new.docx
        office-compare.ps1 -ByDate a.pptx b.pptx    # older file (by modified date) = old version

    The files are never opened directly: private copies are compared, so the originals are
    not locked or changed, and git can clean up its temporary files right away.
#>
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Old,
    [Parameter(Mandatory = $true, Position = 1)][string]$New,
    # Path of the file inside the repository (git difftool's $MERGED); used for naming only.
    [Parameter(Position = 2)][string]$Name,
    # Treat the file with the older modification date as the old version.
    [switch]$ByDate
)
$ErrorActionPreference = 'Stop'

# `git difftool` exports GIT_EXTERNAL_DIFF; drop it (and use --no-ext-diff) so the plain
# `git diff` calls below can't re-launch this script in an endless loop.
Remove-Item Env:GIT_EXTERNAL_DIFF, Env:GIT_DIFF_PATH_COUNTER, Env:GIT_DIFF_PATH_TOTAL -ErrorAction SilentlyContinue
function Show-TextDiff($a, $b) {
    git -c core.autocrlf=false -c core.safecrlf=false --no-pager diff --no-index --no-ext-diff --no-textconv --color -- $a $b
}
function Test-Missing($p) { ($p -eq '/dev/null') -or ($p -eq 'nul') -or -not (Test-Path -LiteralPath $p -PathType Leaf) }
function Get-Kind($ext) {
    switch -Regex ($ext) {
        '^\.(doc|docx|docm)$'  { return 'word' }
        '^\.(ppt|pptx|pptm)$'  { return 'powerpoint' }
        '^\.(xls|xlsx|xlsm)$'  { return 'excel' }
    }
    return $null
}
function New-OfficeApp($progId, $label) {
    try { return New-Object -ComObject $progId }
    catch { throw "$label (desktop app) is needed for this comparison but could not be started: $($_.Exception.Message)" }
}

if ($ByDate -and -not (Test-Missing $Old) -and -not (Test-Missing $New) -and
    (Get-Item -LiteralPath $Old).LastWriteTime -gt (Get-Item -LiteralPath $New).LastWriteTime) {
    $Old, $New = $New, $Old
}

# Standalone use and `git difftool --no-index a b` (where $MERGED == $LOCAL) compare two different
# files: keep each one's own name. A normal difftool run compares two versions of $Name.
$ownNames = (-not $Name) -or ($Name -eq $Old)
if ($ownNames) {
    $oldExt = [IO.Path]::GetExtension($Old).ToLowerInvariant()
    $newExt = [IO.Path]::GetExtension($New).ToLowerInvariant()
    $Name = $New
} else {
    $oldExt = $newExt = [IO.Path]::GetExtension($Name).ToLowerInvariant()
}
$kind = Get-Kind $newExt

if (-not $kind) {
    Show-TextDiff $Old $New
    exit 0
}
if ($ownNames -and (Get-Kind $oldExt) -ne $kind) {
    throw "Cannot compare '$Old' with '$New': they are different kinds of documents."
}
if ((Test-Missing $Old) -or (Test-Missing $New)) {
    $what = if (Test-Missing $Old) { 'added' } else { 'deleted' }
    Write-Host "$Name was $what - nothing to compare."
    exit 0
}

# Private copies in %TEMP%\office-diff\<timestamp>; their names become the window titles.
$tempRoot = Join-Path $env:TEMP 'office-diff'
Get-ChildItem -LiteralPath $tempRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue   # leftovers from earlier comparisons
$work = Join-Path $tempRoot (Get-Date -Format 'yyyyMMdd-HHmmss-fff')
New-Item -ItemType Directory -Path $work -Force | Out-Null

if ($ownNames) {
    $oldBase = [IO.Path]::GetFileNameWithoutExtension($Old)
    $newBase = [IO.Path]::GetFileNameWithoutExtension($New)
} else {
    $oldBase = $newBase = [IO.Path]::GetFileNameWithoutExtension($Name)
}
$oldName = "$oldBase (OLD)"
$newName = "$newBase (NEW)"
$oldCopy = Join-Path $work "$oldName$oldExt"
$newCopy = Join-Path $work "$newName$newExt"
Copy-Item -LiteralPath $Old -Destination $oldCopy
Copy-Item -LiteralPath $New -Destination $newCopy
foreach ($f in $oldCopy, $newCopy) {
    Set-ItemProperty -LiteralPath $f -Name IsReadOnly -Value $false
    Unblock-File -LiteralPath $f -ErrorAction SilentlyContinue   # avoid Protected View on downloaded files
}

switch ($kind) {
    'word' {
        $word = New-OfficeApp 'Word.Application' 'Microsoft Word'
        $word.Visible = $true
        # Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        $orig = $word.Documents.Open($oldCopy, $false, $true, $false)
        $rev  = $word.Documents.Open($newCopy, $false, $true, $false)
        # CompareDocuments(Original, Revised, Destination=wdCompareDestinationNew(2), Granularity=wdGranularityWordLevel(1),
        #   Formatting, CaseChanges, Whitespace, Tables, Headers, Footnotes, Textboxes, Fields, Comments, Moves, RevisedAuthor)
        $result = $word.CompareDocuments($orig, $rev, 2, 1, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, 'New version')
        $orig.Close(0)
        $rev.Close(0)
        $result.Saved = $true   # no "save changes?" prompt when closing the result
        $result.Activate()
        $word.Activate()
    }
    'powerpoint' {
        # Microsoft retired PowerPoint's Compare & Merge in Microsoft 365 version 2502: Merge() and
        # MergeWithBaseline() now throw "no longer supported". Next best: text diff + side by side.
        if ($oldExt -ne '.ppt' -and $newExt -ne '.ppt') {
            $conv = Join-Path $PSScriptRoot 'office-textconv.ps1'
            & $conv -Path $oldCopy -OutFile (Join-Path $work "$oldName.txt")
            & $conv -Path $newCopy -OutFile (Join-Path $work "$newName.txt")
            $label = if ($ownNames) { "$oldBase$oldExt -> $newBase$newExt" } else { $Name }
            Write-Host "Text changes in ${label}:"
            Push-Location $work   # short a/ b/ names in the diff header
            try { Show-TextDiff "$oldName.txt" "$newName.txt" } finally { Pop-Location }
        }
        $pp = New-OfficeApp 'PowerPoint.Application' 'Microsoft PowerPoint'
        $pp.Visible = -1   # msoTrue
        # Presentations.Open(FileName, ReadOnly=msoTrue, Untitled=msoFalse, WithWindow=msoTrue)
        $oldPres = $pp.Presentations.Open($oldCopy, -1, 0, -1)
        $newPres = $pp.Presentations.Open($newCopy, -1, 0, -1)
        $oldWin = $oldPres.Windows[1]; $newWin = $newPres.Windows[1]
        foreach ($w in $oldWin, $newWin) { $w.WindowState = 1 }   # ppWindowNormal
        $pp.Windows.Arrange(1)   # ppArrangeTiled
        if ($oldWin.Left -gt $newWin.Left) {   # tiling order varies: force old = left, new = right
            $l = $oldWin.Left; $oldWin.Left = $newWin.Left; $newWin.Left = $l
        }
        $pp.Activate()
    }
    'excel' {
        # Spreadsheet Compare takes a text file listing the two workbooks, one path per line.
        $list = Join-Path $work 'files.txt'
        Set-Content -LiteralPath $list -Value @($oldCopy, $newCopy) -Encoding UTF8
        $appv = "$env:ProgramFiles\Microsoft Office\root\Client\AppVLP.exe"
        $c2r  = "$env:ProgramFiles\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE"
        $msi  = @("$env:ProgramFiles\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE",
                  "${env:ProgramFiles(x86)}\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE") |
                Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if ((Test-Path -LiteralPath $appv) -and (Test-Path -LiteralPath $c2r)) {
            # Microsoft 365 (Click-to-Run): the exe crashes unless started through the App-V
            # launcher, exactly like its Start-menu shortcut does.
            Start-Process -FilePath $appv -ArgumentList "`"C:\Program Files (x86)\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE`" `"$list`""
        }
        elseif ($msi) {
            Start-Process -FilePath $msi -ArgumentList "`"$list`""
        }
        else {
            Write-Host 'Spreadsheet Compare is not installed (it comes with Office Professional Plus / Microsoft 365 Apps).'
        }
    }
}
