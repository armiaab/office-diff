<#
.SYNOPSIS
    Print the readable text of a .docx or .pptx, so `git diff` can show what changed inside it.

.DESCRIPTION
    Word:       one line per paragraph - body, tables, text boxes, equations, footnotes, endnotes.
    PowerPoint: slides in presentation order, one line per paragraph, plus speaker notes.
    Pictures appear as [image]. Used as a git textconv driver (see install.ps1), or on its own:
        office-textconv.ps1 thesis.docx
        office-textconv.ps1 slides.pptx -OutFile slides.txt
#>
param([Parameter(Mandatory = $true, Position = 0)][string]$Path, [string]$OutFile)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$ns = @{
    w   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    m   = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    a   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    p   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    r   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
    mc  = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
}
$out = New-Object Text.StringBuilder

function Read-Xml([string]$entryName) {
    $entry = $zip.GetEntry($entryName)
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    try {
        $doc = New-Object Xml.XmlDocument
        $doc.Load($stream)
        $mgr = New-Object Xml.XmlNamespaceManager($doc.NameTable)
        foreach ($k in $ns.Keys) { $mgr.AddNamespace($k, $ns[$k]) }
        return @{ Doc = $doc; Ns = $mgr }
    } finally { $stream.Dispose() }
}

function Resolve-Target([string]$fromDir, [string]$target) {
    $parts = New-Object Collections.Generic.List[string]
    foreach ($seg in ("$fromDir/$target" -split '/')) {
        if ($seg -eq '..') { $parts.RemoveAt($parts.Count - 1) }
        elseif ($seg -and $seg -ne '.') { $parts.Add($seg) }
    }
    return ($parts -join '/')
}

function Get-Rels([string]$relsEntry, [string]$fromDir) {
    $map = @{}
    $x = Read-Xml $relsEntry
    if ($x) {
        foreach ($r in $x.Doc.SelectNodes('//rel:Relationship', $x.Ns)) {
            $map[$r.GetAttribute('Id')] = @{ Type = $r.GetAttribute('Type'); Target = (Resolve-Target $fromDir $r.GetAttribute('Target')) }
        }
    }
    return $map
}

function Add-WordParagraphs($x, [string]$heading) {
    $lines = foreach ($para in $x.Doc.SelectNodes('//w:body//w:p | //w:footnote//w:p | //w:endnote//w:p', $x.Ns)) {
        if ($para.SelectSingleNode('ancestor::w:p', $x.Ns)) { continue }   # text boxes: counted by outer paragraph
        $sb = New-Object Text.StringBuilder
        # mc:Fallback holds a legacy duplicate of shapes/text boxes; skip it.
        foreach ($n in $para.SelectNodes('(.//w:t | .//m:t | .//w:tab | .//w:br | .//w:drawing | .//w:pict)[not(ancestor::mc:Fallback)]', $x.Ns)) {
            switch ($n.LocalName) {
                't'     { [void]$sb.Append($n.InnerText) }
                'tab'   { [void]$sb.Append("`t") }
                'br'    { [void]$sb.Append(' ') }
                default { [void]$sb.Append('[image]') }
            }
        }
        $line = $sb.ToString().Trim()
        if ($line) { $line }
    }
    if ($lines) {
        if ($heading) { [void]$out.Append("`n=== $heading ===`n") }
        foreach ($l in $lines) { [void]$out.Append($l).Append("`n") }
    }
}

function Get-DrawingParagraphs($x, [string]$shapeFilter) {
    foreach ($sp in $x.Doc.SelectNodes("//p:cSld//*[self::p:sp or self::p:graphicFrame or self::p:pic]$shapeFilter", $x.Ns)) {
        if ($sp.LocalName -eq 'pic') { '[image]'; continue }
        foreach ($para in $sp.SelectNodes('.//a:p', $x.Ns)) {
            $sb = New-Object Text.StringBuilder
            foreach ($n in $para.SelectNodes('.//a:t | .//a:br', $x.Ns)) {
                if ($n.LocalName -eq 't') { [void]$sb.Append($n.InnerText) } else { [void]$sb.Append(' ') }
            }
            $line = $sb.ToString().Trim()
            if ($line) { $line }
        }
    }
}

try {
    $zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Path).ProviderPath)
} catch [IO.InvalidDataException] {
    $zip = $null
    [void]$out.Append("(not a .docx/.pptx file - no text extracted)`n")
}

if ($zip) {
    try {
        if ($zip.GetEntry('word/document.xml')) {
            Add-WordParagraphs (Read-Xml 'word/document.xml') $null
            foreach ($part in @(@('word/footnotes.xml', 'Footnotes'), @('word/endnotes.xml', 'Endnotes'))) {
                $x = Read-Xml $part[0]
                if ($x) { Add-WordParagraphs $x $part[1] }
            }
        }
        elseif ($zip.GetEntry('ppt/presentation.xml')) {
            $pres = Read-Xml 'ppt/presentation.xml'
            $rels = Get-Rels 'ppt/_rels/presentation.xml.rels' 'ppt'
            $i = 0
            foreach ($sld in $pres.Doc.SelectNodes('//p:sldIdLst/p:sldId', $pres.Ns)) {
                $i++
                $slidePath = $rels[$sld.GetAttribute('id', $ns.r)].Target
                [void]$out.Append("=== Slide $i ===`n")
                $slide = Read-Xml $slidePath
                foreach ($l in (Get-DrawingParagraphs $slide '')) { [void]$out.Append($l).Append("`n") }

                $slideDir  = $slidePath.Substring(0, $slidePath.LastIndexOf('/'))
                $slideFile = $slidePath.Substring($slidePath.LastIndexOf('/') + 1)
                $sRels = Get-Rels "$slideDir/_rels/$slideFile.rels" $slideDir
                $notesRel = $sRels.Values | Where-Object { $_.Type -like '*/notesSlide' } | Select-Object -First 1
                if ($notesRel) {
                    $notes = Read-Xml $notesRel.Target
                    $noteLines = @(Get-DrawingParagraphs $notes '[.//p:ph[@type="body"]]')
                    if ($noteLines.Count) {
                        [void]$out.Append("--- notes ---`n")
                        foreach ($l in $noteLines) { [void]$out.Append($l).Append("`n") }
                    }
                }
                [void]$out.Append("`n")
            }
        }
        else {
            [void]$out.Append("(no Word or PowerPoint content found)`n")
        }
    }
    finally { $zip.Dispose() }
}

# Raw UTF-8 bytes: avoids console code-page mangling of non-ASCII text.
$bytes = (New-Object Text.UTF8Encoding($false)).GetBytes($out.ToString())
if ($OutFile) { [IO.File]::WriteAllBytes($OutFile, $bytes); return }
$stdout = [Console]::OpenStandardOutput()
$stdout.Write($bytes, 0, $bytes.Length)
$stdout.Flush()
