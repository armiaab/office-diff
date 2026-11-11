param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Path,
    [string]$OutFile,
    [string]$CacheKey
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$xmlNames = [Xml.NameTable]::new()
$ns = [Xml.XmlNamespaceManager]::new($xmlNames)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$ns.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
$ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
$ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
$ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
$ns.AddNamespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
$ns.AddNamespace('v', 'urn:schemas-microsoft-com:vml')
$ns.AddNamespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
$ns.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')
$W = $ns.LookupNamespace('w')
$R = $ns.LookupNamespace('r')
$wordTextPieces = './/w:t | .//m:t | .//w:r/w:tab | .//w:r/w:br | .//w:r/w:cr | .//w:r/w:noBreakHyphen | .//w:drawing | .//w:pict'
$wordPiecesQueries = @{}
$headingParagraphs = @{}

function Read-Part([string]$name) {
    $entry = if ($name) { $zip.GetEntry($name) }
    if (-not $entry) { return }
    $doc = [Xml.XmlDocument]::new($xmlNames)
    $stream = $entry.Open()
    try { $doc.Load($stream) } finally { $stream.Dispose() }
    $doc
}

function Resolve-PartName([string]$fromDir, [string]$target) {
    $target = [Uri]::UnescapeDataString($target)
    if ($target.StartsWith('/')) { $fromDir = '' }
    $parts = [Collections.Generic.List[string]]::new()
    foreach ($segment in "$fromDir/$target".Split('/')) {
        if ($segment -eq '..') {
            if ($parts.Count) { $parts.RemoveAt($parts.Count - 1) }
        } elseif ($segment -and $segment -ne '.') {
            $parts.Add($segment)
        }
    }
    $parts -join '/'
}

function Get-Relationships([string]$partName) {
    $slash = $partName.LastIndexOf('/')
    $dir = if ($slash -ge 0) { $partName.Substring(0, $slash) } else { '' }
    $relsPart = (@($dir, '_rels', ($partName.Substring($slash + 1) + '.rels')) | Where-Object { $_ }) -join '/'
    $rels = Read-Part $relsPart
    if (-not $rels) { return }
    foreach ($rel in $rels.SelectNodes('/rel:Relationships/rel:Relationship[not(@TargetMode="External")]', $ns)) {
        [pscustomobject]@{
            Id     = $rel.GetAttribute('Id')
            Type   = $rel.GetAttribute('Type') -replace '^.*/'
            Target = Resolve-PartName $dir $rel.GetAttribute('Target')
        }
    }
}

function Get-TargetsById([string]$partName) {
    $targets = @{}
    foreach ($rel in Get-Relationships $partName) { $targets[$rel.Id] = $rel.Target }
    $targets
}

function Format-Section([string]$title, [object[]]$lines) {
    if ($lines) { ''; "=== $title ==="; $lines }
}

function Format-TableRow([string[]]$cells) {
    if (($cells -join '').Trim()) { '| ' + ($cells -join ' | ') + ' |' }
}

function Get-GraphicMarker($element) {
    $kind = [string]$element.SelectSingleNode('.//a:graphicData/@uri', $ns).Value
    if ($kind -match '/chart(ex)?$') { '[chart]' }
    elseif ($kind -match '/diagram$') { '[diagram]' }
    elseif ($kind -match '/ole$') { '[object]' }
    elseif ($element.SelectSingleNode('.//pic:pic[not(ancestor::w:txbxContent)]', $ns)) { '[image]' }
}

function Get-HeadingLevels($styles) {
    $levels = @{}
    if (-not $styles) { return $levels }
    $ownLevel = @{}
    $basedOn = @{}
    foreach ($style in $styles.SelectNodes('/w:styles/w:style[@w:type="paragraph"]', $ns)) {
        $id = $style.GetAttribute('styleId', $W)
        $level = $style.SelectSingleNode('w:pPr/w:outlineLvl/@w:val', $ns)
        if ($level) { $ownLevel[$id] = [int]$level.Value }
        $parent = $style.SelectSingleNode('w:basedOn/@w:val', $ns)
        if ($parent) { $basedOn[$id] = $parent.Value }
    }
    foreach ($id in @($ownLevel.Keys) + @($basedOn.Keys)) {
        $current = $id
        for ($hops = 0; $current -and -not $ownLevel.ContainsKey($current) -and $hops -lt 32; $hops++) {
            $current = $basedOn[$current]
        }
        if ($current -and $ownLevel.ContainsKey($current) -and $ownLevel[$current] -lt 9) {
            $levels[$id] = $ownLevel[$current]
        }
    }
    $levels
}

function Get-HeadingParagraphs($document, $styleLevels) {
    $levels = @{}
    foreach ($style in $document.SelectNodes('//w:p/w:pPr/w:pStyle', $ns)) {
        $level = $styleLevels[$style.GetAttribute('val', $W)]
        if ($null -ne $level) { $levels[$style.ParentNode.ParentNode] = $level }
    }
    foreach ($outline in $document.SelectNodes('//w:p/w:pPr/w:outlineLvl', $ns)) {
        $levels[$outline.ParentNode.ParentNode] = [int]$outline.GetAttribute('val', $W)
    }
    $levels
}

function Get-WordBlockLines($container, [int]$depth = 1) {
    if (-not $container) { return }
    foreach ($block in $container.SelectNodes('w:p | w:tbl | w:sdt/w:sdtContent | w:customXml', $ns)) {
        switch ($block.LocalName) {
            'p'     { Get-WordParagraphLines $block $depth }
            'tbl'   { Get-WordTableLines $block $depth }
            default { Get-WordBlockLines $block $depth }
        }
    }
}

function Get-WordTableLines($table, [int]$depth) {
    foreach ($row in $table.SelectNodes('w:tr | w:sdt/w:sdtContent/w:tr | w:customXml/w:tr', $ns)) {
        $cells = foreach ($cell in $row.SelectNodes('w:tc | w:sdt/w:sdtContent/w:tc | w:customXml/w:tc', $ns)) {
            @(Get-WordBlockLines $cell $depth) -join ' / '
        }
        Format-TableRow @($cells)
    }
}

function Get-WordParagraphLines($paragraph, [int]$depth) {
    $ownContent = "[count(ancestor::w:p) = $depth][not(ancestor::mc:Fallback or ancestor::w:del or ancestor::w:moveFrom)]"
    $piecesQuery = $wordPiecesQueries[$depth]
    if (-not $piecesQuery) {
        $piecesQuery = $wordPiecesQueries[$depth] = [Xml.XPath.XPathExpression]::Compile("($wordTextPieces)$ownContent", $ns)
    }
    $text = [Text.StringBuilder]::new()
    $hasDrawing = $false
    foreach ($match in $paragraph.CreateNavigator().Select($piecesQuery)) {
        $node = $match.UnderlyingObject
        $piece = switch ($node.LocalName) {
            't'             { $node.InnerText }
            'tab'           { "`t" }
            'noBreakHyphen' { '-' }
            'drawing'       { $hasDrawing = $true; Get-GraphicMarker $node }
            'pict'          { $hasDrawing = $true; if ($node.SelectSingleNode('.//v:imagedata', $ns)) { '[image]' } }
            default         { ' ' }
        }
        if ($piece) { [void]$text.Append($piece) }
    }
    $line = $text.ToString().Trim()
    if ($line) {
        $level = $headingParagraphs[$paragraph]
        if ($null -ne $level -and $level -lt 9) { $line = ('#' * ($level + 1)) + ' ' + $line }
        $line
    }
    if ($hasDrawing) {
        foreach ($textBox in $paragraph.SelectNodes(".//w:txbxContent$ownContent", $ns)) {
            Get-WordBlockLines $textBox ($depth + 1)
        }
    }
}

function Get-WordNoteLines($notesPart, [string]$noteName) {
    $notes = Read-Part $notesPart
    if (-not $notes) { return }
    foreach ($note in $notes.SelectNodes("/*/w:$noteName[not(@w:type) or @w:type='normal']", $ns)) {
        Get-WordBlockLines $note
    }
}

function Get-WordCommentLines($commentsPart) {
    $comments = Read-Part $commentsPart
    if (-not $comments) { return }
    foreach ($comment in $comments.SelectNodes('/w:comments/w:comment', $ns)) {
        $lines = @(Get-WordBlockLines $comment)
        if (-not $lines) { continue }
        $author = $comment.GetAttribute('author', $W)
        if ($author) { $lines[0] = "${author}: $($lines[0])" }
        $lines
    }
}

function Get-WordText($document, [string]$documentPart) {
    $rels = @(Get-Relationships $documentPart)
    $partOfType = { param($type) $rels | Where-Object Type -eq $type | Select-Object -First 1 -ExpandProperty Target }
    $targets = @{}
    foreach ($rel in $rels) { $targets[$rel.Id] = $rel.Target }
    $script:headingParagraphs = Get-HeadingParagraphs $document (Get-HeadingLevels (Read-Part (& $partOfType 'styles')))

    Get-WordBlockLines $document.SelectSingleNode('/w:document/w:body', $ns)

    $headerParts = foreach ($ref in $document.SelectNodes('//w:sectPr/w:headerReference | //w:sectPr/w:footerReference', $ns)) {
        $targets[$ref.GetAttribute('id', $R)]
    }
    Format-Section 'Headers and footers' @(
        foreach ($part in $headerParts | Where-Object { $_ } | Select-Object -Unique) {
            $header = Read-Part $part
            if ($header) { Get-WordBlockLines $header.DocumentElement }
        }
    )
    Format-Section 'Footnotes' @(Get-WordNoteLines (& $partOfType 'footnotes') 'footnote')
    Format-Section 'Endnotes' @(Get-WordNoteLines (& $partOfType 'endnotes') 'endnote')
    Format-Section 'Comments' @(Get-WordCommentLines (& $partOfType 'comments'))
}

function Get-DrawingParagraphText($paragraph) {
    $text = [Text.StringBuilder]::new()
    foreach ($node in $paragraph.SelectNodes('a:r/a:t | a:fld/a:t | a:br | mc:AlternateContent/mc:Choice//m:t', $ns)) {
        if ($node.LocalName -eq 'br') { [void]$text.Append(' ') } else { [void]$text.Append($node.InnerText) }
    }
    $text.ToString().Trim()
}

function Get-TextBodyLines($shape) {
    foreach ($paragraph in $shape.SelectNodes('p:txBody/a:p', $ns)) {
        $line = Get-DrawingParagraphText $paragraph
        if ($line) { $line }
    }
}

function Get-GraphicFrameLines($frame) {
    $rows = $frame.SelectNodes('a:graphic/a:graphicData/a:tbl/a:tr', $ns)
    if (-not $rows.Count) { return Get-GraphicMarker $frame }
    foreach ($row in $rows) {
        $cells = foreach ($cell in $row.SelectNodes('a:tc', $ns)) {
            @(foreach ($paragraph in $cell.SelectNodes('a:txBody/a:p', $ns)) {
                $line = Get-DrawingParagraphText $paragraph
                if ($line) { $line }
            }) -join ' / '
        }
        Format-TableRow @($cells)
    }
}

function Get-PictureMarker($picture) {
    $media = $picture.SelectSingleNode('p:nvPicPr/p:nvPr/a:videoFile | p:nvPicPr/p:nvPr/a:audioFile', $ns)
    if (-not $media) { '[image]' }
    elseif ($media.LocalName -eq 'videoFile') { '[video]' }
    else { '[audio]' }
}

function Get-ShapeLines($shapeTree) {
    if (-not $shapeTree) { return }
    foreach ($shape in $shapeTree.SelectNodes('p:sp | p:grpSp | p:graphicFrame | p:pic | mc:AlternateContent/mc:Choice[1]', $ns)) {
        switch ($shape.LocalName) {
            'sp'           { Get-TextBodyLines $shape }
            'pic'          { Get-PictureMarker $shape }
            'graphicFrame' { Get-GraphicFrameLines $shape }
            default        { Get-ShapeLines $shape }
        }
    }
}

function Get-NotesLines([string]$slidePart) {
    $notesPart = Get-Relationships $slidePart | Where-Object Type -eq 'notesSlide' | Select-Object -First 1
    $notes = if ($notesPart) { Read-Part $notesPart.Target }
    if (-not $notes) { return }
    foreach ($body in $notes.SelectNodes('/p:notes/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type="body"]', $ns)) {
        Get-TextBodyLines $body
    }
}

function Get-PresentationText($presentation, [string]$presentationPart) {
    $targets = Get-TargetsById $presentationPart
    $number = 0
    foreach ($slideId in $presentation.SelectNodes('/p:presentation/p:sldIdLst/p:sldId', $ns)) {
        $number++
        $slidePart = $targets[$slideId.GetAttribute('id', $R)]
        $slide = Read-Part $slidePart
        $hidden = $slide -and ($slide.DocumentElement.GetAttribute('show') -in '0', 'false')
        if ($hidden) { "=== Slide $number (hidden) ===" } else { "=== Slide $number ===" }
        if ($slide) {
            Get-ShapeLines $slide.SelectSingleNode('/p:sld/p:cSld/p:spTree', $ns)
            $notes = @(Get-NotesLines $slidePart)
            if ($notes) { '--- notes ---'; $notes }
        }
        ''
    }
}

function Get-PackageText([string]$file) {
    try {
        $script:zip = [IO.Compression.ZipFile]::OpenRead($file)
    } catch [IO.InvalidDataException] {
        return '(not a .docx/.pptx file, or it is password-protected - no text extracted)'
    }
    try {
        $mainPart = Get-Relationships '' | Where-Object Type -eq 'officeDocument' | Select-Object -First 1 -ExpandProperty Target
        $main = Read-Part $mainPart
        $root = if ($main) { $main.DocumentElement }
        if ($root -and $root.NamespaceURI -eq $W -and $root.LocalName -eq 'document') {
            Get-WordText $main $mainPart
        } elseif ($root -and $root.NamespaceURI -eq $ns.LookupNamespace('p') -and $root.LocalName -eq 'presentation') {
            Get-PresentationText $main $mainPart
        } else {
            '(no Word or PowerPoint content found)'
        }
    } finally {
        $zip.Dispose()
    }
}

$lines = try {
    Get-PackageText (Resolve-Path -LiteralPath $Path).ProviderPath
} catch {
    "(could not extract text: $($_.Exception.Message))"
}

$bytes = [Text.UTF8Encoding]::new($false).GetBytes((@($lines) -join "`n") + "`n")
if ($OutFile) {
    [IO.File]::WriteAllBytes($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile), $bytes)
} else {
    $stdout = [Console]::OpenStandardOutput()
    $stdout.Write($bytes, 0, $bytes.Length)
    $stdout.Flush()
}
