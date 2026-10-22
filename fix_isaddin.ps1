# fix_isaddin.ps1 - Ensure isAddin="true" is present in xl/workbook.xml inside an .xlam
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File fix_isaddin.ps1 <xlamPath>
#
# WPS/Excel SaveAs(format 55) may omit the isAddin attribute, which makes WPS
# report "invalid add-in". This patches the ZIP. If <workbookPr> is missing it is
# CREATED inside <workbook> rather than erroring out (so build can rely on exit code).

param(
    [Parameter(Mandatory=$true)]
    [string]$XlamPath
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not (Test-Path -LiteralPath $XlamPath)) {
    Write-Error "File not found: $XlamPath"
    exit 1
}

try {
    $zip = [System.IO.Compression.ZipFile]::Open($XlamPath, [System.IO.Compression.ZipArchiveMode]::Update)
} catch {
    Write-Error "Cannot open ZIP: $_"
    exit 1
}

try {
    $entry = $zip.GetEntry("xl/workbook.xml")
    if ($null -eq $entry) {
        Write-Error "xl/workbook.xml not found in ZIP"
        exit 1
    }

    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $content = $reader.ReadToEnd()
    $reader.Close()

    if ($content -match 'isAddin="true"') {
        Write-Output "OK: isAddin already present"
        exit 0
    }

    if ($content -match '<workbookPr') {
        # add attribute to existing element (handle both <workbookPr ...> and <workbookPr/>)
        $content = $content -replace '<workbookPr(\s)', '<workbookPr isAddin="true"$1'
        $content = $content -replace '<workbookPr>', '<workbookPr isAddin="true">'
        $content = $content -replace '<workbookPr/>', '<workbookPr isAddin="true"/>'
    } else {
        # create <workbookPr isAddin="true"/> right after the opening <workbook ...> tag
        $newXml = '<workbookPr isAddin="true"/>'
        $content = [regex]::Replace($content, '(<workbook\b[^>]*>)', ('$1' + $newXml), 1)
        if ($content -notmatch 'isAddin="true"') {
            Write-Error "Could not locate <workbook> root to insert workbookPr"
            exit 1
        }
    }

    if ($content -notmatch 'isAddin="true"') {
        Write-Error "Patch produced no isAddin attribute"
        exit 1
    }

    $entry.Delete()
    $newEntry = $zip.CreateEntry("xl/workbook.xml", [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $newEntry.Open()
    $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
    $writer.Write($content)
    $writer.Close()
    $stream.Close()

    Write-Output "OK: isAddin=true ensured in workbook.xml"
    exit 0
} catch {
    Write-Error "Failed: $_"
    exit 1
} finally {
    $zip.Dispose()
}
