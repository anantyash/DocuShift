param (
    [Parameter(Mandatory=$true)]
    [string]$InputPath,

    [Parameter(Mandatory=$true)]
    [string]$OutputPath,

    [string]$Quality = "Standard",

    [string]$EmbedFonts = "1"
)

$ErrorActionPreference = "Stop"

# Ensure absolute paths
$InputPath = [System.IO.Path]::GetFullPath($InputPath)
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

if (-not (Test-Path -Path $InputPath)) {
    $result = @{
        success = $false
        error   = "Input file does not exist: $InputPath"
    }
    Write-Output ($result | ConvertTo-Json -Compress)
    exit 1
}

# Output directory check
$OutputDir = [System.IO.Path]::GetDirectoryName($OutputPath)
if (-not (Test-Path -Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$word = $null
$doc = $null

try {
    # Initialize Word COM Application
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0 # wdAlertsNone
    $word.AutomationSecurity = 3 # msoAutomationSecurityForceDisable

    # Open Word document read-only
    $doc = $word.Documents.Open(
        $InputPath,
        $false, # ConfirmConversions
        $true,  # ReadOnly
        $false, # AddToRecentFiles
        "", "", $true
    )

    $converted = $false

    # Strategy 1: ExportAsFixedFormat (Word 2010+ / 2007 with PDF plugin)
    try {
        $exportQuality = if ($Quality -eq "Minimal") { 1 } else { 0 }
        $doc.ExportAsFixedFormat(
            $OutputPath,
            17, # wdExportFormatPDF
            $false,
            $exportQuality,
            0, 1, 1, 0,
            $true, $true, 1, $true, $true, $false
        )
        $converted = $true
    } catch {
        # Strategy 1 failed, fallback
    }

    # Strategy 2: SaveAs wdFormatPDF (Format Code 17)
    if (-not $converted) {
        try {
            $doc.SaveAs([ref]$OutputPath, [ref]17)
            $converted = $true
        } catch {
            # Strategy 2 failed, fallback
        }
    }

    # Strategy 3: PrintOut to "Microsoft Print to PDF" (Universal Windows Native Print Driver)
    if (-not $converted) {
        try {
            $word.ActivePrinter = "Microsoft Print to PDF"
            $doc.PrintOut($false, $false, [ref]0, [ref]$OutputPath)
            
            # Wait for spooler file lock to clear if needed
            $maxWait = 10
            while ($maxWait -gt 0 -and (-not (Test-Path -Path $OutputPath))) {
                Start-Sleep -Milliseconds 500
                $maxWait--
            }
            if (Test-Path -Path $OutputPath) {
                $converted = $true
            }
        } catch {
            throw $_
        }
    }

    if (-not $converted -or -not (Test-Path -Path $OutputPath)) {
        throw New-Object System.Exception("Word PDF export strategies failed to write output file.")
    }

    # Compute Statistics (Pages)
    $pageCount = 1
    try {
        $pageCount = $doc.ComputeStatistics(2) # 2 = wdStatisticPages
    } catch {
        $pageCount = 1
    }

    $result = @{
        success   = $true
        pdfPath   = $OutputPath
        pageCount = $pageCount
        engine    = "MS Word Automation Engine"
    }
    Write-Output ($result | ConvertTo-Json -Compress)

} catch {
    $result = @{
        success = $false
        error   = $_.Exception.Message
    }
    Write-Output ($result | ConvertTo-Json -Compress)
    exit 1
} finally {
    if ($doc -ne $null) {
        try { $doc.Close($false) } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($word -ne $null) {
        try { $word.Quit() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
