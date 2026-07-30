$InputPath = "e:\Word to PDF converter\temp\uploads\Sample_Test_Doc.docx"
$OutputPath = "e:\Word to PDF converter\temp\outputs\Sample_Test_Doc.pdf"

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

Write-Host "Testing PDF Conversion Methods..."

# Method 1: SaveAs format 17
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($InputPath)
    
    # 17 = wdFormatPDF
    $doc.SaveAs([ref]$OutputPath, [ref]17)
    $doc.Close()
    $word.Quit()
    Write-Host "Method 1 (SaveAs 17): SUCCESS!"
    exit 0
} catch {
    Write-Host "Method 1 Error: "$_.Exception.Message
} finally {
    if ($word) { try { $word.Quit() } catch {} }
}

# Method 2: Print to PDF
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.ActivePrinter = "Microsoft Print to PDF"
    $doc = $word.Documents.Open($InputPath)
    $doc.PrintOut($false, $false, [ref]0, [ref]$OutputPath)
    $doc.Close()
    $word.Quit()
    Write-Host "Method 2 (PrintOut): SUCCESS!"
    exit 0
} catch {
    Write-Host "Method 2 Error: "$_.Exception.Message
} finally {
    if ($word) { try { $word.Quit() } catch {} }
}
