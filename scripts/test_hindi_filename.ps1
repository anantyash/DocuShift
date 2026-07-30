$rootDir = "e:\Word to PDF converter"
$hindiDocPath = Join-Path $rootDir "temp\uploads\हिंदी_परीक्षण_दस्तावेज़.docx"

# Create a test Word doc with Hindi filename
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()

$p = $doc.Paragraphs.Add()
$p.Range.Text = "यह एक हिंदी भाषा परीक्षण दस्तावेज़ है। (Hindi Unicode Test Document)"
$p.Range.Font.Name = "Mangal"
$p.Range.Font.Size = 16
$p.Range.InsertParagraphAfter()

$doc.SaveAs($hindiDocPath)
$doc.Close()
$word.Quit()

Write-Host "Created Hindi test file at: $hindiDocPath"

# Test API conversion
$curlCmd = "curl.exe -X POST -F 'files=@temp/uploads/हिंदी_परीक्षण_दस्तावेज़.docx' http://localhost:3000/api/convert"
$res = Invoke-Expression $curlCmd
Write-Host "API Response:"
Write-Host $res
