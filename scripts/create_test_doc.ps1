$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()

# Header
$p1 = $doc.Paragraphs.Add()
$p1.Range.Text = "DocuShift Pro - Style & Font Preservation Test"
$p1.Range.Font.Name = "Georgia"
$p1.Range.Font.Size = 22
$p1.Range.Font.Bold = $true
$p1.Range.Font.Color = 12800000 # Dark red/purple
$p1.Range.InsertParagraphAfter()

# Subtitle
$p2 = $doc.Paragraphs.Add()
$p2.Range.Text = "This document tests exact font preservation, line spacing, margins, and custom table formatting when converting from Word to PDF."
$p2.Range.Font.Name = "Calibri"
$p2.Range.Font.Size = 13
$p2.Range.Font.Italic = $true
$p2.Range.InsertParagraphAfter()

# Styled Table
$tableRange = $doc.Paragraphs.Add().Range
$table = $doc.Tables.Add($tableRange, 4, 3)
$table.Cell(1,1).Range.Text = "Feature"
$table.Cell(1,2).Range.Text = "Word Source"
$table.Cell(1,3).Range.Text = "PDF Output"

$table.Cell(2,1).Range.Text = "Typography"
$table.Cell(2,2).Range.Text = "Custom Fonts & Sizes"
$table.Cell(2,3).Range.Text = "100% Embedded"

$table.Cell(3,1).Range.Text = "Layout"
$table.Cell(3,2).Range.Text = "Margins & Indents"
$table.Cell(3,3).Range.Text = "Pixel Perfect"

$table.Cell(4,1).Range.Text = "Batch Mode"
$table.Cell(4,2).Range.Text = "Multi-file Upload"
$table.Cell(4,3).Range.Text = "ZIP / Merged PDF"

$outputPath = [System.IO.Path]::GetFullPath("e:\Word to PDF converter\temp\uploads\Sample_Test_Doc.docx")
$doc.SaveAs([ref]$outputPath)
$doc.Close()
$word.Quit()

Write-Host "Created sample doc at: $outputPath"
