$fileId = "f44de415-2025-44d1-be51-3e09b5d83f48"

Write-Host "--- Testing ZIP Endpoint ---"
$zipPayload = @{
    files = @(
        @{ fileId = $fileId; displayName = "Sample_Test_Doc.pdf" }
    )
} | ConvertTo-Json -Depth 5

$zipRes = Invoke-RestMethod -Uri "http://localhost:3000/api/zip" -Method POST -Body $zipPayload -ContentType "application/json"
$zipRes | ConvertTo-Json

Write-Host "--- Testing Merge Endpoint ---"
$mergePayload = @{
    fileIds = @($fileId)
    customTitle = "Merged_Sample_Test"
} | ConvertTo-Json

$mergeRes = Invoke-RestMethod -Uri "http://localhost:3000/api/merge" -Method POST -Body $mergePayload -ContentType "application/json"
$mergeRes | ConvertTo-Json
