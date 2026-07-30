$rootDir = Resolve-Path "e:\Word to PDF converter"
$portableDir = Join-Path $rootDir "DocuShift_Portable"
$binDir = Join-Path $portableDir "bin"

Write-Host "Creating Portable Standalone Bundle..."

# Clean old portable folder
if (Test-Path -Path $portableDir) {
    Remove-Item -Path $portableDir -Recurse -Force
}

# Create directories
New-Item -ItemType Directory -Path $binDir -Force | Out-Null

# Find system node.exe
$systemNode = (Get-Command node -ErrorAction SilentlyContinue).Source
if (-not $systemNode) {
    $systemNode = "C:\Program Files\nodejs\node.exe"
}

if (Test-Path -Path $systemNode) {
    Copy-Item -Path $systemNode -Destination (Join-Path $binDir "node.exe") -Force
    Write-Host "Copied portable node.exe from $systemNode"
} else {
    Write-Warning "Could not find node.exe on host system!"
}

# Copy project files (excluding node_modules, temp, git)
$itemsToCopy = @("package.json", "server.js", "start_app.bat", "public", "src", "scripts")
foreach ($item in $itemsToCopy) {
    $srcPath = Join-Path $rootDir $item
    if (Test-Path -Path $srcPath) {
        Copy-Item -Path $srcPath -Destination $portableDir -Recurse -Force
    }
}

# Copy production node_modules
$nodeModulesSrc = Join-Path $rootDir "node_modules"
if (Test-Path -Path $nodeModulesSrc) {
    Copy-Item -Path $nodeModulesSrc -Destination (Join-Path $portableDir "node_modules") -Recurse -Force
}

Write-Host "========================================================="
Write-Host "SUCCESS! Standalone Portable Package created at:"
Write-Host " -> $portableDir"
Write-Host "You can zip this folder and run it on ANY Windows PC without installing Node.js!"
Write-Host "========================================================="
