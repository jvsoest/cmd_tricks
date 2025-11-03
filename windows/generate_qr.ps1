# QR Code Generator PowerShell Wrapper
# Calls the Python QR code generator script

param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="The URL or text to encode in the QR code")]
    [string]$Url,
    
    [Parameter(Position=1)]
    [string]$Output = "qr_code.png",
    
    [int]$Size = 10,
    
    [int]$Border = 4
)

# Get the script directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptPath
$pythonScript = Join-Path $repoRoot "python\qr_generator.py"

# Check if Python script exists
if (-not (Test-Path $pythonScript)) {
    Write-Error "Python script not found at: $pythonScript"
    exit 1
}

# Build the Python command
$pythonArgs = @(
    $pythonScript,
    $Url,
    "-o", $Output,
    "-s", $Size,
    "-b", $Border
)

# Execute the Python script
Write-Host "Generating QR code for: $Url"
try {
    & python $pythonArgs
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "QR code generated successfully!" -ForegroundColor Green
        
        # Open the generated QR code image
        if (Test-Path $Output) {
            Write-Host "Opening QR code image..."
            Start-Process $Output
        }
    } else {
        Write-Error "Failed to generate QR code"
        exit $LASTEXITCODE
    }
} catch {
    Write-Error "Error executing Python script: $_"
    exit 1
}
