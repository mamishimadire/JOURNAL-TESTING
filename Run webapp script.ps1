param(
    [string]$ProjectPath = "c:\Users\Mamishi.Madire\Desktop\Webapp\CaatsWebApp\CaatsWebApp.csproj",
    [string]$Url = "http://127.0.0.1:5099",
    [switch]$NoBuild
)

$ErrorActionPreference = "Stop"

function Test-AppReady {
    param(
        [string]$HealthUrl,
        [int]$TimeoutSeconds = 45
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            $status = (Invoke-WebRequest -Uri $HealthUrl -UseBasicParsing -TimeoutSec 3).StatusCode
            if ($status -ge 200 -and $status -lt 500) {
                return $true
            }
        }
        catch {
            Start-Sleep -Milliseconds 800
        }
    }

    return $false
}

if (-not (Test-Path $ProjectPath)) {
    throw "Project not found at: $ProjectPath"
}

$projectDir = Split-Path -Parent $ProjectPath
Write-Host "Project: $ProjectPath" -ForegroundColor Cyan
Write-Host "URL: $Url" -ForegroundColor Cyan

# Clear old dotnet processes to avoid locked DLL issues.
Get-Process dotnet -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Push-Location $projectDir
try {
    if (-not $NoBuild) {
        Write-Host "Building project..." -ForegroundColor Yellow
        dotnet build
        if ($LASTEXITCODE -ne 0) {
            throw "Build failed."
        }
    }

    Write-Host "Starting app in background..." -ForegroundColor Yellow
    $escapedUrl = $Url.Replace('"', '""')
    $cmd = "dotnet run --project `"$ProjectPath`" --urls `"$escapedUrl`""
    Start-Process powershell -ArgumentList "-NoExit", "-Command", $cmd -WindowStyle Normal | Out-Null

    Write-Host "Waiting for app to become available..." -ForegroundColor Yellow
    if (Test-AppReady -HealthUrl ($Url.TrimEnd('/') + '/')) {
        Start-Process $Url
        Write-Host "Web app is running and opened in browser." -ForegroundColor Green
    }
    else {
        Write-Host "App process started, but health check timed out. Check the launched PowerShell window for errors." -ForegroundColor Red
    }
}
finally {
    Pop-Location
}
