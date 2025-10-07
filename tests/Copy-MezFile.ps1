# Copy-MezFile.ps1
# PowerShell script to copy a MEZ file from the local project's build folder 
# to the Power BI Custom Connectors directory

param(
    [string]$SourceMezPath,
    [switch]$Verbose
)

function Get-ProjectMezFile {
    <#
    .SYNOPSIS
    Automatically finds the MEZ file in the current project's bin folder
    #>
    
    # Look for MEZ files in common build output locations
    $searchPaths = @(
        "bin\AnyCPU\Debug",
        "bin\AnyCPU\Release", 
        "bin\Debug",
        "bin\Release",
        "bin",
        "."
    )
    
    foreach ($path in $searchPaths) {
        $fullPath = Join-Path -Path (Get-Location) -ChildPath $path
        if (Test-Path $fullPath) {
            $mezFiles = Get-ChildItem -Path $fullPath -Filter "*.mez" -ErrorAction SilentlyContinue
            if ($mezFiles) {
                # Return the most recently modified MEZ file
                $latestMez = $mezFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                return $latestMez.FullName
            }
        }
    }
    
    return $null
}

function Get-PowerBIConnectorsPath {
    <#
    .SYNOPSIS
    Gets the Power BI Custom Connectors directory path
    #>
    
    return "$env:USERPROFILE\Documents\Power BI Desktop\Custom Connectors"
}

function Copy-MezToPowerBI {
    <#
    .SYNOPSIS
    Copies the MEZ file to the Power BI Custom Connectors directory
    #>
    param(
        [string]$SourcePath,
        [string]$DestinationFolder
    )
    
    # Validate source file exists
    if (-Not (Test-Path $SourcePath)) {
        Write-Host "ERROR: The source MEZ file was not found at $SourcePath" -ForegroundColor Red
        return $false
    }
    
    # Create destination folder if it doesn't exist
    if (-Not (Test-Path $DestinationFolder)) {
        Write-Host "Destination folder $DestinationFolder not found. Creating..." -ForegroundColor Yellow
        try {
            New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
            Write-Host "Created directory: $DestinationFolder" -ForegroundColor Green
        }
        catch {
            Write-Host "ERROR: Could not create directory $DestinationFolder" -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    
    # Copy the MEZ file
    $fileName = Split-Path $SourcePath -Leaf
    $destination = Join-Path $DestinationFolder $fileName
    
    try {
        Copy-Item -Path $SourcePath -Destination $destination -Force
        Write-Host "Successfully copied MEZ file:" -ForegroundColor Green
        Write-Host "  From: $SourcePath" -ForegroundColor Cyan
        Write-Host "  To:   $destination" -ForegroundColor Cyan
        
        # Show file details
        $fileInfo = Get-Item $destination
        Write-Host "File size: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Gray
        Write-Host "Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
        
        return $true
    }
    catch {
        Write-Host "ERROR: Could not copy MEZ file to $destination" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Main script execution
Write-Host "Power BI Custom Connector MEZ File Copy Utility" -ForegroundColor Blue
Write-Host "=" * 50 -ForegroundColor Blue

# Determine source MEZ file
$sourceMez = if ($SourceMezPath) {
    # Use provided path (resolve relative paths)
    if ([System.IO.Path]::IsPathRooted($SourceMezPath)) {
        $SourceMezPath
    } else {
        Join-Path -Path (Get-Location) -ChildPath $SourceMezPath | Resolve-Path -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path
    }
} else {
    # Auto-discover MEZ file in project
    Get-ProjectMezFile
}

if (-not $sourceMez) {
    Write-Host "ERROR: No MEZ file found!" -ForegroundColor Red
    if ($SourceMezPath) {
        Write-Host "  Specified path: $SourceMezPath" -ForegroundColor Red
    } else {
        Write-Host "  Searched in common build folders (bin\Debug, bin\Release, etc.)" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Solutions:" -ForegroundColor Yellow
    Write-Host "  1. Build your Power BI connector project first" -ForegroundColor White
    Write-Host "  2. Specify explicit MEZ path: .\Copy-MezFile.ps1 -SourceMezPath 'path\to\file.mez'" -ForegroundColor White
    Write-Host "  3. Ensure MEZ file is in bin\Debug, bin\Release, or current directory" -ForegroundColor White
    exit 1
}

if (-not (Test-Path $sourceMez)) {
    Write-Host "ERROR: MEZ file not found at resolved path: $sourceMez" -ForegroundColor Red
    exit 1
}

# Get destination folder
$destinationFolder = Get-PowerBIConnectorsPath

if ($Verbose) {
    Write-Host ""
    Write-Host "Configuration:" -ForegroundColor Yellow
    Write-Host "  Source MEZ: $sourceMez" -ForegroundColor White
    Write-Host "  Target Dir: $destinationFolder" -ForegroundColor White
    Write-Host ""
}

# Perform the copy operation
$success = Copy-MezToPowerBI -SourcePath $sourceMez -DestinationFolder $destinationFolder

if ($success) {
    Write-Host ""
    Write-Host "MEZ file deployment completed successfully!" -ForegroundColor Green
    Write-Host "You can now use your custom connector in Power BI Desktop." -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Restart Power BI Desktop if it's running" -ForegroundColor White
    Write-Host "  2. Enable 'Allow any extension to load without validation' in Power BI options" -ForegroundColor White
    Write-Host "  3. Look for your connector in Get Data dialog" -ForegroundColor White
} else {
    Write-Host ""
    Write-Host "MEZ file deployment failed!" -ForegroundColor Red
    exit 1
}