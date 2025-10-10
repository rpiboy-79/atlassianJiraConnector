# Copy-MezFile.ps1
# PowerShell script to copy a MEZ file from the local project's build folder 
# to the Power BI Custom Connectors directory

param(
    [string]$SourceMezPath,
    [string]$TargetPath,
    [switch]$CopyToAll,
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
    Gets the Power BI Custom Connectors directory path(s),
    properly handling OneDrive integration (including multiple OneDrives)
    #>
    
    $localDocuments = [Environment]::GetFolderPath('MyDocuments')
    $standardPath = Join-Path $localDocuments "Power BI Desktop\Custom Connectors"

    $userProfile = $env:USERPROFILE
    $oneDriveRoots = Get-ChildItem -Path $userProfile -Directory -Filter "OneDrive*" -ErrorAction SilentlyContinue
    $paths = @()

    # Always include the standard local path as a fallback
    if (Test-Path $standardPath) {
        $paths += $standardPath
    }

    # Collect all possible OneDrive "Documents" folders
    foreach ($odr in $oneDriveRoots) {
        $odDocs = Join-Path $odr.FullName "Documents\Power BI Desktop\Custom Connectors"
        if (Test-Path $odDocs) {
            $paths += $odDocs
        }
    }

    # Return all found paths, or create standard path if none exist
    if ($paths.Count -gt 0) {
        return $paths
    }
    return @($standardPath)
}

function Select-TargetPath {
    <#
    .SYNOPSIS
    Allows user to select target path when multiple OneDrive locations exist
    #>
    param(
        [string[]]$AvailablePaths
    )
    
    if ($AvailablePaths.Count -eq 1) {
        return $AvailablePaths[0]
    }
    
    Write-Host ""
    Write-Host "Multiple Power BI Custom Connectors locations found:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $AvailablePaths.Count; $i++) {
        Write-Host "  [$($i + 1)] $($AvailablePaths[$i])" -ForegroundColor White
    }
    Write-Host ""
    
    do {
        $selection = Read-Host "Select target location (1-$($AvailablePaths.Count)) or press Enter for all"
        
        if ([string]::IsNullOrWhiteSpace($selection)) {
            return $AvailablePaths  # Return all paths
        }
        
        if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $AvailablePaths.Count) {
            return $AvailablePaths[[int]$selection - 1]
        }
        
        Write-Host "Invalid selection. Please enter a number between 1 and $($AvailablePaths.Count), or press Enter for all." -ForegroundColor Red
    } while ($true)
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
        Write-Host ""
        
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

# Get destination folder(s)
$destinationFolders = if ($TargetPath) {
    # Use explicitly provided target path
    @($TargetPath)
} else {
    # Get all available Power BI Custom Connectors paths
    $availablePaths = Get-PowerBIConnectorsPath
    
    if ($CopyToAll) {
        # Copy to all available paths
        $availablePaths
    } else {
        # Let user select or default to first available
        if ($availablePaths.Count -gt 1) {
            $selected = Select-TargetPath -AvailablePaths $availablePaths
            if ($selected -is [array]) {
                $selected  # User chose all
            } else {
                @($selected)  # User chose specific path
            }
        } else {
            $availablePaths
        }
    }
}

if ($Verbose) {
    Write-Host ""
    Write-Host "Configuration:" -ForegroundColor Yellow
    Write-Host "  Source MEZ: $sourceMez" -ForegroundColor White
    Write-Host "  Target Directories:" -ForegroundColor White
    foreach ($folder in $destinationFolders) {
        Write-Host "    - $folder" -ForegroundColor White
    }
    Write-Host ""
}

# Perform the copy operation(s)
$overallSuccess = $true
$successCount = 0

foreach ($destinationFolder in $destinationFolders) {
    Write-Host "Copying to: $destinationFolder" -ForegroundColor Cyan
    $success = Copy-MezToPowerBI -SourcePath $sourceMez -DestinationFolder $destinationFolder
    
    if ($success) {
        $successCount++
    } else {
        $overallSuccess = $false
    }
}

# Final status report
Write-Host ""
if ($overallSuccess -and $successCount -eq $destinationFolders.Count) {
    Write-Host "MEZ file deployment completed successfully to all $successCount location(s)!" -ForegroundColor Green
    Write-Host "You can now use your custom connector in Power BI Desktop." -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Restart Power BI Desktop if it's running" -ForegroundColor White
    Write-Host "  2. Enable 'Allow any extension to load without validation' in Power BI options" -ForegroundColor White
    Write-Host "  3. Look for your connector in Get Data dialog" -ForegroundColor White
} elseif ($successCount -gt 0) {
    Write-Host "MEZ file deployment partially successful ($successCount of $($destinationFolders.Count) locations)!" -ForegroundColor Yellow
    Write-Host "Check error messages above for failed locations." -ForegroundColor Yellow
} else {
    Write-Host "MEZ file deployment failed to all locations!" -ForegroundColor Red
    exit 1
}