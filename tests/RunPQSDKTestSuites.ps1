<#
       .SYNOPSIS
       Script: RunPQSDKTestSuites.ps1
       Runs the pre-built PQ/PQOut format tests in Power Query SDK Test Framework using pqtest.exe compare command.

       .DESCRIPTION
       This script will execute the PQ SDK PQ/PQOut tests present under Sanity, Standard & DataSourceSpecific folders. 
       RunPQSDKTestSuitesSettings.json file is used provide configurations need to this script. Please review the template RunPQSDKTestSuitesSettingsTemplate.json for more info.
       Pre-Requisite: Ensure the credentials are setup for your connector following the instructions here: https://learn.microsoft.com/power-query/power-query-sdk-vs-code#set-credential

       .LINK
       General pqtest.md: https://learn.microsoft.com/power-query/sdk-tools/pqtest-overview
       Compare command specific pqtest.md : https://learn.microsoft.com/power-query/sdk-tools/pqtest-commands-options
       
       .PARAMETER PQTestExePath
       Provide the path to PQTest.exe. Ex: 'C:\\Users\\ContosoUser\\.vscode\\extensions\\powerquery.vscode-powerquery-sdk-0.2.3-win32-x64\\.nuget\\Microsoft.PowerQuery.SdkTools.2.114.4\\tools\\PQTest.exe'
       
       .PARAMETER ExtensionPath
       Provide the path to extension .mez or .pqx file. Ex: 'C:\\dev\\ConnectorName\\ConnectorName.mez'

       .PARAMETER TestSettingsDirectoryPath
       Provide the path to TestSettingsDirectory folder. Ex: 'C:\\dev\\DataConnectors\\testframework\\tests\\ConnectorConfigs\\ConnectorName\\Settings'

       .PARAMETER TestSettingsList
       Provide the list of settings file needed to initialize the Test Result Object. Ex: SanitySettings.json StandardSettings.json
       
       .PARAMETER ValidateQueryFolding
       Optional parameter to specify if query folding needs to be verified as part of the test run

       .PARAMETER DetailedResults
       Optional parameter to specify if detailed results are needed along with summary results
       
       .PARAMETER JSONResults
       Optional parameter to specify if detailed results are needed along with summary results

       .EXAMPLE 
       PS> .\RunPQSDKTestSuites.ps1

       .EXAMPLE 
       PS> .\RunPQSDKTestSuites.ps1 -DetailedResults

       .EXAMPLE 
       PS> .\RunPQSDKTestSuites.ps1 -JSONResults
#>

param(
    [string]$PQTestExePath,
    [string]$ExtensionPath,
    [string]$TestSettingsDirectoryPath,
    [string[]]$TestSettingsList,
    [switch]$ValidateQueryFolding,
    [switch]$DetailedResults,
    [switch]$JSONResults,
    [switch]$HTMLPreview,
    [switch]$SkipConfirmation,
    [switch]$NoOpenHTML
)

function Get-PQTestPath {
    $basePath = Join-Path -Path $env:USERPROFILE -ChildPath ".vscode\extensions"
    $sdkPattern = "powerquery.vscode-powerquery-sdk-*"
    $toolsPattern = "Microsoft.PowerQuery.SdkTools.*"
    
    # Find the SDK extension folder
    $sdkFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $sdkPattern }
    
    if ($sdkFolders) {
        # Use the newest version if multiple exist
        $latestSdk = $sdkFolders | Sort-Object Name -Descending | Select-Object -First 1
        
        # Look for the tools folder
        $nugetPath = Join-Path -Path $latestSdk.FullName -ChildPath ".nuget"
        if (Test-Path $nugetPath) {
            $toolsFolders = Get-ChildItem -Path $nugetPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $toolsPattern }
            
            if ($toolsFolders) {
                $latestTools = $toolsFolders | Sort-Object Name -Descending | Select-Object -First 1
                $pqTestPath = Join-Path -Path $latestTools.FullName -ChildPath "tools\PQTest.exe"
                
                if (Test-Path $pqTestPath) {
                    return $pqTestPath
                }
            }
        }
    }
    
    return $null
}

function Read-PQOutFile {
    param(
        [string]$FilePath
    )
    
    try {
        $content = Get-Content -Path $FilePath -Raw
        
        # Parse the #table syntax using regex
        $tablePattern = '#table\s*\(\s*type\s+table\s+\[([^\]]+)\]\s*,\s*\{(.*)\}\s*\)'
        
        if ($content -match $tablePattern) {
            # Extract column definitions and data
            $columnDefs = $Matches[1]
            $dataRows = $Matches[2]
            
            # Parse column names from "Column1 = type1, Column2 = type2" format
            $columnNames = @()
            $columnMatches = [regex]::Matches($columnDefs, '(\w+)\s*=\s*\w+')
            foreach ($match in $columnMatches) {
                $columnNames += $match.Groups[1].Value
            }
            
            # Parse data rows - each row is like {"val1", "val2", "val3"}
            $parsedRows = @()
            $rowPattern = '\{([^}]+)\}'
            $rowMatches = [regex]::Matches($dataRows, $rowPattern)
            
            foreach ($rowMatch in $rowMatches) {
                $rowData = $rowMatch.Groups[1].Value
                # Split by commas, but handle quoted strings properly
                $values = @()
                $valuePattern = '"([^"]*)"'
                $valueMatches = [regex]::Matches($rowData, $valuePattern)
                foreach ($valueMatch in $valueMatches) {
                    $values += $valueMatch.Groups[1].Value
                }
                
                if ($values.Count -gt 0) {
                    $parsedRows += ,$values
                }
            }
            
            return @{
                ColumnNames = $columnNames
                Rows = $parsedRows
                Success = $true
            }
        } else {
            # Handle simple string results (like your "Connected Successfully")
            return @{
                Content = $content.Trim()
                Success = $true
                IsSimpleResult = $true
            }
        }
    } catch {
        return @{
            Error = $_.Exception.Message
            Success = $false
        }
    }
}

function New-HTMLTestReport {
    param(
        [string]$TestSuiteName,
        [array]$TestResults,
        [string]$OutputPath = "TestResults.html"
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Power Query Test Results - $TestSuiteName</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .header { background: #0078d4; color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .header-content { display: flex; align-items: flex-start; }
        .header-left { flex: 1; }
        .header h1 { margin: 0 0 10px 0; font-size: 1.8em; }
        .header p { margin: 5px 0; font-size: 1.1em; }
        .timestamp { color: white !important; font-size: 0.95em !important; font-weight: normal; opacity: 0.9; }
        .summary { display: flex; gap: 20px; margin-bottom: 20px; }
        .stat-card { background: #f8f9fa; padding: 15px; border-radius: 5px; text-align: center; flex: 1; }
        .pass { color: #107c10; font-weight: bold; }
        .fail { color: #d13438; font-weight: bold; }
        .info { color: #0078d4; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; background: white; }
        th { background: #f8f9fa; border: 1px solid #dee2e6; padding: 12px; text-align: left; font-weight: bold; }
        td { border: 1px solid #dee2e6; padding: 12px; }
        .test-section { margin-bottom: 30px; border: 1px solid #dee2e6; border-radius: 5px; padding: 15px; }
        .test-title { background: #e3f2fd; padding: 10px; border-radius: 5px; font-weight: bold; margin-bottom: 15px; font-size: 1.1em; }
        .pass-cell { background: #d4edda; color: #155724; font-weight: bold; }
        .fail-cell { background: #f8d7da; color: #721c24; font-weight: bold; }
        .info-cell { background: #d1ecf1; color: #0c5460; font-weight: bold; }
        .simple-result { padding: 15px; border-radius: 5px; margin: 10px 0; font-size: 1.2em; font-weight: bold; text-align: center; }
        .simple-success { background: #d4edda; color: #155724; }
        .simple-fail { background: #f8d7da; color: #721c24; }
        .no-results { background: #fff3cd; color: #856404; padding: 15px; border-radius: 5px; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-content">
                <div class="header-left">
                    <h1>TEST RESULTS - Power Query</h1>
                    <p>Test Suite: $TestSuiteName</p>
                    <p class="timestamp">Generated: $timestamp</p>
                </div>
            </div>
        </div>
"@

    # Add test results
    foreach ($testFile in $TestResults) {
        $testName = $testFile.Name -replace '\.query$', ''
        $pqoutPath = $testFile.FullName -replace '\.query\.pq$', '.query.pqout'
        
        $htmlContent += "`n        <div class='test-section'>"
        $htmlContent += "`n            <div class='test-title'>TEST: $testName</div>"
        
        if (Test-Path $pqoutPath) {
            $parsed = Read-PQOutFile -FilePath $pqoutPath
            
            if ($parsed.Success -and !$parsed.IsSimpleResult) {
                # Table data
                $htmlContent += "`n            <table>`n                <thead>`n                    <tr>"
                foreach ($col in $parsed.ColumnNames) {
                    $htmlContent += "`n                        <th>$col</th>"
                }
                $htmlContent += "`n                    </tr>`n                </thead>`n                <tbody>"
                
                foreach ($row in $parsed.Rows) {
                    $htmlContent += "`n                    <tr>"
                    for ($i = 0; $i -lt $parsed.ColumnNames.Count -and $i -lt $row.Count; $i++) {
                        $cellClass = ""
                        $value = $row[$i]
                        
                        if ($parsed.ColumnNames[$i] -eq "Result" -or $parsed.ColumnNames[$i] -eq "Status" -or $parsed.ColumnNames[$i] -eq "ValidationResult") {
                            switch ($value) {
                                "Pass" { 
                                    $cellClass = "pass-cell"
                                    $value = "PASS"
                                }
                                "Fail" { 
                                    $cellClass = "fail-cell"
                                    $value = "FAIL"
                                }
                                "Info" { 
                                    $cellClass = "info-cell"
                                    $value = "INFO"
                                }
                            }
                        }
                        
                        $htmlContent += "`n                        <td class='$cellClass'>$value</td>"
                    }
                    $htmlContent += "`n                    </tr>"
                }
                $htmlContent += "`n                </tbody>`n            </table>"
            } elseif ($parsed.Success -and $parsed.IsSimpleResult) {
                $statusClass = if ($parsed.Content -like "*Success*") { "simple-success" } else { "simple-fail" }
                $htmlContent += "`n            <div class='simple-result $statusClass'>"
                $htmlContent += "`n                Result: $($parsed.Content)"
                $htmlContent += "`n            </div>"
            } else {
                $htmlContent += "`n            <div class='simple-result simple-fail'>"
                $htmlContent += "`n                Error: Could not parse test results"
                $htmlContent += "`n            </div>"
            }
        } else {
            $htmlContent += "`n            <div class='no-results'>"
            $htmlContent += "`n                Warning: No results file found (.pqout)"
            $htmlContent += "`n            </div>"
        }
        
        $htmlContent += "`n        </div>"
    }
    
    $htmlContent += "`n    </div>`n</body>`n</html>"

    # Write HTML file
    $htmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
    return $OutputPath
}

function Remove-OldHTMLReports {
    param(
        [string]$ResultsDirectory,
        [string]$TestSuiteName,
        [int]$MaxFiles = 3
    )
    
    if (!(Test-Path $ResultsDirectory)) {
        return
    }
    
    # Find all HTML files for this specific test suite
    $pattern = "TestResults_$TestSuiteName`_*.html"
    $existingFiles = Get-ChildItem -Path $ResultsDirectory -Filter $pattern | Sort-Object CreationTime -Descending
    
    if ($existingFiles.Count -gt $MaxFiles) {
        $filesToDelete = $existingFiles | Select-Object -Skip $MaxFiles
        
        Write-Host "Cleaning up old HTML reports for $TestSuiteName..." -ForegroundColor Yellow
        foreach ($file in $filesToDelete) {
            try {
                Remove-Item -Path $file.FullName -Force
                Write-Host "  Deleted: $($file.Name)" -ForegroundColor Gray
            } catch {
                Write-Host "  Warning: Could not delete $($file.Name)" -ForegroundColor Yellow
            }
        }
        Write-Host "Kept $MaxFiles most recent reports for $TestSuiteName" -ForegroundColor Green
    }
}

function Show-TestResultsInVSCode {
    param(
        [string]$TestSuitePath,
        [string]$SuiteName,
        [switch]$NoOpenHTML
    )
    
    # Create results directory if it doesn't exist
    $resultsDir = Join-Path -Path (Get-Location) -ChildPath "results"
    if (!(Test-Path $resultsDir)) {
        New-Item -ItemType Directory -Path $resultsDir -Force | Out-Null
        Write-Host "Created Results directory: $resultsDir" -ForegroundColor Green
    }
    
    # Find test files
    $testFiles = Get-ChildItem -Path $TestSuitePath -Filter '*.query.pq' -Recurse
    
    if ($testFiles.Count -eq 0) {
        Write-Host "No test files found in $TestSuitePath" -ForegroundColor Yellow
        return
    }
    
    # Clean up old reports BEFORE creating new one
    $safeTestName = $SuiteName -replace '[^\w\-_]', '_'  # Replace invalid filename chars
    Remove-OldHTMLReports -ResultsDirectory $resultsDir -TestSuiteName $safeTestName -MaxFiles 3
    
    # Generate HTML report with descriptive filename
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $htmlFileName = "TestResults_$safeTestName`_$timestamp.html"
    $htmlPath = Join-Path -Path $resultsDir -ChildPath $htmlFileName
    
    $generatedFile = New-HTMLTestReport -TestSuiteName $SuiteName -TestResults $testFiles -OutputPath $htmlPath
    
    Write-Host "`n" -ForegroundColor Green
    Write-Host ("=" * 80) -ForegroundColor Green
    Write-Host "HTML REPORT CREATED SUCCESSFULLY!" -ForegroundColor Green
    Write-Host ("=" * 80) -ForegroundColor Green
    Write-Host "Location: $generatedFile" -ForegroundColor Cyan
    Write-Host ""
    
    # Only open files if NoOpenHTML is not specified
    if (-not $NoOpenHTML) {
        # Open in default browser (most reliable method)
        try {
            Start-Process $generatedFile
            Write-Host "[SUCCESS] Opened in default browser" -ForegroundColor Green
        } catch {
            Write-Host "[ERROR] Could not auto-open browser" -ForegroundColor Red
        }
        
        # Also try to open the HTML file in VS Code for manual preview
        try {
            & code $generatedFile
            Write-Host "[SUCCESS] Opened HTML file in VS Code" -ForegroundColor Green
        } catch {
            Write-Host "[WARNING] Could not open in VS Code" -ForegroundColor Yellow
        }
    } else {
        Write-Host "[INFO] HTML file opening skipped (NoOpenHTML enabled)" -ForegroundColor Cyan
    }
    
    # Show remaining files for this test suite
    $remainingFiles = Get-ChildItem -Path $resultsDir -Filter "TestResults_$safeTestName`_*.html" | Sort-Object CreationTime -Descending
    Write-Host ""
    Write-Host "Available reports for $SuiteName (newest first):" -ForegroundColor Blue
    foreach ($file in $remainingFiles) {
        $age = if ($file.CreationTime -gt (Get-Date).AddHours(-1)) { "Just created" } 
               elseif ($file.CreationTime -gt (Get-Date).AddDays(-1)) { "Recent" }
               else { "Older" }
        Write-Host "  $($file.Name) [$age]" -ForegroundColor White
    }
    
    if (-not $NoOpenHTML) {
        # Provide clear manual instructions
        Write-Host ""
        Write-Host "MANUAL PREVIEW OPTIONS IN VS CODE:" -ForegroundColor Blue
        Write-Host "  1. Right-click the HTML file → 'Open Preview'" -ForegroundColor White
        Write-Host "  2. With HTML file open: Ctrl+Shift+V" -ForegroundColor White
        Write-Host "  3. Command Palette: 'Live Preview: Show Preview'" -ForegroundColor White
        Write-Host ""
        Write-Host "Or simply use the browser that just opened!" -ForegroundColor Yellow
    }
    Write-Host ("=" * 80) -ForegroundColor Green
}





function Resolve-RelativePath {
    param(
        [string]$Path,
        [string]$BasePath = (Get-Location)
    )
    
    if ([System.IO.Path]::IsPathRooted($Path)) {
        # Already absolute path
        return $Path
    }
    
    # Handle relative paths
    $resolvedPath = Join-Path -Path $BasePath -ChildPath $Path
    
    # Resolve any .. or . references
    try {
        return [System.IO.Path]::GetFullPath($resolvedPath)
    }
    catch {
        return $resolvedPath
    }
}


# Pre-Requisite:   
# Ensure the credentials are setup for your connector following the instructions here: https://learn.microsoft.com/en-us/power-query/power-query-sdk-vs-code#set-credential 

# Retrieving the settings for running the TestSuites from the JSON settings file
$RunPQSDKTestSuitesSettings = Get-Content -Path RunPQSDKTestSuitesSettings.json | ConvertFrom-Json

# GENERATE ABSOLUTE PATH CONFIGURATION FOR MOCK DATA TESTS
if ($TestSettingsList -contains "MockDataTests.json") {
    $MockDataFilePath = Join-Path -Path (Get-Location) -ChildPath "MockData\jira_api_response_acme_software_pb.json"
    
    # Resolve to absolute path
    if (Test-Path $MockDataFilePath) {
        $AbsoluteMockDataPath = (Resolve-Path $MockDataFilePath).Path
        
        # Create configuration object for Power Query to read
        $MockDataConfig = @{
            "MockDataAbsolutePath" = $AbsoluteMockDataPath
            "GeneratedTimestamp" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            "WorkingDirectory" = (Get-Location).Path
        }
        
        # Write config into the TestSuites/MockData folder
        $TestSuiteMockFolder = Join-Path -Path (Get-Location) -ChildPath "TestSuites\MockData"
        $ConfigPath = Join-Path -Path $TestSuiteMockFolder -ChildPath "MockDataConfig.json"

        $MockDataConfig | ConvertTo-Json | Out-File $ConfigPath -Encoding UTF8
        Write-Host "Written config to test folder: $ConfigPath" -ForegroundColor Green
        Write-Host "Generated mock data configuration file with absolute path:" -ForegroundColor Green
        Write-Host "  $AbsoluteMockDataPath" -ForegroundColor Cyan
    } else {
        Write-Host "ERROR: Mock data file not found at: $MockDataFilePath" -ForegroundColor Red
        Write-Host "Please ensure the mock data file exists before running mock data tests." -ForegroundColor Red
    }
}

# Setting the PQTestExePath from settings object if not passed as an argument
if (!$PQTestExePath) { 
    # Try auto-discovery first
    $discoveredPath = Get-PQTestPath
    if ($discoveredPath) {
        $PQTestExePath = $discoveredPath
        Write-Output("Auto-discovered PQTest.exe at: " + $PQTestExePath)
    } else {
        # Fallback to JSON settings
        $PQTestExePath = $RunPQSDKTestSuitesSettings.PQTestExePath
        if ($PQTestExePath) {
            Write-Output("Using PQTest.exe path from settings: " + $PQTestExePath)
        }
    }
}

if (!(Test-Path -Path $PQTestExePath)){
    $PSStyle.Foreground.Red
    Write-Output("PQTestExe path could not be found or is not correctly set.")
    Write-Output("Attempted path: " + $PQTestExePath)
    Write-Output("")
    Write-Output("Solutions:")
    Write-Output("1. Ensure Power Query SDK is installed in VS Code")
    Write-Output("2. Set explicit path in RunPQSDKTestSuitesSettings.json") 
    Write-Output("3. Pass path as -PQTestExePath argument")
    Write-Output("")
    Write-Output("Expected VS Code extension location: " + (Join-Path -Path $env:USERPROFILE -ChildPath ".vscode\extensions\powerquery.vscode-powerquery-sdk-*"))
    $PSStyle.Reset
    exit
}

# Setting the ExtensionPath from settings object if not passed as an argument
if (!$ExtensionPath){ $ExtensionPath = $RunPQSDKTestSuitesSettings.ExtensionPath }
if (!(Test-Path -Path $ExtensionPath)){
    Write-Output("Extension path is not correctly set. Either set it in RunPQSDKTestSuitesSettings.json or pass it as an argument. " + $ExtensionPath)
    exit
}

# Setting the TestSettingsDirectoryPath if not passed as an argument
if (!$TestSettingsDirectoryPath){
    if ($RunPQSDKTestSuitesSettings.TestSettingsDirectoryPath){
        $TestSettingsDirectoryPath = $RunPQSDKTestSuitesSettings.TestSettingsDirectoryPath
    }
    else 
    {
        $GenericTestSettingsDirectoryPath  = Join-Path -Path (Get-Location) -ChildPath ("ConnectorConfigs\generic\Settings")
        $TestSettingsDirectoryPath = Join-Path -Path (Get-Location) -ChildPath ("ConnectorConfigs\" + (Get-Item $ExtensionPath).Basename + "\Settings")

        # Creating the test settings and parameter query file(s) automatically 
        if (!(Test-Path $TestSettingsDirectoryPath)){
            $PSStyle.Foreground.Blue
            Write-Output("Performing the initial setup by creating the test settings and parameter query file(s) automatically...");
            $PSStyle.Reset
            Copy-Item -Path $GenericTestSettingsDirectoryPath -Destination $TestSettingsDirectoryPath -Recurse
            Write-Output("Successfully created test settings file(s) under the directory:`n" + $TestSettingsDirectoryPath);
        }
        $GenericParameterQueriesPath  = Join-Path -Path (Get-Location) -ChildPath ("ConnectorConfigs\generic\ParameterQueries")
        $ExtensionParameterQueriesPath = Join-Path -Path (Get-Location) -ChildPath ("ConnectorConfigs\" + (Get-Item $ExtensionPath).Basename + "\ParameterQueries")
        if (!(Test-Path $ExtensionParameterQueriesPath)){
            Copy-Item -Path $GenericParameterQueriesPath -Destination $ExtensionParameterQueriesPath -Recurse
            Rename-Item -Path (Join-Path -Path $ExtensionParameterQueriesPath -ChildPath "Generic.parameterquery.pq")  -NewName ((Get-Item $ExtensionPath).Basename +".parameterquery.pq")
            Write-Output("Successfully created the parameter query file(s) under the directory:`n" + $ExtensionParameterQueriesPath);
            $PSStyle.Reset

            # Updating the parameter query file(s) location in the test setting file
            foreach ($SettingsFile in (Get-ChildItem $TestSettingsDirectoryPath | ForEach-Object {$_.FullName})){
                $SettingsFileJson = Get-Content -Path $SettingsFile | ConvertFrom-Json
                $SettingsFileJson.ParameterQueryFilePath = $SettingsFileJson.ParameterQueryFilePath.ToLower().Replace("generic", (Get-Item $ExtensionPath).Basename)
                $SettingsFileJson | ConvertTo-Json -depth 100 |  Set-Content $SettingsFile
            }

            # Prompting the user to update the parameter query file(s)
            $PSStyle.Foreground.Magenta
            $parameterQueryUpdated = Read-Host ("Please update the parameter query file(s) generated by replacing with the M query to connect to your data source and retrieve Jira Data.`nAre File(s) updated? [y/n]")
            $PSStyle.Reset

            while($parameterQueryUpdated -ne "y")
            {
                if ($parameterQueryUpdated -eq 'n') {
                    $PSStyle.Foreground.Red
                    Write-Host("Please update the parameter query file(s) generated by replacing with the M query to connect to your data source and retrieve Jira Data and rerun the script.")
                    $PSStyle.Reset
                    exit
                }
                $PSStyle.Foreground.Yellow
                $parameterQueryUpdated = Read-Host "Please update the parameter query file(s) generated by replacing with the M query to connect to your data source and retrieve Jira Data.`nAre File(s) updated? [y/n]"
                $PSStyle.Reset
            }
    }
} 
}
if (!(Test-Path -Path $TestSettingsDirectoryPath)){ 
    Write-Output("Test Settings Directory is not correctly set. Either set it in RunPQSDKTestSuitesSettings.json or pass it as an argument. " + $TestSettingsDirectoryPath)
    exit
}


#Setting the TestSettingsList if not passed as an argument

if (!$TestSettingsList){
    if ($RunPQSDKTestSuitesSettings.TestSettingsList){
        $TestSettingsList = $RunPQSDKTestSuitesSettings.TestSettingsList
    }
    else{
    $TestSettingsList = (Get-ChildItem -Path $TestSettingsDirectoryPath -Name)
    }
}

#Setting the ValidateQueryFolding if not passed as an argument
if (!$ValidateQueryFolding){ 
    if ($RunPQSDKTestSuitesSettings.ValidateQueryFolding -eq "True"){
         $ValidateQueryFolding = $true 
    }
}

#Setting the DetailedResults if not passed as an argument
if (!$DetailedResults){ 
    if ($RunPQSDKTestSuitesSettings.DetailedResults -eq "True"){
        $DetailedResults = $true 
    } 
}

#Setting the JSONResults if not passed as an argument
if (!$JSONResults){
    if ($RunPQSDKTestSuitesSettings.JSONResults -eq "True"){
        $JSONResults = $true
    }
}

$PSStyle.Foreground.Blue
Write-Output("Below are settings for running the TestSuites:")
$PSStyle.Reset
Write-Output ("PQTestExePath: " + $PQTestExePath)
Write-Output ("ExtensionPath: " + $ExtensionPath)
Write-Output ("TestSettingsDirectoryPath: " + $TestSettingsDirectoryPath)
Write-Output ("TestSettingsList: " + $TestSettingsList)
Write-Output ("ValidateQueryFolding: " + $ValidateQueryFolding)
Write-Output ("DetailedResults: " + $DetailedResults)
Write-Output ("JSONResults: "+ $JSONResults)

$ExtensionParameterQueriesPath = Join-Path -Path (Get-Location) -ChildPath ("ConnectorConfigs\" + (Get-Item $ExtensionPath).Basename + "\ParameterQueries")
$DiagnosticFolderPath = Join-Path -Path (Get-Location) -ChildPath ("Diagnostics\" + (Get-Item $ExtensionPath).Basename)

$PSStyle.Foreground.Magenta
Write-Output ("Note: Please verify the settings above and ensure the following:
1. Credentials are setup for the extension following the instructions here: https://learn.microsoft.com/en-us/power-query/power-query-sdk-vs-code#set-credential
2. Parameter query file(s) are updated with the M query to connect to your data source and retrieve Jira data under: 
$ExtensionParameterQueriesPath")
if ($ValidateQueryFolding){ 
Write-Output ("3. Diagnostics folder path that will be used for query folding verfication: 
$DiagnosticFolderPath")
}
$PSStyle.Reset

if ($SkipConfirmation) {
    Write-Host "SkipConfirmation enabled - proceeding automatically" -ForegroundColor Green
} else {
    $confirmation = Read-Host "Do you want to proceed? [y/n]" 
    while($confirmation -ne "y")
    {
        if ($confirmation -eq 'n') {
            $PSStyle.Foreground.Yellow
            Write-Host("Please specify the correct settings in RunPQSDKTestSuitesSettings.json or pass them arguments and re-run the script.")
            $PSStyle.Reset
            exit
        }
        $confirmation = Read-Host "Do you want to proceed? [y/n]"
    }
}

# Creating the DiagnosticFolderPath if ValidateQueryFolding is set to true
if ($ValidateQueryFolding){ 
    $DiagnosticFolderPath = Join-Path -Path (Get-Location) -ChildPath ("Diagnostics\" + (Get-Item $ExtensionPath).Basename)
    if (!(Test-Path $DiagnosticFolderPath)){
        New-Item -ItemType Directory -Force -Path $DiagnosticFolderPath | Out-Null
    }
}

# Created a class to store and display test results
class TestResult { 
    [string] $ParameterQuery; 
    [string] $TestFolder; 
    [string] $TestName; 
    [string] $OutputStatus; 
    [string] $TestStatus; 
    [string] $Duration; 

    TestResult([string]$testFolder, [string]$testName, [string]$outputStatus, [string]$testStatus, [string]$duration){
        # Constructor to Initialize the Test Result Object
        $this.TestFolder = $testFolder
        $this.TestName = $testName;
        $this.OutputStatus = $outputStatus;
        $this.TestStatus = $testStatus;
        $this.Duration = $duration;
    }
}

# Variable to Initialize the Test Result Object
$TestCount = 0
$Passed = 0
$Failed = 0
$TestExecStartTime = Get-Date 

$TestResults = @()
$RawTestResults = @()
$TestResultsObjects = @()

# Run the compare command for each of the TestSetttings Files
foreach ($TestSettings in $TestSettingsList){
    if ($ValidateQueryFolding) { 
        $RawTestResult = & $PQTestExePath compare -p -e $ExtensionPath -sf $TestSettingsDirectoryPath\$TestSettings -dfp $DiagnosticFolderPath
    }
    else {
        $RawTestResult = & $PQTestExePath compare -p -e $ExtensionPath -sf $TestSettingsDirectoryPath\$TestSettings
    }

    $StringTestResult = $RawTestResult -join " " 
    $TestResultsObject = $StringTestResult | ConvertFrom-Json 

    foreach($Result in $TestResultsObject){
           $TestResults += [TestResult]::new($Result.Name.Split("\")[-3] + "\" + $Result.Name.Split("\")[-2], $Result.Name.Split("\")[-1], $Result.Output.Status, $Result.Status, (NEW-TIMESPAN -Start $Result.StartTime  -End $Result.EndTime))   
           $TestCount += 1
            if($Result.Status -eq "Passed"){
                $Passed++ 
            }    
            else{
                $Failed++ 
            }                
    }

    $RawTestResults += $RawTestResult
    $TestResultsObjects += $TestResultsObject
}

# Display the test results
$TestExecEndTime = Get-Date 

if ($DetailedResults)
{
    Write-Output("------------------------------------------------------------------------------------------")
    Write-Output("PQ SDK Test Framework - Test Execution - Detailed Results for Extension: " + $ExtensionPath.Split("\")[-1] )
    Write-Output("------------------------------------------------------------------------------------------")
    $TestResultsObjects
}
if ($JSONResults)
{
    Write-Output("-----------------------------------------------------------------------------------")
    Write-Output("PQ SDK Test Framework - Test Execution - JSON Results for Extension: " + $ExtensionPath.Split("\")[-1] )
    Write-Output("-----------------------------------------------------------------------------------")
    $RawTestResults
}

Write-Output("----------------------------------------------------------------------------------------------")
Write-Output("PQ SDK Test Framework - Test Execution - Test Results Summary for Extension: " + $ExtensionPath.Split("\")[-1] )
Write-Output("----------------------------------------------------------------------------------------------")

$TestResults  | Format-Table -AutoSize -Property TestFolder, TestName, @{
    Label = "OutputStatus" 
    Expression=
    { 
        switch ($_.OutputStatus) {
            { $_ -eq "Passed" } { $color = "$($PSStyle.Foreground.Green)"  }
            { $_ -eq "Failed" } { $color = "$($PSStyle.Foreground.Red)"    }
            { $_ -eq "Error"  } { $color = "$($PSStyle.Foreground.Yellow)" }
    }
    "$color$($_.OutputStatus)$($PSStyle.Reset)" 
}
}, @{
    Label = "TestStatus" 
    Expression=
    { 
        switch ($_.TestStatus) {
            { $_ -eq "Passed" } { $color = "$($PSStyle.Foreground.Green)"  }
            { $_ -eq "Failed" } { $color = "$($PSStyle.Foreground.Red)"    }
            { $_ -eq "Error"  } { $color = "$($PSStyle.Foreground.Yellow)" }
    }
    "$color$($_.TestStatus)$($PSStyle.Reset)"
}
}, 
Duration

Write-Output("----------------------------------------------------------------------------------------------")
Write-Output("Total Tests: " + $TestCount + " | Passed: " + $Passed + " | Failed: " + $Failed +  " | Total Duration: " + "{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s" -f (NEW-TIMESPAN -Start $TestExecStartTime  -End $TestExecEndTime))
Write-Output("----------------------------------------------------------------------------------------------")

# --------------------------------------------------
# ERROR HANDLING: Abort HTML preview if any test failed
# --------------------------------------------------
$failedTests = $TestResultsObjects | Where-Object { $_.Output.Status -ne "Passed" }
if ($failedTests) {
    Write-Host "" -ForegroundColor Yellow #Red
    Write-Host "WARNING: One or more tests output failed. Generating HTML anyway for dashboard review." -ForegroundColor Yellow
    #Write-Host "ERROR: One or more tests failed. Skipping HTML report generation." -ForegroundColor Red
    foreach ($t in $failedTests) {
        Write-Host ("  Test: {0}, OutputStatus: {1}" -f $t.TestName, $t.OutputStatus) -ForegroundColor Yellow #Red
    }
    # return
    # Uncomment lines above to abort HTML generation on ouput failures
}

# Generate HTML preview if requested

if ($HTMLPreview) {
    Write-Output("`nGenerating HTML preview...")
    
    foreach ($TestSettings in $TestSettingsList) {
        $SettingsFileJson = Get-Content -Path (Join-Path $TestSettingsDirectoryPath $TestSettings) | ConvertFrom-Json
        $QueryPath = $SettingsFileJson.QueryFilePath
        
        if ($QueryPath) {
            # Resolve the query path
            $FullQueryPath = if ([System.IO.Path]::IsPathRooted($QueryPath)) { 
                $QueryPath 
            } else { 
                $ResolvedPath = Resolve-RelativePath -Path $QueryPath
                $ResolvedPath
            }
            
            # Generate HTML report with clean test name
            $SuiteName = $TestSettings -replace "\.json$", ""
            Write-Output("Creating HTML report for: $SuiteName")
            Show-TestResultsInVSCode -TestSuitePath $FullQueryPath -SuiteName $SuiteName -NoOpenHTML:$NoOpenHTML
        }
    }
    
    Write-Output("`nHTML preview generation completed!")
}
