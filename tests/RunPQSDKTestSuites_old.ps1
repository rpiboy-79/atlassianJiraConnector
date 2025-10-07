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
       Optional parameter to specify if JSON results should be generated and saved alongside HTML results

       .EXAMPLE 
       PS> .\RunPQSDKTestSuites.ps1

       .EXAMPLE 
       PS> .\RunPQSDKTestSuites.ps1 -DetailedResults

       .EXAMPLE 
       PS> .\RunPQSDKTestSuites.ps1 -JSONResults

       .EXAMPLE
       PS> .\RunPQSDKTestSuites.ps1 -HTMLPreview -JSONResults
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
    [switch]$NoOpenHTML,
    [switch]$CopyMez
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

function Get-ConnectorVersion {
    param(
        [string]$ExtensionPath
    )

    try {
        $connectorName = if ($ExtensionPath) { (Get-Item $ExtensionPath).BaseName } else { "Unknown" }
        $connectorVersion = "1.0.0"  # Default version

        return @{
            Name = $connectorName
            Version = $connectorVersion
        }
    } catch {
        return @{
            Name = "Unknown"
            Version = "1.0.0"
        }
    }
}

function Get-ExpectedResults {
    param(
        [object]$TestResult
    )

    try {
        if ($TestResult.Output -and $TestResult.Output.OutputFilePath) {
            $pqoutPath = $TestResult.Output.OutputFilePath
            if (Test-Path $pqoutPath) {
                $content = Get-Content $pqoutPath -Raw -Encoding UTF8
                Write-Host "Loaded expected results from: $pqoutPath" -ForegroundColor Green
                return $content.Trim()
            } else {
                Write-Host "Expected results file not found: $pqoutPath" -ForegroundColor Yellow
            }
        }
        return $null
    } catch {
        Write-Host "Error reading expected results: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Get-ActualResults {
    param(
        [object]$TestResult
    )

    try {
        # For failed tests, use SerializedSource (what actually happened)
        if ($TestResult.Status -eq "Failed") {
            if ($TestResult.Output -and $TestResult.Output.SerializedSource -and $TestResult.Output.SerializedSource -ne "") {
                Write-Host "Using SerializedSource for failed test: $($TestResult.Name)" -ForegroundColor Yellow
                return $TestResult.Output.SerializedSource
            }
        }

        # For passed tests or failed tests without SerializedSource, read .pqout file
        # This represents the current "expected" result that the test passed against
        if ($TestResult.Output -and $TestResult.Output.OutputFilePath) {
            $pqoutPath = $TestResult.Output.OutputFilePath
            if (Test-Path $pqoutPath) {
                $content = Get-Content $pqoutPath -Raw -Encoding UTF8
                Write-Host "Using .pqout content for actual results: $($TestResult.Name)" -ForegroundColor Blue
                return $content.Trim()
            }
        }

        return $null
    } catch {
        Write-Host "Error reading actual results: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Test-DataDrift {
    param(
        [string]$ExpectedResults,
        [string]$ActualResults
    )

    try {
        # Simple comparison - can be enhanced in future phases
        if ($ExpectedResults -and $ActualResults) {
            $isDifferent = $ExpectedResults -ne $ActualResults
            return @{
                HasDrift = $isDifferent
                ExpectedHash = if ($ExpectedResults) { (Get-FileHash -InputStream ([System.IO.MemoryStream]::new([System.Text.Encoding]::UTF8.GetBytes($ExpectedResults)))).Hash } else { $null }
                ActualHash = if ($ActualResults) { (Get-FileHash -InputStream ([System.IO.MemoryStream]::new([System.Text.Encoding]::UTF8.GetBytes($ActualResults)))).Hash } else { $null }
            }
        }

        return @{
            HasDrift = $false
            ExpectedHash = $null
            ActualHash = $null
        }
    } catch {
        Write-Host "Error detecting data drift: $($_.Exception.Message)" -ForegroundColor Yellow
        return @{
            HasDrift = $false
            ExpectedHash = $null
            ActualHash = $null
        }
    }
}

function Get-BaselineTimestamp {
    param(
        [object]$TestResult
    )

    try {
        if ($TestResult.Output -and $TestResult.Output.OutputFilePath) {
            $pqoutPath = $TestResult.Output.OutputFilePath
            if (Test-Path $pqoutPath) {
                $fileInfo = Get-Item $pqoutPath
                return $fileInfo.CreationTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
            }
        }
        return $null
    } catch {
        return $null
    }
}


function Write-JSONTestResults {
    param(
        [string]$TestSuiteName,
        [array]$TestResultsObjects,
        [array]$RawTestResults,
        [string]$ExtensionPath,
        [string]$ResultsDirectory = "./results"
    )

    try {
        # Create results directory if it doesn't exist
        if (!(Test-Path $ResultsDirectory)) {
            New-Item -ItemType Directory -Path $ResultsDirectory -Force | Out-Null
        }

        # Get connector information safely
        $connectorInfo = Get-ConnectorVersion -ExtensionPath $ExtensionPath

        # Generate timestamp and filename
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $runDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $executionTimestamp = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        $jsonFileName = "TestResults_$TestSuiteName`_$timestamp.json"
        $jsonPath = Join-Path -Path $ResultsDirectory -ChildPath $jsonFileName

        Write-Host "Processing $($TestResultsObjects.Count) test result objects for Enhanced JSON export..." -ForegroundColor Cyan

        # Process test results with enhanced data capture
        $processedResults = @()
        $totalDurationSum = 0

        foreach ($testResult in $TestResultsObjects) {
            try {
                # Safe property extraction with defaults
                $testName = if ($testResult.Name) { 
                    $testResult.Name.Split("\")[-1] -replace '\.query\.pq$', '' 
                } else { 
                    "UnknownTest" 
                }

                # Safe timestamp handling
                $startTime = Get-Date
                $endTime = Get-Date
                $duration = 0

                if ($testResult.StartTime -and $testResult.EndTime) {
                    try {
                        $startTime = [DateTime]$testResult.StartTime
                        $endTime = [DateTime]$testResult.EndTime
                        $duration = ($endTime - $startTime).TotalSeconds
                    } catch {
                        Write-Host "Warning: Could not parse timestamps for test $testName" -ForegroundColor Yellow
                        $duration = 0
                    }
                }

                # Add to total duration sum
                $totalDurationSum += $duration

                # PHASE 1 ENHANCEMENTS: Get both expected and actual results
                $expectedResults = Get-ExpectedResults -TestResult $testResult
                $actualResults = Get-ActualResults -TestResult $testResult
                $baselineTimestamp = Get-BaselineTimestamp -TestResult $testResult
                $driftInfo = Test-DataDrift -ExpectedResults $expectedResults -ActualResults $actualResults

                # Safe property extraction
                $result = if ($testResult.Status) { $testResult.Status } else { "Unknown" }
                $outputStatus = if ($testResult.Output -and $testResult.Output.Status) { 
                    $testResult.Output.Status 
                } else { 
                    "Unknown" 
                }

                # ENHANCED RESULT OBJECT WITH PHASE 1 FEATURES
                $processedResult = @{
                    # Original fields (maintained for backward compatibility)
                    test_name = $testName
                    test_file = if ($testResult.Name) { $testResult.Name } else { $testName }
                    result = $result
                    output_status = $outputStatus
                    start_time = $startTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
                    end_time = $endTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
                    duration_seconds = [Math]::Round($duration, 2)
                    details = if ($testResult.Details) { $testResult.Details } else { "No details available" }
                    method = if ($testResult.Method) { $testResult.Method } else { "Unknown" }

                    # Maintain serialized_source for dashboard compatibility
                    serialized_source = $actualResults  # Dashboard expects this field

                    source_error = if ($testResult.Output -and $testResult.Output.SourceError) { 
                        $testResult.Output.SourceError 
                    } else { 
                        $null 
                    }
                    output_error = if ($testResult.Output -and $testResult.Output.OutputError) { 
                        $testResult.Output.OutputError 
                    } else { 
                        $null 
                    }

                    # PHASE 1 ENHANCEMENTS: TDD + Observability
                    expected_result = $expectedResults
                    actual_result = $actualResults

                    # Historical tracking
                    execution_timestamp = $executionTimestamp
                    baseline_timestamp = $baselineTimestamp

                    # Data drift detection
                    data_drift_detected = $driftInfo.HasDrift
                    expected_hash = $driftInfo.ExpectedHash
                    actual_hash = $driftInfo.ActualHash

                    # Test evolution tracking
                    test_evolution = @{
                        baseline_established = $baselineTimestamp
                        current_execution = $executionTimestamp
                        comparison_method = if ($testResult.Status -eq "Failed" -and $testResult.Output.SerializedSource) { 
                            "TDD_Failed_Comparison" 
                        } else { 
                            "Baseline_Comparison" 
                        }
                        has_historical_data = ($baselineTimestamp -ne $null)
                    }

                    # Enhanced metadata
                    test_metadata = @{
                        test_type = "PQTest_Compare"
                        data_source_state = "Live"
                        generation_method = "Enhanced_Phase1"
                        schema_version = "1.1"
                    }
                }

                $processedResults += $processedResult

            } catch {
                Write-Host "Warning: Error processing test result: $($_.Exception.Message)" -ForegroundColor Yellow
                # Add a minimal result entry for failed processing
                $processedResults += @{
                    test_name = "ProcessingError"
                    test_file = "Unknown"
                    result = "Error" 
                    output_status = "ProcessingFailed"
                    start_time = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
                    end_time = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
                    duration_seconds = 0
                    details = $_.Exception.Message
                    method = "Unknown"
                    serialized_source = $null
                    expected_result = $null
                    actual_result = $null
                    execution_timestamp = $executionTimestamp
                    data_drift_detected = $false
                    source_error = $null
                    output_error = $_.Exception.Message
                }
            }
        }

        Write-Host "Processed $($processedResults.Count) test results successfully with Phase 1 enhancements" -ForegroundColor Green

        # Calculate summary statistics safely
        $totalTests = $processedResults.Count
        $passedTests = ($processedResults | Where-Object { $_.result -eq "Passed" }).Count
        $failedTests = ($processedResults | Where-Object { $_.result -eq "Failed" }).Count
        $errorTests = ($processedResults | Where-Object { $_.result -eq "Error" }).Count
        $testsWithDrift = ($processedResults | Where-Object { $_.data_drift_detected -eq $true }).Count

        # Use the sum we calculated manually (safer than Measure-Object)
        $totalDuration = [Math]::Round($totalDurationSum, 2)

        Write-Host "Statistics: $totalTests total, $passedTests passed, $failedTests failed, $errorTests errors" -ForegroundColor Cyan
        Write-Host "Data Drift: $testsWithDrift tests show differences between expected and actual results" -ForegroundColor Magenta
        Write-Host "Total duration: $totalDuration seconds" -ForegroundColor Cyan

        # ENHANCED JSON STRUCTURE WITH PHASE 1 FEATURES
        $jsonOutput = @{
            # Enhanced metadata with Phase 1 features
            metadata = @{
                # Original metadata (maintained for compatibility)
                run_timestamp = $executionTimestamp
                run_date = $runDate
                test_suite_name = $TestSuiteName
                connector_name = $connectorInfo.Name
                connector_version = $connectorInfo.Version
                extension_path = $ExtensionPath
                total_tests = $totalTests
                passed_tests = $passedTests
                failed_tests = $failedTests
                error_tests = $errorTests
                total_duration_seconds = $totalDuration
                pqtest_version = "Unknown"
                generation_method = "PowerShell_Enhanced_Phase1"

                # Phase 1 enhancements
                schema_version = "1.1"
                execution_timestamp = $executionTimestamp
                baseline_comparison_enabled = $true
                data_drift_detection_enabled = $true
                tests_with_drift = $testsWithDrift
                historical_tracking_enabled = $true

                # Test execution context
                execution_context = @{
                    working_directory = (Get-Location).Path
                    powershell_version = $PSVersionTable.PSVersion.ToString()
                    generation_date = $runDate
                    enhancement_level = "Phase1_TDD_Observability"
                }
            }

            # Enhanced test results with Phase 1 data
            test_results = $processedResults

            # Raw test results (maintained for compatibility)
            raw_pqtest_output = if ($RawTestResults) { $RawTestResults } else { @() }

            # Phase 1 summary data
            summary_analytics = @{
                drift_analysis = @{
                    total_tests_checked = $totalTests
                    tests_with_drift = $testsWithDrift
                    drift_percentage = if ($totalTests -gt 0) { [Math]::Round(($testsWithDrift / $totalTests) * 100, 2) } else { 0 }
                }
                execution_analysis = @{
                    total_execution_time = $totalDuration
                    average_test_time = if ($totalTests -gt 0) { [Math]::Round($totalDuration / $totalTests, 2) } else { 0 }
                    fastest_test = ($processedResults | Sort-Object duration_seconds | Select-Object -First 1).test_name
                    slowest_test = ($processedResults | Sort-Object duration_seconds -Descending | Select-Object -First 1).test_name
                }
            }
        }

        # Write JSON file with error handling
        try {
            $jsonOutput | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
            Write-Host "Enhanced JSON results written to: $jsonPath" -ForegroundColor Green
            Write-Host "Phase 1 features included: Expected/Actual results, Data drift detection, Historical tracking" -ForegroundColor Cyan
            return $jsonPath
        } catch {
            Write-Host "ERROR: Failed to write JSON file: $($_.Exception.Message)" -ForegroundColor Red
            return $null
        }

    } catch {
        Write-Host "ERROR: JSON generation failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Gray
        return $null
    }
}


function Remove-OldJSONReports {
    param(
        [string]$ResultsDirectory,
        [string]$TestSuiteName,
        [int]$MaxFiles = 3
    )

    try {
        if (!(Test-Path $ResultsDirectory)) {
            return
        }

        # Find all JSON files for this specific test suite
        $pattern = "TestResults_$TestSuiteName`_*.json"
        $existingFiles = Get-ChildItem -Path $ResultsDirectory -Filter $pattern -ErrorAction SilentlyContinue | Sort-Object CreationTime -Descending

        if ($existingFiles.Count -gt $MaxFiles) {
            $filesToDelete = $existingFiles | Select-Object -Skip $MaxFiles

            Write-Host "Cleaning up old JSON reports for $TestSuiteName..." -ForegroundColor Yellow
            foreach ($file in $filesToDelete) {
                try {
                    Remove-Item -Path $file.FullName -Force
                    Write-Host "  Deleted: $($file.Name)" -ForegroundColor Gray
                } catch {
                    Write-Host "  Warning: Could not delete $($file.Name): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            Write-Host "Kept $MaxFiles most recent JSON reports for $TestSuiteName" -ForegroundColor Green
        }
    } catch {
        Write-Host "Warning: Error during JSON cleanup: $($_.Exception.Message)" -ForegroundColor Yellow
    }
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

# MAIN SCRIPT EXECUTION STARTS HERE
# Pre-Requisite:   
# Ensure the credentials are setup for your connector following the instructions here: https://learn.microsoft.com/en-us/power-query/power-query-sdk-vs-code#set-credential 

# User Build Path to Power BI Desktop custom connectors folder (update as needed)
$MezTarget = "$env:USERPROFILE\Documents\Power BI Desktop\Custom Connectors"

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

# Generate JSON output if requested

if ($JSONResults) {
    Write-Output("`nGenerating JSON results...")

    foreach ($TestSettings in $TestSettingsList) {
        $SuiteName = $TestSettings -replace "\.json$", ""
        Write-Output("Creating JSON results for: $SuiteName")

        # Filter test results for this specific test suite
        $suiteResults = @()

        # Try multiple filtering approaches
        $suiteResults = $TestResultsObjects | Where-Object { $_.Name -like "*$SuiteName*" }

        if (-not $suiteResults -or $suiteResults.Count -eq 0) {
            # Try alternative filtering
            $suiteResults = $TestResultsObjects | Where-Object { 
                ($_.Name -and $_.Name.Contains($SuiteName)) -or 
                ($_.TestName -and $_.TestName -eq $SuiteName)
            }
        }

        if (-not $suiteResults -or $suiteResults.Count -eq 0) {
            # Last resort: use all results
            Write-Host "Warning: No specific results found for $SuiteName, using all available results" -ForegroundColor Yellow
            $suiteResults = $TestResultsObjects
        }

        Write-Host "Found $($suiteResults.Count) test results for suite $SuiteName" -ForegroundColor Cyan

        $resultsDirectory = "./results"

        # Clean up old JSON reports before creating new one
        Remove-OldJSONReports -ResultsDirectory $resultsDirectory -TestSuiteName $SuiteName -MaxFiles 3

        # Generate JSON results with enhanced error handling
        $jsonPath = Write-JSONTestResults -TestSuiteName $SuiteName -TestResultsObjects $suiteResults -RawTestResults $RawTestResults -ExtensionPath $ExtensionPath -ResultsDirectory $resultsDirectory

        if ($jsonPath -and (Test-Path $jsonPath)) {
            Write-Output("")
            Write-Output("================================================================================")
            Write-Output("JSON RESULTS CREATED SUCCESSFULLY!")
            Write-Output("================================================================================")
            Write-Output("Location: $jsonPath")
            Write-Output("Contains: $($suiteResults.Count) test results")
            Write-Output("Metadata includes: connector version, timestamps, execution details")
            Write-Output("================================================================================")

            # Show file information
            try {
                $fileInfo = Get-Item $jsonPath
                $fileSizeKB = [Math]::Round($fileInfo.Length / 1024, 2)
                Write-Output("File size: $fileSizeKB KB")
                Write-Output("Encoding: UTF-8 (No BOM)")
            } catch {
                Write-Host "Could not get file information: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        } else {
            Write-Output("[ERROR] Failed to create JSON results")
        }
    }

    Write-Output("")
    Write-Output("JSON results generation completed!")
}


# Generate HTML preview if requested

if ($HTMLPreview) {
    Write-Output("`nGenerating HTML preview from LIVE results...")
    
    foreach ($TestSettings in $TestSettingsList) {
        $SuiteName = $TestSettings -replace "\.json$", ""
        Write-Output("Creating HTML report for: $SuiteName")
        
        # Filter test results for this specific test suite
        $suiteResults = $TestResultsObjects | Where-Object { $_.Name -like "*$SuiteName*" -or $_.Name -like "*Navigation*" }
        
        if (-not $suiteResults) {
            # If no specific filtering works, use all results
            $suiteResults = $TestResultsObjects
        }
        
        # Generate HTML report using LIVE results
        $resultsDirectory = ".\results"
        if (!(Test-Path $resultsDirectory)) {
            New-Item -ItemType Directory -Path $resultsDirectory -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $htmlFileName = "TestResults_$($SuiteName)_$timestamp.html"
        $htmlPath = Join-Path $resultsDirectory $htmlFileName
        
        # Clean up old reports
        Remove-OldHTMLReports -ResultsDirectory $resultsDirectory -TestSuiteName $SuiteName
        
        # Generate the HTML report with live data using the ORIGINAL structure
        $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Power Query Test Results - $SuiteName (LIVE RESULTS)</title>
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
        .live-indicator { background: #28a745; color: white; padding: 5px 10px; border-radius: 3px; font-size: 0.8em; margin-left: 10px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-content">
                <div class="header-left">
                    <h1>TEST RESULTS - Power Query<span class="live-indicator">LIVE DATA</span></h1>
                    <p>Test Suite: $SuiteName</p>
                    <p class="timestamp">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
                </div>
            </div>
        </div>
"@

        # Add test results from LIVE execution data using ORIGINAL parsing logic
        foreach ($testResult in $suiteResults) {
            $testName = $testResult.Name.Split("\")[-1] -replace '\.query\.pq$', ''
            
            $htmlContent += "`n        <div class='test-section'>"
            $htmlContent += "`n            <div class='test-title'>TEST: $testName</div>"
            
            # Check if we have SerializedSource data from the live results
            if ($testResult.Output -and $testResult.Output.SerializedSource) {
                $serializedData = $testResult.Output.SerializedSource
                
                # Use EXACT same parsing logic as Read-PQOutFile for consistency
                $tablePattern = '#table\s*\(\s*type\s+table\s+\[([^\]]+)\]\s*,\s*\{(.*)\}\s*\)'
                
                if ($serializedData -match $tablePattern) {
                    # Extract column definitions and data
                    $columnDefs = $Matches[1]
                    $dataRows = $Matches[2]
                    
                    Write-Host "DEBUG: Column definitions: $columnDefs" -ForegroundColor Gray
                    Write-Host "DEBUG: Data rows: $($dataRows.Substring(0, [Math]::Min(100, $dataRows.Length)))..." -ForegroundColor Gray
                    
                    # Parse column names with support for quoted names like #"Test Name"
                    $columnNames = @()
                    $columnMatches = [regex]::Matches($columnDefs, '(?:#"([^"]+)"|(\w+))\s*=\s*\w+')
                    foreach ($match in $columnMatches) {
                        if ($match.Groups[1].Value) {
                            # Quoted column name like #"Test Name"
                            $columnNames += $match.Groups[1].Value
                        } elseif ($match.Groups[2].Value) {
                            # Unquoted column name like Expected  
                            $columnNames += $match.Groups[2].Value
                        }
                    }
                    
                    Write-Host "DEBUG: Extracted columns: $($columnNames -join ', ')" -ForegroundColor Cyan
                    
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
                    
                    Write-Host "DEBUG: Extracted $($parsedRows.Count) data rows" -ForegroundColor Cyan
                    
                    # Generate table HTML with robust error handling
                    $htmlContent += "`n            <table>"
                    $htmlContent += "`n                <thead>"
                    $htmlContent += "`n                    <tr>"
                    
                    # Generate column headers
                    if ($columnNames.Count -gt 0) {
                        foreach ($col in $columnNames) {
                            $encodedCol = [System.Web.HttpUtility]::HtmlEncode($col)
                            $htmlContent += "`n                        <th>$encodedCol</th>"
                        }
                    } else {
                        # Fallback if column parsing failed
                        for ($i = 0; $i -lt 5; $i++) {
                            $htmlContent += "`n                        <th>Column$($i+1)</th>"
                        }
                    }
                    
                    $htmlContent += "`n                    </tr>"
                    $htmlContent += "`n                </thead>"
                    $htmlContent += "`n                <tbody>"
                    
                    # Generate data rows
                    foreach ($row in $parsedRows) {
                        $htmlContent += "`n                    <tr>"
                        
                        # Generate cells for each column
                        $maxCols = [Math]::Max($columnNames.Count, $row.Count)
                        for ($i = 0; $i -lt $maxCols; $i++) {
                            $cellClass = ""
                            $value = if ($i -lt $row.Count) { $row[$i] } else { "" }
                            
                            # Apply styling based on column name and value
                            if ($i -lt $columnNames.Count) {
                                $columnName = $columnNames[$i]
                                if ($columnName -match "Result|Status|Validation") {
                                    switch ($value.ToLower()) {
                                        "pass" { 
                                            $cellClass = "pass-cell"
                                            $value = "PASS"  # Normalize for dashboard compatibility
                                        }
                                        "fail" { 
                                            $cellClass = "fail-cell"
                                            $value = "FAIL"  # Normalize for dashboard compatibility
                                        }
                                        "info" { 
                                            $cellClass = "info-cell"
                                            $value = "INFO"  # Normalize for dashboard compatibility
                                        }
                                    }
                                }
                            }
                            
                            $encodedValue = [System.Web.HttpUtility]::HtmlEncode($value)
                            $htmlContent += "`n                        <td class='$cellClass'>$encodedValue</td>"
                        }
                        $htmlContent += "`n                    </tr>"
                    }
                    $htmlContent += "`n                </tbody>"
                    $htmlContent += "`n            </table>"
                    
                } else {
                    # Simple result or parsing failed
                    $statusClass = if ($testResult.Status -eq "Failed") { "simple-fail" } else { "simple-success" }
                    $htmlContent += "`n            <div class='simple-result $statusClass'>"
                    $htmlContent += "`n                <strong>Raw Data:</strong><br>"
                    $htmlContent += "`n                $([System.Web.HttpUtility]::HtmlEncode($serializedData))"
                    $htmlContent += "`n            </div>"
                }
            } else {
                # No SerializedSource data available
                $htmlContent += "`n            <div class='no-results'>"
                $htmlContent += "`n                No live results data available"
                $htmlContent += "`n            </div>"
            }
            
            $htmlContent += "`n        </div>"
        }
        
        $htmlContent += "`n    </div>`n</body>`n</html>"
        
        # Write HTML file
        $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
        
        if (Test-Path $htmlPath) {
            Write-Output("")
            Write-Output("================================================================================")
            Write-Output("HTML REPORT CREATED SUCCESSFULLY FROM LIVE RESULTS!")
            Write-Output("================================================================================")
            Write-Output("Location: $htmlPath")
            
            # Open in browser and VS Code
            if (-not $NoOpenHTML) {
                try {
                    Start-Process $htmlPath
                    Write-Output("[SUCCESS] Opened in default browser")
                } catch {
                    Write-Output("[WARNING] Could not open in browser: $($_.Exception.Message)")
                }
                
                try {
                    code $htmlPath
                    Write-Output("[SUCCESS] Opened HTML file in VS Code")
                } catch {
                    Write-Output("[INFO] VS Code not available for opening HTML file")
                }
            }
            
            # Show available reports
            $allReports = Get-ChildItem -Path $resultsDirectory -Filter "TestResults_$($SuiteName)_*.html" | Sort-Object CreationTime -Descending
            Write-Output("")
            Write-Output("Available reports for $SuiteName (newest first):")
            foreach ($report in $allReports) {
                $age = if ($report.Name -eq (Split-Path $htmlPath -Leaf)) { " [Just created from LIVE data]" } else { "" }
                Write-Output("  $($report.Name)$age")
            }
            
            Write-Output("")
            Write-Output("MANUAL PREVIEW OPTIONS IN VS CODE:")
            Write-Output("  1. Right-click the HTML file → 'Open Preview'")
            Write-Output("  2. With HTML file open: Ctrl+Shift+V")
            Write-Output("  3. Command Palette: 'Live Preview: Show Preview'")
            Write-Output("")
            Write-Output("Or simply use the browser that just opened!")
            Write-Output("================================================================================")
        } else {
            Write-Output("[ERROR] Failed to create HTML report at $htmlPath")
        }
    }
    
    Write-Output("")
    Write-Output("HTML preview generation completed!")
}


# Optional step: Copy *.mez file to PBI Connector Directory
if ($CopyMez) {
    Write-Host "Requested to copy MEZ file to Power BI Custom Connector folder..." -ForegroundColor Cyan

    if (-Not (Test-Path $ExtensionPath)) {
        Write-Host "ERROR: The source MEZ file was not found at $ExtensionPath" -ForegroundColor Red
        exit 1
    }
    if (-Not (Test-Path $MezTarget)) {
        Write-Host "Destination folder $MezTarget not found. Creating..." -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $MezTarget -Force | Out-Null
    }

    $destination = Join-Path $MezTarget (Split-Path $ExtensionPath -Leaf)
    Copy-Item -Path $ExtensionPath -Destination $destination -Force

    Write-Host "Successfully copied MEZ file to $destination" -ForegroundColor Green
} else {
    Write-Host "Skipping MEZ deploy/copy step. Use -CopyMez to enable." -ForegroundColor Yellow
}
