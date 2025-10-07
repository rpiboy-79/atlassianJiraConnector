<# 
PS 5.1 Compatible Script
This version removes PowerShell 7+ dependencies, fixes encoding handling for PS 5.1, replaces PSStyle coloring,
uses safe .Count and property access, and includes robust array handling. It also includes prior fixes for
System.Web.HttpUtility, safe parameter query settings update, and JSON/HTML generation.
#>

[CmdletBinding()]
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
    [switch]$CopyMez,
    [switch]$PrettyJson,
    [int]$MaxParallel = 1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# PowerShell version compatibility check and flag
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "This script requires PowerShell 5.0 or later. Current version: $($PSVersionTable.PSVersion)"
    exit 1
}
$IsPS5 = $PSVersionTable.PSVersion.Major -eq 5
if ($IsPS5) { Write-Host "Running in PowerShell 5.1 compatibility mode" -ForegroundColor Yellow }

# UTF-8 without BOM encoding for PS 5.1 using .NET API
$UTF8NoBOM = New-Object System.Text.UTF8Encoding($false)

# Helpers
function Write-Color {
    param([Parameter(Mandatory)][string]$Text,[ValidateSet('Green','Red','Yellow','Cyan','Magenta','Gray','White','Blue')][string]$Color='White')
    Write-Host $Text -ForegroundColor $Color
}
function ConvertTo-HtmlEncodedString { param([string]$InputString) if (-not $InputString) { return '' } return $InputString.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;').Replace("'",'&#39;') }
function As-Bool([object]$v) { return [bool]([string]$v -match '^(true|1)$') }
function Get-SafeArrayElement { param([array]$Array,[int]$Index,[string]$DefaultValue='') if (-not $Array -or $Array.Count -eq 0){return $DefaultValue} if($Index -lt 0){$actual=$Array.Count+$Index; if($actual -ge 0 -and $actual -lt $Array.Count){return $Array[$actual]}} else { if($Index -lt $Array.Count){return $Array[$Index]} } return $DefaultValue }
function Get-SafeCount { param([object]$Object) if ($null -eq $Object){return 0} if ($Object -is [array]){return $Object.Count} if ($Object.PSObject.Properties['Count']){return $Object.Count} return 1 }
function Get-SafeProperty { param([object]$Object,[string]$PropertyName,[object]$DefaultValue=$null) if ($null -eq $Object){return $DefaultValue} if ($Object.PSObject.Properties[$PropertyName]){ return $Object.$PropertyName } return $DefaultValue }
function Write-ColoredStatus { param([string]$Status,[string]$Text) switch ($Status){ 'Passed'{Write-Host $Text -ForegroundColor Green}; 'Failed'{Write-Host $Text -ForegroundColor Red}; 'Error'{Write-Host $Text -ForegroundColor Yellow}; default{Write-Host $Text} } }
function Get-SafeHash { param([string]$InputString) if (-not $InputString){return $null} try{ $bytes=[System.Text.Encoding]::UTF8.GetBytes($InputString); $ms=New-Object System.IO.MemoryStream($bytes); $hash=(Get-FileHash -InputStream $ms).Hash; $ms.Dispose(); return $hash } catch { return $null } }

# Paths
$cwd = (Get-Location).Path
$ResultsDirectoryDefault = Join-Path $cwd 'results'

# Discovery
function Confirm-Proceed { param([switch]$Skip) if ($Skip){return $true} if ($env:CI -or $env:BUILD_BUILDNUMBER){return $true} while($true){ $ans=Read-Host 'Do you want to proceed? [y/n]'; if($ans -eq 'y'){return $true} if($ans -eq 'n'){return $false} } }
function Get-PQTestPath {
    $extRoot = Join-Path $env:USERPROFILE '.vscode\extensions'
    $sdk = Get-ChildItem $extRoot -Directory -Filter 'powerquery.vscode-powerquery-sdk-*' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $sdk) { return $null }
    $nuget = Join-Path $sdk.FullName '.nuget'
    $tools = Get-ChildItem $nuget -Directory -Filter 'Microsoft.PowerQuery.SdkTools.*' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $tools) { return $null }
    $path = Join-Path $tools.FullName 'tools\PQTest.exe'
    if (Test-Path $path) { return $path }
    return $null
}
function Get-ConnectorVersion { param([string]$ExtensionPath)
    $connectorName='Unknown'; $connectorVersion='1.0.0'
    try{ if ($ExtensionPath){ $connectorName=(Get-Item $ExtensionPath).BaseName } }catch{}
    $candidateDirs=@(); try{ if($ExtensionPath){ $candidateDirs+= (Split-Path -Path $ExtensionPath -Parent) } }catch{}
    $candidateDirs+= (Join-Path $cwd 'src'); $candidateDirs+= $cwd
    $pqFile=$null
    foreach($d in $candidateDirs){ if(-not (Test-Path $d)){continue}; $pqFile = Get-ChildItem -Path $d -Filter '*.pq' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if($pqFile){break} }
    if($pqFile){ try{ $lines = Get-Content -Path $pqFile.FullName -TotalCount 100 -Encoding UTF8; $pat='^\s*\[\s*Version\s*=\s*\"([0-9]+\.[0-9]+\.[0-9]+)\"\s*\]'; foreach($ln in $lines){ $m=[regex]::Match($ln,$pat); if($m.Success){ $connectorVersion=$m.Groups[1].Value; break } } }catch{} }
    return @{ Name=$connectorName; Version=$connectorVersion }
}

# Read .pqout cached
$__pqoutCache = @{}
function Read-PQOutCached([string]$path){ if(-not $path){return $null} if($__pqoutCache.ContainsKey($path)){return $__pqoutCache[$path]} if(Test-Path $path){ $c=Get-Content $path -Raw -Encoding UTF8; $trim=$c.Trim(); $__pqoutCache[$path]=$trim; return $trim } return $null }

function Get-ExpectedResults { param([object]$TestResult)
    try{ if($TestResult.Output -and $TestResult.Output.OutputFilePath){ $pqoutPath=$TestResult.Output.OutputFilePath; if(Test-Path $pqoutPath){ Write-Host "Loaded expected results from: $pqoutPath" -ForegroundColor Green; return Read-PQOutCached $pqoutPath } else { Write-Host "Expected results file not found: $pqoutPath" -ForegroundColor Yellow } } return $null } catch { Write-Host "Error reading expected results: $($_.Exception.Message)" -ForegroundColor Red; return $null }
}
function Get-ActualResults { param([object]$TestResult)
    try{ if($TestResult.Status -eq 'Failed'){ if($TestResult.Output -and $TestResult.Output.SerializedSource -and $TestResult.Output.SerializedSource -ne ''){ Write-Host "Using SerializedSource for failed test: $($TestResult.Name)" -ForegroundColor Yellow; return $TestResult.Output.SerializedSource } }
         if($TestResult.Output -and $TestResult.Output.OutputFilePath){ $pqoutPath=$TestResult.Output.OutputFilePath; if(Test-Path $pqoutPath){ Write-Host "Using .pqout content for actual results: $($TestResult.Name)" -ForegroundColor Blue; return (Read-PQOutCached $pqoutPath) } }
         return $null } catch { Write-Host "Error reading actual results: $($_.Exception.Message)" -ForegroundColor Red; return $null }
}
function Test-DataDrift { param([string]$ExpectedResults,[string]$ActualResults)
    try{ if($ExpectedResults -and $ActualResults){ $isDifferent = $ExpectedResults -ne $ActualResults; $eh=Get-SafeHash $ExpectedResults; $ah=Get-SafeHash $ActualResults; return @{HasDrift=$isDifferent; ExpectedHash=$eh; ActualHash=$ah} }
         return @{HasDrift=$false; ExpectedHash=$null; ActualHash=$null} } catch { Write-Host "Error detecting data drift: $($_.Exception.Message)" -ForegroundColor Yellow; return @{HasDrift=$false; ExpectedHash=$null; ActualHash=$null} }
}
function Get-BaselineTimestamp { param([object]$TestResult) try{ if($TestResult.Output -and $TestResult.Output.OutputFilePath){ $pqoutPath=$TestResult.Output.OutputFilePath; if(Test-Path $pqoutPath){ $fi=Get-Item $pqoutPath; return $fi.CreationTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ') } } return $null } catch { return $null } }

function Remove-OldJSONReports { param([string]$ResultsDirectory,[string]$TestSuiteName,[int]$MaxFiles=3)
    try{ if(!(Test-Path $ResultsDirectory)){return}; $pattern = "TestResults_$TestSuiteName`_*.json"; Get-ChildItem -Path $ResultsDirectory -Filter $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending | Select-Object -Skip $MaxFiles | ForEach-Object { try{ Remove-Item -Path $_.FullName -Force; Write-Host "  Deleted: $($_.Name)" -ForegroundColor Gray } catch { Write-Host "  Warning: Could not delete $($_.Name): $($_.Exception.Message)" -ForegroundColor Yellow } }; Write-Host "Kept $MaxFiles most recent JSON reports for $TestSuiteName" -ForegroundColor Green } catch { Write-Host "Warning: Error during JSON cleanup: $($_.Exception.Message)" -ForegroundColor Yellow }
}
function Remove-OldHTMLReports { param([string]$ResultsDirectory,[string]$TestSuiteName,[int]$MaxFiles=3)
    try{ if(!(Test-Path $ResultsDirectory)){return}; $pattern = "TestResults_$TestSuiteName`_*.html"; Get-ChildItem -Path $ResultsDirectory -Filter $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending | Select-Object -Skip $MaxFiles | ForEach-Object { try{ Remove-Item -Path $_.FullName -Force; Write-Host "  Deleted: $($_.Name)" -ForegroundColor Gray } catch { Write-Host "  Warning: Could not delete $($_.Name)" -ForegroundColor Yellow } }; Write-Host "Kept $MaxFiles most recent reports for $TestSuiteName" -ForegroundColor Green } catch { Write-Host "Warning: Error during HTML cleanup: $($_.Exception.Message)" -ForegroundColor Yellow }
}

function Write-JSONTestResults { param([string]$TestSuiteName,[array]$TestResultsObjects,[array]$RawTestResults,[string]$ExtensionPath,[string]$ResultsDirectory=$ResultsDirectoryDefault,[switch]$Pretty)
    try{
        if(!(Test-Path $ResultsDirectory)){ New-Item -ItemType Directory -Path $ResultsDirectory -Force | Out-Null }
        $connectorInfo = Get-ConnectorVersion -ExtensionPath $ExtensionPath
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $runDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $executionTimestamp = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        $jsonFileName = "TestResults_$TestSuiteName" + "_$timestamp.json"
        $jsonPath = Join-Path -Path $ResultsDirectory -ChildPath $jsonFileName

        $testResultsCount = Get-SafeCount $TestResultsObjects
        Write-Host "Processing $testResultsCount test result objects for Enhanced JSON export..." -ForegroundColor Cyan

        $processedResults = @()
        $totalDurationSum = 0
        $resultsToProcess = if ($null -eq $TestResultsObjects){ @() } elseif ($TestResultsObjects -is [array]){ $TestResultsObjects } else { @($TestResultsObjects) }

        foreach($testResult in $resultsToProcess){ try{
            $testName = if ($testResult.Name){ $testResult.Name.Split('\')[-1] -replace '\.query\.pq$','' } else { 'UnknownTest' }
            $startTime = Get-Date; $endTime = Get-Date; $duration=0
            if ($testResult.StartTime -and $testResult.EndTime) { try{ $startTime=[DateTime]$testResult.StartTime; $endTime=[DateTime]$testResult.EndTime; $duration=($endTime - $startTime).TotalSeconds } catch { $duration=0 } }
            $totalDurationSum += $duration
            $expectedResults = Get-ExpectedResults -TestResult $testResult
            $actualResults   = Get-ActualResults   -TestResult $testResult
            $baselineTimestamp = Get-BaselineTimestamp -TestResult $testResult
            $driftInfo = Test-DataDrift -ExpectedResults $expectedResults -ActualResults $actualResults
            $result = if ($testResult.Status){ $testResult.Status } else { 'Unknown' }
            $outputStatus = if ($testResult.Output -and $testResult.Output.Status){ $testResult.Output.Status } else { 'Unknown' }
            $processedResult = @{
                test_name           = $testName
                test_file           = if ($testResult.Name){ $testResult.Name } else { $testName }
                result              = $result
                output_status       = $outputStatus
                start_time          = $startTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
                end_time            = $endTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
                duration_seconds    = [Math]::Round($duration,2)
                details             = if ($testResult.Details){ $testResult.Details } else { 'No details available' }
                method              = if ($testResult.Method){ $testResult.Method } else { 'Unknown' }
                serialized_source   = $actualResults
                source_error        = if ($testResult.Output -and $testResult.Output.SourceError){ $testResult.Output.SourceError } else { $null }
                output_error        = if ($testResult.Output -and $testResult.Output.OutputError){ $testResult.Output.OutputError } else { $null }
                expected_result     = $expectedResults
                actual_result       = $actualResults
                execution_timestamp = $executionTimestamp
                baseline_timestamp  = $baselineTimestamp
                data_drift_detected = $driftInfo.HasDrift
                expected_hash       = $driftInfo.ExpectedHash
                actual_hash         = $driftInfo.ActualHash
                test_evolution      = @{
                    baseline_established = $baselineTimestamp
                    current_execution    = $executionTimestamp
                    comparison_method    = if ($testResult.Status -eq 'Failed' -and ($testResult.Output -and $testResult.Output.SerializedSource)) { 'TDD_Failed_Comparison' } else { 'Baseline_Comparison' }
                    has_historical_data  = ($null -ne $baselineTimestamp)
                }
                test_metadata       = @{
                    test_type          = 'PQTest_Compare'
                    data_source_state  = 'Live'
                    generation_method  = 'Enhanced_Phase1'
                    schema_version     = '1.1'
                }
            }
            $processedResults += $processedResult
        } catch { Write-Host "Warning: Error processing test result: $($_.Exception.Message)" -ForegroundColor Yellow; $processedResults += @{ test_name='ProcessingError'; test_file='Unknown'; result='Error'; output_status='ProcessingFailed'; start_time=(Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffZ'); end_time=(Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffZ'); duration_seconds=0; details=$_.Exception.Message; method='Unknown'; serialized_source=$null; expected_result=$null; actual_result=$null; execution_timestamp=$executionTimestamp; data_drift_detected=$false; source_error=$null; output_error=$_.Exception.Message } } }

        $processedCount = Get-SafeCount $processedResults
        Write-Host "Processed $processedCount test results successfully with Phase 1 enhancements" -ForegroundColor Green

        $totalTests     = $processedCount
        $passedTests    = Get-SafeCount (@($processedResults | Where-Object { $_.result -eq 'Passed' }))
        $failedTests    = Get-SafeCount (@($processedResults | Where-Object { $_.result -eq 'Failed' }))
        $errorTests     = Get-SafeCount (@($processedResults | Where-Object { $_.result -eq 'Error'  }))
        $testsWithDrift = Get-SafeCount (@($processedResults | Where-Object { $_.data_drift_detected -eq $true }))
        $totalDuration  = [Math]::Round($totalDurationSum, 2)

        $jsonOutput = @{
            metadata = @{
                run_timestamp                = $executionTimestamp
                run_date                     = $runDate
                test_suite_name              = $TestSuiteName
                connector_name               = $connectorInfo.Name
                connector_version            = $connectorInfo.Version
                extension_path               = $ExtensionPath
                total_tests                  = $totalTests
                passed_tests                 = $passedTests
                failed_tests                 = $failedTests
                error_tests                  = $errorTests
                total_duration_seconds       = $totalDuration
                pqtest_version               = 'Unknown'
                generation_method            = 'PowerShell_Enhanced_Phase1_PS51'
                schema_version               = '1.1'
                execution_timestamp          = $executionTimestamp
                baseline_comparison_enabled  = $true
                data_drift_detection_enabled = $true
                tests_with_drift             = $testsWithDrift
                historical_tracking_enabled  = $true
                execution_context            = @{
                    working_directory  = $cwd
                    powershell_version = $PSVersionTable.PSVersion.ToString()
                    generation_date    = $runDate
                    enhancement_level  = 'Phase1_TDD_Observability'
                }
            }
            test_results      = $processedResults
            raw_pqtest_output = if ($RawTestResults) { $RawTestResults } else { @() }
            summary_analytics = @{
                drift_analysis = @{
                    total_tests_checked = $totalTests
                    tests_with_drift    = $testsWithDrift
                    drift_percentage    = if ($totalTests -gt 0) { [Math]::Round(($testsWithDrift / $totalTests) * 100, 2) } else { 0 }
                }
                execution_analysis = @{
                    total_execution_time = $totalDuration
                    average_test_time    = if ($totalTests -gt 0) { [Math]::Round($totalDuration / $totalTests, 2) } else { 0 }
                    fastest_test         = ( @($processedResults | Sort-Object duration_seconds | Select-Object -First 1) )[0].test_name
                    slowest_test         = ( @($processedResults | Sort-Object duration_seconds -Descending | Select-Object -First 1) )[0].test_name
                }
            }
        }

        try{
            $jsonString = if ($Pretty){ $jsonOutput | ConvertTo-Json -Depth 10 } else { $jsonOutput | ConvertTo-Json -Depth 10 -Compress }
            [System.IO.File]::WriteAllText($jsonPath, $jsonString, $UTF8NoBOM)
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

# Summary class (unchanged shape)
class TestResult { [string]$ParameterQuery; [string]$TestFolder; [string]$TestName; [string]$OutputStatus; [string]$TestStatus; [string]$Duration; TestResult([string]$testFolder,[string]$testName,[string]$outputStatus,[string]$testStatus,[string]$duration){ $this.TestFolder=$testFolder; $this.TestName=$testName; $this.OutputStatus=$outputStatus; $this.TestStatus=$testStatus; $this.Duration=$duration } }

try{
    $MezTarget = Join-Path $env:USERPROFILE 'Documents\Power BI Desktop\Custom Connectors'

    # Load settings file
    $SettingsFilePath = Join-Path $cwd 'RunPQSDKTestSuitesSettings.json'
    $RunPQSDKTestSuitesSettings = $null
    if (Test-Path $SettingsFilePath) { $RunPQSDKTestSuitesSettings = Get-Content -Path $SettingsFilePath -Encoding UTF8 | ConvertFrom-Json }

    # Mock data config generation
    if ($TestSettingsList -and ($TestSettingsList -contains 'MockDataTests.json')) {
        $MockDataFilePath = Join-Path $cwd 'MockData\jira_api_response_acme_software_pb.json'
        if (Test-Path $MockDataFilePath) {
            $AbsoluteMockDataPath = (Resolve-Path $MockDataFilePath).Path
            $MockDataConfig = @{ MockDataAbsolutePath=$AbsoluteMockDataPath; GeneratedTimestamp=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); WorkingDirectory=$cwd }
            $TestSuiteMockFolder = Join-Path $cwd 'TestSuites\MockData'
            if (!(Test-Path $TestSuiteMockFolder)) { New-Item -ItemType Directory -Path $TestSuiteMockFolder -Force | Out-Null }
            $ConfigPath = Join-Path $TestSuiteMockFolder 'MockDataConfig.json'
            $configJson = $MockDataConfig | ConvertTo-Json -Depth 5
            [System.IO.File]::WriteAllText($ConfigPath, $configJson, $UTF8NoBOM)
            Write-Color "Written config to test folder: $ConfigPath" Green
            Write-Color "Generated mock data configuration file with absolute path:" Green
            Write-Host "  $AbsoluteMockDataPath" -ForegroundColor Cyan
        } else {
            Write-Color "ERROR: Mock data file not found at: $MockDataFilePath" Red
            Write-Color "Please ensure the mock data file exists before running mock data tests." Red
        }
    }

    # Resolve PQTestExePath
    if (-not $PQTestExePath) {
        $discoveredPath = Get-PQTestPath
        if ($discoveredPath) { $PQTestExePath = $discoveredPath; Write-Host ("Auto-discovered PQTest.exe at: " + $PQTestExePath) }
        elseif ($RunPQSDKTestSuitesSettings -and $RunPQSDKTestSuitesSettings.PQTestExePath) { $PQTestExePath = $RunPQSDKTestSuitesSettings.PQTestExePath; Write-Host ("Using PQTest.exe path from settings: " + $PQTestExePath) }
    }
    if (-not (Test-Path -Path $PQTestExePath)) {
        Write-Color "PQTestExe path could not be found or is not correctly set." Red
        Write-Host ("Attempted path: " + $PQTestExePath)
        Write-Host "1. Ensure Power Query SDK is installed in VS Code"
        Write-Host "2. Set explicit path in RunPQSDKTestSuitesSettings.json"
        Write-Host "3. Pass path as -PQTestExePath argument"
        Write-Host ("Expected VS Code extension location: " + (Join-Path $env:USERPROFILE '.vscode\extensions\powerquery.vscode-powerquery-sdk-*'))
        exit 1
    }

    # Resolve ExtensionPath
    if (-not $ExtensionPath -and $RunPQSDKTestSuitesSettings) { $ExtensionPath = $RunPQSDKTestSuitesSettings.ExtensionPath }
    if (-not (Test-Path -Path $ExtensionPath)) { Write-Host ("Extension path is not correctly set. Either set it in RunPQSDKTestSuitesSettings.json or pass it as an argument. " + $ExtensionPath); exit 1 }

    # Resolve TestSettingsDirectoryPath and setup defaults
    if (-not $TestSettingsDirectoryPath) {
        if ($RunPQSDKTestSuitesSettings -and $RunPQSDKTestSuitesSettings.TestSettingsDirectoryPath) {
            $TestSettingsDirectoryPath = $RunPQSDKTestSuitesSettings.TestSettingsDirectoryPath
        } else {
            $GenericTestSettingsDirectoryPath  = Join-Path $cwd 'ConnectorConfigs\generic\Settings'
            $connectorBaseName = (Get-Item $ExtensionPath).BaseName
            $TestSettingsDirectoryPath = Join-Path $cwd ("ConnectorConfigs\{0}\Settings" -f $connectorBaseName)
            if (-not (Test-Path $TestSettingsDirectoryPath)) { Write-Color "Performing the initial setup by creating the test settings and parameter query file(s) automatically..." Blue; Copy-Item -Path $GenericTestSettingsDirectoryPath -Destination $TestSettingsDirectoryPath -Recurse -Force; Write-Host ("Successfully created test settings file(s) under: " + $TestSettingsDirectoryPath) }
            $GenericParameterQueriesPath   = Join-Path $cwd 'ConnectorConfigs\generic\ParameterQueries'
            $ExtensionParameterQueriesPath = Join-Path $cwd ("ConnectorConfigs\{0}\ParameterQueries" -f $connectorBaseName)
            if (-not (Test-Path $ExtensionParameterQueriesPath)) { Copy-Item -Path $GenericParameterQueriesPath -Destination $ExtensionParameterQueriesPath -Recurse -Force; $genericParam = Join-Path $ExtensionParameterQueriesPath 'Generic.parameterquery.pq'; if (Test-Path $genericParam) { Rename-Item -Path $genericParam -NewName ("{0}.parameterquery.pq" -f $connectorBaseName) }; Write-Host ("Successfully created the parameter query file(s) under: " + $ExtensionParameterQueriesPath) }
            foreach ($SettingsFile in (Get-ChildItem $TestSettingsDirectoryPath -File | ForEach-Object { $_.FullName })) { try{ $SettingsFileJson = Get-Content -Path $SettingsFile -Encoding UTF8 | ConvertFrom-Json; if ($SettingsFileJson.ParameterQueryFilePath){ $SettingsFileJson.ParameterQueryFilePath = $SettingsFileJson.ParameterQueryFilePath.ToLower().Replace('generic',$connectorBaseName); $settingsJson = $SettingsFileJson | ConvertTo-Json -Depth 100; [System.IO.File]::WriteAllText($SettingsFile, $settingsJson, $UTF8NoBOM) } } catch { $errorMessage = $_.Exception.Message; Write-Host "Warning: Could not update ParameterQueryFilePath in $SettingsFile`: $errorMessage" -ForegroundColor Yellow } }
            if (-not (Confirm-Proceed -Skip:$SkipConfirmation)) { Write-Color "Aborted by user." Yellow; exit 1 }
        }
    }
    if (-not (Test-Path -Path $TestSettingsDirectoryPath)) { Write-Host ("Test Settings Directory is not correctly set. Either set it in RunPQSDKTestSuitesSettings.json or pass it as an argument. " + $TestSettingsDirectoryPath); exit 1 }

    # Resolve TestSettingsList
    if (-not $TestSettingsList) { if ($RunPQSDKTestSuitesSettings -and $RunPQSDKTestSuitesSettings.TestSettingsList) { $TestSettingsList = $RunPQSDKTestSuitesSettings.TestSettingsList } else { $TestSettingsList = (Get-ChildItem -Path $TestSettingsDirectoryPath -Filter '*.json' -Name) } }

    # Resolve booleans
    if (-not $ValidateQueryFolding) { if ($RunPQSDKTestSuitesSettings) { $ValidateQueryFolding = As-Bool $RunPQSDKTestSuitesSettings.ValidateQueryFolding } }
    if (-not $DetailedResults)      { if ($RunPQSDKTestSuitesSettings) { $DetailedResults      = As-Bool $RunPQSDKTestSuitesSettings.DetailedResults } }
    if (-not $JSONResults)          { if ($RunPQSDKTestSuitesSettings) { $JSONResults          = As-Bool $RunPQSDKTestSuitesSettings.JSONResults } }

    $connectorBaseName = (Get-Item $ExtensionPath).BaseName
    $ExtensionParameterQueriesPath = Join-Path $cwd ("ConnectorConfigs\{0}\ParameterQueries" -f $connectorBaseName)
    $DiagnosticFolderPath = Join-Path $cwd ("Diagnostics\{0}" -f $connectorBaseName)

    Write-Color "Below are settings for running the TestSuites:" Blue
    Write-Host ("PQTestExePath: " + $PQTestExePath)
    Write-Host ("ExtensionPath: " + $ExtensionPath)
    Write-Host ("TestSettingsDirectoryPath: " + $TestSettingsDirectoryPath)
    Write-Host ("TestSettingsList: " + ($TestSettingsList -join ', '))
    Write-Host ("ValidateQueryFolding: " + [bool]$ValidateQueryFolding)
    Write-Host ("DetailedResults: " + [bool]$DetailedResults)
    Write-Host ("JSONResults: " + [bool]$JSONResults)

    Write-Host ("Note: Please verify the settings above and ensure the following:") -ForegroundColor Magenta
    Write-Host ("1. Credentials are setup for the extension following the instructions here: https://learn.microsoft.com/power-query/power-query-sdk-vs-code#set-credential")
    Write-Host ("2. Parameter query file(s) are updated under: ")
    Write-Host ("   " + $ExtensionParameterQueriesPath)
    if ($ValidateQueryFolding) { Write-Host ("3. Diagnostics folder path for query folding verification:"); Write-Host ("   " + $DiagnosticFolderPath) }

    if (-not (Confirm-Proceed -Skip:$SkipConfirmation)) { Write-Color "Please specify the correct settings in RunPQSDKTestSuitesSettings.json or pass them as arguments and re-run the script." Yellow; exit 1 }

    if ($ValidateQueryFolding -and -not (Test-Path $DiagnosticFolderPath)) { New-Item -ItemType Directory -Force -Path $DiagnosticFolderPath | Out-Null }

    $TestCount=0; $Passed=0; $Failed=0; $TestExecStartTime=Get-Date
    $TestResults=@(); $RawTestResults=@(); $TestResultsObjects=@()

    function Invoke-CompareForSettings { param([string]$SettingsFileName)
        $args=@('compare','-p','-e',$ExtensionPath,'-sf',(Join-Path $TestSettingsDirectoryPath $SettingsFileName))
        if ($ValidateQueryFolding) { $args+=@('-dfp',$DiagnosticFolderPath) }
        & $PQTestExePath @args
    }

    foreach ($TestSettings in $TestSettingsList) {
        $RawTestResult = Invoke-CompareForSettings -SettingsFileName $TestSettings
        if ($RawTestResult) {
            $StringTestResult = $RawTestResult -join ' '
            $TestResultsObject = $StringTestResult | ConvertFrom-Json
            $resultsToProcess = if ($null -eq $TestResultsObject){ @() } elseif ($TestResultsObject -is [array]){ $TestResultsObject } else { @($TestResultsObject) }
            foreach($Result in $resultsToProcess){
                $parts = if ($Result.Name){ $Result.Name -split '\\' } else { @('', '', 'UnknownTest') }
                $partsCount = Get-SafeCount $parts
                $folder = if ($partsCount -ge 3) { (Get-SafeArrayElement $parts -3 'Unknown') + '\' + (Get-SafeArrayElement $parts -2 'Unknown') } else { 'Unknown\Unknown' }
                $test = Get-SafeArrayElement $parts -1 'UnknownTest'
                $outputStatus = if ($Result.Output -and $Result.Output.Status){ $Result.Output.Status } else { 'Unknown' }
                $testStatus   = if ($Result.Status){ $Result.Status } else { 'Unknown' }
                $duration     = if ($Result.StartTime -and $Result.EndTime){ try{ (New-TimeSpan -Start $Result.StartTime -End $Result.EndTime).ToString() } catch { '00:00:00' } } else { '00:00:00' }
                $TestResults += [TestResult]::new($folder,$test,$outputStatus,$testStatus,$duration)
                $TestCount++
                if ($Result.Status -eq 'Passed') { $Passed++ } else { $Failed++ }
            }
            $RawTestResults += $RawTestResult
            $TestResultsObjects += $resultsToProcess
        }
    }

    $TestExecEndTime = Get-Date

    if ($DetailedResults) {
        Write-Host ("------------------------------------------------------------------------------------------")
        Write-Host ("PQ SDK Test Framework - Test Execution - Detailed Results for Extension: " + $ExtensionPath.Split('\')[-1])
        Write-Host ("------------------------------------------------------------------------------------------")
        $TestResultsObjects
    }
    if ($JSONResults) {
        Write-Host ("-----------------------------------------------------------------------------------")
        Write-Host ("PQ SDK Test Framework - Test Execution - JSON Results for Extension: " + $ExtensionPath.Split('\')[-1])
        Write-Host ("-----------------------------------------------------------------------------------")
        $RawTestResults
    }

    Write-Host ("----------------------------------------------------------------------------------------------")
    Write-Host ("PQ SDK Test Framework - Test Execution - Test Results Summary for Extension: " + $ExtensionPath.Split('\')[-1])
    Write-Host ("----------------------------------------------------------------------------------------------")
    # Simple PS 5.1 compatible table (no PSStyle colors)
    $TestResults | Format-Table -AutoSize -Property TestFolder, TestName, OutputStatus, TestStatus, Duration

    Write-Host ("----------------------------------------------------------------------------------------------")
    Write-Host ("Total Tests: $TestCount | Passed: $Passed | Failed: $Failed | Total Duration: " + ("{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s" -f (New-TimeSpan -Start $TestExecStartTime -End $TestExecEndTime)))
    Write-Host ("----------------------------------------------------------------------------------------------")

    $failedTests = @($TestResultsObjects | Where-Object { $_.Output.Status -ne 'Passed' })
    if (Get-SafeCount $failedTests) {
        Write-Host "" -ForegroundColor Yellow
        Write-Host "WARNING: One or more tests output failed. Generating HTML anyway for dashboard review." -ForegroundColor Yellow
        foreach ($t in $failedTests) { Write-Host ("  Test: {0}, OutputStatus: {1}" -f $t.Name, $t.Output.Status) -ForegroundColor Yellow }
    }

    if ($JSONResults) {
        Write-Host "`nGenerating JSON results..."
        foreach ($TestSettings in $TestSettingsList) {
            $SuiteName = $TestSettings -replace '\.json$', ''
            Write-Host ("Creating JSON results for: $SuiteName")
            $suiteResults = @($TestResultsObjects | Where-Object { (Get-SafeProperty $_ 'Name' '') -like "*$SuiteName*" })
            if (-not (Get-SafeCount $suiteResults)) {
                $suiteResults = @($TestResultsObjects | Where-Object { $n=Get-SafeProperty $_ 'Name' ''; $m=Get-SafeProperty $_ 'Method' ''; ($n -and $n.Contains($SuiteName)) -or ($m -and $m.Contains($SuiteName)) })
            }
            if (-not (Get-SafeCount $suiteResults)) { Write-Color "Warning: No specific results found for $SuiteName, using all available results" Yellow; $suiteResults = @($TestResultsObjects) }
            Write-Host ("Found {0} test results for suite {1}" -f (Get-SafeCount $suiteResults), $SuiteName) -ForegroundColor Cyan
            $resultsDirectory = $ResultsDirectoryDefault
            Remove-OldJSONReports -ResultsDirectory $resultsDirectory -TestSuiteName $SuiteName -MaxFiles 3
            $jsonPath = Write-JSONTestResults -TestSuiteName $SuiteName -TestResultsObjects $suiteResults -RawTestResults $RawTestResults -ExtensionPath $ExtensionPath -ResultsDirectory $resultsDirectory -Pretty:$PrettyJson
            if ($jsonPath -and (Test-Path $jsonPath)) {
                Write-Host ""
                Write-Host "================================================================================"
                Write-Host "JSON RESULTS CREATED SUCCESSFULLY!"
                Write-Host "================================================================================"
                Write-Host ("Location: {0}" -f $jsonPath)
                Write-Host ("Contains: {0} test results" -f (Get-SafeCount $suiteResults))
                Write-Host "Metadata includes: connector version, timestamps, execution details"
                Write-Host "================================================================================"
                try{ $fi=Get-Item $jsonPath; $kb=[Math]::Round($fi.Length/1024,2); Write-Host ("File size: {0} KB" -f $kb); Write-Host "Encoding: UTF-8 (No BOM)" } catch { Write-Host ("Could not get file information: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
            } else {
                Write-Host "[ERROR] Failed to create JSON results"
            }
        }
        Write-Host ""; Write-Host "JSON results generation completed!"
    }

    if ($HTMLPreview) {
        Write-Host "`nGenerating HTML preview from LIVE results..."
        foreach ($TestSettings in $TestSettingsList) {
            $SuiteName = $TestSettings -replace '\.json$', ''
            Write-Host ("Creating HTML report for: {0}" -f $SuiteName)
            $suiteResults = $TestResultsObjects | Where-Object { (Get-SafeProperty $_ 'Name' '') -like "*$SuiteName*" -or (Get-SafeProperty $_ 'Name' '') -like '*Navigation*' }
            if (-not $suiteResults) { $suiteResults = $TestResultsObjects }
            $resultsDirectory = $ResultsDirectoryDefault; if (!(Test-Path $resultsDirectory)) { New-Item -ItemType Directory -Path $resultsDirectory -Force | Out-Null }
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $htmlFileName = "TestResults_{0}_{1}.html" -f $SuiteName, $timestamp
            $htmlPath = Join-Path $resultsDirectory $htmlFileName
            Remove-OldHTMLReports -ResultsDirectory $resultsDirectory -TestSuiteName $SuiteName
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('<!DOCTYPE html>'); [void]$sb.AppendLine('<html>'); [void]$sb.AppendLine('<head>'); [void]$sb.AppendLine("    <title>Power Query Test Results - $SuiteName (LIVE RESULTS)</title>"); [void]$sb.AppendLine('    <meta charset="UTF-8">'); [void]$sb.AppendLine('    <style>');
            [void]$sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background: #f5f5f5; }"); [void]$sb.AppendLine('        .container { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }'); [void]$sb.AppendLine('        .header { background: #0078d4; color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; }'); [void]$sb.AppendLine('        .header-content { display: flex; align-items: flex-start; }'); [void]$sb.AppendLine('        .header-left { flex: 1; }'); [void]$sb.AppendLine('        .header h1 { margin: 0 0 10px 0; font-size: 1.8em; }'); [void]$sb.AppendLine('        .header p { margin: 5px 0; font-size: 1.1em; }'); [void]$sb.AppendLine('        .timestamp { color: white !important; font-size: 0.95em !important; font-weight: normal; opacity: 0.9; }'); [void]$sb.AppendLine('        .summary { display: flex; gap: 20px; margin-bottom: 20px; }'); [void]$sb.AppendLine('        .stat-card { background: #f8f9fa; padding: 15px; border-radius: 5px; text-align: center; flex: 1; }'); [void]$sb.AppendLine('        .pass { color: #107c10; font-weight: bold; }'); [void]$sb.AppendLine('        .fail { color: #d13438; font-weight: bold; }'); [void]$sb.AppendLine('        .info { color: #0078d4; font-weight: bold; }'); [void]$sb.AppendLine('        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; background: white; }'); [void]$sb.AppendLine('        th { background: #f8f9fa; border: 1px solid #dee2e6; padding: 12px; text-align: left; font-weight: bold; }'); [void]$sb.AppendLine('        td { border: 1px solid #dee2e6; padding: 12px; }'); [void]$sb.AppendLine('        .test-section { margin-bottom: 30px; border: 1px solid #dee2e6; border-radius: 5px; padding: 15px; }'); [void]$sb.AppendLine('        .test-title { background: #e3f2fd; padding: 10px; border-radius: 5px; font-weight: bold; margin-bottom: 15px; font-size: 1.1em; }'); [void]$sb.AppendLine('        .pass-cell { background: #d4edda; color: #155724; font-weight: bold; }'); [void]$sb.AppendLine('        .fail-cell { background: #f8d7da; color: #721c24; font-weight: bold; }'); [void]$sb.AppendLine('        .info-cell { background: #d1ecf1; color: #0c5460; font-weight: bold; }'); [void]$sb.AppendLine('        .simple-result { padding: 15px; border-radius: 5px; margin: 10px 0; font-size: 1.2em; font-weight: bold; text-align: center; }'); [void]$sb.AppendLine('        .simple-success { background: #d4edda; color: #155724; }'); [void]$sb.AppendLine('        .simple-fail { background: #f8d7da; color: #721c24; }'); [void]$sb.AppendLine('        .no-results { background: #fff3cd; color: #856404; padding: 15px; border-radius: 5px; text-align: center; }'); [void]$sb.AppendLine('        .live-indicator { background: #28a745; color: white; padding: 5px 10px; border-radius: 3px; font-size: 0.8em; margin-left: 10px; }'); [void]$sb.AppendLine('    </style>'); [void]$sb.AppendLine('</head>'); [void]$sb.AppendLine('<body>'); [void]$sb.AppendLine('    <div class="container">'); [void]$sb.AppendLine('        <div class="header">'); [void]$sb.AppendLine('            <div class="header-content">'); [void]$sb.AppendLine('                <div class="header-left">'); [void]$sb.AppendLine("                    <h1>TEST RESULTS - Power Query<span class=""live-indicator"">LIVE DATA</span></h1>"); [void]$sb.AppendLine("                    <p>Test Suite: $SuiteName</p>"); [void]$sb.AppendLine("                    <p class=""timestamp"">Generated: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))</p>"); [void]$sb.AppendLine('                </div>'); [void]$sb.AppendLine('            </div>'); [void]$sb.AppendLine('        </div>');
            foreach($testResult in $suiteResults){ $testName = if ($testResult.Name){ $testResult.Name.Split('\')[-1] -replace '\.query\.pq$','' } else { 'UnknownTest' }; [void]$sb.AppendLine("        <div class='test-section'>"); [void]$sb.AppendLine("            <div class='test-title'>TEST: $(ConvertTo-HtmlEncodedString $testName)</div>"); if ($testResult.Output -and $testResult.Output.SerializedSource) { $serializedData=$testResult.Output.SerializedSource; $tablePattern = '#table\s*\(\s*type\s+table\s+\[([^\]]+)\]\s*,\s*\{(.*)\}\s*\)'; if ($serializedData -match $tablePattern) { $columnDefs=$Matches[1]; $dataRows=$Matches[2]; $columnNames=@(); $columnMatches=[regex]::Matches($columnDefs,'(?:#\"([^\"]+)\"|(\w+))\s*=\s*\w+'); foreach($match in $columnMatches){ if($match.Groups[1].Value){$columnNames+=$match.Groups[1].Value} elseif($match.Groups[2].Value){$columnNames+=$match.Groups[2].Value} } $parsedRows=@(); $rowPattern='\{([^}]+)\}'; $rowMatches=[regex]::Matches($dataRows,$rowPattern); foreach($rowMatch in $rowMatches){ $rowData=$rowMatch.Groups[1].Value; $values=@(); $valuePattern='\"([^\"]*)\"'; $valueMatches=[regex]::Matches($rowData,$valuePattern); foreach($valueMatch in $valueMatches){ $values+=$valueMatch.Groups[1].Value } if($values.Count -gt 0){ $parsedRows+=,$values } }
                    [void]$sb.AppendLine('            <table>'); [void]$sb.AppendLine('                <thead>'); [void]$sb.AppendLine('                    <tr>'); $columnCount=Get-SafeCount $columnNames; if($columnCount -gt 0){ foreach($col in $columnNames){ [void]$sb.AppendLine("                        <th>$(ConvertTo-HtmlEncodedString $col)</th>") } } else { for($i=0;$i -lt 5;$i++){ [void]$sb.AppendLine("                        <th>Column$($i+1)</th>") } } [void]$sb.AppendLine('                    </tr>'); [void]$sb.AppendLine('                </thead>'); [void]$sb.AppendLine('                <tbody>'); foreach($row in $parsedRows){ [void]$sb.AppendLine('                    <tr>'); $rowCount=Get-SafeCount $row; $maxCols=[Math]::Max($columnCount,$rowCount); for($i=0;$i -lt $maxCols;$i++){ $cellClass=''; $value= if($i -lt $rowCount){ $row[$i] } else { '' }; if($i -lt $columnCount){ $columnName=$columnNames[$i]; if($columnName -match 'Result|Status|Validation'){ switch($value.ToLower()){ 'pass'{$cellClass='pass-cell'; $value='PASS'}; 'fail'{$cellClass='fail-cell'; $value='FAIL'}; 'info'{$cellClass='info-cell'; $value='INFO'} } } } [void]$sb.AppendLine("                        <td class='$cellClass'>$(ConvertTo-HtmlEncodedString $value)</td>") } [void]$sb.AppendLine('                    </tr>') } [void]$sb.AppendLine('                </tbody>'); [void]$sb.AppendLine('            </table>') } else { $statusClass = if ($testResult.Status -eq 'Failed') { 'simple-fail' } else { 'simple-success' }; [void]$sb.AppendLine("            <div class='simple-result $statusClass'>"); [void]$sb.AppendLine("                <strong>Raw Data:</strong><br>"); [void]$sb.AppendLine("                $(ConvertTo-HtmlEncodedString $serializedData)"); [void]$sb.AppendLine('            </div>') } } else { [void]$sb.AppendLine("            <div class='no-results'>No live results data available</div>") } [void]$sb.AppendLine('        </div>') }
            [void]$sb.AppendLine('    </div>'); [void]$sb.AppendLine('</body>'); [void]$sb.AppendLine('</html>')
            [System.IO.File]::WriteAllText($htmlPath, $sb.ToString(), $UTF8NoBOM)
            if (Test-Path $htmlPath) {
                Write-Host ""; Write-Host "================================================================================"; Write-Host "HTML REPORT CREATED SUCCESSFULLY FROM LIVE RESULTS!"; Write-Host "================================================================================"; Write-Host ("Location: {0}" -f $htmlPath)
                if (-not $NoOpenHTML) { try{ Start-Process $htmlPath; Write-Host "[SUCCESS] Opened in default browser" } catch { Write-Host ("[WARNING] Could not open in browser: {0}" -f $_.Exception.Message) } ; try{ code $htmlPath; Write-Host "[SUCCESS] Opened HTML file in VS Code" } catch { Write-Host "[INFO] VS Code not available for opening HTML file" } }
                $allReports = Get-ChildItem -Path $resultsDirectory -Filter ("TestResults_{0}_*.html" -f $SuiteName) | Sort-Object CreationTime -Descending
                Write-Host ""; Write-Host ("Available reports for {0} (newest first):" -f $SuiteName); foreach($report in $allReports){ $age = if ($report.Name -eq (Split-Path $htmlPath -Leaf)) { " [Just created from LIVE data]" } else { "" }; Write-Host ("  {0}{1}" -f $report.Name,$age) }
                Write-Host ""; Write-Host "MANUAL PREVIEW OPTIONS IN VS CODE:"; Write-Host "  1. Right-click the HTML file → 'Open Preview'"; Write-Host "  2. With HTML file open: Ctrl+Shift+V"; Write-Host "  3. Command Palette: 'Live Preview: Show Preview'"; Write-Host ""; Write-Host "Or simply use the browser that just opened!"; Write-Host "================================================================================"
            } else { Write-Host ("[ERROR] Failed to create HTML report at {0}" -f $htmlPath) }
        }
        Write-Host ""; Write-Host "HTML preview generation completed!"
    }

    if ($CopyMez) {
        Write-Color "Requested to copy MEZ file to Power BI Custom Connector folder..." Cyan
        if (-not (Test-Path $ExtensionPath)) { Write-Color "ERROR: The source MEZ file was not found at $ExtensionPath" Red; exit 1 }
        if (-not (Test-Path $MezTarget)) { Write-Color "Destination folder $MezTarget not found. Creating..." Yellow; New-Item -ItemType Directory -Path $MezTarget -Force | Out-Null }
        $destination = Join-Path $MezTarget (Split-Path $ExtensionPath -Leaf)
        Copy-Item -Path $ExtensionPath -Destination $destination -Force
        Write-Color "Successfully copied MEZ file to $destination" Green
    } else { Write-Color "Skipping MEZ deploy/copy step. Use -CopyMez to enable." Yellow }
}
catch{ Write-Color ("ERROR: {0}" -f $_.Exception.Message) Red; if ($_.InvocationInfo) { Write-Host ("At: {0}:{1}" -f $_.InvocationInfo.ScriptName, $_.InvocationInfo.ScriptLineNumber) -ForegroundColor Gray } ; exit 1 }
