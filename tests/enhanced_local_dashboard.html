<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power Query Test Results Dashboard - Enhanced</title>
    <style>
        /* Enhanced styles maintaining original look and feel */
        :root {
            --primary-color: #0078d4;
            --success-color: #107c10;
            --danger-color: #d13438;
            --warning-color: #f9a825;
            --info-color: #0078d4;
            --light-bg: #f8f9fa;
            --border-color: #dee2e6;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f5f5;
            color: #333;
            line-height: 1.6;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }

        .header {
            background: var(--primary-color);
            color: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .header h1 {
            font-size: 2rem;
            margin-bottom: 10px;
        }

        .header p {
            opacity: 0.9;
            font-size: 1.1rem;
        }

        .status-bar {
            background: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .status-indicator {
            padding: 8px 16px;
            border-radius: 4px;
            font-weight: bold;
            text-transform: uppercase;
            font-size: 0.9rem;
        }

        .status-connected { background: #d4edda; color: #155724; }
        .status-error { background: #f8d7da; color: #721c24; }
        .status-loading { background: #fff3cd; color: #856404; }

        .refresh-btn {
            background: var(--primary-color);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 4px;
            cursor: pointer;
            font-weight: bold;
            transition: all 0.3s ease;
        }

        .refresh-btn:hover {
            background: #106ebe;
            transform: translateY(-2px);
        }

        .test-suites {
            display: grid;
            gap: 30px;
        }

        .test-suite {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        .suite-header {
            background: var(--light-bg);
            padding: 20px;
            border-bottom: 1px solid var(--border-color);
        }

        .suite-title {
            font-size: 1.5rem;
            font-weight: bold;
            color: var(--primary-color);
            margin-bottom: 10px;
        }

        .suite-metadata {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }

        .metadata-item {
            background: white;
            padding: 10px;
            border-radius: 4px;
            border-left: 4px solid var(--primary-color);
        }

        .metadata-label {
            font-size: 0.8rem;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 4px;
        }

        .metadata-value {
            font-weight: bold;
            font-size: 1rem;
        }

        .suite-summary {
            display: flex;
            gap: 20px;
            margin-bottom: 15px;
        }

        .summary-stat {
            background: var(--light-bg);
            padding: 15px;
            border-radius: 6px;
            text-align: center;
            flex: 1;
        }

        .stat-number {
            font-size: 2rem;
            font-weight: bold;
            margin-bottom: 5px;
        }

        .stat-label {
            font-size: 0.9rem;
            color: #666;
            text-transform: uppercase;
        }

        .stat-passed .stat-number { color: var(--success-color); }
        .stat-failed .stat-number { color: var(--danger-color); }
        .stat-total .stat-number { color: var(--primary-color); }

        .test-results-table {
            margin: 20px;
        }

        .results-table {
            width: 100%;
            border-collapse: collapse;
        }

        .results-table th,
        .results-table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }

        .results-table th {
            background: var(--light-bg);
            font-weight: bold;
            color: #333;
        }

        .results-table tbody tr:hover {
            background: #f8f9fa;
        }

        .result-badge {
            padding: 4px 8px;
            border-radius: 4px;
            font-weight: bold;
            font-size: 0.8rem;
            text-transform: uppercase;
        }

        .result-pass { background: #d4edda; color: #155724; }
        .result-fail { background: #f8d7da; color: #721c24; }
        .result-info { background: #d1ecf1; color: #0c5460; }

        .loading-spinner {
            display: inline-block;
            width: 20px;
            height: 20px;
            border: 2px solid #f3f3f3;
            border-top: 2px solid var(--primary-color);
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .error-message {
            background: #f8d7da;
            color: #721c24;
            padding: 15px;
            border-radius: 4px;
            margin: 20px;
        }

        .no-data-message {
            background: #fff3cd;
            color: #856404;
            padding: 20px;
            text-align: center;
            border-radius: 4px;
            margin: 20px;
        }

        .file-type-indicator {
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 0.7rem;
            font-weight: bold;
            text-transform: uppercase;
        }

        .file-type-json { background: #e8f5e8; color: #2e7d2e; }
        .file-type-html { background: #fff3cd; color: #856404; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 Power Query Test Results Dashboard</h1>
            <p>Enhanced with JSON API Support - Real-time Test Results & History</p>
        </div>

        <div class="status-bar">
            <div id="connection-status" class="status-indicator status-loading">
                <span class="loading-spinner"></span> Connecting...
            </div>
            <button id="refresh-btn" class="refresh-btn" onclick="loadTestResults()">
                🔄 Refresh Results
            </button>
        </div>

        <div id="test-suites" class="test-suites">
            <div class="no-data-message">
                <div class="loading-spinner" style="margin-bottom: 10px;"></div>
                Loading test results...
            </div>
        </div>
    </div>

    <script>
        // Enhanced JavaScript with JSON API support
        let currentTestFiles = [];
        let connectionStatus = 'connecting';

        async function checkServerConnection() {
            try {
                const response = await fetch('/api/test-files');
                if (response.ok) {
                    updateConnectionStatus('connected');
                    return true;
                } else {
                    updateConnectionStatus('error');
                    return false;
                }
            } catch (error) {
                updateConnectionStatus('error');
                return false;
            }
        }

        function updateConnectionStatus(status) {
            const statusElement = document.getElementById('connection-status');
            connectionStatus = status;

            switch (status) {
                case 'connected':
                    statusElement.className = 'status-indicator status-connected';
                    statusElement.innerHTML = '✅ Connected';
                    break;
                case 'error':
                    statusElement.className = 'status-indicator status-error';
                    statusElement.innerHTML = '❌ Connection Error';
                    break;
                case 'loading':
                    statusElement.className = 'status-indicator status-loading';
                    statusElement.innerHTML = '<span class="loading-spinner"></span> Loading...';
                    break;
            }
        }

        async function loadTestResults() {
            updateConnectionStatus('loading');

            const connected = await checkServerConnection();
            if (!connected) {
                showError('Unable to connect to dashboard server. Please ensure the server is running.');
                return;
            }

            try {
                const response = await fetch('/api/test-files');
                const data = await response.json();

                if (data && data.files) {
                    currentTestFiles = data.files;
                    displayTestResults(data.files);
                } else {
                    showNoData('No test results found.');
                }

                updateConnectionStatus('connected');
            } catch (error) {
                console.error('Error loading test results:', error);
                showError(`Error loading test results: ${error.message}`);
                updateConnectionStatus('error');
            }
        }

        async function loadJSONTestDetails(filename) {
            try {
                const response = await fetch(`/api/json-test-details/${filename}`);
                const data = await response.json();
                return data;
            } catch (error) {
                console.error('Error loading JSON test details:', error);
                return null;
            }
        }

        function displayTestResults(files) {
            const container = document.getElementById('test-suites');

            if (!files || files.length === 0) {
                showNoData('No test results available.');
                return;
            }

            // Group files by test suite
            const suiteGroups = {};
            files.forEach(file => {
                const suiteName = file.test_suite_name || 'Unknown Suite';
                if (!suiteGroups[suiteName]) {
                    suiteGroups[suiteName] = [];
                }
                suiteGroups[suiteName].push(file);
            });

            let html = '';
            for (const [suiteName, suiteFiles] of Object.entries(suiteGroups)) {
                html += generateTestSuiteHTML(suiteName, suiteFiles);
            }

            container.innerHTML = html;

            // Load detailed results for JSON files
            loadDetailedResults();
        }

        async function loadDetailedResults() {
            const jsonFiles = currentTestFiles.filter(f => f.type === 'json');

            for (const file of jsonFiles) {
                try {
                    const details = await loadJSONTestDetails(file.filename);
                    if (details) {
                        updateTestSuiteWithDetails(file.filename, details);
                    }
                } catch (error) {
                    console.error(`Error loading details for ${file.filename}:`, error);
                }
            }
        }

        function generateTestSuiteHTML(suiteName, files) {
            // Use the most recent file for main display
            const latestFile = files[0];
            const hasJson = files.some(f => f.type === 'json');

            return `
                <div class="test-suite" id="suite-${suiteName.replace(/[^a-zA-Z0-9]/g, '_')}">
                    <div class="suite-header">
                        <div class="suite-title">
                            ${suiteName}
                            <span class="file-type-indicator file-type-${latestFile.type}">
                                ${latestFile.type.toUpperCase()}
                            </span>
                        </div>
                        <div class="suite-metadata">
                            <div class="metadata-item">
                                <div class="metadata-label">Run Date</div>
                                <div class="metadata-value">${formatDateTime(latestFile.run_date || latestFile.timestamp)}</div>
                            </div>
                            <div class="metadata-item">
                                <div class="metadata-label">Connector</div>
                                <div class="metadata-value">${latestFile.connector_name || 'Unknown'}</div>
                            </div>
                            <div class="metadata-item">
                                <div class="metadata-label">Version</div>
                                <div class="metadata-value">${latestFile.connector_version || 'N/A'}</div>
                            </div>
                            <div class="metadata-item">
                                <div class="metadata-label">Available Reports</div>
                                <div class="metadata-value">${files.length} file(s)</div>
                            </div>
                        </div>
                        ${hasJson ? generateSuiteSummary(latestFile) : ''}
                    </div>
                    <div class="test-results-table" id="results-${suiteName.replace(/[^a-zA-Z0-9]/g, '_')}">
                        <div class="loading-spinner" style="margin: 20px;"></div>
                        Loading test details...
                    </div>
                </div>
            `;
        }

        function generateSuiteSummary(file) {
            return `
                <div class="suite-summary">
                    <div class="summary-stat stat-total">
                        <div class="stat-number">${file.total_tests || 0}</div>
                        <div class="stat-label">Total Tests</div>
                    </div>
                    <div class="summary-stat stat-passed">
                        <div class="stat-number">${file.passed_tests || 0}</div>
                        <div class="stat-label">Passed</div>
                    </div>
                    <div class="summary-stat stat-failed">
                        <div class="stat-number">${file.failed_tests || 0}</div>
                        <div class="stat-label">Failed</div>
                    </div>
                    <div class="summary-stat">
                        <div class="stat-number">${formatDuration(file.total_duration)}</div>
                        <div class="stat-label">Duration</div>
                    </div>
                </div>
            `;
        }

        function updateTestSuiteWithDetails(filename, details) {
            // Find the corresponding suite container
            const file = currentTestFiles.find(f => f.filename === filename);
            if (!file) return;

            const suiteName = file.test_suite_name || 'Unknown Suite';
            const resultContainer = document.getElementById(`results-${suiteName.replace(/[^a-zA-Z0-9]/g, '_')}`);

            if (!resultContainer) return;

            if (details.test_results && details.test_results.length > 0) {
                resultContainer.innerHTML = generateTestResultsTable(details.test_results);
            } else {
                resultContainer.innerHTML = '<div class="no-data-message">No detailed test results available.</div>';
            }
        }

        function generateTestResultsTable(testResults) {
            let html = `
                <table class="results-table">
                    <thead>
                        <tr>
                            <th>Test Name</th>
                            <th>Result</th>
                            <th>Expected</th>
                            <th>Actual</th>
                            <th>Details</th>
                            <th>Duration</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            for (const test of testResults) {
                const resultClass = getResultClass(test.result);
                const parsedData = test.parsed_data;

                // If we have table data, show individual test rows
                if (parsedData && parsedData.type === 'table' && parsedData.rows) {
                    for (const row of parsedData.rows) {
                        html += `
                            <tr>
                                <td>${escapeHtml(row[0] || '')}</td>
                                <td><span class="result-badge ${getResultClass(row[1])}">${escapeHtml(row[1] || '')}</span></td>
                                <td>${escapeHtml(row[2] || '')}</td>
                                <td>${escapeHtml(row[3] || '')}</td>
                                <td>${escapeHtml(row[4] || '')}</td>
                                <td>${formatDuration(test.duration_seconds)}</td>
                            </tr>
                        `;
                    }
                } else {
                    // Show test-level information
                    html += `
                        <tr>
                            <td>${escapeHtml(test.test_name)}</td>
                            <td><span class="result-badge ${resultClass}">${escapeHtml(test.result)}</span></td>
                            <td>-</td>
                            <td>-</td>
                            <td>${escapeHtml(test.details || 'No details available')}</td>
                            <td>${formatDuration(test.duration_seconds)}</td>
                        </tr>
                    `;
                }
            }

            html += `
                    </tbody>
                </table>
            `;

            return html;
        }

        function getResultClass(result) {
            if (!result) return '';

            const r = result.toString().toLowerCase();
            if (r === 'pass' || r === 'passed') return 'result-pass';
            if (r === 'fail' || r === 'failed') return 'result-fail';
            if (r === 'info') return 'result-info';
            return '';
        }

        function formatDateTime(dateStr) {
            if (!dateStr) return 'Unknown';
            try {
                const date = new Date(dateStr);
                return date.toLocaleString();
            } catch {
                return dateStr;
            }
        }

        function formatDuration(seconds) {
            if (!seconds) return '0s';
            if (seconds < 60) return `${seconds.toFixed(1)}s`;
            const minutes = Math.floor(seconds / 60);
            const remainingSeconds = seconds % 60;
            return `${minutes}m ${remainingSeconds.toFixed(1)}s`;
        }

        function escapeHtml(text) {
            if (!text) return '';
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        function showError(message) {
            const container = document.getElementById('test-suites');
            container.innerHTML = `
                <div class="error-message">
                    <strong>Error:</strong> ${escapeHtml(message)}
                </div>
            `;
        }

        function showNoData(message) {
            const container = document.getElementById('test-suites');
            container.innerHTML = `
                <div class="no-data-message">
                    ${escapeHtml(message)}
                </div>
            `;
        }

        // Initialize dashboard
        document.addEventListener('DOMContentLoaded', function() {
            console.log('🚀 Enhanced Power Query Dashboard loaded');
            loadTestResults();

            // Auto-refresh every 30 seconds
            setInterval(loadTestResults, 30000);
        });

        // Add keyboard shortcut for refresh (Ctrl+R or Cmd+R)
        document.addEventListener('keydown', function(e) {
            if ((e.ctrlKey || e.metaKey) && e.key === 'r') {
                e.preventDefault();
                loadTestResults();
            }
        });
    </script>
</body>
</html>