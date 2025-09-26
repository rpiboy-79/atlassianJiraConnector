// TestCanConnect.query.pq
// Tests basic connectivity to Jira
(parameter)=>
let
    // Get the connection from parameter query
    Source = parameter,
    // Try to take first row to verify connection works
    FirstRow = try Table.FirstN(Source, 1) otherwise null,
    // Check if we got data
    HasData = FirstRow <> null,
    // Return simple test result
    Result = if HasData then "Connected Successfully" else "Connection Failed"
in
    Result

// Test_RowCount.query.pq
// Simple test that just counts rows (like Microsoft's FirstN example)

/*(parameter) =>
let
    // Just count the rows - this should create a tiny .pqout file
    Count = Table.RowCount(parameter),
    
    // Return single value in table format (matches Microsoft pattern)
    Result = #table({"RowCount"}, {{Count}})
in
    Result

*/
/*
// Test_CanConnect.query.pq - Simple connection validation test
// Returns minimal data for baseline comparison per Microsoft standards

let
    // Test configuration
    company = "drofus",
    
    // Connect via the loaded connector
    Source = try atlassianJiraConnector.Contents(company) otherwise null,
    
    // Create a simple validation table instead of returning the full navigation
    ValidationResult = if Source = null then
        #table({"ConnectionStatus"}, {{"Failed"}})
    else if not (Source is table) then
        #table({"ConnectionStatus"}, {{"Invalid Response Type"}})
    else
        let
            rowCount = Table.RowCount(Source),
            hasNameColumn = List.Contains(Table.ColumnNames(Source), "Name"),
            names = if hasNameColumn then Table.Column(Source, "Name") else {},
            hasIssues = List.Contains(names, "Issues")
        in
            #table(
                {"ConnectionStatus", "NavigationRows", "HasIssuesEntry"}, 
                {{"Success", rowCount, hasIssues}}
            )
in
    ValidationResult

*/

/*
// Test_CanConnect.query.pq - Comprehensive connection test with tabular results
// Returns test results in table format as expected by Microsoft Test Framework

let
    // Test configuration
    company = "drofus",
    
    // Test connection and gather results
    Source = try atlassianJiraConnector.Contents(company) otherwise null,
    
    // Analyze the results
    TestResults = 
        if Source = null then
            // Connection failed
            #table(
                {"Test", "Result", "Details"},
                {
                    {"Connection", "Failed", "Could not connect to Jira"},
                    {"Navigation Table", "Failed", "No data returned"},
                    {"Has Issues Entry", "Failed", "No navigation structure"}
                }
            )
        else if not (Source is table) then
            // Wrong data type returned
            #table(
                {"Test", "Result", "Details"},
                {
                    {"Connection", "Passed", "Connector responded"},
                    {"Navigation Table", "Failed", "Returned " & Value.Type(Source)},
                    {"Has Issues Entry", "Failed", "Not a table"}
                }
            )
        else
            // Connection successful, analyze table structure
            let
                columns = Table.ColumnNames(Source),
                hasName = List.Contains(columns, "Name"),
                hasData = List.Contains(columns, "Data"),
                rowCount = Table.RowCount(Source),
                names = if hasName then Table.Column(Source, "Name") else {},
                hasIssues = List.Contains(names, "Issues")
            in
                #table(
                    {"Test", "Result", "Details"},
                    {
                        {"Connection", "Passed", "Connected to " & company},
                        {"Navigation Table", if hasName and hasData then "Passed" else "Failed", 
                         "Columns: " & Text.Combine(columns, ", ")},
                        {"Row Count", if rowCount > 0 then "Passed" else "Failed", 
                         Text.From(rowCount) & " rows"},
                        {"Has Issues Entry", if hasIssues then "Passed" else "Failed",
                         "Available: " & Text.Combine(names, ", ")}
                    }
                )
in
    TestResults
*/


/*
// Simple connection test - bypasses Extension.Test complexity
let
    // Test configuration
    TestConfig = [
        BaseUrl = "https://drofus.atlassian.net",
        Company = "drofus"
    ],
    
    // Direct connector test - bypasses Extension.Test module
    Source = atlassianJiraConnector.Contents(TestConfig[Company]),
    
    // Validate we got a navigation table
    IsTable = Source is table,
    HasColumns = if IsTable then List.Contains(Table.ColumnNames(Source), "Name") else false,
    HasData = if IsTable then Table.RowCount(Source) > 0 else false,
    
    // Return test result
    TestResult = IsTable and HasColumns and HasData
in
    TestResult
*/

// Import Extension.Test functions
/*jiraDataRetrievalSimple = #"Extension.Test"[jiraDataRetrievalSimple];

// Test configuration
let
    TestConfig = [
        BaseUrl = "https://drofus.atlassian.net"
    ],
    result = jiraDataRetrievalSimple(
        TestConfig[BaseUrl],
        "rest/api/3/project",
        null, null, "5"
    )
in
    List.Count(result) > 0*/
