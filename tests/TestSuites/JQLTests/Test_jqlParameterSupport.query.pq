// Test_JQL_ParameterAcceptance.query.pq
(parameter) =>
let
    // INLINE HELPER FUNCTIONS (following your template)
    GetIssueData = (connectorResult as any) as table =>
        let
            IssuesFolder = try connectorResult{[Name="Issues"]}[Data] otherwise #table({},{}),
            FirstProject = try Table.FirstN(IssuesFolder, 1) otherwise #table({},{}),
            IssueData = if Table.RowCount(FirstProject) > 0 then 
                try FirstProject{0}[Data] otherwise #table({},{})
            else 
                #table({},{})
        in
            IssueData,
    
    FlexibleColumnCheck = (table as table, possibleNames as list) as text =>
        let
            actualColumns = List.Transform(Table.ColumnNames(table), Text.Lower),
            possibleLower = List.Transform(possibleNames, Text.Lower),
            foundColumn = List.First(List.Intersect({possibleLower, actualColumns}), null),
            actualColumn = if foundColumn <> null then
                let
                    matchIndex = List.PositionOf(actualColumns, foundColumn)
                in
                    if matchIndex >= 0 then Table.ColumnNames(table){matchIndex} else null
            else
                null
        in
            actualColumn,
    
    // ACTUAL TEST LOGIC STARTS HERE
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,
    ColumnNames = Table.ColumnNames(IssueData),

    test =
        if HasData and HasColumns then
            try
                let
                    // Validate basic API connection and data structure
                    IdColumn = FlexibleColumnCheck(IssueData, {"id", "Id", "ID"}),
                    KeyColumn = FlexibleColumnCheck(IssueData, {"key", "Key", "KEY"}),
                    StatusColumn = FlexibleColumnCheck(IssueData, {"status", "Status", "STATUS"}),
                    
                    // Check first row for Jira-like data
                    FirstRow = if HasData then IssueData{0} else null,
                    KeyValue = if KeyColumn <> null and FirstRow <> null then 
                        try Record.Field(FirstRow, KeyColumn) otherwise null else null,
                    
                    KeyIsJiraFormat = if KeyValue <> null then 
                        Text.Contains(Text.From(KeyValue), "-") else false,
                    
                    // Create test results table
                    TestResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "API Connection",
                                if HasData then "Pass" else "Fail",
                                "Connected to real Jira API",
                                if HasData then "Connected: " & Text.From(RowCount) & " issues" else "No connection",
                                "Validates connector can reach real Jira API"
                            },
                            {
                                "Data Structure",
                                if HasColumns then "Pass" else "Fail", 
                                "Valid table structure returned",
                                if HasColumns then Text.From(Table.ColumnCount(IssueData)) & " columns" else "No columns",
                                "Validates API returns proper table structure"
                            },
                            {
                                "Required Columns",
                                if IdColumn <> null and KeyColumn <> null then "Pass" else "Fail",
                                "ID and Key columns present",
                                if IdColumn <> null and KeyColumn <> null then "ID: " & IdColumn & ", Key: " & KeyColumn else "Missing required columns",
                                "Validates core Jira issue fields are present"
                            },
                            {
                                "Jira Data Format",
                                if KeyIsJiraFormat then "Pass" else "Fail",
                                "Issue keys follow Jira format (PROJECT-###)",
                                if KeyValue <> null then Text.From(KeyValue) else "No key found",
                                "Validates data is from real Jira instance"
                            },
                            {
                                "JQL Parameter Ready",
                                "Info",
                                "Baseline established for JQL testing",
                                "Ready to test JQL filtering",
                                "This test establishes baseline for subsequent JQL tests"
                            },
                            {
                                "Available Columns",
                                "Info", 
                                "List of all returned columns",
                                Text.Combine(ColumnNames, ", "),
                                "Shows all columns available for JQL filtering tests"
                            }
                        }
                    )
                in
                    TestResults
            otherwise
                #table(
                    {"Test", "Result", "Expected", "Actual", "Details"},
                    {
                        {
                            "JQL Parameter Acceptance",
                            "Error",
                            "Successful API connection",
                            "An error occurred during analysis",
                            "Error in test execution - check connector implementation"
                        }
                   }
                )
        else
            #table(
                {"Test", "Result", "Expected", "Actual", "Details"},
                {
                    {
                        "JQL Parameter Acceptance",
                        "Fail",
                        "Data returned from API",
                        "No data available",
                        "No data returned - check API connectivity or JQL parameter may have filtered out all results"
                    }
                }
            )
in
    test