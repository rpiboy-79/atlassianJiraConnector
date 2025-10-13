// Test_OrderBy_ParameterAcceptance.query.pq
(parameter) =>
let
    // INLINE HELPER FUNCTIONS (copy from your template)
    testName = "ORDER BY Parameter Acceptance",
    
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
    
    GetFieldValue = (row as record, fieldName as text) as any =>
        try 
            Record.Field(row, fieldName) 
        otherwise 
            null,
    
    // ACTUAL TEST LOGIC STARTS HERE
    
    // Step 1: Extract connector data (should use ORDER BY parameter)
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    // Step 2: Validate basic structure
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,
    ColumnNames = Table.ColumnNames(IssueData),

    // Step 2A: Wrap Test in error handling and no-data handling
    test =
        if HasData and HasColumns then
            try
                let
                    // Step 3: Validate connector accepts ORDER BY without crashing
                    IdColumn = FlexibleColumnCheck(IssueData, {"id", "Id", "ID"}),
                    KeyColumn = FlexibleColumnCheck(IssueData, {"key", "Key", "KEY"}),
                    CreatedColumn = FlexibleColumnCheck(IssueData, {"created", "Created", "CREATED"}),
                    
                    // Step 4: Basic validation that data was returned (ORDER BY didn't break connector)
                    TestPassed = HasData and HasColumns and IdColumn <> null,

                    ValidateIDColumn = if IdColumn <> null
                        then "ID: " & IdColumn
                        else "Missing ID column",
                    ValidateKeyColumn = if KeyColumn <> null
                        then "Key: " & KeyColumn
                        else "Missing Key column",
                    
                    // Step 5: Create results table using your format
                    TestResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "Connector Accepts ORDER BY",
                                if HasData then "Pass" else "Fail",
                                "Connector processes ORDER BY parameter without error",
                                if HasData then "Data returned successfully" else "No data returned",
                                "Validates that ORDER BY parameter doesn't break connector"
                            },
                            {
                                "Required Columns Present", 
                                if IdColumn <> null and KeyColumn <> null then "Pass" else "Fail",
                                "Essential columns (id, key) are present",
                                ValidateIDColumn & ", " & ValidateKeyColumn,
                                "Ensures basic Jira data structure is intact"
                            },
                            {
                                "Created Column Available",
                                if CreatedColumn <> null then "Pass" else "Fail", 
                                "Created column present for date sorting tests",
                                if CreatedColumn <> null then "Created column: " & CreatedColumn else "Created column not found",
                                "Verifies date field is available for ORDER BY created tests"
                            },
                            {
                                "Data Count",
                                "Info",
                                "Issues returned for ORDER BY validation",
                                Text.From(RowCount) & " issues returned",
                                "Shows data volume available for sorting validation"
                            },
                            {
                                "Available Columns",
                                "Info",
                                "All columns available for ORDER BY testing",
                                Text.Combine(ColumnNames, ", "),
                                "Lists all columns that can be used in ORDER BY tests"
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
                            testName,
                            "Error",
                            "Successful ORDER BY parameter processing",
                            "An error occurred during analysis",
                            "Error in test execution - check ORDER BY implementation"
                        }
                   }
                )
        else
            #table(
                {"Test", "Result", "Expected", "Actual", "Details"},
                {
                    {
                        testName,
                        "Fail",
                        "Data returned from API with ORDER BY", 
                        "No data available",
                        "ORDER BY parameter may have caused API failure or filtering error"
                    }
                }
            )
in
    test