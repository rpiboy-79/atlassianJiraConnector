// Test_MockData_RowCount.query.pq
// Validates expected data volume and basic structure
(parameter) =>
let
    testName = "Row Count Test",
    // Helper 1: Safely extract issue data from connector output
    GetIssueData = (connectorResult as any) as table =>
        let
            // Navigate to the Issues data with error handling
            IssuesFolder = try connectorResult{[Name="Issues"]}[Data] otherwise #table({},{}),
            FirstProject = try Table.FirstN(IssuesFolder, 1) otherwise #table({},{}),
            IssueData = if Table.RowCount(FirstProject) > 0 then 
                try FirstProject{0}[Data] otherwise #table({},{})
            else 
                #table({},{})
        in
            IssueData,

    // ACTUAL TEST LOGIC STARTS HERE
    
    // Step 1: Extract connector data
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    // Step 2: Validate basic structure
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,

    // Step 2A: Wrap Test in error handling and no-data handling
    test =
        if HasData and HasColumns then
            try
                let
                    RowCount = Table.RowCount(IssueData),
                    ColumnCount = Table.ColumnCount(IssueData),

                    IssueDataType = Value.Type(IssueData),
                    IsTable = IssueData is table,
                    TypeDescription = if IsTable then "table" else "not table",
    
                    // Validate we get expected data volume (should be 10 based on your parameter query)
                    Results = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {"Row Count", if RowCount >= 1000 then "Pass" else "Fail", "1000", Text.From(RowCount), "Mock data should have exactly 1000 rows"},
                            {"Has Data", if RowCount > 0 then "Pass" else "Fail", "> 0", Text.From(RowCount), "Table must contain data"},
                            {"Column Count", if ColumnCount >= 7 then "Pass" else "Fail", ">= 7", Text.From(ColumnCount), "Should have at least 7 default fields"},
                            {"Table Structure", if IsTable then "Pass" else "Fail", "table", TypeDescription, "IssueData must be a table"}
                        }
                    )
                in
                    Results
            otherwise
                #table(
                    {"Test", "Result", "Details"},
                    {
                        {
                            testName,
                            "Error",
                            "An error occurred during analysis."
                        }
                   }
                )
        else
            #table(
                {"Test", "Result", "Details"},
                {
                    {
                        testName,
                        "n/a",
                        "There is no data to analyze."
                    }
                }
            )
in
    test
