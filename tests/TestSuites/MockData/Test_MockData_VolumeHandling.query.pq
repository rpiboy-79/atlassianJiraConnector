// Test_MockData_VolumeHandling.query.pq
// Tests connector behavior with larger data sets
(parameter) =>
let
    testName = "Volume Handling Test",
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
                    // Perform typical operations
                    DistinctKeys = Table.RowCount(Table.Distinct(Table.SelectColumns(IssueData, {"key"}))),
                    GroupedByType = Table.Group(IssueData, {"issuetype"}, {{"Count", Table.RowCount, type number}}),
                    TypeCount = Table.RowCount(GroupedByType),
                    // Test basic operations on the full dataset
                    StartTime = DateTimeZone.UtcNow(),
                    TimingTest = Table.RowCount(Table.SelectRows(IssueData, each [id] <> null)),
                    EndTime = DateTimeZone.UtcNow(),
                    
                    DurationInSeconds = Duration.TotalSeconds(EndTime - StartTime),
                    DurationInMs = DurationInSeconds * 1000,
                    DurationOk = DurationInMs < 5000,
                    
                    Results = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "Data Volume",
                                if RowCount = 1000 then "Pass" else "Fail",
                                "1000 rows",
                                Text.From(RowCount) & " rows",
                                "Should handle expected volume"
                            },
                            {
                                "Key Uniqueness",
                                if DistinctKeys = RowCount then "Pass" else "Fail",
                                "All unique",
                                Text.From(DistinctKeys) & "/" & Text.From(RowCount),
                                "All keys should be unique"
                            },
                            {"Grouping Operations", if TypeCount > 0 then "Pass" else "Fail", "> 0 groups", Text.From(TypeCount) & " issue types", "Should support grouping operations"},
                            //{"Processing Time", if DurationInMs < 5000 then "Pass" else "Fail", "< 5 seconds", Text.From(DurationInMs) & "ms", "Should process data efficiently"}
                            {
                                "Calculated Duration",
                                if DurationOk then "Pass" else "Fail",
                                "value greater less than 5000",
                                if DurationOk then "value greater than 0" else Text.From(DurationInSeconds),
                                "Duration to count all rows"
                            }
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
