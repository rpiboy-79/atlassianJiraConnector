// Test_MockData_DataProfiling.query.pq
// Profiles data quality and identifies patterns
(parameter) =>
let
    // INLINE HELPER FUNCTIONS (copy these into each test file)
    
    //fill in test name here
    testName = "Data Profiling",
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
    
    // Helper 2: Flexible preference based column detection (case-insensitive), i.e. can search for a column based on a list of possible name variations e.g. {"id", "Id", "ID", "Indentifier"}
    FlexibleColumnCheck = (table as table, possibleNames as list) as text =>
        let
            actualColumns = List.Transform(Table.ColumnNames(table), Text.Lower),
            possibleLower = List.Transform(possibleNames, Text.Lower),
            foundColumn = List.First(List.Intersect({possibleLower, actualColumns}), null),
            // Map back to actual column name
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
    
    // Step 1: Extract connector data
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    // Step 2: Validate basic structure
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,
    ColumnNames = Table.ColumnNames(IssueData),
    
    test =
        if HasData and HasColumns then
            try
                let
                    //Profile key Text fields
                    IdProfile = [
                        Field = "id",
                        UniqueCount = Table.RowCount(Table.Distinct(Table.SelectColumns(IssueData, {"id"}))),
                        NullCount = Table.RowCount(Table.SelectRows(IssueData, each [id] = null)),
                        EmptyCount = Table.RowCount(Table.SelectRows(IssueData, each [id] = "")),
                        Result = if UniqueCount = RowCount and NullCount = 0 then "Pass" else "Fail"
                    ],
                    
                    KeyProfile = [
                        Field = "key", 
                        UniqueCount = Table.RowCount(Table.Distinct(Table.SelectColumns(IssueData, {"key"}))),
                        NullCount = Table.RowCount(Table.SelectRows(IssueData, each [key] = null)),
                        HasValidFormat = Table.RowCount(Table.SelectRows(IssueData, each Text.Contains([key], "-"))),
                        Result = if UniqueCount = RowCount and NullCount = 0 and HasValidFormat = RowCount then "Pass" else "Fail"
                    ],
                    
                    SummaryProfile = [
                        Field = "summary",
                        NullCount = Table.RowCount(Table.SelectRows(IssueData, each [summary] = null)),
                        EmptyCount = Table.RowCount(Table.SelectRows(IssueData, each [summary] = "")),
                        AvgLength = Number.Round(List.Average(List.Transform(Table.Column(IssueData, "summary"), each Text.Length(_ ?? ""))), 2),
                        Result = if NullCount + EmptyCount < RowCount * 0.1 then "Pass" else "Fail"  // Less than 10% null/empty
                    ],
                    
                    CreatedProfile = [
                        Field = "created",
                        NullCount = Table.RowCount(Table.SelectRows(IssueData, each [created] = null)),
                        ValidDateCount = Table.RowCount(Table.SelectRows(IssueData, each try DateTime.From([created]) is datetime otherwise false)),
                        Result = if NullCount = 0 and ValidDateCount = RowCount then "Pass" else "Fail"
                    ],
                    
                    // Convert profiles to test results
                    Results = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {"ID Uniqueness", IdProfile[Result], "All unique", Text.From(IdProfile[UniqueCount]) & "/" & Text.From(RowCount), "All IDs should be unique and non-null"},
                            {"Key Format", KeyProfile[Result], "PROJECT-NUMBER", Text.From(KeyProfile[HasValidFormat]) & "/" & Text.From(RowCount) & " valid", "All keys should follow Jira format"},
                            {"Summary Quality", SummaryProfile[Result], "< 10% null/empty", Text.From(SummaryProfile[NullCount] + SummaryProfile[EmptyCount]) & " null/empty", "Most issues should have summaries"},
                            {"Summary Length", "Info", "Variable", Text.From(SummaryProfile[AvgLength]) & " avg chars", "Average summary length"},
                            {"Created Dates", CreatedProfile[Result], "All valid dates", Text.From(CreatedProfile[ValidDateCount]) & "/" & Text.From(RowCount) & " valid", "All created dates should be parseable"}
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