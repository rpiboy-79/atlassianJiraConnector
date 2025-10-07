// Test_MockData_DataTypes.query.pq
// Validates field data types match expectations
(parameter) =>
let
    //fill in test name here
    testName = "Data Types Validation",
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
                    Schema = Table.Schema(IssueData),
                    
                    // Expected data types for key fields
                    ExpectedTypes = #table(
                        {"Test Name", "Expected"},
                        {
                            {"id", "text"},
                            {"key", "text"}, 
                            {"summary", "text"},
                            {"created", "datetimezone"},
                            {"description", "any"},  // Complex type, should be record or text
                            {"issuetype", "any"},    // Complex type, should be record 
                            {"components", "any"}    // Complex type, should be list or null
                        }
                    ),
                    
                    // Join with actual schema
                    TypeComparison = Table.NestedJoin(
                        ExpectedTypes, {"Test Name"}, 
                        Schema, {"Name"}, 
                        "ActualInfo", JoinKind.LeftOuter
                    ),
                    
                    // Validate types
                    Results = Table.AddColumn(TypeComparison, "Result", 
                        each if [ActualInfo] = null then "Missing"
                        else if [Expected] = "any" then "Pass"  // Complex types just need to exist
                        else if Text.Contains(Text.Lower([ActualInfo]{0}[TypeName]), Text.Lower([Expected])) then "Pass"
                        else "Fail"
                    ),
                    
                    AddActual = Table.AddColumn(Results, "Actual", each [ActualInfo]{0}[TypeName]),
                    AddDetails = Table.AddColumn(AddActual, "Details", each "Column data Type should match expected, 'Any' type passes 100%"),

                    FinalResults = Table.SelectColumns(AddDetails, {"Test Name", "Result", "Expected", "Actual", "Details"})
                in
                    FinalResults
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
