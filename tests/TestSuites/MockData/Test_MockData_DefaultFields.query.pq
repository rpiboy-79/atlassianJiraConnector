// Test_MockData_DefaultFields.query.pq  
// Validates default field set matches connector specification
(parameter) =>
let
    testName = "Default Fields",
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
    ColumnNames = Table.ColumnNames(IssueData),

    // Step 2A: Wrap Test in error handling and no-data handling
    test =
        if HasData and HasColumns then
            try
                let
    
                    // Default fields as defined in your connector
                    ExpectedDefaultFields = {"id", "key", "summary", "description", "issuetype", "created", "components"},
                    
                    // Check each expected field exists
                    FieldTests = List.Transform(
                        ExpectedDefaultFields,
                        each [
                            Field = _,
                            Exists = List.Contains(ColumnNames, _),
                            Result = if Exists then "Pass" else "Fail"
                        ]
                    ),
                    
                    // Convert to test results table
                    Results = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        List.Transform(FieldTests, each {
                            "Field: " & [Field], 
                            [Result], 
                            "Present", 
                            if [Exists] then "Found" else "Missing",
                            "Default connector field must be present"
                        }) & {
                            {"Extra Fields", "Info", "Variable", Text.Combine(List.Difference(ColumnNames, ExpectedDefaultFields), ", "), "Additional fields beyond defaults"}
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
