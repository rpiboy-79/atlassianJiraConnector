// Test_FieldTypes.query.pq  
// Validates field types match expectations
(parameter) =>
let
    testName = "Field Types Test",
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
                    // Get column information
                    ColumnInfo = Table.Schema(IssueData),
                    
                    // Define expected types for key fields
                    ExpectedTypes = #table(
                        {"ColumnName", "ExpectedType"},
                        {
                            {"id", "text"},
                            {"key", "text"}, 
                            {"summary", "text"},
                            {"created", "datetimezone"}
                        }
                    ),
                    
                    // Join actual with expected types
                    TypeValidation = Table.NestedJoin(
                        ExpectedTypes, {"ColumnName"}, 
                        ColumnInfo, {"Name"}, 
                        "ActualInfo", JoinKind.LeftOuter
                    ),

                    // Add validation columns
                    WithValidation =
                        Table.AddColumn(
                            TypeValidation,
                            "Result", 
                            each if [ActualInfo] = null 
                                then "Missing"
                                else
                                    if 
                                        Text.Contains(
                                            Text.Lower([ActualInfo]{0}[TypeName]),
                                            Text.Lower([ExpectedType])
                                        )
                                    then "Pass"
                                    else "Fail"
                        ),
                    
                    ExpandActualInfo = Table.ExpandTableColumn(WithValidation, "ActualInfo", {"TypeName"}, {"Actual"}),
                    Results = Table.SelectColumns(ExpandActualInfo, {"ColumnName", "ExpectedType", "Result", "Actual"}),
                    TableRenameColumns = Table.RenameColumns(Results, {{"ColumnName", "Test Name"}, {"ExpectedType", "Expected"}}),
                    AddDetails = Table.AddColumn(TableRenameColumns, "Details", each "Validates known column Types"),
                    OrderColumns = Table.ReorderColumns(AddDetails, {"Test Name", "Result", "Expected", "Actual", "Details"})

                in
                    OrderColumns
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
    
