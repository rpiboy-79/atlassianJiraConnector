// Test_MockData_ComplexFields.query.pq
// Analyzes complex fields that your connector handles as 'any' type
(parameter) =>
let
    //fill in test name here
    testName = "Complex Field Analysis",
    // INLINE HELPER FUNCTIONS (copy these into each test file)
    
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
    
    test =
        if HasData and HasColumns then
            try
                let
                    HasDescription = List.Contains(ColumnNames, "description"),
                    HasComponents = List.Contains(ColumnNames, "components"),
                    HasIssueType = List.Contains(ColumnNames, "issuetype"),

                    // Safely analyze description field complexity
                    DescriptionAnalysis = 
                        if HasDescription then
                            let
                                DescColumn = Table.Column(IssueData, "description"),
                                RecordCount = List.Count(List.Select(DescColumn, each _ is record)),
                                TextCount = List.Count(List.Select(DescColumn, each _ is text)),
                                NullCount = List.Count(List.Select(DescColumn, each _ = null))
                            in
                                [RecordCount = RecordCount, TextCount = TextCount, NullCount = NullCount]
                        else
                            [RecordCount = 0, TextCount = 0, NullCount = RowCount],

                    // Safely analyze issuetype field
                    IssueTypeAnalysis = 
                        if HasIssueType then
                            let
                                IssueTypeColumn = Table.Column(IssueData, "issuetype"),
                                RecordCount = List.Count(List.Select(IssueTypeColumn, each _ is record)),
                                UniqueTypes = Table.RowCount(Table.Distinct(Table.SelectColumns(IssueData, {"issuetype"})))
                            in
                                [RecordCount = RecordCount, UniqueTypes = UniqueTypes]
                        else
                            [RecordCount = 0, UniqueTypes = 0],
                    
                    // Safely analyze components field
                    ComponentsAnalysis = 
                        if HasComponents then
                            let
                                ComponentsColumn = Table.Column(IssueData, "components"),
                                ListCount = List.Count(List.Select(ComponentsColumn, each _ is list)),
                                NullCount = List.Count(List.Select(ComponentsColumn, each _ = null))
                            in
                                [ListCount = ListCount, NullCount = NullCount]
                        else
                            [ListCount = 0, NullCount = RowCount],

                    //results output
                    Results = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {"Description Type", if DescriptionAnalysis[RecordCount] > 0 then "Pass" else "Fail", "Complex records", Text.From(DescriptionAnalysis[RecordCount]) & " records, " & Text.From(DescriptionAnalysis[TextCount]) & " text", "Description should be complex Jira content"},
                            {"IssueType Structure", if IssueTypeAnalysis[RecordCount] > 0 then "Pass" else "Fail", "Record objects", Text.From(IssueTypeAnalysis[RecordCount]) & "/" & Text.From(RowCount) & " are records", "Issue types should be structured objects"},
                            {"IssueType Variety", if IssueTypeAnalysis[UniqueTypes] >= 2 then "Pass" else "Info", ">= 2 types", Text.From(IssueTypeAnalysis[UniqueTypes]) & " unique types", "Should have multiple issue types"},
                            {"Components Structure", if ComponentsAnalysis[ListCount] > 0 then "Pass" else "Info", "List objects", Text.From(ComponentsAnalysis[ListCount]) & "/" & Text.From(RowCount) & " have components", "Components should be arrays when present"},
                            {"Components Coverage", "Info", "Variable", Text.From(RowCount - ComponentsAnalysis[NullCount]) & "/" & Text.From(RowCount) & " have components", "Percentage of issues with components"}
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

    

    
    
    
    
    
    
    
