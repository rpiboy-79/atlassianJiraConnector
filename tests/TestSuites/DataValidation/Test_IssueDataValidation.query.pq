// Test_IssueDataValidation.query.pq
// Validates actual issue data structure and content from live connection
(parameter) =>
let
    testName = "Issue Data Validation",
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
                    // Check for expected Jira fields based on your parameter query
                    ExpectedFields = {"id", "key", "summary", "description", "issuetype", "created", "componnents"},
                    HasExpectedFields = List.Transform(ExpectedFields, each List.Contains(ColumnNames, _)),
                    
                    // Sample some data for validation
                    SampleIssue = if RowCount > 0 then IssueData{0} else null,

                    // JIRA KEY FORMAT VALIDATION (computed outside table)
                    KeyValidation = 
                        if SampleIssue <> null and Record.HasFields(SampleIssue, "key") then
                            let
                                KeyValue = SampleIssue[key],
                                HasHyphen = Text.Contains(KeyValue, "-"),
                                ProjectPart = if HasHyphen then Text.BeforeDelimiter(KeyValue, "-") else "",
                                NumberPart = if HasHyphen then Text.AfterDelimiter(KeyValue, "-") else "",
                                
                                // Project validation (relaxed - no case sensitivity)
                                ProjectLength = Text.Length(ProjectPart),
                                ProjectLengthOK = ProjectLength >= 2 and ProjectLength <= 10,
                                
                                // Number validation
                                NumberIsValid = try Number.From(NumberPart) > 0 otherwise false,
                                
                                IsValid = HasHyphen and ProjectLengthOK and NumberIsValid
                            in
                                [
                                    IsValid = IsValid,
                                    KeyValue = KeyValue,
                                    ProjectPart = ProjectPart,
                                    NumberPart = NumberPart
                                ]
                        else
                            [
                                IsValid = false,
                                KeyValue = "N/A",
                                ProjectPart = "N/A", 
                                NumberPart = "N/A"
                            ],
                    
                    // Create validation results table
                    ValidationResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {"Issue Count", if RowCount > 0 then "Pass" else "Fail", "> 0", Text.From(RowCount), "Number of issues returned"},
                            {"Has ID Field", if HasExpectedFields{0} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{0}), "ID field present"},
                            {"Has Key Field", if HasExpectedFields{1} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{1}), "Key field present"},
                            {"Has Summary Field", if HasExpectedFields{2} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{2}), "Summary field present"},
                            {"Has Issue Type Field", if HasExpectedFields{3} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{3}), "Issue type field present"},
                            {"Has Created Field", if HasExpectedFields{4} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{4}), "Created field present"},
                            {"Sample Issue Key Format", if KeyValidation[IsValid] then "Pass" else "Fail", "PROJECT-NUMBER", KeyValidation[KeyValue], "Jira key format validation (PROJECT-####)"},
                            {"Total Columns", "Info", "Variable", Text.From(List.Count(ColumnNames)), Text.Combine(ColumnNames, ", ")}
                        }
                    )
                in
                    ValidationResults
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