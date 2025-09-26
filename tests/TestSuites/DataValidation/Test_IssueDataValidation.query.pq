// Test_IssueDataValidation.query.pq
// Validates actual issue data structure and content from live connection
(parameter) =>
let
    // Get navigation table from connector
    Source = parameter,
    
    // Navigate to Issues folder 
    IssuesFolder = Source{[Name="Issues"]}[Data],
    
    // Get first project's issues (assuming we have at least one project)
    FirstProject = Table.FirstN(IssuesFolder, 1),
    ProjectIssues = if Table.RowCount(FirstProject) > 0 then FirstProject{0}[Data] else #table({},{}),
    
    // Validate data structure
    IssueColumns = Table.ColumnNames(ProjectIssues),
    IssueCount = Table.RowCount(ProjectIssues),
    
    // Check for expected Jira fields based on your parameter query
    ExpectedFields = {"id", "key", "summary", "description", "issuetype", "created", "componnents"},
    HasExpectedFields = List.Transform(ExpectedFields, each List.Contains(IssueColumns, _)),
    
    // Sample some data for validation
    SampleIssue = if IssueCount > 0 then ProjectIssues{0} else null,

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
            {"Issue Count", if IssueCount > 0 then "Pass" else "Fail", "> 0", Text.From(IssueCount), "Number of issues returned"},
            {"Has ID Field", if HasExpectedFields{0} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{0}), "ID field present"},
            {"Has Key Field", if HasExpectedFields{1} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{1}), "Key field present"},
            {"Has Summary Field", if HasExpectedFields{2} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{2}), "Summary field present"},
            {"Has Issue Type Field", if HasExpectedFields{3} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{3}), "Issue type field present"},
            {"Has Created Field", if HasExpectedFields{4} then "Pass" else "Fail", "true", Text.From(HasExpectedFields{4}), "Created field present"},
            {"Sample Issue Key Format", if KeyValidation[IsValid] then "Pass" else "Fail", "PROJECT-NUMBER", KeyValidation[KeyValue], "Jira key format validation (PROJECT-####)"},
            {"Total Columns", "Info", "Variable", Text.From(List.Count(IssueColumns)), Text.Combine(IssueColumns, ", ")}
        }
    )
in
    ValidationResults
