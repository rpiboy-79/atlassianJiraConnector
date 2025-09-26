// Test_FieldTypes.query.pq  
// Validates field types match expectations
(parameter) =>
let
    // Navigate to first project's issues
    Source = parameter,
    IssuesFolder = Source{[Name="Issues"]}[Data],
    FirstProject = Table.FirstN(IssuesFolder, 1),
    ProjectIssues = if Table.RowCount(FirstProject) > 0 then FirstProject{0}[Data] else #table({},{}),
    
    // Get column information
    ColumnInfo = Table.Schema(ProjectIssues),
    
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
            "ValidationResult", 
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
    
    FinalResults = Table.SelectColumns(WithValidation, {"ColumnName", "ExpectedType", "ValidationResult"})
in
    FinalResults
