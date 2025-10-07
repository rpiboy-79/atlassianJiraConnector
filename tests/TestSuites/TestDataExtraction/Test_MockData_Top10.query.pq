// Test_MockData_ComplexFields.query.pq
// Analyzes complex fields that your connector handles as 'any' type
(parameter) =>
let
    Source = parameter,
    IssuesFolder = Source{[Name="Issues"]}[Data],
    
    // Get first project's issues (assuming we have at least one project)
    FirstProject = Table.FirstN(IssuesFolder, 1),
    IssueData = FirstProject{0}[Data],
    /*ColNames = Table.ColumnNames(IssueData),
    SampleRows = Table.FirstN(IssueData, 5),
    Results = Table.Combine({ColNames, SampleRows})


in
    SampleRows*/


    
    // Check what columns exist after first expansion
    ColumnsFromFirstProject = Table.ColumnNames(FirstProject),
    ColumnsAfterFirstExpansion = Table.ColumnNames(IssueData),
    HasFieldsColumn = List.Contains(ColumnsAfterFirstExpansion, "fields"),
    FirstRowSample = if Table.RowCount(IssueData) > 0 then IssueData{0} else null,
    
    Results = [
        FirstColumns = ColumnsFromFirstProject,
        Columns = ColumnsAfterFirstExpansion,
        HasFieldsColumn = HasFieldsColumn, 
        FirstRowSample = FirstRowSample
    ]
in
    Results
