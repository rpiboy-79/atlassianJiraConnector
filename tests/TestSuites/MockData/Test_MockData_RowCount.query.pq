// Test_MockData_RowCount.query.pq
// Validates expected data volume and basic structure
(parameter) =>
let
    Source = parameter,
    RowCount = Table.RowCount(Source),
    ColumnCount = Table.ColumnCount(Source),

    SourceType = Value.Type(Source),
    IsTable = Source is table,
    TypeDescription = if IsTable then "table" else "not table",
    
    // Validate we get expected data volume (should be 10 based on your parameter query)
    Results = #table(
        {"Test", "Result", "Expected", "Actual", "Details"},
        {
            {"Row Count", if RowCount = 10 then "Pass" else "Fail", "10", Text.From(RowCount), "Mock data should have exactly 10 rows"},
            {"Has Data", if RowCount > 0 then "Pass" else "Fail", "> 0", Text.From(RowCount), "Table must contain data"},
            {"Column Count", if ColumnCount >= 7 then "Pass" else "Fail", ">= 7", Text.From(ColumnCount), "Should have at least 7 default fields"},
            {"Table Structure", if IsTable then "Pass" else "Fail", "table", TypeDescription, "Source must be a table"}
        }
    )
in
    Results
