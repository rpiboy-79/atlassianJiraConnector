// Test_MockData_DataProfiling.query.pq
// Profiles data quality and identifies patterns
(parameter) =>
let
    Source = parameter,
    RowCount = Table.RowCount(Source),
    
    // Profile key text fields
    IdProfile = [
        Field = "id",
        UniqueCount = Table.RowCount(Table.Distinct(Table.SelectColumns(Source, {"id"}))),
        NullCount = Table.RowCount(Table.SelectRows(Source, each [id] = null)),
        EmptyCount = Table.RowCount(Table.SelectRows(Source, each [id] = "")),
        Result = if UniqueCount = RowCount and NullCount = 0 then "Pass" else "Fail"
    ],
    
    KeyProfile = [
        Field = "key", 
        UniqueCount = Table.RowCount(Table.Distinct(Table.SelectColumns(Source, {"key"}))),
        NullCount = Table.RowCount(Table.SelectRows(Source, each [key] = null)),
        HasValidFormat = Table.RowCount(Table.SelectRows(Source, each Text.Contains([key], "-"))),
        Result = if UniqueCount = RowCount and NullCount = 0 and HasValidFormat = RowCount then "Pass" else "Fail"
    ],
    
    SummaryProfile = [
        Field = "summary",
        NullCount = Table.RowCount(Table.SelectRows(Source, each [summary] = null)),
        EmptyCount = Table.RowCount(Table.SelectRows(Source, each [summary] = "")),
        AvgLength = Number.Round(List.Average(List.Transform(Table.Column(Source, "summary"), each Text.Length(_ ?? ""))), 2),
        Result = if NullCount + EmptyCount < RowCount * 0.1 then "Pass" else "Fail"  // Less than 10% null/empty
    ],
    
    CreatedProfile = [
        Field = "created",
        NullCount = Table.RowCount(Table.SelectRows(Source, each [created] = null)),
        ValidDateCount = Table.RowCount(Table.SelectRows(Source, each try DateTime.From([created]) is datetime otherwise false)),
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
