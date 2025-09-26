// Test_MockData_VolumeHandling.query.pq
// Tests connector behavior with larger data sets
(parameter) =>
let
    Source = parameter,
    
    
    // Perform typical operations
    RowCount = Table.RowCount(Source),
    DistinctKeys = Table.RowCount(Table.Distinct(Table.SelectColumns(Source, {"key"}))),
    GroupedByType = Table.Group(Source, {"issuetype"}, {{"Count", Table.RowCount, type number}}),
    TypeCount = Table.RowCount(GroupedByType),
    // Test basic operations on the full dataset
    StartTime = DateTimeZone.UtcNow(),
    TimingTest = Table.RowCount(Table.SelectRows(Source, each [id] <> null)),
    EndTime = DateTimeZone.UtcNow(),
    
    DurationInSeconds = Duration.TotalSeconds(EndTime - StartTime),
    DurationInMs = DurationInSeconds * 1000,
    DurationOk = DurationInMs < 5000,
    
    Results = #table(
        {"Test", "Result", "Expected", "Actual", "Details"},
        {
            {
                "Data Volume",
                if RowCount = 10 then "Pass" else "Fail",
                "10 rows",
                Text.From(RowCount) & " rows",
                "Should handle expected volume"
            },
            {
                "Key Uniqueness",
                if DistinctKeys = RowCount then "Pass" else "Fail",
                "All unique",
                Text.From(DistinctKeys) & "/" & Text.From(RowCount),
                "All keys should be unique"
            },
            {"Grouping Operations", if TypeCount > 0 then "Pass" else "Fail", "> 0 groups", Text.From(TypeCount) & " issue types", "Should support grouping operations"},
            //{"Processing Time", if DurationInMs < 5000 then "Pass" else "Fail", "< 5 seconds", Text.From(DurationInMs) & "ms", "Should process data efficiently"}
            {
                "Calculated Duration",
                if DurationOk then "Pass" else "Fail",
                "value greater less than 5000",
                if DurationOk then "value greater than 0" else Text.From(DurationInSeconds),
                "Duration to count all rows"
            }
        }
    )
in
    Results
