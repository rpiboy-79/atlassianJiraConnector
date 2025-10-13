// Test_OrderBy_CreatedDescending.query.pq
(parameter) =>
let
    // INLINE HELPER FUNCTIONS (copy from template)
    testName = "ORDER BY created DESC",
    
    GetIssueData = (connectorResult as any) as table =>
        let
            IssuesFolder = try connectorResult{[Name="Issues"]}[Data] otherwise #table({},{}),
            FirstProject = try Table.FirstN(IssuesFolder, 1) otherwise #table({},{}),
            IssueData = if Table.RowCount(FirstProject) > 0 then 
                try FirstProject{0}[Data] otherwise #table({},{})
            else 
                #table({},{})
        in
            IssueData,
    
    FlexibleColumnCheck = (table as table, possibleNames as list) as text =>
        let
            actualColumns = List.Transform(Table.ColumnNames(table), Text.Lower),
            possibleLower = List.Transform(possibleNames, Text.Lower),
            foundColumn = List.First(List.Intersect({possibleLower, actualColumns}), null),
            actualColumn = if foundColumn <> null then
                let
                    matchIndex = List.PositionOf(actualColumns, foundColumn)
                in
                    if matchIndex >= 0 then Table.ColumnNames(table){matchIndex} else null
            else
                null
        in
            actualColumn,
    
    GetFieldValue = (row as record, fieldName as text) as any =>
        try 
            Record.Field(row, fieldName) 
        otherwise 
            null,
    
    // Expected ORDER BY: created DESC (newest issues first)
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,

    test =
        if HasData and HasColumns then
            try
                let
                    CreatedColumn = FlexibleColumnCheck(IssueData, {"created", "Created", "CREATED"}),
                    
                    // Validate descending order by checking first few vs last few created dates
                    SortingAnalysis = if CreatedColumn <> null and RowCount >= 3 then
                        let
                            CreatedValues = Table.Column(IssueData, CreatedColumn),
                            
                            // Get first 3 and last 3 created dates (as text for comparison)
                            FirstThree = List.FirstN(CreatedValues, 3),
                            LastThree = List.LastN(CreatedValues, 3), 
                            
                            // Convert to comparable date values where possible
                            FirstThreeDates = List.Transform(FirstThree, each 
                                try DateTimeZone.FromText(Text.From(_)) otherwise null),
                            LastThreeDates = List.Transform(LastThree, each
                                try DateTimeZone.FromText(Text.From(_)) otherwise null),
                            
                            // Remove nulls for comparison
                            ValidFirstDates = List.RemoveNulls(FirstThreeDates),
                            ValidLastDates = List.RemoveNulls(LastThreeDates),
                            
                            // Check if first dates are >= last dates (descending order)
                            IsDescending = if List.Count(ValidFirstDates) > 0 and List.Count(ValidLastDates) > 0 then
                                List.First(ValidFirstDates) >= List.Last(ValidLastDates)
                            else
                                null, // Can't determine sorting
                            
                            // Get earliest and latest for display
                            EarliestDate = if List.Count(ValidFirstDates) > 0 then List.Min(ValidFirstDates) else null,
                            LatestDate = if List.Count(ValidFirstDates) > 0 then List.Max(ValidFirstDates) else null
                        in
                            [
                                HasCreatedColumn = true,
                                CanAnalyzeSorting = IsDescending <> null,
                                IsDescendingSorted = IsDescending,
                                TotalIssues = RowCount,
                                EarliestDate = EarliestDate,
                                LatestDate = LatestDate,
                                FirstThreeCreated = FirstThree,
                                LastThreeCreated = LastThree
                            ]
                    else
                        [
                            HasCreatedColumn = CreatedColumn <> null,
                            CanAnalyzeSorting = false,
                            IsDescendingSorted = false,
                            TotalIssues = RowCount,
                            EarliestDate = null,
                            LatestDate = null,
                            FirstThreeCreated = {},
                            LastThreeCreated = {}
                        ],
                    
                    TestResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "Created Column Present",
                                if SortingAnalysis[HasCreatedColumn] then "Pass" else "Fail",
                                "Created column available for sorting",
                                if CreatedColumn <> null then "Created column: " & CreatedColumn else "Created column missing",
                                "Validates that created field is returned by API for ORDER BY created DESC"
                            },
                            {
                                "Sufficient Data for Sort Test",
                                if RowCount >= 3 then "Pass" else "Warn",
                                "At least 3 issues for sorting validation",
                                Text.From(RowCount) & " issues returned",
                                "Need multiple issues to validate descending sort order"
                            },
                            {
                                "Descending Sort Validation", 
                                if SortingAnalysis[IsDescendingSorted] = true then "Pass" 
                                else if SortingAnalysis[IsDescendingSorted] = false then "Fail" 
                                else "Info",
                                "Issues sorted by created date descending (newest first)",
                                if SortingAnalysis[IsDescendingSorted] = true then "Confirmed descending order"
                                else if SortingAnalysis[IsDescendingSorted] = false then "Not in descending order"
                                else "Cannot determine sort order",
                                "Validates ORDER BY created DESC is working at API level"
                            },
                            {
                                "Date Range Analysis",
                                "Info",
                                "Range of created dates in result set",
                                if SortingAnalysis[EarliestDate] <> null then
                                    "From: " & DateTimeZone.ToText(SortingAnalysis[EarliestDate]) & " To: " & DateTimeZone.ToText(SortingAnalysis[LatestDate])
                                else "Date analysis not available",
                                "Shows the date range of issues returned with ORDER BY created DESC"
                            },
                            {
                                "Sample Created Dates",
                                "Info", 
                                "First 3 created dates (should be newest)",
                                Text.Combine(List.Transform(SortingAnalysis[FirstThreeCreated], each Text.From(_)), ", "),
                                "Shows actual created dates to manually verify descending order"
                            }
                        }
                    )
                in
                    TestResults
            otherwise
                #table(
                    {"Test", "Result", "Expected", "Actual", "Details"},
                    {
                        {
                            testName,
                            "Error",
                            "Successful created DESC sorting validation",
                            "An error occurred during analysis",
                            "Error in test execution - check ORDER BY created DESC implementation"
                        }
                   }
                )
        else
            #table(
                {"Test", "Result", "Expected", "Actual", "Details"},
                {
                    {
                        testName,
                        "Fail",
                        "Issues sorted by created date descending",
                        "No data returned from API",
                        "ORDER BY created DESC may have failed or filtered out all results"
                    }
                }
            )
in
    test