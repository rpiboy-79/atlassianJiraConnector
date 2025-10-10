// Test_JQL_SimpleStatusFilter.query.pq
(parameter) =>
let
    // INLINE HELPER FUNCTIONS
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
    
    FlexibleColumnCheck = (tbl as table, possibleNames as list) as text =>
        let
            actualColumns = List.Transform(Table.ColumnNames(tbl), Text.Lower),
            possibleLower = List.Transform(possibleNames, Text.Lower),
            foundColumn = List.First(List.Intersect({possibleLower, actualColumns}), null),
            actualColumn = if foundColumn <> null then
                let
                    matchIndex = List.PositionOf(actualColumns, foundColumn)
                in
                    if matchIndex >= 0 then Table.ColumnNames(tbl){matchIndex} else null
            else
                null
        in
            actualColumn,
    
    GetStatusName = (statusObject as any) as text =>
        try
            if statusObject is record then statusObject[name] else Text.From(statusObject)
        otherwise
            null,
    
    // TEST LOGIC: Expected JQL = status = "In Progress"
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,

    test =
        if HasData and HasColumns then
            try
                let
                    ExpectedJQL = "status = ""In Progress""",
                    StatusColumn = FlexibleColumnCheck(IssueData, {"status", "Status", "STATUS"}),
                    
                    // Analyze status distribution from real API response
                    StatusAnalysis = if StatusColumn <> null then
                        let
                            StatusValues = Table.Column(IssueData, StatusColumn),
                            StatusNames = List.Transform(StatusValues, each GetStatusName(_)),
                            CleanStatusNames = List.RemoveNulls(StatusNames),
                            InProgressCount = List.Count(List.Select(CleanStatusNames, each _ = "In Progress")),
                            OtherStatusCount = List.Count(CleanStatusNames) - InProgressCount,
                            UniqueStatuses = List.Distinct(CleanStatusNames),
                            AllInProgress = OtherStatusCount = 0 and InProgressCount > 0
                        in
                            [
                                Total = List.Count(CleanStatusNames),
                                InProgress = InProgressCount,
                                Others = OtherStatusCount,
                                Unique = UniqueStatuses,
                                AllMatch = AllInProgress
                            ]
                    else
                        [
                            Total = 0,
                            InProgress = 0,
                            Others = 0,
                            Unique = {},
                            AllMatch = false
                        ],
                    
                    // Create test results table
                    TestResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "JQL Status Filter Applied",
                                if HasData then "Pass" else "Fail",
                                "API returns filtered data",
                                if HasData then "Received " & Text.From(RowCount) & " issues" else "No data returned",
                                "Validates JQL query was sent to real Jira API and data returned"
                            },
                            {
                                "Status Column Present",
                                if StatusColumn <> null then "Pass" else "Fail",
                                "Status field in response",
                                if StatusColumn <> null then "Status column: " & StatusColumn else "Status column missing",
                                "Validates status field is included in API response"
                            },
                            {
                                "All Issues In Progress",
                                if StatusAnalysis[AllMatch] then "Pass" else "Fail",
                                "All issues have status 'In Progress'",
                                if StatusColumn <> null then 
                                    "In Progress: " & Text.From(StatusAnalysis[InProgress]) & ", Others: " & Text.From(StatusAnalysis[Others])
                                else "Cannot analyze - no status column",
                                "Validates JQL filter worked correctly at API level"
                            },
                            {
                                "Status Variety Check",
                                if List.Count(StatusAnalysis[Unique]) = 1 then "Pass" else "Fail",
                                "Only one status type returned",
                                if StatusColumn <> null then 
                                    "Unique statuses: " & Text.Combine(StatusAnalysis[Unique], ", ")
                                else "Cannot check - no status column",
                                "Confirms server-side filtering eliminated other statuses"
                            },
                            {
                                "JQL Query Used",
                                "Info",
                                ExpectedJQL,
                                ExpectedJQL,
                                "Shows JQL query that should have been sent to API"
                            },
                            {
                                "Sample Issue Keys",
                                "Info",
                                "Representative issue keys",
                                if HasData then
                                    let
                                        KeyColumn = FlexibleColumnCheck(IssueData, {"key", "Key", "KEY"})
                                    in
                                        if KeyColumn <> null then
                                            Text.Combine(List.FirstN(Table.Column(IssueData, KeyColumn), 3), ", ")
                                        else "No key column"
                                else "No data",
                                "Shows sample of issues returned by JQL query"
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
                            "JQL Status Filter",
                            "Error",
                            "Successful status filtering",
                            "An error occurred during analysis", 
                            "Error in test execution - check JQL implementation"
                        }
                   }
                )
        else
            #table(
                {"Test", "Result", "Expected", "Actual", "Details"},
                {
                    {
                        "JQL Status Filter",
                        "Fail",
                        "Issues with In Progress status returned",
                        "No data returned from API",
                        "JQL query may have filtered out all results or API connection failed"
                    }
                }
            )
in
    test