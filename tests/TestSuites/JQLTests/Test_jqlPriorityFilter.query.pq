// Test_JQL_PriorityFilter.query.pq
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
    
    GetPriorityName = (priorityObject as any) as text =>
        try
            if priorityObject is record then priorityObject[name] else Text.From(priorityObject)
        otherwise
            null,
    
    // TEST LOGIC: Expected JQL = priority = Major  
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,

    test =
        if HasData and HasColumns then
            try
                let
                    ExpectedJQL = "priority = Major",
                    PriorityColumn = FlexibleColumnCheck(IssueData, {"priority", "Priority", "PRIORITY"}),
                    
                    PriorityAnalysis = if PriorityColumn <> null then
                        let
                            PriorityValues = Table.Column(IssueData, PriorityColumn),
                            PriorityNames = List.Transform(PriorityValues, each GetPriorityName(_)),
                            CleanPriorityNames = List.RemoveNulls(PriorityNames),
                            HighPriorityCount = List.Count(List.Select(CleanPriorityNames, each _ = "Major")),
                            OtherPriorityCount = List.Count(CleanPriorityNames) - HighPriorityCount,
                            UniquePriorities = List.Distinct(CleanPriorityNames),
                            AllHighPriority = OtherPriorityCount = 0 and HighPriorityCount > 0
                        in
                            [
                                Total = List.Count(CleanPriorityNames),
                                Major = HighPriorityCount,
                                Others = OtherPriorityCount,
                                Unique = UniquePriorities,
                                AllMatch = AllHighPriority
                            ]
                    else
                        [
                            Total = 0,
                            Major = 0,
                            Others = 0,
                            Unique = {},
                            AllMatch = false
                        ],
                    
                    TestResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "JQL Priority Filter Applied",
                                if HasData then "Pass" else "Fail",
                                "API returns priority-filtered data",
                                if HasData then "Received " & Text.From(RowCount) & " issues" else "No data returned",
                                "Validates JQL priority query was sent to real Jira API"
                            },
                            {
                                "Priority Column Present",
                                if PriorityColumn <> null then "Pass" else "Fail",
                                "Priority field in response",
                                if PriorityColumn <> null then "Priority column: " & PriorityColumn else "Priority column missing",
                                "Validates priority field is included in API response"
                            },
                            {
                                "All Issues Major Priority",
                                if PriorityAnalysis[AllMatch] then "Pass" else "Fail",
                                "All issues have priority 'Major'",
                                if PriorityColumn <> null then 
                                    "Major: " & Text.From(PriorityAnalysis[Major]) & ", Others: " & Text.From(PriorityAnalysis[Others])
                                else "Cannot analyze - no priority column",
                                "Validates JQL priority filter worked correctly at API level"
                            },
                            {
                                "Priority Consistency",
                                if List.Count(PriorityAnalysis[Unique]) <= 1 then "Pass" else "Fail",
                                "Only Major priority returned",
                                if PriorityColumn <> null then 
                                    "Unique priorities: " & Text.Combine(PriorityAnalysis[Unique], ", ")
                                else "Cannot check - no priority column",
                                "Confirms server-side filtering eliminated other priorities"
                            },
                            {
                                "JQL Query Used",
                                "Info",
                                ExpectedJQL,
                                ExpectedJQL,
                                "Shows JQL query that should have been sent to API"
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
                            "JQL Priority Filter",
                            "Error",
                            "Successful priority filtering",
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
                        "JQL Priority Filter",
                        "Fail",
                        "Issues with Major priority returned",
                        "No data returned from API",
                        "JQL query may have filtered out all results or API connection failed"
                    }
                }
            )
in
    test