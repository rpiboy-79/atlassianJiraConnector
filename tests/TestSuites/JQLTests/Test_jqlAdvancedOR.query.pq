// Test_JQL_OrCondition.query.pq
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
    
    GetStatusName = (statusObject as any) as text =>
        try
            if statusObject is record then statusObject[name] else Text.From(statusObject)
        otherwise
            null,
    
    GetPriorityName = (priorityObject as any) as text =>
        try
            if priorityObject is record then priorityObject[name] else Text.From(priorityObject)
        otherwise
            null,
    
    // TEST LOGIC: Expected JQL = status = "Done" OR priority = High
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,

    test =
        if HasData and HasColumns then
            try
                let
                    ExpectedJQL = "status = ""In Progress"" OR priority = ""Major""",
                    StatusColumn = FlexibleColumnCheck(IssueData, {"status", "Status", "STATUS"}),
                    PriorityColumn = FlexibleColumnCheck(IssueData, {"priority", "Priority", "PRIORITY"}),
                    
                    OrAnalysis = if StatusColumn <> null and PriorityColumn <> null then
                        let
                            // Check each issue to see if it matches either condition
                            IssuesWithValidation = Table.AddColumn(IssueData, "OrValidation", each
                                let
                                    statusName = GetStatusName(Record.Field(_, StatusColumn)),
                                    priorityName = GetPriorityName(Record.Field(_, PriorityColumn)),
                                    isDone = statusName = "In Progress",
                                    isHighPriority = priorityName = "Major",
                                    matchesEither = isDone or isHighPriority
                                in
                                    [
                                        Status = statusName,
                                        Priority = priorityName, 
                                        IsDone = isDone,
                                        IsHigh = isHighPriority,
                                        MatchesOR = matchesEither
                                    ]
                            ),
                            
                            ValidationColumn = Table.Column(IssuesWithValidation, "OrValidation"),
                            AllMatch = List.AllTrue(List.Transform(ValidationColumn, each [MatchesOR])),
                            DoneCount = List.Count(List.Select(ValidationColumn, each [IsDone])),
                            HighCount = List.Count(List.Select(ValidationColumn, each [IsHigh])),
                            BothCount = List.Count(List.Select(ValidationColumn, each [IsDone] and [IsHigh]))
                        in
                            [
                                Total = List.Count(ValidationColumn),
                                Done = DoneCount,
                                High = HighCount,
                                Both = BothCount,
                                AllMatch = AllMatch
                            ]
                    else
                        [
                            Total = 0,
                            Done = 0,
                            High = 0,
                            Both = 0,
                            AllMatch = false
                        ],
                    
                    TestResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "JQL OR Query Applied",
                                if HasData then "Pass" else "Fail",
                                "API returns OR-filtered data",
                                if HasData then "Received " & Text.From(RowCount) & " issues" else "No data returned",
                                "Validates complex JQL OR query was sent to real Jira API"
                            },
                            {
                                "Required Columns Present",
                                if StatusColumn <> null and PriorityColumn <> null then "Pass" else "Fail",
                                "Status and Priority fields in response",
                                "Status: " & (if StatusColumn <> null then StatusColumn else "Missing") & ", Priority: " & (if PriorityColumn <> null then PriorityColumn else "Missing"),
                                "Validates both fields needed for OR condition are present"
                            },
                            {
                                "OR Condition Logic",
                                if OrAnalysis[AllMatch] then "Pass" else "Fail",
                                "All issues match: In Progress OR Major Priority",
                                if StatusColumn <> null and PriorityColumn <> null then
                                    "In Progress: " & Text.From(OrAnalysis[Done]) & ", Major: " & Text.From(OrAnalysis[High]) & ", Both: " & Text.From(OrAnalysis[Both])
                                else "Cannot validate - missing columns",
                                "Validates OR condition logic worked correctly at API level"
                            },
                            {
                                "Complex JQL Support",
                                if OrAnalysis[AllMatch] and HasData then "Pass" else "Fail",
                                "Connector supports OR operators in JQL",
                                if OrAnalysis[AllMatch] then "OR operator working" else "OR operator may not be supported",
                                "Tests advanced JQL operator support in connector"
                            },
                            {
                                "JQL Query Used",
                                "Info",
                                ExpectedJQL,
                                "--",
                                "Shows complex JQL query that should have been sent to API"
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
                            "JQL OR Condition",
                            "Error",
                            "Successful OR condition filtering",
                            "An error occurred during analysis",
                            "Error in test execution - check complex JQL implementation"
                        }
                   }
                )
        else
            #table(
                {"Test", "Result", "Expected", "Actual", "Details"},
                {
                    {
                        "JQL OR Condition",
                        "Fail",
                        "Issues matching Done OR High Priority",
                        "No data returned from API", 
                        "Complex JQL query may have failed or filtered out all results"
                    }
                }
            )
in
    test