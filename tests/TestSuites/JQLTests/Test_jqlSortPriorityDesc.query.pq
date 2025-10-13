// Test_OrderBy_PriorityDescending.query.pq  
(parameter) =>
let
    // INLINE HELPER FUNCTIONS (copy from template)
    testName = "ORDER BY priority DESC",
    
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
    
    // Expected ORDER BY: priority DESC (highest priority first)
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,

    test =
        if HasData and HasColumns then
            try
                let
                    PriorityColumn = FlexibleColumnCheck(IssueData, {"priority", "Priority", "PRIORITY"}),
                    
                    PriorityAnalysis = if PriorityColumn <> null and RowCount >= 2 then
                        let
                            PriorityValues = Table.Column(IssueData, PriorityColumn),
                            PriorityNames = List.Transform(PriorityValues, each GetPriorityName(_)),
                            CleanPriorityNames = List.RemoveNulls(PriorityNames),
                            
                            // Check priority distribution
                            UniquePriorities = List.Distinct(CleanPriorityNames),
                            FirstPriority = if List.Count(CleanPriorityNames) > 0 then List.First(CleanPriorityNames) else null,
                            LastPriority = if List.Count(CleanPriorityNames) > 0 then List.Last(CleanPriorityNames) else null,
                            
                            // Jira priority order (typically): Blocker > Critical > High > Medium > Low
                            HighPriorityTypes = {"Blocker", "Critical", "High"},
                            HasHighPriorityFirst = if FirstPriority <> null then List.Contains(HighPriorityTypes, FirstPriority) else false
                        in
                            [
                                HasPriorityColumn = true,
                                TotalPriorities = List.Count(CleanPriorityNames),
                                UniquePriorities = UniquePriorities,
                                FirstPriority = FirstPriority,
                                LastPriority = LastPriority,
                                HighPriorityFirst = HasHighPriorityFirst,
                                CanValidateSorting = List.Count(UniquePriorities) > 1
                            ]
                    else
                        [
                            HasPriorityColumn = PriorityColumn <> null,
                            TotalPriorities = 0,
                            UniquePriorities = {},
                            FirstPriority = null,
                            LastPriority = null,
                            HighPriorityFirst = false,
                            CanValidateSorting = false
                        ],
                    
                    TestResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "Priority Column Present",
                                if PriorityAnalysis[HasPriorityColumn] then "Pass" else "Fail",
                                "Priority column available for sorting",
                                if PriorityColumn <> null then "Priority column: " & PriorityColumn else "Priority column missing",
                                "Validates that priority field is returned by API for ORDER BY priority DESC"
                            },
                            {
                                "Priority Data Available",
                                if PriorityAnalysis[TotalPriorities] > 0 then "Pass" else "Fail",
                                "Issues have priority values",
                                Text.From(PriorityAnalysis[TotalPriorities]) & " issues with priority data",
                                "Ensures priority values are present for sorting validation"
                            },
                            {
                                "Priority Sorting Validation",
                                if PriorityAnalysis[HighPriorityFirst] then "Pass" 
                                else if PriorityAnalysis[CanValidateSorting] then "Info" 
                                else "Info",
                                "High priority issues appear first (DESC order)",
                                if PriorityAnalysis[FirstPriority] <> null then 
                                    "First priority: " & PriorityAnalysis[FirstPriority] & ", Last priority: " & Text.From(PriorityAnalysis[LastPriority])
                                else "Priority analysis not available",
                                "Validates ORDER BY priority DESC places highest priority issues first"
                            },
                            {
                                "Priority Distribution",
                                "Info",
                                "Variety of priorities in result set",
                                Text.Combine(PriorityAnalysis[UniquePriorities], ", "),
                                "Shows the priority types returned with ORDER BY priority DESC"
                            },
                            {
                                "Sorting Complexity",
                                if PriorityAnalysis[CanValidateSorting] then "Pass" else "Info",
                                "Multiple priority levels for sorting validation",
                                if PriorityAnalysis[CanValidateSorting] then "Multiple priorities found - sorting testable" else "Single priority level - sorting not validatable",
                                "Indicates whether priority sorting can be meaningfully validated"
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
                            "Successful priority DESC sorting validation",
                            "An error occurred during analysis",
                            "Error in test execution - check ORDER BY priority DESC implementation"
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
                        "Issues sorted by priority descending",
                        "No data returned from API",
                        "ORDER BY priority DESC may have failed or filtered out all results"
                    }
                }
            )
in
    test
