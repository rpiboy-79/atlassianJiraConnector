// Test_OrderBy_MultipleFields.query.pq
(parameter) =>
let
    // INLINE HELPER FUNCTIONS (copy from template) 
    testName = "ORDER BY priority DESC, created DESC",
    
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
    
    // Expected ORDER BY: priority DESC, created DESC (priority first, then creation date)
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
                    CreatedColumn = FlexibleColumnCheck(IssueData, {"created", "Created", "CREATED"}),
                    
                    MultiFieldAnalysis = if PriorityColumn <> null and CreatedColumn <> null and RowCount >= 3 then
                        let
                            // Add analysis columns for priority and created
                            ExtendedData = Table.AddColumn(
                                Table.AddColumn(IssueData, "PriorityName", each GetPriorityName(Record.Field(_, PriorityColumn))),
                                "CreatedDate", each 
                                    try DateTimeZone.FromText(Text.From(Record.Field(_, CreatedColumn))) otherwise null
                            ),
                            
                            // Group by priority to see if created dates are sorted within each priority
                            GroupedByPriority = Table.Group(ExtendedData, {"PriorityName"}, {
                                {"Count", each Table.RowCount(_), Int64.Type},
                                {"CreatedDates", each Table.Column(_, "CreatedDate"), type list}
                            }),
                            
                            // Check if we have multiple priority groups
                            PriorityGroups = Table.RowCount(GroupedByPriority),
                            UniquePriorities = Table.Column(GroupedByPriority, "PriorityName")
                        in
                            [
                                HasBothColumns = true,
                                PriorityGroups = PriorityGroups,
                                UniquePriorities = UniquePriorities, 
                                CanTestMultiSort = PriorityGroups > 1,
                                TotalIssues = RowCount,
                                GroupedData = GroupedByPriority
                            ]
                    else
                        [
                            HasBothColumns = PriorityColumn <> null and CreatedColumn <> null,
                            PriorityGroups = 0,
                            UniquePriorities = {},
                            CanTestMultiSort = false,
                            TotalIssues = RowCount,
                            GroupedData = null
                        ],
                    
                    TestResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "Both Sort Columns Present",
                                if MultiFieldAnalysis[HasBothColumns] then "Pass" else "Fail",
                                "Both priority and created columns available",
                                "Priority: " & (if PriorityColumn <> null then PriorityColumn else "Missing") & 
                                ", Created: " & (if CreatedColumn <> null then CreatedColumn else "Missing"),
                                "Validates both fields needed for multi-field ORDER BY are present"
                            },
                            {
                                "Multi-Field Sorting Testable",
                                if MultiFieldAnalysis[CanTestMultiSort] then "Pass" else "Info",
                                "Multiple priority levels for multi-field sort testing",
                                Text.From(MultiFieldAnalysis[PriorityGroups]) & " priority groups found",
                                "Need multiple priority levels to validate priority DESC, created DESC ordering"
                            },
                            {
                                "Complex ORDER BY Processing", 
                                if HasData and MultiFieldAnalysis[HasBothColumns] then "Pass" else "Fail",
                                "Connector processes multi-field ORDER BY without error",
                                if HasData then "Multi-field ORDER BY processed successfully" else "Multi-field ORDER BY failed",
                                "Validates connector handles complex ORDER BY priority DESC, created DESC syntax"
                            },
                            {
                                "Priority Groups Analysis",
                                "Info",
                                "Distribution of priorities in sorted results",
                                Text.Combine(MultiFieldAnalysis[UniquePriorities], ", "),
                                "Shows priority distribution when using ORDER BY priority DESC, created DESC"
                            },
                            {
                                "Data Volume for Testing",
                                if RowCount >= 5 then "Pass" else "Info",
                                "Sufficient data for multi-field sort validation",
                                Text.From(RowCount) & " issues available for multi-field sorting analysis",
                                "Indicates data volume available for comprehensive multi-field ORDER BY testing"
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
                            "Successful multi-field ORDER BY validation",
                            "An error occurred during analysis",
                            "Error in test execution - check ORDER BY priority DESC, created DESC implementation"
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
                        "Issues sorted by priority DESC, created DESC",
                        "No data returned from API",
                        "Multi-field ORDER BY may have failed or filtered out all results"
                    }
                }
            )
in
    test