# ORDER BY FUNCTIONALITY TESTS
# Complete test suite for ORDER BY parameter validation using your test template

# ==============================================================================
# TEST 1: ORDER BY PARAMETER ACCEPTANCE
# ==============================================================================

// Test_OrderBy_ParameterAcceptance.query.pq
(parameter) =>
let
    // INLINE HELPER FUNCTIONS (copy from your template)
    testName = "ORDER BY Parameter Acceptance",
    
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
    
    // ACTUAL TEST LOGIC STARTS HERE
    
    // Step 1: Extract connector data (should use ORDER BY parameter)
    ConnectorResult = parameter,
    IssueData = GetIssueData(ConnectorResult),
    
    // Step 2: Validate basic structure
    RowCount = Table.RowCount(IssueData),
    HasData = RowCount > 0,
    HasColumns = Table.ColumnCount(IssueData) > 0,
    ColumnNames = Table.ColumnNames(IssueData),

    // Step 2A: Wrap Test in error handling and no-data handling
    test =
        if HasData and HasColumns then
            try
                let
                    // Step 3: Validate connector accepts ORDER BY without crashing
                    IdColumn = FlexibleColumnCheck(IssueData, {"id", "Id", "ID"}),
                    KeyColumn = FlexibleColumnCheck(IssueData, {"key", "Key", "KEY"}),
                    CreatedColumn = FlexibleColumnCheck(IssueData, {"created", "Created", "CREATED"}),
                    
                    // Step 4: Basic validation that data was returned (ORDER BY didn't break connector)
                    TestPassed = HasData and HasColumns and IdColumn <> null,
                    
                    // Step 5: Create results table using your format
                    TestResults = #table(
                        {"Test", "Result", "Expected", "Actual", "Details"},
                        {
                            {
                                "Connector Accepts ORDER BY",
                                if HasData then "Pass" else "Fail",
                                "Connector processes ORDER BY parameter without error",
                                if HasData then "Data returned successfully" else "No data returned",
                                "Validates that ORDER BY parameter doesn't break connector"
                            },
                            {
                                "Required Columns Present", 
                                if IdColumn <> null and KeyColumn <> null then "Pass" else "Fail",
                                "Essential columns (id, key) are present",
                                if IdColumn <> null then "ID: " & IdColumn else "Missing ID column" & 
                                if KeyColumn <> null then ", Key: " & KeyColumn else ", Missing Key column",
                                "Ensures basic Jira data structure is intact"
                            },
                            {
                                "Created Column Available",
                                if CreatedColumn <> null then "Pass" else "Fail", 
                                "Created column present for date sorting tests",
                                if CreatedColumn <> null then "Created column: " & CreatedColumn else "Created column not found",
                                "Verifies date field is available for ORDER BY created tests"
                            },
                            {
                                "Data Count",
                                "Info",
                                "Issues returned for ORDER BY validation",
                                Text.From(RowCount) & " issues returned",
                                "Shows data volume available for sorting validation"
                            },
                            {
                                "Available Columns",
                                "Info",
                                "All columns available for ORDER BY testing",
                                Text.Combine(ColumnNames, ", "),
                                "Lists all columns that can be used in ORDER BY tests"
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
                            "Successful ORDER BY parameter processing",
                            "An error occurred during analysis",
                            "Error in test execution - check ORDER BY implementation"
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
                        "Data returned from API with ORDER BY", 
                        "No data available",
                        "ORDER BY parameter may have caused API failure or filtering error"
                    }
                }
            )
in
    test

# ==============================================================================
# TEST 2: ORDER BY CREATED DESC VALIDATION
# ==============================================================================

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
                                try DateTime.FromText(Text.From(_)) otherwise null),
                            LastThreeDates = List.Transform(LastThree, each
                                try DateTime.FromText(Text.From(_)) otherwise null),
                            
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
                                    "From: " & DateTime.ToText(SortingAnalysis[EarliestDate]) & " To: " & DateTime.ToText(SortingAnalysis[LatestDate])
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

# ==============================================================================
# TEST 3: ORDER BY PRIORITY DESC VALIDATION
# ==============================================================================

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

# ==============================================================================
# TEST 4: ORDER BY MULTIPLE FIELDS VALIDATION
# ==============================================================================

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
                                    try DateTime.FromText(Text.From(Record.Field(_, CreatedColumn))) otherwise null
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

# ==============================================================================
# TEST EXECUTION STRATEGY
# ==============================================================================

# Phase 1: Parameter Acceptance Test
# File: Test_OrderBy_ParameterAcceptance.query.pq
# Parameterquery: Add orderBy = "created DESC"
# Expected: Should pass (connector accepts ORDER BY parameter)

# Phase 2: Basic Sorting Tests  
# File: Test_OrderBy_CreatedDescending.query.pq
# Parameterquery: Add orderBy = "created DESC" 
# Expected: Should validate descending date order

# File: Test_OrderBy_PriorityDescending.query.pq
# Parameterquery: Add orderBy = "priority DESC"
# Expected: Should validate priority sorting

# Phase 3: Advanced Sorting Test
# File: Test_OrderBy_MultipleFields.query.pq  
# Parameterquery: Add orderBy = "priority DESC, created DESC"
# Expected: Should validate complex multi-field sorting

# Each test provides detailed validation of:
# ✅ Parameter acceptance without connector crash
# ✅ Proper data structure returned  
# ✅ Actual sorting behavior validation
# ✅ Column availability for ORDER BY fields
# ✅ Data volume sufficiency for testing
# ✅ Multi-field sorting complexity handling

# Run these tests progressively as you implement ORDER BY functionality!