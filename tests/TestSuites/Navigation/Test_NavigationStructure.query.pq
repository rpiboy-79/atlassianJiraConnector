// Test_NavigationStructure.query.pq
// Validates the navigation table structure
(parameter) =>
let
    // Get the root navigation table
    Source = parameter,
    
    // Validate navigation structure
    NavColumns = Table.ColumnNames(Source),
    NavRowCount = Table.RowCount(Source),
    
    // Check for expected navigation entries
    NavNames = Table.Column(Source, "Name"),
    HasIssuesEntry = List.Contains(NavNames, "Issues"),
    
    // Test Issues navigation
    IssuesEntry = if HasIssuesEntry then Source{[Name="Issues"]} else null,
    IssuesData = if IssuesEntry <> null then IssuesEntry[Data] else null,
    IssuesIsTable = IssuesData <> null and IssuesData is table,
    IssuesIsTableType =
        if not IssuesIsTable
        then Value.Type(IssuesData)
        else "has not content",
    
    // Count project entries in Issues navigation
    ProjectCount = if IssuesIsTable then Table.RowCount(IssuesData) else 0,

    NavHasColName = List.Contains(NavColumns, "Name"),
    NavHasColData = List.Contains(NavColumns, "Data"),
    
    // Create test results
    TestResults = #table(
        {"Test", "Result", "Expected", "Actual", "Details"},
        {
            {
                "Navigation Table Exists",
                if Source is table then "Pass" else "Fail",
                "Root navigation is table",
                if Source is table then "Table" else "Not a table: " & Text.From(Value.Type(Source)),
                "Verifies existence of navigation table"
            },
            {
                "Has Name Column",
                if NavHasColName then "Pass" else "Fail",
                "Navigation has Name column",
                if NavHasColName then "Name column present" else "Name column missing: " & Text.Combine(NavColumns, ", "),
                "Validates structure of Navigation table"
            },
            {
                "Has Data Column",
                if NavHasColData then "Pass" else "Fail",
                "Navigation has Data column",
                if NavHasColData then "Data column present" else "Data column missing: " & Text.Combine(NavColumns, ", "),
                "Validates structure of Navigation table"
            },
            {
                "Has Issues Entry",
                if HasIssuesEntry then "Pass" else "Fail",
                "Issues folder present",
                if HasIssuesEntry then "Issues entry found" else "Issues entry missing: " & Text.Combine(NavNames, ", "),
                "Validates peresence of Issues navigation"
            },
            {
                "Issues Data Is Table",
                if IssuesIsTable then "Pass" else "Fail",
                "Issues navigation is valid table",
                if IssuesIsTable then "Issues is table" else "Issues is not a valid table: " & Text.From(IssuesIsTableType),
                "Validates Issues navigation data structure"
            },
            {
                "Project Count",
                if ProjectCount > 0 then "Pass" else "Fail",
                "> 0 projects",
                Text.From(ProjectCount) & " projects found",
                "Counts number of projects in Issues navigation"
            },
            {
                "Available Navigation Entries",
                "Info",
                "List of Entries in top-level navigation",
                Text.Combine(NavNames, ", "),
                "Lists all available entries at the top of the navigation structure"
            }
        }
    )
in
    TestResults
