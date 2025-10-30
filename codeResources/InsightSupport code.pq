/*  Add OAuth Detection:  */
// In jiraNavTbl function, detect auth type
let
    Credentials = try Extension.CurrentCredential() otherwise null,
    
    AuthType = if Credentials <> null then
        if Record.HasFields(Credentials, "access_token") then "OAuth"
        else if Record.HasFields(Credentials, "Key") then "ApiToken"
        else "BasicAuth"
    else "Unknown",
    
    // Only fetch Insights if OAuth
    useInsights = (AuthType = "OAuth"),


/*  Conditionally Add Insights Folder to Navigation:  */
// After existing baseEntries definition:
entriesWithInsights = if useInsights then
    List.Combine({
        baseEntries,
        {
            {
                "Insights (Product Discovery)",
                GetInsightsNavigation(Credentials[access_token], Credentials[cloud_id], requestedProjects),
                "Folder",
                "Folder",
                false
            }
        }
    })
else
    baseEntries,

// Then use entriesWithInsights instead of baseEntries for final navigation




/* Helper Function for Insights Navigation */
GetInsightsNavigation = (accessToken as text, cloudId as text, projectKeys as list) as table =>
    let
        insightsData = jiraInsights.GetInsights(cloudId, accessToken, projectKeys),
        
        // Group by workspace/project for navigation
        groupedByProject = Table.Group(insightsData, {"WorkspaceName", "ProjectName"}, {
            {"Insights", each _, type table}
        }),
        
        // Create navigation entries
        navEntries = Table.TransformRows(groupedByProject, 
            each {
                [WorkspaceName] & " / " & [ProjectName],
                [Insights],
                "Table",
                "Table",
                true
            }
        ),
        
        navTable = Table.FromRows(navEntries, {"Name", "Data", "ItemKind", "ItemName", "IsLeaf"}),
        result = Table.ToNavigationTable(navTable, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")
    in
        result;

/* ###second query### */
// In your jiraNavTbl function, add these sections:

// ... existing code ...

let
    // ... your existing variable definitions ...
    
    // ✅ ADD: Detect OAuth authentication
    Credentials = try Extension.CurrentCredential() otherwise null,
    
    AuthType = if Credentials <> null then
        if Record.HasFields(Credentials, "access_token") and Record.HasFields(Credentials, "cloud_id") then 
            "OAuth"
        else if Record.HasFields(Credentials, "Key") then 
            "ApiToken"
        else 
            "BasicAuth"
    else 
        "Unknown",
    
    useInsights = (AuthType = "OAuth"),
    
    // ... your existing allProjects logic ...
    
    // ✅ ADD: Fetch Insights if OAuth (after getting projects)
    insightsData = if useInsights and testOverride = null then
        try jiraInsights.GetInsightsForProjects(getProjects, Credentials[access_token])
        otherwise #table({}, {}) // Empty table on error
    else
        #table({}, {}),
    
    // ... your existing issues logic ...
    
    // ✅ MODIFY: Add Insights to navigation entries
    insightsEntry = if useInsights and Table.RowCount(insightsData) > 0 then
        {
            {
                "Insights (Product Discovery)",
                CreateInsightsNavigationTable(insightsData),
                "Folder",
                "Folder",
                false
            }
        }
    else
        {},
    
    // Combine all navigation entries
    baseEntries = {
        {
            "Issues",
            createIssuesNavigationTable(issuesForNav),
            "Folder",
            "Folder",
            false
        }
    },
    
    allEntries = List.Combine({
        baseEntries,
        insightsEntry,
        // ... your existing validation report logic ...
    }),
    
    source = Table.FromRows(allEntries, {"Name", "Data", "ItemKind", "ItemName", "IsLeaf"}),
    navTable = Table.ToNavigationTable(source, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")
in
    navTable;

// ✅ ADD: Helper function to create Insights navigation
CreateInsightsNavigationTable = (insightsData as table) as table =>
    let
        // Group by project
        groupedByProject = Table.Group(insightsData, {"ProjectKey", "ProjectName"}, {
            {"Insights", each _, type table}
        }),
        
        // Create navigation rows
        navRows = Table.TransformRows(groupedByProject, 
            each {
                [ProjectKey] & " - " & [ProjectName],
                [Insights],
                "Table",
                "Table",
                true
            }
        ),
        
        navTable = Table.FromRows(navRows, {"Name", "Data", "ItemKind", "ItemName", "IsLeaf"}),
        result = Table.ToNavigationTable(navTable, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")
    in
        result;
