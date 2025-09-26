// This file contains your Data Connector logic
[Version = "1.0.2"]
section atlassianJiraConnector;

[DataSource.Kind="Atlassian_Jira_Connector.Kind", Publish="Atlassian_Jira_Connector.Publish"]
shared Jira = (
    #"Company URL Identifier (as specified in your Jira URL)" as text,
    optional #"Field List (comma separated, e.g., summary,created)" as text,
    optional #"Project Keys (comma separated, e.g., PROJ1,PROJ2)" as text,
    optional #"Max Number of Results (default: 1000)" as text,
    optional TestDataOverride as nullable record
) => jiraNavTbl(
    #"Company URL Identifier (as specified in your Jira URL)",
    #"Field List (comma separated, e.g., summary,created)",
    #"Project Keys (comma separated, e.g., PROJ1,PROJ2)",
    #"Max Number of Results (default: 1000)",
    TestDataOverride
);

// Data Source Kind description
Atlassian_Jira_Connector.Kind = [
    Authentication = [
        Key = [],
        UsernamePassword = []
        // Windows = [],
        // Anonymous = []
    ],
    TestConnection = (dataSourcePath as text) as list =>
        { dataSourcePath }
];

atlassianJiraConnector.Icons = [
    Icon16 = { Extension.Contents("atlassianJiraConnector16.png"), Extension.Contents("atlassianJiraConnector20.png"), Extension.Contents("atlassianJiraConnector24.png"), Extension.Contents("atlassianJiraConnector32.png") },
    Icon32 = { Extension.Contents("atlassianJiraConnector32.png"), Extension.Contents("atlassianJiraConnector40.png"), Extension.Contents("atlassianJiraConnector48.png"), Extension.Contents("atlassianJiraConnector64.png") }
];

// Data Source UI publishing description
Atlassian_Jira_Connector.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
    LearnMoreUrl = "https://github.com/rpiboy-79/atlassianJiraConnector",
    SourceImage = atlassianJiraConnector.Icons,
    SourceTypeImage = atlassianJiraConnector.Icons
];

// TEST DATA OVERRIDE FUNCTIONS (Optimized for your JSON structure)
// Function to extract issues from test data
ExtractTestIssuesData = (testData as record) as list =>
    let
        // Your JSON structure has "issues" as the main array
        issuesArray = if Record.HasFields(testData, "issues") then 
            testData[issues]
        else 
            // Fallback for other possible structures
            if Record.HasFields(testData, "values") then testData[values]
            else if Record.HasFields(testData, "results") then testData[results]
            else if Value.Is(testData, type list) then testData
            else {}
    in
        issuesArray;

// Function to create mock projects from test issues (Optimized for your data)
CreateMockProjects = (testIssues as list) as table =>
    let
        // Extract project data from the first few issues to get project info
        projectRecords = List.Transform(
            List.FirstN(testIssues, 10), // Just sample first 10 issues
            each 
                if Record.HasFields(_, "fields") and Record.HasFields([fields], "project") then
                    [fields][project]
                else null
        ),
        
        // Remove nulls and get unique projects
        validProjects = List.RemoveNulls(projectRecords),
        uniqueProjects = List.Distinct(validProjects, each [key]),
        
        // Convert to table - your test data already has proper project structure
        projectTable = Table.FromRecords(uniqueProjects)
    in
        if Table.RowCount(projectTable) > 0 then 
            projectTable 
        else 
            // Fallback mock project if extraction fails
            #table(
                {"key", "name", "id", "projectTypeKey"}, 
                {{"PB", "Product Bugs", "10000", "software"}}
            );


// Discover API limits for maxResults
DiscoverMaxResults = (url as text, jqlQuery as text, optional testOverride as nullable record) =>
    let
        result = if testOverride <> null then
            1000 //return default for test mode
        else
            let
                testMaxResults = "5000",
                queryHeaders = [Accept = "application/json"],
                queryParameters = [
                    jql = jqlQuery,
                    fields = "id,key",
                    maxResults = testMaxResults,
                    startAt = "0"
                ],
                source = Web.Contents(
                    url, [
                        RelativePath = "rest/api/3/search/jql",
                        Headers = queryHeaders,
                        Query = queryParameters
                    ]
                ),
                jsonDoc = Json.Document(source),
                actualMaxResults = jsonDoc[maxResults]
            in
                actualMaxResults
    in
        result;


//Query Fields List
issuesFieldsList = {
    "id",
    "key",
    "summary",
    "description",
    "issuetype",
    "created",
    "components"
    };

// Known Jira field schema definitions
JiraKnownFields = #table({"Name", "Type"}, {
    {"id", type text},
    {"key", type text},
    {"self", type text},
    {"summary", type text},
    {"created", type datetimezone},
    {"updated", type datetimezone}
});

// Phase 1: Simple data retrieval (works with API tokens)
// Consider adding filtering to allow manaul pagination by user in PowerBI.
jiraDataRetrievalSimple = (
    url as text,
    queryPath as text,
    actualJql as nullable text,
    fieldsText as nullable text,
    maxResultsText as text,
    optional testOverride as nullable record
) =>
    let
        dataArray = if testOverride <> null then
            // TEST MODE: Use provided test data
            ExtractTestIssuesData(testOverride)
        else
            let
                // Build query parameters (only include non-null values)
                baseParams = [
                    jql = actualJql,
                    fields = fieldsText,
                    maxResults = maxResultsText
                ],
                queryParams = Record.RemoveFields(baseParams, 
                    List.Select(Record.FieldNames(baseParams), each Record.Field(baseParams, _) = null)),

                // Single API call
                source = Web.Contents(url, [
                    RelativePath = queryPath,
                    Headers = [Accept = "application/json"],
                    Query = queryParams
                ]),
                
                response = Json.Document(source),
                
                // Extract data array
                extractedArray = if Record.HasFields(response, "issues") then response[issues]
                        else if Record.HasFields(response, "values") then response[values]  
                        else if Record.HasFields(response, "results") then response[results]
                        else if Value.Is(response, type list) then response
                        else {}
            in
                extractedArray
    in
        dataArray;


// Phase 2: Full pagination data retrieval (for OAuth)
jiraDataRetrievalPaginated = (
    url as text,
    queryPath as text,
    actualJql as nullable text,
    fieldsText as nullable text,
    maxResultsText as text,
    maxResultsNumber as number,
    optional testOverride as nullable record
) =>
    let
        result = if testOverride <> null then
            // TEST MODE: Use provided test data
            ExtractTestIssuesData(testOverride)
        else
            let
                // Build query parameters dynamically, excluding null values
                BuildQueryParameters = (startAt as number) =>
                    let
                        baseParams = [
                            startAt = Number.ToText(startAt),
                            maxResults = maxResultsText
                        ],
                        withJql = if actualJql <> null then Record.AddField(baseParams, "jql", actualJql) else baseParams,
                        withFields = if fieldsText <> null then Record.AddField(withJql, "fields", fieldsText) else withJql
                    in
                        withFields,

                GetPage = (startAt as number) =>
                    let
                        queryHeaders = [Accept = "application/json"],
                        queryParameters = BuildQueryParameters(startAt),
                        source = Web.Contents(
                            url, [
                                RelativePath = queryPath,
                                Headers = queryHeaders,
                                Query = queryParameters
                            ]
                        ),
                        jsonDoc = Json.Document(source)
                    in
                        jsonDoc,

                firstResponse = GetPage(0),
                isDirectArrayResponse = Value.Is(firstResponse, type list),

                GetAllPages = if isDirectArrayResponse then
                    {firstResponse}
                else
                    let
                        ExtractDataArray = (pageResponse as record) =>
                            if Record.HasFields(pageResponse, "issues") then pageResponse[issues]
                            else if Record.HasFields(pageResponse, "values") then pageResponse[values]  
                            else if Record.HasFields(pageResponse, "results") then pageResponse[results]
                            else null,
                        
                        isPaginated = Record.HasFields(firstResponse, "isLast") or Record.HasFields(firstResponse, "total")
                    in
                        if isPaginated then
                            List.Generate(
                                () => [Page = firstResponse, StartAt = 0],
                                each [Page] <> null and 
                                    ExtractDataArray([Page]) <> null and
                                    ([Page][isLast]? <> true),
                                each 
                                    let
                                        nextStartAt = [StartAt] + maxResultsNumber
                                    in
                                        [Page = GetPage(nextStartAt), StartAt = nextStartAt],
                                each ExtractDataArray([Page])
                            )
                        else
                            {ExtractDataArray(firstResponse)}
            in
                List.Combine(List.RemoveNulls(GetAllPages))
    in
        result;

// Infer types from field names and sample data
InferJiraFieldType = (fieldName as text, sampleValue as any) as type =>
    let
        lowerFieldName = Text.Lower(fieldName),
        
        // Custom field pattern recognition
        typeFromName = if Text.StartsWith(lowerFieldName, "customfield_") then
            if Text.Contains(lowerFieldName, "date") then type datetimezone
            else if Text.Contains(lowerFieldName, "number") then type number
            else if Text.Contains(lowerFieldName, "user") then type any
            else type any
        else type any,
            
        // Sample value analysis
        typeFromValue = if typeFromName = type any then
            if sampleValue is text then
                if try DateTime.FromText(sampleValue) <> null otherwise false then
                    type datetimezone
                else if try Number.FromText(sampleValue) <> null otherwise false then
                    type number
                else type text
            else if sampleValue is number then type number
            else if sampleValue is list then type any
            else if sampleValue is record then type any
            else if sampleValue is table then type any
            else type any
        else typeFromName
    in
        typeFromValue;


// Apply schema only to fields we can confidently type
ApplyJiraSchema = (tbl as table) as table =>
    let
        tableColumns = Table.ColumnNames(tbl),
        sampleRow = if Table.RowCount(tbl) > 0 then tbl{0} else null,

        // Build schema for existing columns only
        schemaTransforms = List.Transform(tableColumns, (colName) =>
            let
                knownTypeRow = Table.SelectRows(JiraKnownFields, each [Name] = colName),
                knownType = if Table.RowCount(knownTypeRow) > 0 then 
                    knownTypeRow{0}[Type] else null,
                    
                inferredType = if knownType = null and sampleRow <> null then
                    InferJiraFieldType(colName, Record.Field(sampleRow, colName))
                else knownType,
                
                // Only transform if we have confidence in the type
                finalType = if inferredType <> null and
                                inferredType <> type any and
                                inferredType <> type record and     // SKIP generic records
                                inferredType <> type nullable record// SKIP nullable records
                                then
                                inferredType
                            else null
            in
                if finalType <> null then {colName, finalType} else null
        ),
        
        // Remove null entries and apply transformations
        validTransforms = List.RemoveNulls(schemaTransforms),

        //original final step
        result = if List.Count(validTransforms) > 0 then
            Table.TransformColumnTypes(tbl, validTransforms)
        else tbl
    in
        result;

//Universal Dynamic Jira Query Function
jiraQuery = (
    url as text, 
    queryPath as text, 
    optional jqlQuery as nullable text,
    optional fieldList as nullable list,
    optional maxResults as nullable text,
    optional discoverLimits as nullable logical,
    optional useSimpleRetrieval as nullable logical,
    optional testOverride as nullable record
) =>
    let
    // Handle optional parameters with defaults
        actualJql = if jqlQuery = null then null else jqlQuery,
        actualFieldsList = if fieldList = null then null else fieldList,
        actualMaxResults = if maxResults = null then "1000" else maxResults,
        shouldDiscoverLimits = if discoverLimits = null then false else discoverLimits,
        useSimple = if useSimpleRetrieval = null then true else useSimpleRetrieval, // Default to Simple
        fieldsText = if actualFieldsList = null then null else Text.Combine(actualFieldsList, ","),

    // Dynamic limit discovery (only if requested and JQL is provided)
        optimalMaxResults = if shouldDiscoverLimits and actualJql <> null and not useSimple then
            try DiscoverMaxResults(url, actualJql) otherwise Number.FromText(actualMaxResults)
            else if useSimple then 1000 // adjust for Simple
            else Number.FromText(actualMaxResults),
        
        maxResultsText = Number.ToText(optimalMaxResults),
        maxResultsNumber = Number.FromText(maxResultsText),
        
    // DATA RETRIEVAL: Choose between simple or paginated
        CombinedData = if useSimple then
            jiraDataRetrievalSimple(url, queryPath, actualJql, fieldsText, maxResultsText, testOverride)
        else
            jiraDataRetrievalPaginated(url, queryPath, actualJql, fieldsText, maxResultsText, maxResultsNumber, testOverride),

    // TRANSFORMATION LOGIC:
        ConvertToTable = Table.FromList(CombinedData, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        
        // Dynamic field expansion
        ExpandedTable = 
            let
            // First expand the main issue fields (id, key, self, fields, etc.)
                firstRecord = if Table.RowCount(ConvertToTable) > 0 then ConvertToTable{0}[Column1] else null,
                topLevelFields = if firstRecord <> null then Record.FieldNames(firstRecord) else {},
                
                expandedTop = if List.Count(topLevelFields) > 0 then
                    Table.ExpandRecordColumn(ConvertToTable, "Column1", topLevelFields, topLevelFields)
                else
                    ConvertToTable,
                
            // Check if we have a 'fields' column to expand (Jira issue data)
                hasFieldsColumn = List.Contains(Table.ColumnNames(expandedTop), "fields"),
                
                expandedFields = if hasFieldsColumn and actualFieldsList <> null then
                // Expand the nested 'fields' record with the requested field list
                    Table.ExpandRecordColumn(
                        expandedTop, 
                        "fields", 
                        List.RemoveItems(actualFieldsList, {"id", "key"}), // Remove id,key as they're already top-level
                        List.RemoveItems(actualFieldsList, {"id", "key"})
                    )
                else if hasFieldsColumn then
                // Dynamic expansion of fields if we don't have a specific field list
                    let
                        sampleFieldsRecord = if Table.RowCount(expandedTop) > 0 then 
                            expandedTop{0}[fields] else null,
                        fieldsToExpand = if sampleFieldsRecord <> null then 
                            Record.FieldNames(sampleFieldsRecord) else {}
                    in
                        if List.Count(fieldsToExpand) > 0 then
                            Table.ExpandRecordColumn(expandedTop, "fields", fieldsToExpand, fieldsToExpand)
                        else
                            expandedTop
                else
                    expandedTop
            in
                expandedFields,
        FinalTable = ApplyJiraSchema(ExpandedTable)
    in
        FinalTable;

// function to create the Issues sub-navigation table
createIssuesNavigationTable = (projectTables as list) as table =>
    let
        source = Table.FromRows(
            projectTables,
            {"Name", "Data", "ItemKind", "ItemName", "IsLeaf"}
        ),
        navTable = Table.ToNavigationTable(
            source, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf"
        )
    in
        navTable;

// Navigation Table
jiraNavTbl = (
    company as text,
    optional fields as text,
    optional projectKeys as text,
    optional maxResults as text,
    optional testOverride as nullable record
) =>
    let
        fieldsDynamic =
            if fields = null
            then issuesFieldsList
            else
                List.Transform(
                    List.RemoveNulls(List.Transform(
                        Text.Split(fields, ","),
                        each if Text.Trim(_) = "" then null else Text.Trim(Text.Upper(_))
                    )),
                    each _
                ),

        requestedProjects = 
            if projectKeys = null then 
                {} 
            else 
                List.Transform(
                    List.RemoveNulls(List.Transform(
                        Text.Split(projectKeys, ","), 
                        each if Text.Trim(_) = "" then null else Text.Trim(Text.Upper(_))
                    )),
                    each _
                ),
        URL = ValidateUrlScheme("https://" & company & ".atlassian.net"),
            
        allProjects = if testOverride <> null then
            let
                testIssues = ExtractTestIssuesData(testOverride)
            in
                CreateMockProjects(testIssues)
        else
            jiraQuery(URL, "rest/api/3/project", null, null, null, null, false),//uses advanced query since project retrieval is not paginated to begin with

        projectValidation = 
            let
                validProjectKeys = List.Transform(Table.Column(allProjects, "key"), each Text.Upper(_)),
                
                // Separate valid and invalid keys
                validKeys = List.Select(requestedProjects, each List.Contains(validProjectKeys, _)),
                invalidKeys = List.Select(requestedProjects, each not List.Contains(validProjectKeys, _)),
                
                // Use all projects if no filter specified, otherwise use validated subset
                finalKeys = if List.Count(requestedProjects) = 0 then validProjectKeys else validKeys,
                
                filteredProjects = Table.SelectRows(allProjects, 
                    each List.Contains(finalKeys, Text.Upper([key]))
                ),
                
                // Create validation report
                validationReport = #table(
                    {"Metric", "Value", "Details"},
                    {
                        {"Total Available Projects", Table.RowCount(allProjects), ""},
                        {"Requested Projects", List.Count(requestedProjects), Text.Combine(requestedProjects, ", ")},
                        {"Valid Requested", List.Count(validKeys), Text.Combine(validKeys, ", ")},
                        {"Invalid Requested", List.Count(invalidKeys), Text.Combine(invalidKeys, ", ")},
                        {"Final Project Count", Table.RowCount(filteredProjects), ""}
                    }
                )
            in
                [
                    Projects = filteredProjects,
                    ValidKeys = validKeys,
                    InvalidKeys = invalidKeys,
                    Report = validationReport,
                    HasInvalidKeys = List.Count(invalidKeys) > 0
                ],
        
        getProjects = projectValidation[Projects],

        issues =
            let
                getIssues = Table.AddColumn(getProjects, "Issues", 
                    each jiraQuery(
                        URL, 
                        "rest/api/3/search/jql",
                        "project=" & [key],
                        fieldsDynamic,
                        maxResults,
                        false, // Don't discover limits in Phase 1
                        true,   // Use simple retrieval
                        testOverride // PASS TEST OVERRIDE
                    )
                ),
                removeErrors = Table.RemoveRowsWithErrors(getIssues, {"Issues"})
            in
                removeErrors,
        issuesForNav = Table.TransformRows(issues, 
                    each {
                        [name],
                        [Issues],                   // Data: The actual issues table
                        "Table",
                        "Table",
                        true
                    }
                ),
        
        
        baseEntries = {
            {
                "Issues",
                createIssuesNavigationTable(issuesForNav),
                "Folder",
                "Folder",
                false
            }
        },

        entriesWithValidation =
            if projectValidation[HasInvalidKeys] or List.Count(requestedProjects) > 0 then
                List.Combine({
                    baseEntries,
                    {
                        {
                            "Project Filter Report",
                            projectValidation[Report],
                            "Table",
                            "Table",
                            true
                        }
                    }
                })
            else
                baseEntries,
        
        source = Table.FromRows(
            entriesWithValidation,
            {"Name", "Data", "ItemKind", "ItemName", "IsLeaf"}
        ),
        navTable = Table.ToNavigationTable(
            source, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf"
        )
    in
        navTable;

// MSFT reccomended helper functions
// Navigation Tables Function from Microsoft docs:
Table.ToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(table),
        newTableType = Type.AddTableKey(tableType, keyColumns, true) meta 
        [
            NavigationTable.NameColumn = nameColumn, 
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn, 
            Preview.DelayColumn = itemNameColumn, 
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(table, newTableType)
    in
        navigationTable;

// URL Validation function, required for certifcation per MSFT docs:
ValidateUrlScheme = (url as text) as text =>
    if (Uri.Parts(url)[Scheme] <> "https") then error "Url scheme must be HTTPS" else url;