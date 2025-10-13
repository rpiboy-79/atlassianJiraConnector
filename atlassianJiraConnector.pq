// This file contains your Data Connector logic
[Version = "1.0.5"]
section atlassianJiraConnector;

// Helper function to load .pqm files
Extension.LoadFunction = (fileName as text) =>
  let
      binary = Extension.Contents(fileName),
      asText = Text.FromBinary(binary)
  in
      try
        Expression.Evaluate(asText, #shared)
      catch (e) =>
        error [
            Reason = "Extension.LoadFunction Failure",
            Message.Format = "Loading '#{0}' failed - '#{1}': '#{2}'",
            Message.Parameters = {fileName, e[Reason], e[Message]},
            Detail = [File = fileName, Error = e]
        ];

//Load Module Groups
JiraAPI = Extension.LoadFunction("jiraAPI.pqm");
JiraSchema = Extension.LoadFunction("jiraSchema.pqm");
JiraTestSupport = Extension.LoadFunction("jiraTestSupport.pqm");
helperFunctions = Extension.LoadFunction("helperFunctions.pqm");

// Load individual API functions
buildJqlQuery = JiraAPI[buildJqlQuery];
DiscoverMaxResults = JiraAPI[DiscoverMaxResults];
jiraDataRetrievalSimple = JiraAPI[jiraDataRetrievalSimple];
jiraDataRetrievalPaginated = JiraAPI[jiraDataRetrievalPaginated];

// Load schema functions
issuesFieldsList = JiraSchema[issuesFieldsList];
JiraKnownFields = JiraSchema[JiraKnownFields];
InferJiraFieldType = JiraSchema[InferJiraFieldType];
ApplyJiraSchema = JiraSchema[ApplyJiraSchema];

// Load test support functions
ExtractTestIssuesData = JiraAPI[ExtractTestIssuesData];
CreateMockProjects = JiraTestSupport[CreateMockProjects];
Table.ToNavigationTable = helperFunctions[Table.ToNavigationTable];
ValidateUrlScheme = helperFunctions[ValidateUrlScheme];

// Main entry point function
[DataSource.Kind="Atlassian_Jira_Connector.Kind", Publish="Atlassian_Jira_Connector.Publish"]
shared Jira = (
    #"Company URL Identifier (as specified in your Jira URL)" as text,
    optional #"Field List (comma separated, e.g., summary,created)" as text,
    optional #"Project Keys (comma separated, e.g., PROJ1,PROJ2)" as text,
    optional #"Max Number of Results (default: 1000)" as text,
    optional #"JQL Query String (e.g., status=""Open"")" as text,
    optional #"JQL Order By String (e.g. priority DESC, Created DESC)" as text
) => jiraNavTbl(
    #"Company URL Identifier (as specified in your Jira URL)",
    #"Field List (comma separated, e.g., summary,created)",
    #"Project Keys (comma separated, e.g., PROJ1,PROJ2)",
    #"Max Number of Results (default: 1000)",
    #"JQL Query String (e.g., status=""Open"")",
    #"JQL Order By String (e.g. priority DESC, Created DESC)",
    null
);

// This function is for SDK testing only. It is not published.
shared JiraTest = (
    #"Company URL Identifier (as specified in your Jira URL)" as text,
    optional #"Field List (comma separated, e.g., summary,created)" as text,
    optional #"Project Keys (comma separated, e.g., PROJ1,PROJ2)" as text,
    optional #"Max Number of Results (default: 1000)" as text,
    optional #"JQL Query String (e.g., status=""Open"")" as text,
    optional #"JQL Order By String (e.g. priority DESC, Created DESC)" as text,
    optional TestDataOverride as nullable record
) =>
    jiraNavTbl(
        #"Company URL Identifier (as specified in your Jira URL)",
        #"Field List (comma separated, e.g., summary,created)",
        #"Project Keys (comma separated, e.g., PROJ1,PROJ2)",
        #"Max Number of Results (default: 1000)",
        #"JQL Query String (e.g., status=""Open"")",
        #"JQL Order By String (e.g. priority DESC, Created DESC)",
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

// Data Source UI publishing description
Atlassian_Jira_Connector.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
    LearnMoreUrl = "https://github.com/rpiboy-79/atlassianJiraConnector",
    SourceImage = atlassianJiraConnector.Icons,
    SourceTypeImage = atlassianJiraConnector.Icons
];

// icon resources
atlassianJiraConnector.Icons = [
    Icon16 = { Extension.Contents("atlassianJiraConnector16.png"), Extension.Contents("atlassianJiraConnector20.png"), Extension.Contents("atlassianJiraConnector24.png"), Extension.Contents("atlassianJiraConnector32.png") },
    Icon32 = { Extension.Contents("atlassianJiraConnector32.png"), Extension.Contents("atlassianJiraConnector40.png"), Extension.Contents("atlassianJiraConnector48.png"), Extension.Contents("atlassianJiraConnector64.png") }
];

//Universal Dynamic Jira Query Function
jiraQuery =
    (
        url as text, 
        queryPath as text, 
        optional jqlProjectKey as nullable text,
        optional fieldList as nullable list,
        optional jqlString as nullable text,
        optional jqlOrderByString as nullable text,
        optional maxResults as nullable text,
        optional discoverLimits as nullable logical,
        optional useSimpleRetrieval as nullable logical,
        optional testOverride as nullable record
    ) =>
        let
        // Handle optional parameters with defaults
            actualJqlProjectKey = if jqlProjectKey = null then null else jqlProjectKey,
            actualFieldsList = if fieldList = null then null else fieldList,
            actualJqlString = if jqlString = null then null else jqlString,
            actualJqlOrderBy = if jqlOrderByString = null then null else jqlOrderByString,
            actualMaxResults = if maxResults = null then "1000" else maxResults,
            shouldDiscoverLimits = if discoverLimits = null then false else discoverLimits,
            useSimple = if useSimpleRetrieval = null then true else useSimpleRetrieval, // Default to Simple
            useTestData = if testOverride = null then null else testOverride,
            fieldsText = if actualFieldsList = null then null else Text.Combine(actualFieldsList, ","),

        //somewhere in here build full query
        jqlQuery = buildJqlQuery(actualJqlProjectKey, actualJqlString, actualJqlOrderBy),

        // Dynamic limit discovery (only if requested and JQL is provided)
            optimalMaxResults =
                if shouldDiscoverLimits and actualJqlProjectKey <> null
                    then
                        try DiscoverMaxResults(url, actualJqlProjectKey, useTestData)
                        otherwise Number.FromText(actualMaxResults)
                    else Number.FromText(actualMaxResults), // adjust for Simple
                        
            
            maxResultsText = Number.ToText(optimalMaxResults),
            maxResultsNumber = Number.FromText(maxResultsText),
            
        // DATA RETRIEVAL: Choose between simple or paginated
            CombinedData = if useSimple then
                jiraDataRetrievalSimple(
                    url,
                    queryPath,
                    jqlQuery,
                    fieldsText,
                    maxResultsText,
                    useTestData
                )
            else
                jiraDataRetrievalPaginated(
                    url,
                    queryPath,
                    jqlQuery,
                    fieldsText,
                    maxResultsText,
                    maxResultsNumber,//why do we need both values? Right now they could be different, again why?
                    useTestData
                ),

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
createIssuesNavigationTable = 
    (projectTables as list) as table =>
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
jiraNavTbl =
    (
        company as text,
        optional fields as text,
        optional projectKeys as text,
        optional maxResults as text,
        optional jqlString as text,
        optional jqlOrderBy as text,
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
                            each if Text.Trim(_) = "" then null else Text.Trim(_)
                        )),
                        each _
                    ),

            //if projectKeys = null, then check JQLstring for "project = " if project = is TRUE then get string after project = before AND/OR/(
            //how to extend for multiple project =, use list...
            //pass to project Validation so that we validate only the projects included in the JQLstring
            //allows a user to skip project keys and provide a fully qualified JQL string, but will prevent attempting to iterate over projects where the fully custom JQL string won't work
            //pass result to requested Projects either null or result of extraction from JQL query


            //is this called anywhere besides projectValidation?
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

            allProjects =
                if testOverride <> null then
                    let
                        testIssues = ExtractTestIssuesData(testOverride)
                    in
                        CreateMockProjects(testIssues)
                else
                    jiraQuery(
                        URL, //url
                        "rest/api/3/project", //queryPath
                        null, //jqlProjectKey
                        null, //fieldList
                        null, //jqlString
                        null, //jqlOrderByString
                        null, //maxResults
                        null, //discoverLimits
                        false, //useSimpleRetrieval
                        null //testOverride
                    ),//uses advanced query since project retrieval is not paginated to begin with

            


            projectValidation = 
                let
                    validProjectKeys = Table.Column(allProjects, "key"),
                    
                    // Separate valid and invalid keys
                    validKeys = List.Select(requestedProjects, each List.Contains(validProjectKeys, _)),
                    invalidKeys = List.Select(requestedProjects, each not List.Contains(validProjectKeys, _)),
                    
                    // Use all projects if no filter specified, otherwise use validated subset
                    finalKeys =
                        if List.Count(requestedProjects) = 0
                        then validProjectKeys
                        else validKeys,
                    
                    filteredProjects = Table.SelectRows(allProjects, 
                        each List.Contains(finalKeys, [key])
                    ),
                    
                    // Create validation report
                    validationReport = #table(
                        {"Metric", "Value", "Details"},
                        {
                            {"Total Available Projects", Table.RowCount(allProjects), ""},
                            {"Requested Projects", List.Count(requestedProjects), Text.Combine(requestedProjects, ", ")},
                            {"Valid Requested", List.Count(validKeys), Text.Combine(validKeys, ", ")},
                            {"Invalid Requested", List.Count(invalidKeys), Text.Combine(invalidKeys, ", ")},
                            {"Final Project Count", Table.RowCount(filteredProjects), ""},
                            {"User Supplied JQL String", if jqlString = null then "(none)" else jqlString, "Additional JQL to be applied to issue retrieval"}
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
                            jqlString,
                            jqlOrderBy, //jqlOrderByString
                            maxResults,
                            false, // Don't discover limits in Phase 1
                            true,   // Use simple retrieval
                            testOverride // PASS TEST OVERRIDE
                        )
                    ),
                    issuesWithErrorHandling = Table.TransformColumns(getIssues, {"Issues", each 
                        try _
                        catch (e) => 
                            #table(
                                {"Error Type", "Details"}, 
                                {
                                    {"Failed to retrieve issues", e[Reason]? },
                                    {"Error Message", e[Message]? },
                                    {"Check project permissions", "Verify project key and access rights"}
                                }
                            )
                    })
                in
                    issuesWithErrorHandling,
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