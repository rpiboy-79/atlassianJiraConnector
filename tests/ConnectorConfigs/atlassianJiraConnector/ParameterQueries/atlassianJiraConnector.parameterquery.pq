// atlassianJiraConnector.parameterquery.pq

let
    // Update company identifier to your instance
    Company = "drofus",
    // Optional: specify fields, project keys, max results
    FieldList = null,
    ProjectKeys = "CFR,PB",
    MaxResults = "100",
    JQLString = null,
    JQLOrderBy = null,
        
    Source = Jira (
        Company,
        FieldList,
        ProjectKeys,
        MaxResults,
        JQLString,
        JQLOrderBy
    )
in
    Source