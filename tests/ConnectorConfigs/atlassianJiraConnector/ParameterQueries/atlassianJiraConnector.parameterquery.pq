// atlassianJiraConnector.parameterquery.pq

let
    // Update company identifier to your instance
    Company = "drofus",
    // Optional: specify fields, project keys, max results
    FieldList = null,
    ProjectKeys = "CFR,PB",
    MaxResults = "100",
    
    Source = Jira (
        Company,
        FieldList,
        ProjectKeys,
        MaxResults
    )
in
    Source