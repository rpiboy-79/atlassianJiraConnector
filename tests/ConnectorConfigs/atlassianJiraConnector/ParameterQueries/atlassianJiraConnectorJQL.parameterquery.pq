// atlassianJiraConnector.parameterquery.pq

let
    // Update company identifier to your instance
    Company = "drofus",
    // Optional: specify fields, project keys, max results
    FieldList = "id,key,summary,description,issuetype,created,components,status,priority",
    ProjectKeys = "CFR,PB",
    MaxResults = "100",
    JQLString = "(status=""In Progress"") OR (priority =""Major"")",
    
    Source = Jira (
        Company,
        FieldList,
        ProjectKeys,
        MaxResults,
        JQLString
    )
in
    Source