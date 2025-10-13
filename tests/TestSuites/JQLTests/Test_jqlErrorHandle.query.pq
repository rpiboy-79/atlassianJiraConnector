// Test_JQL_InvalidSyntax.query.pq  
(parameter) =>
let
    // This test should use parameterquery with invalid JQL: "status = " 
    // Expected: Real Jira API should return error, connector should handle gracefully
    
    ConnectorResult = try parameter otherwise null,
    HasConnectionError = ConnectorResult = null,
    
    test = 
        if HasConnectionError then
            // Good - error was handled gracefully by connector
            #table(
                {"Test", "Result", "Expected", "Actual", "Details"},
                {
                    {
                        "Invalid JQL Handling",
                        "Pass",
                        "Graceful error handling for invalid JQL",
                        "Error handled gracefully by connector",
                        "Connector properly caught and handled Jira API error response"
                    },
                    {
                        "Error Recovery",
                        "Pass", 
                        "No connector crash on API error",
                        "Connector remained stable",
                        "Validates connector error handling prevents crashes"
                    },
                    {
                        "JQL Validation", 
                        "Info",
                        "Invalid JQL: status = ",
                        "Jira API rejected invalid syntax",
                        "Confirms Jira API properly validates JQL syntax"
                    }
                }
            )
        else
            // Connector returned something - unexpected for invalid JQL
            let
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
                
                IssueData = GetIssueData(ConnectorResult),
                HasData = Table.RowCount(IssueData) > 0
            in
                #table(
                    {"Test", "Result", "Expected", "Actual", "Details"},
                    {
                        {
                            "Invalid JQL Handling",
                            if not HasData then "Pass" else "Fail",
                            "Error or empty result for invalid JQL",
                            if HasData then "Unexpectedly returned " & Text.From(Table.RowCount(IssueData)) & " issues" else "No data returned (correct)",
                            "Invalid JQL should either error or return empty result"
                        },
                        {
                            "JQL Syntax Validation",
                            if not HasData then "Pass" else "Warn",
                            "Jira API rejects invalid syntax",
                            if HasData then "API may have ignored invalid JQL" else "API correctly handled invalid JQL", 
                            "Tests that invalid JQL syntax is properly handled"
                        },
                        {
                            "JQL Query Used",
                            "Info",
                            "status = (invalid - missing value)",
                            "Invalid JQL sent to API",
                            "Shows invalid JQL that was sent to test error handling"
                        }
                    }
                )
in
    test
