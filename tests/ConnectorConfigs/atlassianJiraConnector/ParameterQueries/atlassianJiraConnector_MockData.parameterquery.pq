// atlassianJiraConnector_MockData.parameterquery.pq
// Uses JSON mock data for transformation testing

let
    //update this step to match your local repo location
    AbsolutePathPart = "C:\Users\RobertManna\BitBucket\JiraConnector",
    // Hardcoded path
    JsonContent = Json.Document(File.Contents( AbsolutePathPart & "\atlassianJiraConnector\tests\MockData\jira_api_response_acme_software_pb.json" )),
    FieldList = Text.FromBinary(File.Contents( AbsolutePathPart & "\atlassianJiraConnector\tests\MockData\FieldList.txt" )),

    // Create test override record
    TestOverride = JsonContent,
    
    // Create parameter record for connector
    ConnectorParams = [
        // Required connector parameters
        #"Company URL Identifier (as specified in your Jira URL)" = "acmesoftware", 
        #"Field List (comma separated, e.g., summary,created)" = FieldList,
        #"Project Keys (comma separated, e.g., PROJ1,PROJ2)" = "PB",
        #"Max Number of Results (default: 1000)" = "1000",
        
        // Test data override
        TestDataOverride = TestOverride
    ]
in
    ConnectorParams
