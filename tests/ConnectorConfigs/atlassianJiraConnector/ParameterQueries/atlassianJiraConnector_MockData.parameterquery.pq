// atlassianJiraConnector_MockData.parameterquery.pq
// Uses JSON mock data for transformation testing

let
//update this step to match your local repo location
    AbsolutePathPart = "C:\Users\RobertManna\BitBucket\JiraConnector",

// ##Pointer to Null Dataset##
    //JsonContent = Json.Document(File.Contents( AbsolutePathPart & "\atlassianJiraConnector\tests\MockData\jira_api_response_null.json" )),
    
// ##Pointer to Acme Dataset##
    JsonContent = Json.Document(File.Contents( AbsolutePathPart & "\atlassianJiraConnector\tests\MockData\jira_api_response_acme_software_pb.json" )),

// ##External Field List##
    FieldList = Text.FromBinary(File.Contents( AbsolutePathPart & "\atlassianJiraConnector\tests\MockData\FieldList.txt" )),

    jqlString = "status=""Open""",
    jqlOrderBy = null,
    
//Call unpublished test interface function
    Source = JiraTest(
        "acmesoftware",
        FieldList,
        "PB", 
        "1000",
        JsonContent,
        jqlString,
        jqlOrderBy
    )
in
    Source
