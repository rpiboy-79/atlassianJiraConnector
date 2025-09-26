let
    // This lets PQTest bypass your connector logic and call TestConnection directly
    Test = Extension.InvokeFunction(
        Extension.AsFunction(Atlassian_Jira_Connector.Kind, "TestConnection"),
        "{""Company URL Identifier (as specified in your Jira URL)"":" & """" & "drofus" & """" & "}"
    )
in
    Test