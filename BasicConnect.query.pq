let
    Source = Jira("drofus", null, "CFR, PB", "2000", null, null),

    issues = Source[JPDIssues]
in
    issues

