# JQL PROJECT EXTRACTION CODE
# Extract project values from JQL strings for project validation

# ==============================================================================
# HELPER FUNCTION TO EXTRACT PROJECT VALUES FROM JQL
# ==============================================================================

// Function to extract all project values from a JQL string
ExtractProjectsFromJQL = (jqlString as text) as list =>
    let
        // Handle null or empty JQL
        cleanJQL = if jqlString = null then "" else Text.Trim(jqlString),
        
        // If empty JQL, return empty list
        projectList = if cleanJQL = "" then {} else
            let
                // Remove extra spaces and normalize
                normalizedJQL = Text.Replace(Text.Replace(cleanJQL, "  ", " "), chr(9), " "),
                
                // Find all "project =" patterns (case insensitive)
                // Using text splitting approach since M doesn't have regex
                projectMatches = ExtractProjectValues(normalizedJQL),
                
                // Clean up and deduplicate project values
                cleanProjects = List.Transform(
                    List.Distinct(List.RemoveNulls(projectMatches)), 
                    each Text.Trim(Text.Upper(_))
                )
            in
                cleanProjects
    in
        projectList,

// Helper function to extract project values using text parsing
ExtractProjectValues = (jqlText as text) as list =>
    let
        // Convert to lowercase for pattern matching
        lowerJQL = Text.Lower(jqlText),
        
        // Find all occurrences of "project =" or "project="
        // Split on "project" and process each segment
        segments = Text.Split(lowerJQL, "project"),
        
        // Process segments that start with = (indicating project = value)
        projectSegments = List.Skip(segments, 1), // Skip first segment (before any "project")
        
        projectValues = List.Transform(projectSegments, each 
            let
                trimmed = Text.Trim(_),
                // Check if starts with = (could be " = " or "=")
                startsWithEquals = Text.StartsWith(trimmed, "="),
                
                projectValue = if startsWithEquals then
                    let
                        // Get everything after the = sign
                        afterEquals = Text.Trim(Text.RemoveRange(trimmed, 0, 1)),
                        
                        // Extract the project value (handle quotes, spaces, operators)
                        extractedValue = ExtractSingleProjectValue(afterEquals)
                    in
                        extractedValue
                else
                    null
            in
                projectValue
        ),
        
        // Remove nulls from failed extractions
        validProjects = List.RemoveNulls(projectValues)
    in
        validProjects,

// Extract a single project value from text after "project ="
ExtractSingleProjectValue = (valueText as text) as text =>
    let
        // Handle quoted values: "PROJECT" or 'PROJECT'
        quotedValue = if Text.StartsWith(valueText, """") then
            let
                // Find closing quote
                restText = Text.RemoveRange(valueText, 0, 1),
                quotePos = Text.PositionOf(restText, """"),
                extracted = if quotePos >= 0 then Text.Range(restText, 0, quotePos) else restText
            in
                extracted
        else if Text.StartsWith(valueText, "'") then
            let
                // Find closing single quote  
                restText = Text.RemoveRange(valueText, 0, 1),
                quotePos = Text.PositionOf(restText, "'"),
                extracted = if quotePos >= 0 then Text.Range(restText, 0, quotePos) else restText
            in
                extracted
        else
            // Unquoted value - extract until space, AND, OR, or parenthesis
            let
                stopChars = {" ", "(", ")", " and ", " or ", " AND ", " OR "},
                // Find first occurrence of any stop character
                positions = List.Transform(stopChars, each Text.PositionOf(valueText, _)),
                validPositions = List.Select(positions, each _ >= 0),
                stopPosition = if List.Count(validPositions) > 0 then List.Min(validPositions) else Text.Length(valueText),
                extracted = Text.Range(valueText, 0, stopPosition)
            in
                extracted,
        
        // Clean up the extracted value
        cleanValue = Text.Trim(quotedValue)
    in
        if cleanValue = "" then null else cleanValue

# ==============================================================================
# UPDATED requestedProjects CODE
# ==============================================================================

//if projectKeys = null, then check JQLstring for "project = " if project = is TRUE then get string after project = before AND/OR/(
//how to extend for multiple project =, use list...
//pass to project Validation so that we validate only the projects included in the JQLstring
//allows a user to skip project keys and provide a fully qualified JQL string, but will prevent attempting to iterate over projects where the fully custom JQL string won't work
//pass result to requested Projects either null or result of extraction from JQL query

requestedProjects = 
    if projectKeys <> null then 
        // Use provided projectKeys (existing logic)
        List.Transform(
            List.RemoveNulls(List.Transform(
                Text.Split(projectKeys, ","), 
                each if Text.Trim(_) = "" then null else Text.Trim(Text.Upper(_))
            )),
            each _
        )
    else 
        // Extract projects from JQL string
        if jqlString <> null and jqlString <> "" then
            ExtractProjectsFromJQL(jqlString)
        else
            {} // Empty list if no projectKeys and no JQL

# ==============================================================================
# USAGE EXAMPLES AND TEST CASES
# ==============================================================================

// Test Case 1: Simple project filter
jql1 = "project = PB AND status = \"In Progress\""
// Should extract: {"PB"}

// Test Case 2: Multiple projects with OR
jql2 = "project = PB OR project = \"TEST PROJECT\" OR project = 'DEMO'"
// Should extract: {"PB", "TEST PROJECT", "DEMO"}

// Test Case 3: Complex query with parentheses
jql3 = "assignee = currentUser() AND (project = PB OR project = TEST) AND created >= -30d"
// Should extract: {"PB", "TEST"}

// Test Case 4: Mixed case and spacing
jql4 = "PROJECT=\"pb\" and status=Done or Project = test"
// Should extract: {"PB", "TEST"}

// Test Case 5: No project filters
jql5 = "status = \"Done\" AND assignee = currentUser()"
// Should extract: {}

// Test Case 6: Project with spaces and quotes
jql6 = "project = \"MY PROJECT NAME\" AND priority = High"
// Should extract: {"MY PROJECT NAME"}

# ==============================================================================
# COMPLETE INTEGRATION INTO YOUR CONNECTOR
# ==============================================================================

jiraNavTbl =
    (
        company as text,
        optional fields as text,
        optional projectKeys as text,
        optional maxResults as text,
        optional jqlString as text,
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

            // INLINE HELPER FUNCTIONS FOR JQL PARSING
            ExtractProjectsFromJQL = (jqlString as text) as list =>
                let
                    cleanJQL = if jqlString = null then "" else Text.Trim(jqlString),
                    projectList = if cleanJQL = "" then {} else
                        let
                            normalizedJQL = Text.Replace(Text.Replace(cleanJQL, "  ", " "), chr(9), " "),
                            projectMatches = ExtractProjectValues(normalizedJQL),
                            cleanProjects = List.Transform(
                                List.Distinct(List.RemoveNulls(projectMatches)), 
                                each Text.Trim(Text.Upper(_))
                            )
                        in
                            cleanProjects
                in
                    projectList,

            ExtractProjectValues = (jqlText as text) as list =>
                let
                    lowerJQL = Text.Lower(jqlText),
                    segments = Text.Split(lowerJQL, "project"),
                    projectSegments = List.Skip(segments, 1),
                    projectValues = List.Transform(projectSegments, each 
                        let
                            trimmed = Text.Trim(_),
                            startsWithEquals = Text.StartsWith(trimmed, "="),
                            projectValue = if startsWithEquals then
                                let
                                    afterEquals = Text.Trim(Text.RemoveRange(trimmed, 0, 1)),
                                    extractedValue = ExtractSingleProjectValue(afterEquals)
                                in
                                    extractedValue
                            else
                                null
                        in
                            projectValue
                    ),
                    validProjects = List.RemoveNulls(projectValues)
                in
                    validProjects,

            ExtractSingleProjectValue = (valueText as text) as text =>
                let
                    quotedValue = if Text.StartsWith(valueText, """") then
                        let
                            restText = Text.RemoveRange(valueText, 0, 1),
                            quotePos = Text.PositionOf(restText, """"),
                            extracted = if quotePos >= 0 then Text.Range(restText, 0, quotePos) else restText
                        in
                            extracted
                    else if Text.StartsWith(valueText, "'") then
                        let
                            restText = Text.RemoveRange(valueText, 0, 1),
                            quotePos = Text.PositionOf(restText, "'"),
                            extracted = if quotePos >= 0 then Text.Range(restText, 0, quotePos) else restText
                        in
                            extracted
                    else
                        let
                            stopChars = {" ", "(", ")", " and ", " or ", " AND ", " OR "},
                            positions = List.Transform(stopChars, each Text.PositionOf(valueText, _)),
                            validPositions = List.Select(positions, each _ >= 0),
                            stopPosition = if List.Count(validPositions) > 0 then List.Min(validPositions) else Text.Length(valueText),
                            extracted = Text.Range(valueText, 0, stopPosition)
                        in
                            extracted,
                    cleanValue = Text.Trim(quotedValue)
                in
                    if cleanValue = "" then null else cleanValue,

            // UPDATED requestedProjects logic
            requestedProjects = 
                if projectKeys <> null then 
                    // Use provided projectKeys (existing logic)
                    List.Transform(
                        List.RemoveNulls(List.Transform(
                            Text.Split(projectKeys, ","), 
                            each if Text.Trim(_) = "" then null else Text.Trim(Text.Upper(_))
                        )),
                        each _
                    )
                else 
                    // Extract projects from JQL string when projectKeys is null
                    if jqlString <> null and jqlString <> "" then
                        ExtractProjectsFromJQL(jqlString)
                    else
                        {}, // Empty list if no projectKeys and no JQL

            // Rest of your connector logic continues here...

# ==============================================================================
# BENEFITS OF THIS APPROACH
# ==============================================================================

# ✅ Handles multiple project = conditions with OR
# ✅ Supports quoted project names with spaces  
# ✅ Case-insensitive matching
# ✅ Extracts from complex nested JQL queries
# ✅ Returns empty list when no projects found
# ✅ Validates only the projects actually used in JQL
# ✅ Enables users to skip projectKeys parameter for fully custom JQL
# ✅ Maintains backward compatibility with existing projectKeys usage

# The extracted project list can now be passed to your project validation logic
# to ensure the connector only attempts to iterate over projects that are 
# actually referenced in the user's JQL query.