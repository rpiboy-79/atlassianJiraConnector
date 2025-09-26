// Test_MockData_DefaultFields.query.pq  
// Validates default field set matches connector specification
(parameter) =>
let
    Source = parameter,
    ColumnNames = Table.ColumnNames(Source),
    
    // Default fields as defined in your connector
    ExpectedDefaultFields = {"id", "key", "summary", "description", "issuetype", "created", "components"},
    
    // Check each expected field exists
    FieldTests = List.Transform(
        ExpectedDefaultFields,
        each [
            Field = _,
            Exists = List.Contains(ColumnNames, _),
            Result = if Exists then "Pass" else "Fail"
        ]
    ),
    
    // Convert to test results table
    Results = #table(
        {"Test", "Result", "Expected", "Actual", "Details"},
        List.Transform(FieldTests, each {
            "Field: " & [Field], 
            [Result], 
            "Present", 
            if [Exists] then "Found" else "Missing",
            "Default connector field must be present"
        }) & {
            {"Extra Fields", "Info", "Variable", Text.Combine(List.Difference(ColumnNames, ExpectedDefaultFields), ", "), "Additional fields beyond defaults"}
        }
    )
in
    Results
