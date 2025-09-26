// Test_MockData_DataTypes.query.pq
// Validates field data types match expectations
(parameter) =>
let
    Source = parameter,
    Schema = Table.Schema(Source),
    
    // Expected data types for key fields
    ExpectedTypes = #table(
        {"ColumnName", "ExpectedType"},
        {
            {"id", "text"},
            {"key", "text"}, 
            {"summary", "text"},
            {"created", "datetimezone"},
            {"description", "any"},  // Complex type, should be record or text
            {"issuetype", "any"},    // Complex type, should be record 
            {"components", "any"}    // Complex type, should be list or null
        }
    ),
    
    // Join with actual schema
    TypeComparison = Table.NestedJoin(
        ExpectedTypes, {"ColumnName"}, 
        Schema, {"Name"}, 
        "ActualInfo", JoinKind.LeftOuter
    ),
    
    // Validate types
    Results = Table.AddColumn(TypeComparison, "ValidationResult", 
        each if [ActualInfo] = null then "Missing"
        else if [ExpectedType] = "any" then "Pass"  // Complex types just need to exist
        else if Text.Contains(Text.Lower([ActualInfo]{0}[TypeName]), Text.Lower([ExpectedType])) then "Pass"
        else "Fail"
    ),
    
    FinalResults = Table.SelectColumns(Results, {"ColumnName", "ExpectedType", "ValidationResult"})
in
    FinalResults
