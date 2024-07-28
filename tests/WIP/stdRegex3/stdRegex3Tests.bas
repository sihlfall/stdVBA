Attribute VB_Name = "stdRegex3Tests"
'@lang VBA

Sub testAll()
    Test.Topic "stdRegex3"

    Dim sHaystack As String
    sHaystack = "D-22-4BU - London: London is the capital and largest city of England and the United Kingdom." & vbCrLf & _
                "D-48-8AO - Birmingham: Birmingham is a city and metropolitan borough in the West Midlands, England" & vbCrLf & _
                "A-22-9AO - Paris: Paris is the capital and most populous city of France. Also contains A-22-9AP."

    'On Error Resume Next
    Dim sPattern As String: sPattern = "(?<Code>(?<Country>[A-Z])-(?:\d{2}-(\d[A-Z]{2})))"
    Dim rx As stdRegex3: Set rx = stdRegex3.Create(sPattern, "")

    Test.Assert "Property Access - Pattern", rx.pattern = sPattern
    Test.Assert "Property Access - Flags", rx.flags = ""
    Test.Assert "Test() Method", rx.Test(sHaystack)
    Test.Assert "Test() Method", Not stdRegex.Create("^[A-Z]{3}").Test(sHaystack) 'Ensure this returns false

    'Match should return a dictionary containing the 1st match only
    Dim oMatch As Object: Set oMatch = rx.Match(sHaystack)
    Test.Assert "Match returns Dictionary", typeName(oMatch) = "Dictionary"
    Test.Assert "Match Dictionary contains named captures 1", oMatch.Exists("Code")
    Test.Assert "Match Dictionary contains named captures 2", oMatch.Exists("Country")
    Test.Assert "Match Dictionary contains named captures 3", oMatch("Code") = "D-22-4BU"
    Test.Assert "Match Dictionary contains named captures 4", oMatch("Country") = "D"
    Test.Assert "Match Dictionary contains numbered captures 1", oMatch(0) = "D-22-4BU"
    Test.Assert "Match Dictionary contains numbered captures 2", oMatch(1) = "D-22-4BU"
    Test.Assert "Match Dictionary contains numbered captures 3", oMatch(2) = "D"
    Test.Assert "Match Dictionary contains numbered captures 4 & ensure non-capturing group not captured", oMatch(3) = "4BU"
    Test.Assert "Match contains count of submatches", oMatch("$COUNT") = 3

    'MatchAll should return a Collection of Dictionaries, and contain all matches in the haystack
    Dim oMatches As Object: Set oMatches = rx.MatchAll(sHaystack)
    Test.Assert "MatchAll returns Collection", typeName(oMatches) = "Collection"
    Test.Assert "MatchAll contains all matches", oMatches.count = 4
    Test.Assert "MatchAll contains Dictionaries", typeName(oMatches(1)) = "Dictionary"
    Test.Assert "MatchAll Dictionaries are populated 0", oMatches(1)(0) = "D-22-4BU"
    Test.Assert "MatchAll Dictionaries are populated 1", oMatches(1)(1) = "D-22-4BU"
    Test.Assert "MatchAll Dictionaries are populated 2", oMatches(1)(2) = "D"
    Test.Assert "MatchAll Dictionaries are populated 3", oMatches(1)(3) = "4BU"
    Test.Assert "MatchAll Dictionaries are populated 4", oMatches(2)(0) = "D-48-8AO"
    Test.Assert "MatchAll Dictionaries are populated 5", oMatches(3)(0) = "A-22-9AO"
    Test.Assert "MatchAll Dictionaries are populated 6", oMatches(4)(0) = "A-22-9AP"


End Sub
