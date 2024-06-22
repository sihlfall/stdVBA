VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stdRegex3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TThis
    pattern As String
    bytecode() As Long
End Type

Private This As TThis

'Creates a new regex object
'@param pattern - The regex pattern
'@param Flags - The regex flags
'@returns - The regex object
'@remark - The flags are:
' - `i` - Case insensitive
' - `m` - Multiline
' - `g` - Global
Public Function Create(ByRef pattern As String, ByRef flags As String) As stdRegex3
    Dim obj As stdRegex3
    Set Create = New stdRegex3
    Create.protInit pattern, flags
End Function

'Initializes the regex object
'@protected
'@param pattern - The regex pattern
'@param Flags - The regex flags
'@remark - This is a protected method and should not be called directly. Use `Create` instead.
Friend Sub protInit(ByRef pattern As String, ByRef flags As String)
    With This
        .pattern = pattern
        stdRegex3Implementation.Compile .bytecode, pattern
    End With
End Sub

Public Property Get pattern() As String
    pattern = This.pattern
End Property

Public Property Get Flags() As String
    Flags = ""
End Property

'Test whether a string matches the regex
'@param str - The string to test
'@return - `True` if the string matches the regex, `False` otherwise
Public Function Test(ByRef str As String) As Boolean
    Dim captures() As Long
    With This
        Test = stdRegex3Implementation.DfsMatch(captures, .bytecode, str) <> -1
    End With
End Function

'Execute the regex against a string
'@param sHaystack - The string to execute the regex against
'@return - A collection of matches
Public Function Match(ByRef sHaystack As String) As Object
    Dim captures() As Long
    Dim oRet As Object
    Dim j As Long, n As Long

    If Me Is stdRegex3 Then
        Error.Raise ("Method called on class not object")
        Exit Function
    End If
  
    'Execute regex on haystack provided
    If stdRegex3Implementation.DfsMatch(captures, This.bytecode, sHaystack) = -1 Then
        Set Match = Nothing
        Exit Function
    End If

    'Create dictionary to store this match's data
    Set oRet = CreateObject("Scripting.Dictionary")
    
    'Basic properties
    n = UBound(captures) \ 2
    oRet("$COUNT") = n

    For j = 0 To n
        oRet(j) = Mid$(sHaystack, 1 + captures(2 * j), captures(2 * j + 1) - captures(2 * j))
    Next

    Set Match = oRet
End Function

Public Function MatchAll(ByRef sHaystack As String) As Collection
    Dim captures() As Long
    Dim oRet As Collection
    Dim oDict As Object
    Dim j As Long, n As Long, cur As Long
    Dim context As stdRegex3Implementation.DfsMatcherContext

    If Me Is stdRegex3 Then
        Error.Raise "Method called on class not object"
        Exit Function
    End If

    Set oRet = New Collection

    cur = 0
    Do
        cur = stdRegex3Implementation.DfsMatchFrom(context, captures, This.bytecode, sHaystack, cur)
        If cur = -1 Then Exit Do

        Set oDict = CreateObject("Scripting.Dictionary")
        n = UBound(captures) \ 2
        oDict("$COUNT") = n

        For j = 0 To n
            oDict(j) = Mid$(sHaystack, captures(2 * j) + 1, captures(2 * j + 1) - captures(2 * j))
        Next

        oRet.Add oDict
    Loop

    Set MatchAll = oRet
End Function
