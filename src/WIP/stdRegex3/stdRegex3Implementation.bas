Attribute VB_Name="stdRegex3Implementation"
Option Private Module
Option Explicit


Private Type ArrayBuffer
    buffer() As Long
    capacity As Long
    length As Long
End Type

Private Const MINIMUM_CAPACITY As Long = 16


Private Const REGEX_ERR = vbObjectError + 3000

Private Const REGEX_ERR_INVALID_REGEXP_ESCAPE = REGEX_ERR + 1
Private Const REGEX_ERR_INVALID_ESCAPE = REGEX_ERR + 2
Private Const REGEX_ERR_INVALID_RANGE = REGEX_ERR + 3
Private Const REGEX_ERR_UNTERMINATED_CHARCLASS = REGEX_ERR + 4
Private Const REGEX_ERR_INVALID_REGEXP_GROUP = REGEX_ERR + 6
Private Const REGEX_ERR_INVALID_REGEXP_CHARACTER = REGEX_ERR + 7
Private Const REGEX_ERR_INVALID_QUANTIFIER = REGEX_ERR + 8
Private Const REGEX_ERR_UNEXPECTED_END_OF_PATTERN = REGEX_ERR + 9
Private Const REGEX_ERR_UNEXPECTED_REGEXP_TOKEN = REGEX_ERR + 10
Private Const REGEX_ERR_UNEXPECTED_CLOSING_PAREN = REGEX_ERR + 11
Private Const REGEX_ERR_INVALID_QUANTIFIER_NO_ATOM = REGEX_ERR + 12

Private Const REGEX_ERR_INTERNAL_LOGIC_ERR = REGEX_ERR + 20

Private Const REGEX_ERR_INVALID_INPUT = REGEX_ERR + 50


' regexp opcodes
Private Enum ReOpType
    REOP_MATCH = 1
    REOP_CHAR = 2
    REOP_PERIOD = 3
    REOP_RANGES = 4 ' nranges [must be >= 1], chfrom, chto, chfrom, chto, ...
    REOP_INVRANGES = 5
    REOP_JUMP = 6
    REOP_SPLIT1 = 7 ' prefer direct
    REOP_SPLIT2 = 8 ' prefer jump
    REOP_SAVE = 11
    REOP_LOOKPOS = 13
    REOP_LOOKNEG = 14
    REOP_BACKREFERENCE = 15
    REOP_ASSERT_START = 16
    REOP_ASSERT_END = 17
    REOP_ASSERT_WORD_BOUNDARY = 18
    REOP_ASSERT_NOT_WORD_BOUNDARY = 19
    REOP_REPEAT_EXACTLY_INIT = 20 ' <none>
    REOP_REPEAT_EXACTLY_START = 21 ' quantity [must be >= 1], offset
    REOP_REPEAT_EXACTLY_END = 22 ' quantity [must be >= 1], offset
    REOP_REPEAT_MAX_HUMBLE_INIT = 23 ' <none>
    REOP_REPEAT_MAX_HUMBLE_START = 24 ' quantity, offset
    REOP_REPEAT_MAX_HUMBLE_END = 25 ' quantitiy, offset
    REOP_REPEAT_GREEDY_MAX_INIT = 26 ' <none>
    REOP_REPEAT_GREEDY_MAX_START = 27 ' quantity, offset
    REOP_REPEAT_GREEDY_MAX_END = 28 ' quantitiy, offset
    REOP_CHECK_LOOKAHEAD = 29 ' <none>
    REOP_CHECK_LOOKBEHIND = 30 ' <none>
    REOP_END_LOOKPOS = 31 ' <none>
    REOP_END_LOOKNEG = 32 ' <none>
    REOP_FAIL = 33
End Enum

' Todo: Introduce special value or restrict max. explicit quantifier value to MAX_LONG - 1
Private Const RE_QUANTIFIER_INFINITE As Long = &H7FFFFFFF


Private UnicodeCanonLookupTable(0 To 65535) As Integer
Private UnicodeCanonRunsTable(0 To 302) As Long

Private UnicodeInitialized As Boolean


' Guarantee: All AST_ASSERT_LOOKAHEAD and LOOKBEHIND constants are > 0.
'   Parser relies on this.
Private Const MIN_AST_CODE As Long = 0

Private Const AST_EMPTY As Long = 0
Private Const AST_STRING As Long = 1
Private Const AST_DISJ As Long = 2
Private Const AST_CONCAT As Long = 3
Private Const AST_CHAR As Long = 4
Private Const AST_CAPTURE As Long = 5
Private Const AST_REPEAT_EXACTLY As Long = 6
Private Const AST_PERIOD As Long = 7
Private Const AST_ASSERT_START As Long = 8
Private Const AST_ASSERT_END As Long = 9
Private Const AST_ASSERT_WORD_BOUNDARY As Long = 10
Private Const AST_ASSERT_NOT_WORD_BOUNDARY As Long = 11
Private Const AST_MATCH As Long = 12
Private Const AST_ZEROONE_GREEDY As Long = 13
Private Const AST_ZEROONE_HUMBLE As Long = 14
Private Const AST_STAR_GREEDY As Long = 15
Private Const AST_STAR_HUMBLE As Long = 16
Private Const AST_REPEAT_MAX_GREEDY As Long = 17
Private Const AST_REPEAT_MAX_HUMBLE As Long = 18
Private Const AST_RANGES As Long = 19
Private Const AST_INVRANGES As Long = 20
Private Const AST_ASSERT_POS_LOOKAHEAD As Long = 21
Private Const AST_ASSERT_NEG_LOOKAHEAD As Long = 22
Private Const AST_ASSERT_POS_LOOKBEHIND As Long = 23
Private Const AST_ASSERT_NEG_LOOKBEHIND As Long = 24
Private Const AST_FAIL As Long = 25
Private Const AST_BACKREFERENCE As Long = 26

Private Const MAX_AST_CODE As Long = 27

Private Const LONGTYPE_FIRST_BIT As Long = &H80000000
Private Const LONGTYPE_ALL_BUT_FIRST_BIT As Long = Not LONGTYPE_FIRST_BIT

Private Const NODE_TYPE As Long = 0
Private Const NODE_LCHILD As Long = 1
Private Const NODE_RCHILD As Long = 2

' The initial stack capacity must be >= 2 * max([1 + entry.esfs for entry in AST_TABLE]),
'   since when increasing the stack capacity, we increase by (current size \ 2) and
'   we assume that this will be sufficient for the next stack frame.
Private Const INITIAL_STACK_CAPACITY As Long = 8 '16

Private AST_TABLE_INITIALIZED As Boolean ' default-initialized to False

' .nc: number of children; negative values have special meaning
'   -2: is AST_STRING
'   -1: is AST_RANGES or AST_INVRANGES
' .blen: length of bytecode generated for this node (meaningful only if .nc >= 0)
' .esfs: extra stack space required when generating bytecode for this node
'   Only nodes with children are permitted to require extra stack space.
'   Hence .esfs > 0 must imply .nc >= 1.
Private Type AstTableEntry
    nc As Integer
    blen As Integer
    esfs As Integer
End Type

Private AST_TABLE(MIN_AST_CODE To MAX_AST_CODE) As AstTableEntry


Private RangeTablesInitialized As Boolean ' auto-initialized to False
Private RangeTableDigit(0 To 1) As Long
Private RangeTableWhite(0 To 21) As Long
Private RangeTableWordchar(0 To 7) As Long
Private RangeTableNotDigit(0 To 3) As Long
Private RangeTableNotWhite(0 To 23) As Long
Private RangeTableNotWordChar(0 To 9) As Long

Private Const MIN_LONG = &H80000000
Private Const MAX_LONG = &H7FFFFFFF



' TODO: Make sure unexpected end of input is treated as an error everywhere.

' inputArray an array of bytes, representing two-byte characters.
' iEnd points to the first byte of the last character, i.e. iEnd = UBound(inputArray) - 1.
'   iEnd should be guaranteed to be an even number.
' iCurrent points to the first byte of the current character.
' currentCharacter contains the value of the current character (16 bit between 0 and 2^16 - 1, or ENDOFINPUT).
' After the end of the input has been reached, currentCharacter = ENDOFINPUT and iCurrent = iEnd.
' All fields are intended to be private to this module.
' TODO: Probably character values from MIN_INT (negative) to MAX_INT would be more efficient. Stick to one solution.
Private Type LexerContext
    iCurrent As Long
    iEnd As Long
    inputArray() As Byte
    currentCharacter As Long
End Type

Private Type ReToken
    t As Long ' token type
    greedy As Boolean
    num As Long ' numeric value (character, count)
    qmin As Long
    qmax As Long
End Type

Private Const RETOK_EOF As Long = 0
Private Const RETOK_DISJUNCTION As Long = 1
Private Const RETOK_QUANTIFIER As Long = 2
Private Const RETOK_ASSERT_START As Long = 3
Private Const RETOK_ASSERT_END As Long = 4
Private Const RETOK_ASSERT_WORD_BOUNDARY As Long = 5
Private Const RETOK_ASSERT_NOT_WORD_BOUNDARY As Long = 6
Private Const RETOK_ASSERT_START_POS_LOOKAHEAD As Long = 7
Private Const RETOK_ASSERT_START_NEG_LOOKAHEAD As Long = 8
Private Const RETOK_ATOM_PERIOD As Long = 9
Private Const RETOK_ATOM_CHAR As Long = 10
Private Const RETOK_ATOM_DIGIT As Long = 11                   ' assumptions in regexp compiler
Private Const RETOK_ATOM_NOT_DIGIT As Long = 12               ' -""-
Private Const RETOK_ATOM_WHITE As Long = 13                   ' -""-
Private Const RETOK_ATOM_NOT_WHITE As Long = 14               ' -""-
Private Const RETOK_ATOM_WORD_CHAR As Long = 15               ' -""-
Private Const RETOK_ATOM_NOT_WORD_CHAR As Long = 16           ' -""-
Private Const RETOK_ATOM_BACKREFERENCE As Long = 17
Private Const RETOK_ATOM_START_CAPTURE_GROUP As Long = 18
Private Const RETOK_ATOM_START_NONCAPTURE_GROUP As Long = 19
Private Const RETOK_ATOM_START_CHARCLASS As Long = 20
Private Const RETOK_ATOM_START_CHARCLASS_INVERTED As Long = 21
Private Const RETOK_ASSERT_START_POS_LOOKBEHIND As Long = 22
Private Const RETOK_ASSERT_START_NEG_LOOKBEHIND As Long = 23
Private Const RETOK_ATOM_END As Long = 24 ' closing parenthesis (ends (POS|NEG)_LOOK(AHEAD|BEHIND), CAPTURE_GROUP, NONCAPTURE_GROUP)

Private EMPTY_RE_TOKEN As ReToken ' effectively a constant -- zeroed out by default

' Returned by input reading function after end of input has been reached
' Since our characters are 16 bit and Long is 32 bit, the value -1 is free to use for us.
' TODO: Should probably be MIN_LONG
Private Const ENDOFINPUT As Long = -1

Private Const MAX_LONG_DIV_10 As Long = MAX_LONG \ 10

Private Const UNICODE_EXCLAMATION As Long = 33  ' !
Private Const UNICODE_DOLLAR As Long = 36  ' $
Private Const UNICODE_LPAREN As Long = 40  ' (
Private Const UNICODE_RPAREN As Long = 41  ' )
Private Const UNICODE_STAR As Long = 42  ' *
Private Const UNICODE_PLUS As Long = 43  ' +
Private Const UNICODE_COMMA As Long = 44  ' ,
Private Const UNICODE_MINUS As Long = 45  ' -
Private Const UNICODE_PERIOD As Long = 46  ' .
Private Const UNICODE_0 As Long = 48  ' 0
Private Const UNICODE_1 As Long = 49  ' 1
Private Const UNICODE_7 As Long = 55  ' 7
Private Const UNICODE_9 As Long = 57  ' 9
Private Const UNICODE_COLON As Long = 58  ' :
Private Const UNICODE_LT As Long = 60  ' <
Private Const UNICODE_EQUALS As Long = 61  ' =
Private Const UNICODE_QUESTION As Long = 63  ' ?
Private Const UNICODE_UC_A As Long = 65  ' A
Private Const UNICODE_UC_B As Long = 66  ' B
Private Const UNICODE_UC_D As Long = 68  ' D
Private Const UNICODE_UC_F As Long = 70  ' F
Private Const UNICODE_UC_S As Long = 83  ' S
Private Const UNICODE_UC_W As Long = 87  ' W
Private Const UNICODE_UC_Z As Long = 90  ' Z
Private Const UNICODE_LBRACKET As Long = 91  ' [
Private Const UNICODE_BACKSLASH As Long = 92  ' \
Private Const UNICODE_RBRACKET As Long = 93  ' ]
Private Const UNICODE_CARET As Long = 94  ' ^
Private Const UNICODE_LC_A As Long = 97  ' a
Private Const UNICODE_LC_B As Long = 98  ' b
Private Const UNICODE_LC_C As Long = 99  ' c
Private Const UNICODE_LC_D As Long = 100  ' d
Private Const UNICODE_LC_F As Long = 102  ' f
Private Const UNICODE_LC_N As Long = 110  ' n
Private Const UNICODE_LC_R As Long = 114  ' r
Private Const UNICODE_LC_S As Long = 115  ' s
Private Const UNICODE_LC_T As Long = 116  ' t
Private Const UNICODE_LC_U As Long = 117  ' u
Private Const UNICODE_LC_V As Long = 118  ' v
Private Const UNICODE_LC_W As Long = 119  ' w
Private Const UNICODE_LC_X As Long = 120  ' x
Private Const UNICODE_LC_Z As Long = 122  ' z
Private Const UNICODE_LCURLY As Long = 123  ' {
Private Const UNICODE_PIPE As Long = 124  ' |
Private Const UNICODE_RCURLY As Long = 125  ' }
Private Const UNICODE_CP_ZWNJ As Long = &H200C& ' zero-width non-joiner
Private Const UNICODE_CP_ZWJ As Long = &H200D&  ' zero-width joiner


Private Const RE_FLAG_NONE As Long = 0
Private Const RE_FLAG_GLOBAL As Long = 1
Private Const RE_FLAG_IGNORE_CASE As Long = 2
Private Const RE_FLAG_MULTILINE As Long = 4



Private Const DFS_FLAGS_NONE As Long = 0

Private Const DEFAULT_STEPS_LIMIT As Long = 10000
Private Const DEFAULT_MINIMUM_THREADSTACK_CAPACITY As Long = 16
Private Const Q_NONE As Long = -2


Private Const DFS_MATCHER_STACK_MINIMUM_CAPACITY As Long = 16

Private Type DfsMatcherStackFrame
    master As Long
    capturesStackState As Long
    qStackLength As Long
    pc As Long
    sp As Long
    pcLandmark As Long
    spDelta As Long
    q As Long
    qTop As Long
End Type

Private Type DfsMatcherStack
    buffer() As DfsMatcherStackFrame
    capacity As Long
    length As Long
End Type

Public Type DfsMatcherContext
    matcherStack As DfsMatcherStack
    capturesStack As ArrayBuffer
    qstack As ArrayBuffer
    nCapturePoints As Long
    
    master As Long
    
    capturesRequireCoW As Boolean
    qTop As Long
End Type



Private Sub AppendLong(ByRef lab As ArrayBuffer, ByVal v As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 1
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        .buffer(.length) = v
        .length = requiredCapacity
    End With
End Sub

Private Sub AppendTwo(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 2
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        .buffer(.length) = v1
        .buffer(.length + 1) = v2
        .length = requiredCapacity
    End With
End Sub

Private Sub AppendThree(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 3
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        .buffer(.length) = v1
        .buffer(.length + 1) = v2
        .buffer(.length + 2) = v3
        .length = requiredCapacity
    End With
End Sub

Private Sub AppendFour(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 4
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        .buffer(.length) = v1
        .buffer(.length + 1) = v2
        .buffer(.length + 2) = v3
        .buffer(.length + 3) = v4
        .length = requiredCapacity
    End With
End Sub

Private Sub AppendFive(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long, v5 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 5
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        .buffer(.length) = v1
        .buffer(.length + 1) = v2
        .buffer(.length + 2) = v3
        .buffer(.length + 3) = v4
        .buffer(.length + 4) = v5
        .length = requiredCapacity
    End With
End Sub

Private Sub AppendEight(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long, v5 As Long, v6 As Long, v7 As Long, v8 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 8
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        .buffer(.length) = v1
        .buffer(.length + 1) = v2
        .buffer(.length + 2) = v3
        .buffer(.length + 3) = v4
        .buffer(.length + 4) = v5
        .buffer(.length + 5) = v6
        .buffer(.length + 6) = v7
        .buffer(.length + 7) = v8
        .length = requiredCapacity
    End With
End Sub

Private Sub AppendFill(ByRef lab As ArrayBuffer, ByVal cnt As Long, ByVal v As Long)
    Dim requiredCapacity As Long, i As Long
    With lab
        requiredCapacity = .length + cnt
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        i = .length
        Do While i < requiredCapacity: .buffer(i) = v: i = i + 1: Loop
        .length = requiredCapacity
    End With
End Sub

Private Sub AppendSlice(ByRef lab As ArrayBuffer, ByVal offset As Long, ByVal length As Long)
    Dim requiredCapacity As Long, i As Long, j As Long, upper As Long
    With lab
        requiredCapacity = .length + length
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        upper = offset + length: i = offset: j = .length
        Do While i < upper
            .buffer(j) = .buffer(i)
            i = i + 1: j = j + 1
        Loop
        .length = requiredCapacity
    End With
End Sub

Private Sub AppendUnspecified(ByRef lab As ArrayBuffer, ByVal n As Long)
    Dim requiredCapacity As Long
    With lab
        .length = .length + n
        requiredCapacity = .length
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
    End With
End Sub

Private Sub AppendPrefixedPairsArray(ByRef lab As ArrayBuffer, ByVal prefix As Long, ByRef ary() As Long)
    ' prefix, number of pairs, pairs
    Dim requiredCapacity As Long, lb As Long, ub As Long, i As Long, j As Long
    With lab
        lb = LBound(ary)
        ub = UBound(ary)
        i = .length
        .length = .length + ub - lb + 3
        requiredCapacity = .length
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        
        .buffer(i) = prefix: i = i + 1
        .buffer(i) = (ub - lb + 1) \ 2: i = i + 1
        For j = lb To ub
            .buffer(i) = ary(j)
            i = i + 1
        Next
    End With
End Sub

Private Sub AppendPrefixedPairsArrayBackwards(ByRef lab As ArrayBuffer, ByVal prefix As Long, ByRef ary() As Long)
    ' prefix, number of pairs, pairs
    Dim requiredCapacity As Long, lb As Long, ub As Long, i As Long, j As Long
    With lab
        lb = LBound(ary)
        ub = UBound(ary)
        i = .length
        .length = .length + ub - lb + 3
        requiredCapacity = .length
        If requiredCapacity > .capacity Then IncreaseCapacity lab, requiredCapacity
        
        .buffer(i) = prefix: i = i + 1
        .buffer(i) = (ub - lb + 1) \ 2: i = i + 1
        j = ub - 1
        Do
            .buffer(i) = ary(j): i = i + 1
            .buffer(i) = ary(j + 1): i = i + 1
            j = j - 2
        Loop Until j < lb
    End With
End Sub

Private Sub IncreaseCapacity(ByRef lab As ArrayBuffer, requiredCapacity As Long)
    Dim cap As Long
    With lab
        cap = .capacity
        If cap <= MINIMUM_CAPACITY Then cap = MINIMUM_CAPACITY
        Do Until cap >= requiredCapacity
            cap = cap + cap \ 2
        Loop
        ReDim Preserve .buffer(0 To cap - 1) As Long
        .capacity = cap
    End With
End Sub
Sub UnicodeInitialize()
    InitializeUnicodeCanonLookupTable UnicodeCanonLookupTable
    InitializeUnicodeCanonRunsTable UnicodeCanonRunsTable
    UnicodeInitialized = True
End Sub

Function ReCanonicalizeChar(ByVal Codepoint As Long)
    If Codepoint And &H10000 Then ' Codepoint not in [0;&HFFFF&]
        ReCanonicalizeChar = Codepoint
    Else
        ReCanonicalizeChar = (Codepoint + UnicodeCanonLookupTable(Codepoint) + &H10000) And &HFFFF&
    End If
End Function

Private Sub InitializeUnicodeCanonLookupTable(ByRef t() As Integer)
    Dim i As Long
    
    For i = 97 To 122: t(i) = -32: Next i
    t(181) = 743
    For i = 224 To 246: t(i) = -32: Next i
    For i = 248 To 254: t(i) = -32: Next i
    t(255) = 121
    For i = 257 To 303 Step 2: t(i) = -1: Next i
    t(307) = -1
    t(309) = -1
    t(311) = -1
    For i = 314 To 328 Step 2: t(i) = -1: Next i
    For i = 331 To 375 Step 2: t(i) = -1: Next i
    t(378) = -1
    t(380) = -1
    t(382) = -1
    t(384) = 195
    t(387) = -1
    t(389) = -1
    t(392) = -1
    t(396) = -1
    t(402) = -1
    t(405) = 97
    t(409) = -1
    t(410) = 163
    t(414) = 130
    t(417) = -1
    t(419) = -1
    t(421) = -1
    t(424) = -1
    t(429) = -1
    t(432) = -1
    t(436) = -1
    t(438) = -1
    t(441) = -1
    t(445) = -1
    t(447) = 56
    t(453) = -1
    t(454) = -2
    t(456) = -1
    t(457) = -2
    t(459) = -1
    t(460) = -2
    For i = 462 To 476 Step 2: t(i) = -1: Next i
    t(477) = -79
    For i = 479 To 495 Step 2: t(i) = -1: Next i
    t(498) = -1
    t(499) = -2
    t(501) = -1
    For i = 505 To 543 Step 2: t(i) = -1: Next i
    For i = 547 To 563 Step 2: t(i) = -1: Next i
    t(572) = -1
    t(575) = 10815
    t(576) = 10815
    t(578) = -1
    For i = 583 To 591 Step 2: t(i) = -1: Next i
    t(592) = 10783
    t(593) = 10780
    t(594) = 10782
    t(595) = -210
    t(596) = -206
    t(598) = -205
    t(599) = -205
    t(601) = -202
    t(603) = -203
    t(604) = -23217
    t(608) = -205
    t(609) = -23221
    t(611) = -207
    t(613) = -23256
    t(614) = -23228
    t(616) = -209
    t(617) = -211
    t(618) = -23228
    t(619) = 10743
    t(620) = -23231
    t(623) = -211
    t(625) = 10749
    t(626) = -213
    t(629) = -214
    t(637) = 10727
    t(640) = -218
    t(642) = -23229
    t(643) = -218
    t(647) = -23254
    t(648) = -218
    t(649) = -69
    t(650) = -217
    t(651) = -217
    t(652) = -71
    t(658) = -219
    t(669) = -23275
    t(670) = -23278
    t(837) = 84
    t(881) = -1
    t(883) = -1
    t(887) = -1
    t(891) = 130
    t(892) = 130
    t(893) = 130
    t(940) = -38
    t(941) = -37
    t(942) = -37
    t(943) = -37
    For i = 945 To 961: t(i) = -32: Next i
    t(962) = -31
    For i = 963 To 971: t(i) = -32: Next i
    t(972) = -64
    t(973) = -63
    t(974) = -63
    t(976) = -62
    t(977) = -57
    t(981) = -47
    t(982) = -54
    t(983) = -8
    For i = 985 To 1007 Step 2: t(i) = -1: Next i
    t(1008) = -86
    t(1009) = -80
    t(1010) = 7
    t(1011) = -116
    t(1013) = -96
    t(1016) = -1
    t(1019) = -1
    For i = 1072 To 1103: t(i) = -32: Next i
    For i = 1104 To 1119: t(i) = -80: Next i
    For i = 1121 To 1153 Step 2: t(i) = -1: Next i
    For i = 1163 To 1215 Step 2: t(i) = -1: Next i
    For i = 1218 To 1230 Step 2: t(i) = -1: Next i
    t(1231) = -15
    For i = 1233 To 1327 Step 2: t(i) = -1: Next i
    For i = 1377 To 1414: t(i) = -48: Next i
    For i = 4304 To 4346: t(i) = 3008: Next i
    t(4349) = 3008
    t(4350) = 3008
    t(4351) = 3008
    For i = 5112 To 5117: t(i) = -8: Next i
    t(7296) = -6254
    t(7297) = -6253
    t(7298) = -6244
    t(7299) = -6242
    t(7300) = -6242
    t(7301) = -6243
    t(7302) = -6236
    t(7303) = -6181
    t(7304) = -30270
    t(7545) = -30204
    t(7549) = 3814
    t(7566) = -30152
    For i = 7681 To 7829 Step 2: t(i) = -1: Next i
    t(7835) = -59
    For i = 7841 To 7935 Step 2: t(i) = -1: Next i
    For i = 7936 To 7943: t(i) = 8: Next i
    For i = 7952 To 7957: t(i) = 8: Next i
    For i = 7968 To 7975: t(i) = 8: Next i
    For i = 7984 To 7991: t(i) = 8: Next i
    For i = 8000 To 8005: t(i) = 8: Next i
    t(8017) = 8
    t(8019) = 8
    t(8021) = 8
    t(8023) = 8
    For i = 8032 To 8039: t(i) = 8: Next i
    t(8048) = 74
    t(8049) = 74
    t(8050) = 86
    t(8051) = 86
    t(8052) = 86
    t(8053) = 86
    t(8054) = 100
    t(8055) = 100
    t(8056) = 128
    t(8057) = 128
    t(8058) = 112
    t(8059) = 112
    t(8060) = 126
    t(8061) = 126
    t(8112) = 8
    t(8113) = 8
    t(8126) = -7205
    t(8144) = 8
    t(8145) = 8
    t(8160) = 8
    t(8161) = 8
    t(8165) = 7
    t(8526) = -28
    For i = 8560 To 8575: t(i) = -16: Next i
    t(8580) = -1
    For i = 9424 To 9449: t(i) = -26: Next i
    For i = 11312 To 11358: t(i) = -48: Next i
    t(11361) = -1
    t(11365) = -10795
    t(11366) = -10792
    t(11368) = -1
    t(11370) = -1
    t(11372) = -1
    t(11379) = -1
    t(11382) = -1
    For i = 11393 To 11491 Step 2: t(i) = -1: Next i
    t(11500) = -1
    t(11502) = -1
    t(11507) = -1
    For i = 11520 To 11557: t(i) = -7264: Next i
    t(11559) = -7264
    t(11565) = -7264
    For i = 42561 To 42605 Step 2: t(i) = -1: Next i
    For i = 42625 To 42651 Step 2: t(i) = -1: Next i
    For i = 42787 To 42799 Step 2: t(i) = -1: Next i
    For i = 42803 To 42863 Step 2: t(i) = -1: Next i
    t(42874) = -1
    t(42876) = -1
    For i = 42879 To 42887 Step 2: t(i) = -1: Next i
    t(42892) = -1
    t(42897) = -1
    t(42899) = -1
    t(42900) = 48
    For i = 42903 To 42921 Step 2: t(i) = -1: Next i
    For i = 42933 To 42943 Step 2: t(i) = -1: Next i
    t(42947) = -1
    t(43859) = -928
    For i = 43888 To 43967: t(i) = 26672: Next i
    For i = 65345 To 65370: t(i) = -32: Next i
End Sub

Private Sub InitializeUnicodeCanonRunsTable(ByRef t() As Long)
    t(0) = 97
    t(1) = 123
    t(2) = 181
    t(3) = 182
    t(4) = 224
    t(5) = 247
    t(6) = 248
    t(7) = 255
    t(8) = 256
    t(9) = 257
    t(10) = 304
    t(11) = 307
    t(12) = 312
    t(13) = 314
    t(14) = 329
    t(15) = 331
    t(16) = 376
    t(17) = 378
    t(18) = 383
    t(19) = 384
    t(20) = 385
    t(21) = 387
    t(22) = 390
    t(23) = 392
    t(24) = 393
    t(25) = 396
    t(26) = 397
    t(27) = 402
    t(28) = 403
    t(29) = 405
    t(30) = 406
    t(31) = 409
    t(32) = 410
    t(33) = 411
    t(34) = 414
    t(35) = 415
    t(36) = 417
    t(37) = 422
    t(38) = 424
    t(39) = 425
    t(40) = 429
    t(41) = 430
    t(42) = 432
    t(43) = 433
    t(44) = 436
    t(45) = 439
    t(46) = 441
    t(47) = 442
    t(48) = 445
    t(49) = 446
    t(50) = 447
    t(51) = 448
    t(52) = 453
    t(53) = 454
    t(54) = 455
    t(55) = 456
    t(56) = 457
    t(57) = 458
    t(58) = 459
    t(59) = 460
    t(60) = 461
    t(61) = 462
    t(62) = 477
    t(63) = 478
    t(64) = 479
    t(65) = 496
    t(66) = 498
    t(67) = 499
    t(68) = 500
    t(69) = 501
    t(70) = 502
    t(71) = 505
    t(72) = 544
    t(73) = 547
    t(74) = 564
    t(75) = 572
    t(76) = 573
    t(77) = 575
    t(78) = 577
    t(79) = 578
    t(80) = 579
    t(81) = 583
    t(82) = 592
    t(83) = 593
    t(84) = 594
    t(85) = 595
    t(86) = 596
    t(87) = 597
    t(88) = 598
    t(89) = 600
    t(90) = 601
    t(91) = 602
    t(92) = 603
    t(93) = 604
    t(94) = 605
    t(95) = 608
    t(96) = 609
    t(97) = 610
    t(98) = 611
    t(99) = 612
    t(100) = 613
    t(101) = 614
    t(102) = 615
    t(103) = 616
    t(104) = 617
    t(105) = 618
    t(106) = 619
    t(107) = 620
    t(108) = 621
    t(109) = 623
    t(110) = 624
    t(111) = 625
    t(112) = 626
    t(113) = 627
    t(114) = 629
    t(115) = 630
    t(116) = 637
    t(117) = 638
    t(118) = 640
    t(119) = 641
    t(120) = 642
    t(121) = 643
    t(122) = 644
    t(123) = 647
    t(124) = 648
    t(125) = 649
    t(126) = 650
    t(127) = 652
    t(128) = 653
    t(129) = 658
    t(130) = 659
    t(131) = 669
    t(132) = 670
    t(133) = 671
    t(134) = 837
    t(135) = 838
    t(136) = 881
    t(137) = 884
    t(138) = 887
    t(139) = 888
    t(140) = 891
    t(141) = 894
    t(142) = 940
    t(143) = 941
    t(144) = 944
    t(145) = 945
    t(146) = 962
    t(147) = 963
    t(148) = 972
    t(149) = 973
    t(150) = 975
    t(151) = 976
    t(152) = 977
    t(153) = 978
    t(154) = 981
    t(155) = 982
    t(156) = 983
    t(157) = 984
    t(158) = 985
    t(159) = 1008
    t(160) = 1009
    t(161) = 1010
    t(162) = 1011
    t(163) = 1012
    t(164) = 1013
    t(165) = 1014
    t(166) = 1016
    t(167) = 1017
    t(168) = 1019
    t(169) = 1020
    t(170) = 1072
    t(171) = 1104
    t(172) = 1120
    t(173) = 1121
    t(174) = 1154
    t(175) = 1163
    t(176) = 1216
    t(177) = 1218
    t(178) = 1231
    t(179) = 1232
    t(180) = 1233
    t(181) = 1328
    t(182) = 1377
    t(183) = 1415
    t(184) = 4304
    t(185) = 4347
    t(186) = 4349
    t(187) = 4352
    t(188) = 5112
    t(189) = 5118
    t(190) = 7296
    t(191) = 7297
    t(192) = 7298
    t(193) = 7299
    t(194) = 7301
    t(195) = 7302
    t(196) = 7303
    t(197) = 7304
    t(198) = 7305
    t(199) = 7545
    t(200) = 7546
    t(201) = 7549
    t(202) = 7550
    t(203) = 7566
    t(204) = 7567
    t(205) = 7681
    t(206) = 7830
    t(207) = 7835
    t(208) = 7836
    t(209) = 7841
    t(210) = 7936
    t(211) = 7944
    t(212) = 7952
    t(213) = 7958
    t(214) = 7968
    t(215) = 7976
    t(216) = 7984
    t(217) = 7992
    t(218) = 8000
    t(219) = 8006
    t(220) = 8017
    t(221) = 8024
    t(222) = 8032
    t(223) = 8040
    t(224) = 8048
    t(225) = 8050
    t(226) = 8054
    t(227) = 8056
    t(228) = 8058
    t(229) = 8060
    t(230) = 8062
    t(231) = 8112
    t(232) = 8114
    t(233) = 8126
    t(234) = 8127
    t(235) = 8144
    t(236) = 8146
    t(237) = 8160
    t(238) = 8162
    t(239) = 8165
    t(240) = 8166
    t(241) = 8526
    t(242) = 8527
    t(243) = 8560
    t(244) = 8576
    t(245) = 8580
    t(246) = 8581
    t(247) = 9424
    t(248) = 9450
    t(249) = 11312
    t(250) = 11359
    t(251) = 11361
    t(252) = 11362
    t(253) = 11365
    t(254) = 11366
    t(255) = 11367
    t(256) = 11368
    t(257) = 11373
    t(258) = 11379
    t(259) = 11380
    t(260) = 11382
    t(261) = 11383
    t(262) = 11393
    t(263) = 11492
    t(264) = 11500
    t(265) = 11503
    t(266) = 11507
    t(267) = 11508
    t(268) = 11520
    t(269) = 11558
    t(270) = 11559
    t(271) = 11560
    t(272) = 11565
    t(273) = 11566
    t(274) = 42561
    t(275) = 42606
    t(276) = 42625
    t(277) = 42652
    t(278) = 42787
    t(279) = 42800
    t(280) = 42803
    t(281) = 42864
    t(282) = 42874
    t(283) = 42877
    t(284) = 42879
    t(285) = 42888
    t(286) = 42892
    t(287) = 42893
    t(288) = 42897
    t(289) = 42900
    t(290) = 42901
    t(291) = 42903
    t(292) = 42922
    t(293) = 42933
    t(294) = 42944
    t(295) = 42947
    t(296) = 42948
    t(297) = 43859
    t(298) = 43860
    t(299) = 43888
    t(300) = 43968
    t(301) = 65345
    t(302) = 65371
End Sub

Private Sub InitializeAstTable()
    With AST_TABLE(AST_EMPTY):          .nc = 0:  .blen = 0: .esfs = 0: End With
    With AST_TABLE(AST_STRING):         .nc = -2: .blen = 2: .esfs = 0: End With
    With AST_TABLE(AST_DISJ):           .nc = 2:  .blen = 4: .esfs = 1: End With
    With AST_TABLE(AST_CONCAT):         .nc = 2:  .blen = 0: .esfs = 0: End With
    With AST_TABLE(AST_CHAR):           .nc = 0:  .blen = 2: .esfs = 0: End With
    With AST_TABLE(AST_CAPTURE):          .nc = 1:  .blen = 4: .esfs = 0: End With
    With AST_TABLE(AST_REPEAT_EXACTLY): .nc = 1:  .blen = 7: .esfs = 1: End With
    With AST_TABLE(AST_PERIOD):         .nc = 0:  .blen = 1: .esfs = 0: End With
    With AST_TABLE(AST_ASSERT_START):   .nc = 0:  .blen = 1: .esfs = 0: End With
    With AST_TABLE(AST_ASSERT_END):     .nc = 0:  .blen = 1: .esfs = 0: End With
    With AST_TABLE(AST_ASSERT_WORD_BOUNDARY):     .nc = 0: .blen = 1: .esfs = 0: End With
    With AST_TABLE(AST_ASSERT_NOT_WORD_BOUNDARY): .nc = 0: .blen = 1: .esfs = 0: End With
    With AST_TABLE(AST_MATCH):          .nc = 0:  .blen = 1: .esfs = 0: End With
    With AST_TABLE(AST_ZEROONE_GREEDY): .nc = 1:  .blen = 2: .esfs = 1: End With
    With AST_TABLE(AST_ZEROONE_HUMBLE): .nc = 1:  .blen = 2: .esfs = 1: End With
    With AST_TABLE(AST_STAR_GREEDY):    .nc = 1:  .blen = 4: .esfs = 1: End With
    With AST_TABLE(AST_STAR_HUMBLE):    .nc = 1:  .blen = 4: .esfs = 1: End With
    With AST_TABLE(AST_REPEAT_MAX_GREEDY): .nc = 1: .blen = 7: .esfs = 1: End With
    With AST_TABLE(AST_REPEAT_MAX_HUMBLE): .nc = 1: .blen = 7: .esfs = 1: End With
    With AST_TABLE(AST_RANGES):         .nc = -1: .blen = 2: .esfs = 0: End With
    With AST_TABLE(AST_INVRANGES):      .nc = -1: .blen = 2: .esfs = 0: End With
    With AST_TABLE(AST_ASSERT_POS_LOOKAHEAD):   .nc = 1:  .blen = 4: .esfs = 2: End With
    With AST_TABLE(AST_ASSERT_NEG_LOOKAHEAD):   .nc = 1:  .blen = 4: .esfs = 2: End With
    With AST_TABLE(AST_ASSERT_POS_LOOKBEHIND):  .nc = 1:  .blen = 4: .esfs = 2: End With
    With AST_TABLE(AST_ASSERT_NEG_LOOKBEHIND):  .nc = 1:  .blen = 4: .esfs = 2: End With
    With AST_TABLE(AST_FAIL):           .nc = 0:  .blen = 1: .esfs = 0: End With
    With AST_TABLE(AST_BACKREFERENCE):  .nc = 0:  .blen = 2: .esfs = 0: End With
    
    AST_TABLE_INITIALIZED = True
End Sub

Private Sub AstToBytecode(ByRef ast() As Long, ByRef bytecode() As Long)
    Dim bytecodePtr As Long
    Dim curNode As Long, prevNode As Long
    Dim stack() As Long, sp As Long
    Dim direction As Long ' 0 = left before right, -1 = right before left
    Dim returningFromFirstChild As Long ' 0 = no, LONGTYPE_FIRST_BIT = yes
    
    ' temporaries, do not survive over more than one iteration
    Dim opcode1 As Long, opcode2 As Long, opcode3 As Long, tmp As Long, tmpCnt As Long, _
        e As Long, j As Long, patchPos As Long, maxSave As Long
    
    If Not AST_TABLE_INITIALIZED Then InitializeAstTable
    
    PrepareStackAndBytecodeBuffer ast, stack, bytecode
    
    sp = 0
    
    prevNode = -1
    curNode = ast(0) ' first word contains index of root
    bytecodePtr = 1
    maxSave = -1
    direction = 0
    returningFromFirstChild = 0

ContinueLoop:
        Select Case ast(curNode + NODE_TYPE)
        Case AST_STRING
            tmpCnt = ast(curNode + 1) ' assert(tmpCnt >= 1)
            j = curNode + 2 + ((tmpCnt - 1) And direction)
            e = curNode + 1 + tmpCnt - ((tmpCnt - 1) And direction)
            tmp = 1 + 2 * direction
            Do
                bytecode(bytecodePtr) = REOP_CHAR: bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(j): bytecodePtr = bytecodePtr + 1
                If j = e Then Exit Do
                j = j + tmp
            Loop
            GoTo TurnToParent
        Case AST_RANGES
            opcode1 = REOP_RANGES
            GoTo HandleRanges
        Case AST_INVRANGES
            opcode1 = REOP_INVRANGES
            GoTo HandleRanges
        Case AST_CHAR
            bytecode(bytecodePtr) = REOP_CHAR: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = ast(curNode + 1): bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_PERIOD
            bytecode(bytecodePtr) = REOP_PERIOD: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_MATCH
            bytecode(bytecodePtr) = REOP_MATCH: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_START
            bytecode(bytecodePtr) = REOP_ASSERT_START: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_END
            bytecode(bytecodePtr) = REOP_ASSERT_END: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_WORD_BOUNDARY
            bytecode(bytecodePtr) = REOP_ASSERT_WORD_BOUNDARY: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_NOT_WORD_BOUNDARY
            bytecode(bytecodePtr) = REOP_ASSERT_NOT_WORD_BOUNDARY: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_DISJ
            If returningFromFirstChild Then ' previous was left child
                sp = sp - 1: patchPos = stack(sp)
                bytecode(bytecodePtr) = REOP_JUMP: bytecodePtr = bytecodePtr + 1
                stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
                bytecode(patchPos) = bytecodePtr - patchPos - 1
                
                GoTo TurnToRightChild
            ElseIf prevNode = ast(curNode + NODE_RCHILD) Then ' previous was right child
                sp = sp - 1: patchPos = stack(sp)
                bytecode(patchPos) = bytecodePtr - patchPos - 1
            
                GoTo TurnToParent
            Else ' previous was parent
                bytecode(bytecodePtr) = REOP_SPLIT1: bytecodePtr = bytecodePtr + 1
                stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
                GoTo TurnToLeftChild
            End If
        Case AST_CONCAT
            If returningFromFirstChild Then ' previous was first child
                GoTo TurnToSecondChild
            ElseIf prevNode = ast(curNode + NODE_RCHILD + direction) Then ' previous was second child
                GoTo TurnToParent
            Else ' previous was parent
                GoTo TurnToFirstChild
            End If
        Case AST_CAPTURE
            If returningFromFirstChild Then
                bytecode(bytecodePtr) = REOP_SAVE: bytecodePtr = bytecodePtr + 1
                tmp = ast(curNode + 2) * 2 + 1
                If tmp > maxSave Then maxSave = tmp
                bytecode(bytecodePtr) = tmp + direction: bytecodePtr = bytecodePtr + 1
                GoTo TurnToParent
            Else
                bytecode(bytecodePtr) = REOP_SAVE: bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(curNode + 2) * 2 - direction: bytecodePtr = bytecodePtr + 1
                GoTo TurnToLeftChild
            End If
        Case AST_REPEAT_EXACTLY
            opcode1 = REOP_REPEAT_EXACTLY_INIT: opcode2 = REOP_REPEAT_EXACTLY_START: opcode3 = REOP_REPEAT_EXACTLY_END
            GoTo HandleRepeatQuantified
        Case AST_REPEAT_MAX_GREEDY
            opcode1 = REOP_REPEAT_GREEDY_MAX_INIT: opcode2 = REOP_REPEAT_GREEDY_MAX_START: opcode3 = REOP_REPEAT_GREEDY_MAX_END
            GoTo HandleRepeatQuantified
        Case AST_REPEAT_MAX_HUMBLE
            opcode1 = REOP_REPEAT_MAX_HUMBLE_INIT: opcode2 = REOP_REPEAT_MAX_HUMBLE_START: opcode3 = REOP_REPEAT_MAX_HUMBLE_END
            GoTo HandleRepeatQuantified
        Case AST_ZEROONE_GREEDY
            opcode1 = REOP_SPLIT1
            GoTo HandleZeroone
        Case AST_ZEROONE_HUMBLE
            opcode1 = REOP_SPLIT2
            GoTo HandleZeroone
        Case AST_STAR_GREEDY
            opcode1 = REOP_SPLIT1
            GoTo HandleStar
        Case AST_STAR_HUMBLE
            opcode1 = REOP_SPLIT2
            GoTo HandleStar
        Case AST_ASSERT_POS_LOOKAHEAD
            opcode1 = REOP_LOOKPOS: opcode2 = REOP_END_LOOKPOS
            GoTo HandleLookahead
        Case AST_ASSERT_NEG_LOOKAHEAD
            opcode1 = REOP_LOOKNEG: opcode2 = REOP_END_LOOKNEG
            GoTo HandleLookahead
        Case AST_ASSERT_POS_LOOKBEHIND
            opcode1 = REOP_LOOKPOS: opcode2 = REOP_END_LOOKPOS
            GoTo HandleLookbehind
        Case AST_ASSERT_NEG_LOOKBEHIND
            opcode1 = REOP_LOOKNEG: opcode2 = REOP_END_LOOKNEG
            GoTo HandleLookbehind
        Case AST_EMPTY
            GoTo TurnToParent
        Case AST_FAIL
            bytecode(bytecodePtr) = REOP_FAIL: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_BACKREFERENCE
            bytecode(bytecodePtr) = REOP_BACKREFERENCE: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = ast(curNode + 1): bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        End Select
        
        Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR ' unreachable
        
HandleRanges: ' requires: opcode1
        tmpCnt = ast(curNode + 1)
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        j = curNode
        e = curNode + 1 + 2 * tmpCnt
        Do ' copy everything, including first word, which is the length
            j = j + 1
            bytecode(bytecodePtr) = ast(j): bytecodePtr = bytecodePtr + 1
        Loop Until j = e
        GoTo TurnToParent

HandleRepeatQuantified: ' requires: opcode1, opcode2, opcode 3
        tmpCnt = ast(curNode + 2)
        If returningFromFirstChild Then
            sp = sp - 1: patchPos = stack(sp)
            bytecode(bytecodePtr) = opcode3: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
            tmp = bytecodePtr - patchPos
            bytecode(bytecodePtr) = tmp: bytecodePtr = bytecodePtr + 1
            bytecode(patchPos) = tmp
            GoTo TurnToParent
        Else
            bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = opcode2: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
            stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
            GoTo TurnToLeftChild
        End If

HandleZeroone: ' requires: opcode1
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleStar:
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        tmp = bytecodePtr - patchPos + 1
        bytecode(bytecodePtr) = REOP_JUMP: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = -(tmp + 2): bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = tmp
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleLookahead: ' requires opcode1, opcode2
    If returningFromFirstChild Then
        sp = sp - 1: direction = stack(sp)
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = opcode2: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = REOP_CHECK_LOOKAHEAD: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        stack(sp) = direction: sp = sp + 1
        direction = 0
        GoTo TurnToLeftChild
    End If

HandleLookbehind: ' requires opcode1, opcode2
    If returningFromFirstChild Then
        sp = sp - 1: direction = stack(sp)
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = opcode2: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = REOP_CHECK_LOOKBEHIND: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        stack(sp) = direction: sp = sp + 1
        direction = -1
        GoTo TurnToLeftChild
    End If

TurnToParent:
    prevNode = curNode
    If sp = 0 Then GoTo BreakLoop
    sp = sp - 1: tmp = stack(sp)
    curNode = tmp And LONGTYPE_ALL_BUT_FIRST_BIT: returningFromFirstChild = tmp And LONGTYPE_FIRST_BIT
    GoTo ContinueLoop
TurnToLeftChild:
    prevNode = curNode
    stack(sp) = curNode Or LONGTYPE_FIRST_BIT: sp = sp + 1
    curNode = ast(curNode + NODE_LCHILD): returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToRightChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    curNode = ast(curNode + NODE_RCHILD): returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToFirstChild:
    prevNode = curNode
    stack(sp) = curNode Or LONGTYPE_FIRST_BIT: sp = sp + 1
    curNode = ast(curNode + NODE_LCHILD - direction): returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToSecondChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    curNode = ast(curNode + NODE_RCHILD + direction): returningFromFirstChild = 0
    GoTo ContinueLoop
    
BreakLoop:
    bytecode(0) = maxSave
    bytecode(bytecodePtr) = REOP_MATCH
End Sub

' In this function, we allocate stack frames of the same size as we will in bytecode generation.
' Thus we simulate stack usage and make sure the stack is sufficiently large for bytecode generation.
' This means we can abstain from checking stack capacities later.
Private Sub PrepareStackAndBytecodeBuffer(ByRef ast() As Long, ByRef stack() As Long, ByRef bytecode() As Long)
    Dim sp As Long, prevNode As Long, curNode As Long, esfs As Long, stackCapacity As Long
    Dim tmp As Long
    Dim bytecodeLength As Long
    Dim returningFromFirstChild As Long ' 0 = no, LONGTYPE_FIRST_BIT = yes
    
    stackCapacity = INITIAL_STACK_CAPACITY
    ReDim stack(0 To INITIAL_STACK_CAPACITY - 1) As Long

    sp = 0
    
    prevNode = -1
    curNode = ast(0) ' first word contains index of root
    returningFromFirstChild = 0

    bytecodeLength = 0
    
ContinueLoop:
        With AST_TABLE(ast(curNode + NODE_TYPE))
            esfs = .esfs
            
            Select Case .nc
            Case -2
                bytecodeLength = bytecodeLength + 2 * ast(curNode + 1)
                GoTo TurnToParent
            Case -1
                bytecodeLength = bytecodeLength + 2 + 2 * ast(curNode + 1)
                GoTo TurnToParent
            Case 0
                bytecodeLength = bytecodeLength + .blen
                GoTo TurnToParent
            Case 1
                If returningFromFirstChild Then
                    GoTo TurnToParent
                Else
                    bytecodeLength = bytecodeLength + .blen
                    GoTo TurnToLeftChild
                End If
            Case 2
                If returningFromFirstChild Then ' previous was left child
                    GoTo TurnToRightChild
                ElseIf prevNode = ast(curNode + NODE_RCHILD) Then ' previous was right child
                    GoTo TurnToParent
                Else ' previous was parent
                    bytecodeLength = bytecodeLength + .blen
                    GoTo TurnToLeftChild
                End If
            End Select
        End With
        
        Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR ' unreachable

TurnToParent:
    sp = sp - esfs
    If sp = 0 Then GoTo BreakLoop
    prevNode = curNode
    sp = sp - 1: tmp = stack(sp)
    returningFromFirstChild = tmp And LONGTYPE_FIRST_BIT: curNode = tmp And LONGTYPE_ALL_BUT_FIRST_BIT
    GoTo ContinueLoop
TurnToLeftChild:
    If sp >= stackCapacity - esfs Then
        stackCapacity = stackCapacity + stackCapacity \ 2
        ReDim Preserve stack(0 To stackCapacity - 1) As Long
    End If
    prevNode = curNode
    sp = sp + esfs: stack(sp) = curNode Or LONGTYPE_FIRST_BIT: sp = sp + 1
    returningFromFirstChild = 0: curNode = ast(curNode + NODE_LCHILD)
    GoTo ContinueLoop
TurnToRightChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    returningFromFirstChild = 0: curNode = ast(curNode + NODE_RCHILD)
    GoTo ContinueLoop
    
BreakLoop:
    ' Actual bytecode length is bytecodeLength + 2 due to intial nCaptures and final REOP_MATCH.
    ReDim bytecode(0 To bytecodeLength + 1) As Long
End Sub
Private Sub InitializeRangeTables()
    RangeTableDigit(0) = &H30&
    RangeTableDigit(1) = &H39&
    
    '---------------------------------
    
    RangeTableWhite(0) = &H9&
    RangeTableWhite(1) = &HD&
    RangeTableWhite(2) = &H20&
    RangeTableWhite(3) = &H20&
    RangeTableWhite(4) = &HA0&
    RangeTableWhite(5) = &HA0&
    RangeTableWhite(6) = &H1680&
    RangeTableWhite(7) = &H1680&
    RangeTableWhite(8) = &H180E&
    RangeTableWhite(9) = &H180E&
    RangeTableWhite(10) = &H2000&
    RangeTableWhite(11) = &H200A
    RangeTableWhite(12) = &H2028&
    RangeTableWhite(13) = &H2029&
    RangeTableWhite(14) = &H202F
    RangeTableWhite(15) = &H202F
    RangeTableWhite(16) = &H205F&
    RangeTableWhite(17) = &H205F&
    RangeTableWhite(18) = &H3000&
    RangeTableWhite(19) = &H3000&
    RangeTableWhite(20) = &HFEFF&
    RangeTableWhite(21) = &HFEFF&
    
    '---------------------------------
    
    RangeTableWordchar(0) = &H30&
    RangeTableWordchar(1) = &H39&
    RangeTableWordchar(2) = &H41&
    RangeTableWordchar(3) = &H5A&
    RangeTableWordchar(4) = &H5F&
    RangeTableWordchar(5) = &H5F&
    RangeTableWordchar(6) = &H61&
    RangeTableWordchar(7) = &H7A&
    
    '---------------------------------
    
    RangeTableNotDigit(0) = MIN_LONG
    RangeTableNotDigit(1) = &H2F&
    RangeTableNotDigit(2) = &H3A&
    RangeTableNotDigit(3) = MAX_LONG
    
    '---------------------------------
    
    RangeTableNotWhite(0) = MIN_LONG
    RangeTableNotWhite(1) = &H8&
    RangeTableNotWhite(2) = &HE&
    RangeTableNotWhite(3) = &H1F&
    RangeTableNotWhite(4) = &H21&
    RangeTableNotWhite(5) = &H9F&
    RangeTableNotWhite(6) = &HA1&
    RangeTableNotWhite(7) = &H167F&
    RangeTableNotWhite(8) = &H1681&
    RangeTableNotWhite(9) = &H180D&
    RangeTableNotWhite(10) = &H180F&
    RangeTableNotWhite(11) = &H1FFF&
    RangeTableNotWhite(12) = &H200B&
    RangeTableNotWhite(13) = &H2027&
    RangeTableNotWhite(14) = &H202A&
    RangeTableNotWhite(15) = &H202E&
    RangeTableNotWhite(16) = &H2030&
    RangeTableNotWhite(17) = &H205E&
    RangeTableNotWhite(18) = &H2060&
    RangeTableNotWhite(19) = &H2FFF&
    RangeTableNotWhite(20) = &H3001&
    RangeTableNotWhite(21) = &HFEFE&
    RangeTableNotWhite(22) = &HFF00&
    RangeTableNotWhite(23) = MAX_LONG
    
    '---------------------------------
    
    RangeTableNotWordChar(0) = MIN_LONG
    RangeTableNotWordChar(1) = &H2F&
    RangeTableNotWordChar(2) = &H3A&
    RangeTableNotWordChar(3) = &H40&
    RangeTableNotWordChar(4) = &H5B&
    RangeTableNotWordChar(5) = &H5E&
    RangeTableNotWordChar(6) = &H60&
    RangeTableNotWordChar(7) = &H60&
    RangeTableNotWordChar(8) = &H7B&
    RangeTableNotWordChar(9) = MAX_LONG
    
    '---------------------------------
    
    RangeTablesInitialized = True
End Sub

Sub EmitPredefinedRange(ByRef outBuffer As ArrayBuffer, ByRef source() As Long)
    Dim i As Long, j As Long, sourceLower As Long, sourceUpper As Long
    
    With outBuffer
        i = .length
        sourceLower = LBound(source)
        sourceUpper = UBound(source)
        AppendUnspecified outBuffer, sourceUpper - sourceLower + 1
        j = sourceUpper - 1
        Do While j >= sourceLower
            .buffer(i) = source(j): i = i + 1
            .buffer(i) = source(j + 1): i = i + 1
            j = j - 2
        Loop
    End With
End Sub

Function UnicodeIsIdentifierPart(x As Long) As Boolean
    UnicodeIsIdentifierPart = False
End Function

Sub RegexpGenerateRanges(ByRef outBuffer As ArrayBuffer, _
    ByVal ignoreCase As Boolean, ByVal r1 As Long, ByVal r2 As Long _
)
    Dim a As Long, b As Long, m As Long, d As Long, ub As Long, lastDelta As Long
    Dim r As Long, rc As Long

    If Not ignoreCase Then
        AppendLong outBuffer, r1
        AppendLong outBuffer, r2
        Exit Sub
    End If
        
    rc = ReCanonicalizeChar(r1)
    lastDelta = rc - r1
    AppendLong outBuffer, rc

    a = LBound(UnicodeCanonRunsTable) - 1
    ub = UBound(UnicodeCanonRunsTable)
    
    If UnicodeCanonRunsTable(ub) > r1 Then
        ' Find the index of the first element larger than r1.
        ' The index is guaranteed to be in the interval (a;b].
        b = ub
        Do
            d = b - a
            If d = 1 Then Exit Do
            m = a + d \ 2
            If UnicodeCanonRunsTable(m) > r1 Then b = m Else a = m
        Loop
        
        ' Now b is the index of the first element larger than r1.
        Do
            r = UnicodeCanonRunsTable(b)
            If r > r2 Then Exit Do
            AppendLong outBuffer, r - 1 + lastDelta
            
            rc = ReCanonicalizeChar(r)
            AppendLong outBuffer, rc
            lastDelta = rc - r

            If b = ub Then Exit Do
            b = b + 1
        Loop
    End If
    
    AppendLong outBuffer, r2 + lastDelta
End Sub
Private Sub Initialize(ByRef lexCtx As LexerContext, ByRef inputStr As String)
    With lexCtx
        .inputArray = inputStr
        
        ' Be careful, if inputStr = "", then .iEnd = -2
        .iEnd = UBound(.inputArray) - 1
        If .iEnd And 1 Then Err.Raise REGEX_ERR_INVALID_INPUT
    
        .iCurrent = -2
        .currentCharacter = Not ENDOFINPUT ' value does not matter, as long as it does not equal ENDOFINPUT
    End With
    Advance lexCtx
End Sub

'
'  Parse a RegExp token.  The grammar is described in E5 Section 15.10.
'  Terminal constructions (such as quantifiers) are parsed directly here.
'
'  0xffffffffU is used as a marker for "infinity" in quantifiers.  Further,
'  _MAX_RE_QUANT_DIGITS limits the maximum number of digits that
'  will be accepted for a quantifier.
'
Private Sub ParseReToken(ByRef lexCtx As LexerContext, ByRef outToken As ReToken)
    Dim x As Long
    
    ' used only locally
    Dim i As Long, val1 As Long, val2 As Long, digits As Long, tmp As Long

    outToken = EMPTY_RE_TOKEN

    x = Advance(lexCtx)
    Select Case x
    Case UNICODE_PIPE
        outToken.t = RETOK_DISJUNCTION
    Case UNICODE_CARET
        outToken.t = RETOK_ASSERT_START
    Case UNICODE_DOLLAR
        outToken.t = RETOK_ASSERT_END
    Case UNICODE_QUESTION
        With outToken
            .qmin = 0
            .qmax = 1
            If lexCtx.currentCharacter = UNICODE_QUESTION Then
                Advance lexCtx
                .t = RETOK_QUANTIFIER
                .greedy = False
            Else
                .t = RETOK_QUANTIFIER
                .greedy = True
            End If
        End With
    Case UNICODE_STAR
        With outToken
            .qmin = 0
            .qmax = RE_QUANTIFIER_INFINITE
            If lexCtx.currentCharacter = UNICODE_QUESTION Then
                Advance lexCtx
                .t = RETOK_QUANTIFIER
                .greedy = False
            Else
                .t = RETOK_QUANTIFIER
                .greedy = True
            End If
        End With
    Case UNICODE_PLUS
        With outToken
            .qmin = 1
            .qmax = RE_QUANTIFIER_INFINITE
            If lexCtx.currentCharacter = UNICODE_QUESTION Then
                Advance lexCtx
                .t = RETOK_QUANTIFIER
                .greedy = False
            Else
                .t = RETOK_QUANTIFIER
                .greedy = True
            End If
        End With
    Case UNICODE_LCURLY
        ' Production allows 'DecimalDigits', including leading zeroes
        val1 = 0
        val2 = RE_QUANTIFIER_INFINITE
        
        digits = 0

        Do
            x = Advance(lexCtx)
            If (x >= UNICODE_0) And (x <= UNICODE_9) Then
                digits = digits + 1
                ' Be careful to prevent overflow
                If val1 > MAX_LONG_DIV_10 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                val1 = val1 * 10
                tmp = x - UNICODE_0
                If MAX_LONG - val1 < tmp Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                val1 = val1 + tmp
            ElseIf x = UNICODE_COMMA Then
                If val2 <> RE_QUANTIFIER_INFINITE Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                If lexCtx.currentCharacter = UNICODE_RCURLY Then
                    ' form: { DecimalDigits , }, val1 = min count
                    If digits = 0 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                    outToken.qmin = val1
                    outToken.qmax = RE_QUANTIFIER_INFINITE
                    Advance lexCtx
                    Exit Do
                End If
                val2 = val1
                val1 = 0
                digits = 0 ' not strictly necessary because of lookahead '}' above
            ElseIf x = UNICODE_RCURLY Then
                If digits = 0 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                If val2 <> RE_QUANTIFIER_INFINITE Then
                    ' val2 = min count, val1 = max count
                    outToken.qmin = val2
                    outToken.qmax = val1
                Else
                    ' val1 = count
                    outToken.qmin = val1
                    outToken.qmax = val1
                End If
                Exit Do
            Else
                Err.Raise REGEX_ERR_INVALID_QUANTIFIER
            End If
        Loop
        If lexCtx.currentCharacter = UNICODE_QUESTION Then
            outToken.greedy = False
            Advance lexCtx
        Else
            outToken.greedy = True
        End If
        outToken.t = RETOK_QUANTIFIER
    Case UNICODE_PERIOD
        outToken.t = RETOK_ATOM_PERIOD
    Case UNICODE_BACKSLASH
        ' The E5.1 specification does not seem to allow IdentifierPart characters
        ' to be used as identity escapes.  Unfortunately this includes '$', which
        ' cannot be escaped as '\$'; it needs to be escaped e.g. as '\u0024'.
        ' Many other implementations (including V8 and Rhino, for instance) do
        ' accept '\$' as a valid identity escape, which is quite pragmatic, and
        ' ES2015 Annex B relaxes the rules to allow these (and other) real world forms.
        x = Advance(lexCtx)
        Select Case x
        Case UNICODE_LC_B
            outToken.t = RETOK_ASSERT_WORD_BOUNDARY
        Case UNICODE_UC_B
            outToken.t = RETOK_ASSERT_NOT_WORD_BOUNDARY
        Case UNICODE_LC_F
            outToken.num = &HC&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_N
            outToken.num = &HA&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_T
            outToken.num = &H9&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_R
            outToken.num = &HD&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_V
            outToken.num = &HB&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_C
            x = Advance(lexCtx)
            If (x >= UNICODE_LC_A And x <= UNICODE_LC_Z) Or (x >= UNICODE_UC_A And x <= UNICODE_UC_Z) Then
                outToken.num = x \ 32
                outToken.t = RETOK_ATOM_CHAR
            Else
                Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
            End If
        Case UNICODE_LC_X
            outToken.num = LexerParseEscapeX(lexCtx)
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_U
            ' Todo: What does the following mean?
            ' The token value is the Unicode codepoint without
            ' it being decode into surrogate pair characters
            ' here.  The \u{H+} is only allowed in Unicode mode
            ' which we don't support yet.
            outToken.num = LexerParseEscapeU(lexCtx)
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_D
            outToken.t = RETOK_ATOM_DIGIT
        Case UNICODE_UC_D
            outToken.t = RETOK_ATOM_NOT_DIGIT
        Case UNICODE_LC_S
            outToken.t = RETOK_ATOM_WHITE
        Case UNICODE_UC_S
            outToken.t = RETOK_ATOM_NOT_WHITE
        Case UNICODE_LC_W
            outToken.t = RETOK_ATOM_WORD_CHAR
        Case UNICODE_UC_W
            outToken.t = RETOK_ATOM_NOT_WORD_CHAR
        Case UNICODE_0
            x = Advance(lexCtx)
            
            ' E5 Section 15.10.2.11
            If x >= UNICODE_0 And x <= UNICODE_9 Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
            outToken.num = 0
            outToken.t = RETOK_ATOM_CHAR
        Case Else
            If x >= UNICODE_1 And x <= UNICODE_9 Then
                val1 = 0
                i = 0
                Do
                    ' We have to be careful here to make sure there will be no overflow.
                    ' 2^31 - 1 backreferences is a bit ridiculous, though.
                    If val1 > MAX_LONG_DIV_10 Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
                    val1 = val1 * 10
                    tmp = x - UNICODE_0
                    If MAX_LONG - val1 < tmp Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
                    val1 = val1 + tmp
                    x = lexCtx.currentCharacter
                    If x < UNICODE_0 Or x > UNICODE_9 Then Exit Do
                    Advance lexCtx
                    i = i + 1
                Loop
                outToken.t = RETOK_ATOM_BACKREFERENCE
                outToken.num = val1
            ElseIf (x >= 0 And Not UnicodeIsIdentifierPart(0)) Or x = UNICODE_CP_ZWNJ Or x = UNICODE_CP_ZWJ Then
                ' For ES5.1 identity escapes are not allowed for identifier
                ' parts.  This conflicts with a lot of real world code as this
                ' doesn't e.g. allow escaping a dollar sign as /\$/.
                outToken.num = x
                outToken.t = RETOK_ATOM_CHAR
            Else
                Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
            End If
        End Select
    Case UNICODE_LPAREN
        If lexCtx.currentCharacter = UNICODE_QUESTION Then
            Advance lexCtx
            x = Advance(lexCtx)
            Select Case x
            Case UNICODE_EQUALS
                ' (?=
                outToken.t = RETOK_ASSERT_START_POS_LOOKAHEAD
            Case UNICODE_EXCLAMATION
                ' (?!
                outToken.t = RETOK_ASSERT_START_NEG_LOOKAHEAD
            Case UNICODE_COLON
                ' (?:
                outToken.t = RETOK_ATOM_START_NONCAPTURE_GROUP
            Case UNICODE_LT
                x = Advance(lexCtx)
                If x = UNICODE_EQUALS Then
                    outToken.t = RETOK_ASSERT_START_POS_LOOKBEHIND
                ElseIf x = UNICODE_EXCLAMATION Then
                    outToken.t = RETOK_ASSERT_START_NEG_LOOKBEHIND
                Else
                    Err.Raise REGEX_ERR_INVALID_REGEXP_GROUP
                End If
            Case Else
                Err.Raise REGEX_ERR_INVALID_REGEXP_GROUP
            End Select
        Else
            ' (
            outToken.t = RETOK_ATOM_START_CAPTURE_GROUP
        End If
    Case UNICODE_RPAREN
        outToken.t = RETOK_ATOM_END
    Case UNICODE_LBRACKET
        ' To avoid creating a heavy intermediate value for the list of ranges,
        ' only the start token ('[' or '[^') is parsed here.  The regexp
        ' compiler parses the ranges itself.
        If lexCtx.currentCharacter = UNICODE_CARET Then
            Advance lexCtx
            outToken.t = RETOK_ATOM_START_CHARCLASS_INVERTED
        Else
            outToken.t = RETOK_ATOM_START_CHARCLASS
        End If
    Case UNICODE_RCURLY, UNICODE_RBRACKET
        ' Although these could be parsed as PatternCharacters unambiguously (here),
        ' * E5 Section 15.10.1 grammar explicitly forbids these as PatternCharacters.
        Err.Raise REGEX_ERR_INVALID_REGEXP_CHARACTER
    Case ENDOFINPUT
        ' EOF
        outToken.t = RETOK_EOF
    Case Else
        ' PatternCharacter, all excluded characters are matched by cases above
        outToken.t = RETOK_ATOM_CHAR
        outToken.num = x
    End Select
End Sub

Private Sub ParseReRanges(lexCtx As LexerContext, ByRef outBuffer As ArrayBuffer, ByRef nranges As Long, ByVal ignoreCase As Boolean)
    Dim start As Long, ch As Long, x As Long, dash As Boolean, y As Long, bufferStart As Long
    
    bufferStart = outBuffer.length
    
    ' start is -2 at the very beginning of the range expression,
    '   -1 when we have not seen a possible "start" character,
    '   and it equals the possible start character if we have seen one
    start = -2
    dash = False
    
    Do
ContinueLoop:
        x = Advance(lexCtx)

        If x < 0 Then GoTo FailUntermCharclass
        
        Select Case x
        Case UNICODE_RBRACKET
            If start >= 0 Then
                RegexpGenerateRanges outBuffer, ignoreCase, start, start
                Exit Do
            ElseIf start = -1 Then
                Exit Do
            Else ' start = -2
                ' ] at the very beginning of a range expression is interpreted literally,
                '   since empty ranges are not permitted.
                '   This corresponds to what RE2 does.
                ch = x
            End If
        Case UNICODE_MINUS
            If start >= 0 Then
                If Not dash Then
                    If lexCtx.currentCharacter <> UNICODE_RBRACKET Then
                        ' '-' as a range indicator
                        dash = True
                        GoTo ContinueLoop
                    End If
                End If
            End If
            ' '-' verbatim
            ch = x
        Case UNICODE_BACKSLASH
            '
            '  The escapes are same as outside a character class, except that \b has a
            '  different meaning, and \B and backreferences are prohibited (see E5
            '  Section 15.10.2.19).  However, it's difficult to share code because we
            '  handle e.g. "\n" very differently: here we generate a single character
            '  range for it.
            '

            ' XXX: ES2015 surrogate pair handling.

            x = Advance(lexCtx)

            Select Case x
            Case UNICODE_LC_B
                ' Note: '\b' in char class is different than outside (assertion),
                ' '\B' is not allowed and is caught by the duk_unicode_is_identifier_part()
                ' check below.
                '
                ch = &H8&
            Case x = UNICODE_LC_F
                ch = &HC&
            Case UNICODE_LC_N
                ch = &HA&
            Case UNICODE_LC_T
                ch = &H9&
            Case UNICODE_LC_R
                ch = &HD&
            Case UNICODE_LC_V
                ch = &HB&
            Case UNICODE_LC_C
                x = Advance(lexCtx)
                If ((x >= UNICODE_LC_A And x <= UNICODE_LC_Z) Or (x >= UNICODE_UC_A And x <= UNICODE_UC_Z)) Then
                    ch = x Mod 32
                Else
                    GoTo FailEscape
                End If
            Case UNICODE_LC_X
                ch = LexerParseEscapeX(lexCtx)
            Case UNICODE_LC_U
                ch = LexerParseEscapeU(lexCtx)
            Case UNICODE_LC_D
                EmitPredefinedRange outBuffer, RangeTableDigit
                ch = -1
            Case UNICODE_UC_D
                EmitPredefinedRange outBuffer, RangeTableNotDigit
                ch = -1
            Case UNICODE_LC_S
                EmitPredefinedRange outBuffer, RangeTableWhite
                ch = -1
            Case UNICODE_UC_S
                EmitPredefinedRange outBuffer, RangeTableNotWhite
                ch = -1
            Case UNICODE_LC_W
                EmitPredefinedRange outBuffer, RangeTableWordchar
                ch = -1
            Case UNICODE_UC_W
                EmitPredefinedRange outBuffer, RangeTableNotWordChar
                ch = -1
            Case Else
                If x < 0 Then GoTo FailEscape
                If x <= UNICODE_7 Then
                    If x >= UNICODE_0 Then
                        ' \0 or octal escape from \0 up to \377
                        ch = LexerParseLegacyOctal(lexCtx, x)
                    Else
                        ' IdentityEscape: ES2015 Annex B allows almost all
                        ' source characters here.  Match anything except
                        ' EOF here.
                        ch = x
                    End If
                Else
                    ' IdentityEscape: ES2015 Annex B allows almost all
                    ' source characters here.  Match anything except
                    ' EOF here.
                    ch = x
                End If
            End Select
        Case Else
            ' character represents itself
            ch = x
        End Select

        ' ch is a literal character here or -1 if parsed entity was
        ' an escape such as "\s".
        '

        If ch < 0 Then
            ' multi-character sets not allowed as part of ranges, see
            ' E5 Section 15.10.2.15, abstract operation CharacterRange.
            '
            If start >= 0 Then
                If dash Then
                    GoTo FailRange
                Else
                    RegexpGenerateRanges outBuffer, ignoreCase, start, start
                End If
            End If
            start = -1
            ' dash is already 0
        Else
            If start >= 0 Then
                If dash Then
                    If start > ch Then GoTo FailRange
                    RegexpGenerateRanges outBuffer, ignoreCase, start, ch
                    start = -1
                    dash = 0
                Else
                    RegexpGenerateRanges outBuffer, ignoreCase, start, start
                    start = ch
                    ' dash is already 0
                End If
            Else
                start = ch
            End If
        End If
    Loop

    If outBuffer.length - 2 > bufferStart Then
        ' We have at least 2 intervals.
        HeapsortPairs outBuffer.buffer, bufferStart, outBuffer.length - 2
        outBuffer.length = 2 + Unionize(outBuffer.buffer, bufferStart, outBuffer.length - 2)
    End If
    
    nranges = (outBuffer.length - bufferStart) \ 2
    
    Exit Sub

FailEscape:
    Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE

FailRange:
    Err.Raise REGEX_ERR_INVALID_RANGE

FailUntermCharclass:
    Err.Raise REGEX_ERR_UNTERMINATED_CHARCLASS
End Sub

' input: array of pairs [x0, y0, x1, y1, ..., x(n-1), y(n-1)]
'   sorts pairs by first entry, second entry irrelevant for order
Private Sub HeapsortPairs(ByRef ary() As Long, ByVal b As Long, ByVal t As Long)
    Dim bb As Long
    Dim parent As Long, child As Long
    Dim smallestValueX As Long, smallestValueY As Long, tmpX As Long, tmpY As Long
    
    ' build heap
    ' bb marks the next element to be added to the heap
    bb = t - 2
    Do Until bb < b
        child = bb
        Do Until child = t
            parent = child + 2 + 2 * ((t - child) \ 4)
            If ary(parent) <= ary(child) Then Exit Do
            tmpX = ary(parent): tmpY = ary(parent + 1)
            ary(parent) = ary(child): ary(parent + 1) = ary(child + 1)
            ary(child) = tmpX: ary(child + 1) = tmpY
            child = parent
        Loop
        bb = bb - 2
    Loop

    ' demount heap
    ' bb marks the lower end of the remaining heap
    bb = b
    Do While bb < t
        smallestValueX = ary(t): smallestValueY = ary(t + 1)
        
        parent = t
        Do
            child = parent - t + parent - 2
            
            ' if there are no children, we are finished
            If child < bb Then Exit Do
            
            ' if there are two children, prefer the one with the smaller value
            If child > bb Then child = child + 2 * (ary(child - 2) <= ary(child))
            
            ary(parent) = ary(child): ary(parent + 1) = ary(child + 1)
            parent = child
        Loop
        
        ' now position parent is free
        
        ' if parent <> bb, free bb rather than parent
        ' by swapping the values in parent and bb and repairing the heap bottom-up
        If parent > bb Then
            ary(parent) = ary(bb): ary(parent + 1) = ary(bb + 1)
            child = parent
            Do Until child = t
                parent = child + 2 + 2 * ((t - child) \ 4)
                If ary(parent) <= ary(child) Then Exit Do
                tmpX = ary(parent): tmpY = ary(parent + 1)
                ary(parent) = ary(child): ary(parent + 1) = ary(child + 1)
                ary(child) = tmpX: ary(child + 1) = tmpY
                child = parent
            Loop
        End If
        
        ' now position bb is free
        
        ary(bb) = smallestValueX: ary(bb + 1) = smallestValueY
        bb = bb + 2
    Loop
End Sub

' assert: t > b
' return: index of first element of last pair
Private Function Unionize(ByRef ary() As Long, ByVal b As Long, ByVal t As Long)
    Dim i As Long, j As Long, lower As Long, upper As Long, nextLower As Long, nextUpper As Long
    
    lower = ary(b): upper = ary(b + 1)
    j = b
    For i = b + 2 To t Step 2
        nextLower = ary(i): nextUpper = ary(i + 1)
        If nextLower <= upper + 1 Then
            If nextUpper > upper Then upper = nextUpper
        Else
            ary(j) = lower: j = j + 1: ary(j) = upper: j = j + 1
            lower = nextLower: upper = nextUpper
        End If
    Next
    ary(j) = lower: ary(j + 1) = upper
    Unionize = j
End Function

' Parse a Unicode escape of the form \xHH.
Private Function LexerParseEscapeX(ByRef lexCtx As LexerContext) As Long
    Dim dig As Long, escval As Long, x As Long
    
    x = Advance(lexCtx)
    dig = HexvalValidate(x)
    If dig < 0 Then GoTo FailEscape
    escval = dig
    
    x = Advance(lexCtx)
    dig = HexvalValidate(x)
    If dig < 0 Then GoTo FailEscape
    escval = escval * 16 + dig
    
    LexerParseEscapeX = escval
    Exit Function
    
FailEscape:
    Err.Raise REGEX_ERR_INVALID_ESCAPE
End Function

' Parse a Unicode escape of the form \uHHHH, or \u{H+}.
Private Function LexerParseEscapeU(ByRef lexCtx As LexerContext) As Long
    Dim dig As Long, escval As Long, x As Long
    
    If lexCtx.currentCharacter = UNICODE_LCURLY Then
        Advance lexCtx
        
        escval = 0
        x = Advance(lexCtx)
        If x = UNICODE_RCURLY Then GoTo FailEscape ' Empty escape \u{}
        Do
            dig = HexvalValidate(x)
            If dig < 0 Then GoTo FailEscape
            If escval > &H10FFF Then GoTo FailEscape
            escval = escval * 16 + dig
            
            x = Advance(lexCtx)
        Loop Until x = UNICODE_RCURLY
        LexerParseEscapeU = escval
    Else
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = dig
        
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = escval * 16 + dig
        
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = escval * 16 + dig
        
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = escval * 16 + dig
        
        LexerParseEscapeU = escval
    End If
    Exit Function
    
FailEscape:
    Err.Raise REGEX_ERR_INVALID_ESCAPE
End Function

' If ch is a hex digit, return its value.
' If ch is not a hex digit, return -1.
Private Function HexvalValidate(ByVal ch As Long) As Long
    Const HEX_DELTA_L As Long = UNICODE_LC_A - 10
    Const HEX_DELTA_U As Long = UNICODE_UC_A - 10

    HexvalValidate = -1
    If ch <= UNICODE_UC_F Then
        If ch <= UNICODE_9 Then
            If ch >= UNICODE_0 Then HexvalValidate = ch - UNICODE_0
        Else
            If ch >= UNICODE_UC_A Then HexvalValidate = ch - HEX_DELTA_U
        End If
    Else
        If ch <= UNICODE_LC_F Then
            If ch >= UNICODE_LC_A Then HexvalValidate = ch - HEX_DELTA_L
        End If
    End If
End Function

' Parse legacy octal escape of the form \N{1,3}, e.g. \0, \5, \0377.  Maximum
' allowed value is \0377 (U+00FF), longest match is used.  Used for both string
' RegExp octal escape parsing.
' x is the first digit, which must have already been validated to be in [0-7] by the caller.
'
Private Function LexerParseLegacyOctal(ByRef lexCtx As LexerContext, ByVal x As Long)
    Dim cp As Long, tmp As Long, i As Long

    cp = x - UNICODE_0

    tmp = lexCtx.currentCharacter
    If tmp < UNICODE_0 Then GoTo ExitFunction
    If tmp > UNICODE_7 Then GoTo ExitFunction

    cp = cp * 8 + (tmp - UNICODE_0)
    Advance lexCtx

    If cp > 31 Then GoTo ExitFunction
    
    tmp = lexCtx.currentCharacter
    If tmp < UNICODE_0 Then GoTo ExitFunction
    If tmp > UNICODE_7 Then GoTo ExitFunction

    cp = cp * 8 + (tmp - UNICODE_0)
    Advance lexCtx

ExitFunction:
    LexerParseLegacyOctal = cp
End Function

Private Function Advance(ByRef lexCtx As LexerContext) As Long
    Dim lower As Long, upper As Long, w As Long

    With lexCtx
        Advance = .currentCharacter
        If .currentCharacter = ENDOFINPUT Then Exit Function
        If .iCurrent = .iEnd Then
            .currentCharacter = ENDOFINPUT
        Else
            .iCurrent = .iCurrent + 2
            lower = .inputArray(.iCurrent)
            upper = .inputArray(.iCurrent + 1)
            w = 256 * upper + lower
            .currentCharacter = w
        End If
    End With
End Function

Public Sub Compile(ByRef outBytecode() As Long, ByRef s As String, Optional ByVal flags = RE_FLAG_NONE)
    Dim lex As LexerContext
    Dim ast As ArrayBuffer
    
    If Not UnicodeInitialized Then UnicodeInitialize
    If Not RangeTablesInitialized Then InitializeRangeTables
    
    Initialize lex, s
    Parse lex, flags, ast
    AstToBytecode ast.buffer, outBytecode
End Sub

Private Sub PerformPotentialConcat(ByRef ast As ArrayBuffer, ByRef potentialConcat2 As Long, ByRef potentialConcat1 As Long)
    Dim tmp As Long
    If potentialConcat2 <> -1 Then
        tmp = ast.length
        AppendThree ast, AST_CONCAT, potentialConcat2, potentialConcat1
        potentialConcat2 = -1
        potentialConcat1 = tmp
    End If
End Sub

Private Sub Parse(ByRef lex As LexerContext, ByVal reFlags As Long, ByRef ast As ArrayBuffer)
    Dim currToken As ReToken
    Dim currentAstNode As Long
    Dim potentialConcat2 As Long, potentialConcat1 As Long, pendingDisjunction As Long, currentDisjunction As Long
    Dim nCaptures As Long
    Dim tmp As Long, i As Long, qmin As Long, qmax As Long, n1 As Long, n2 As Long
    Dim parseStack As ArrayBuffer
    
    nCaptures = 0
    
    pendingDisjunction = -1
    potentialConcat2 = -1
    potentialConcat1 = -1
    currentDisjunction = -1
    
    AppendLong ast, 0 ' first word will be index of the root node, to be patched in the end

    Do
ContinueLoop:
        ParseReToken lex, currToken
        
        Select Case currToken.t
        Case RETOK_DISJUNCTION
            If potentialConcat1 = -1 Then
                currentAstNode = ast.length
                AppendLong ast, AST_EMPTY
                potentialConcat1 = currentAstNode
            Else
                PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            End If
            
            currentAstNode = ast.length
            AppendThree ast, AST_DISJ, potentialConcat1, -1
            potentialConcat1 = -1

            If pendingDisjunction <> -1 Then
                ast.buffer(pendingDisjunction + 2) = currentAstNode
            Else
                currentDisjunction = currentAstNode
            End If
        
            pendingDisjunction = currentAstNode
        
        Case RETOK_QUANTIFIER
            If potentialConcat1 = -1 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER_NO_ATOM
            
            qmin = currToken.qmin
            qmax = currToken.qmax
            
            If qmin > qmax Then
                currentAstNode = ast.length
                AppendLong ast, AST_FAIL
                potentialConcat1 = currentAstNode
                GoTo ContinueLoop
            End If
            
            If qmin = 0 Then
                n1 = -1
            ElseIf qmin = 1 Then
                n1 = potentialConcat1
            ElseIf qmin > 1 Then
                currentAstNode = ast.length
                AppendThree ast, AST_REPEAT_EXACTLY, potentialConcat1, qmin
                n1 = currentAstNode
            Else
                Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR
            End If
            
            If qmax = RE_QUANTIFIER_INFINITE Then
                If currToken.greedy Then tmp = AST_STAR_GREEDY Else tmp = AST_STAR_HUMBLE
                currentAstNode = ast.length
                AppendTwo ast, tmp, potentialConcat1
                n2 = currentAstNode
            ElseIf qmax - qmin = 1 Then
                If currToken.greedy Then tmp = AST_ZEROONE_GREEDY Else tmp = AST_ZEROONE_HUMBLE
                currentAstNode = ast.length
                AppendTwo ast, tmp, potentialConcat1
                n2 = currentAstNode
            ElseIf qmax - qmin > 1 Then
                If currToken.greedy Then tmp = AST_REPEAT_MAX_GREEDY Else tmp = AST_REPEAT_MAX_HUMBLE
                currentAstNode = ast.length
                AppendThree ast, tmp, potentialConcat1, qmax - qmin
                n2 = currentAstNode
            ElseIf qmax = qmin Then
                n2 = -1
            Else
                Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR
            End If
            
            If n1 = -1 Then
                If n2 = -1 Then
                    currentAstNode = ast.length
                    AppendLong ast, AST_EMPTY
                    potentialConcat1 = currentAstNode
                Else
                    potentialConcat1 = n2
                End If
            Else
                If n2 = -1 Then
                    potentialConcat1 = n1
                Else
                    currentAstNode = ast.length
                    AppendThree ast, AST_CONCAT, n1, n2
                    potentialConcat1 = currentAstNode
                End If
            End If
            
        Case RETOK_ATOM_START_CAPTURE_GROUP
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            nCaptures = nCaptures + 1
            AppendFour parseStack, _
                nCaptures, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
                
        Case RETOK_ATOM_START_NONCAPTURE_GROUP
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            AppendFour parseStack, _
                -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
            
         Case RETOK_ASSERT_START_POS_LOOKAHEAD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            AppendFour parseStack, _
                -(AST_ASSERT_POS_LOOKAHEAD - MIN_AST_CODE) - 2, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        Case RETOK_ASSERT_START_NEG_LOOKAHEAD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            nCaptures = nCaptures + 1
            AppendFour parseStack, _
                -(AST_ASSERT_NEG_LOOKAHEAD - MIN_AST_CODE) - 2, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
         Case RETOK_ASSERT_START_POS_LOOKBEHIND
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            AppendFour parseStack, _
                -(AST_ASSERT_POS_LOOKBEHIND - MIN_AST_CODE) - 2, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        Case RETOK_ASSERT_START_NEG_LOOKBEHIND
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            nCaptures = nCaptures + 1
            AppendFour parseStack, _
                -(AST_ASSERT_NEG_LOOKBEHIND - MIN_AST_CODE) - 2, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        
        Case RETOK_ATOM_END
            If parseStack.length = 0 Then Err.Raise REGEX_ERR_UNEXPECTED_CLOSING_PAREN
            
            ' Close disjunction
            If potentialConcat1 = -1 Then
                currentAstNode = ast.length
                AppendLong ast, AST_EMPTY
                potentialConcat1 = currentAstNode
            Else
                PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            End If
            
            If pendingDisjunction = -1 Then
                currentDisjunction = potentialConcat1
            Else
                ast.buffer(pendingDisjunction + 2) = potentialConcat1
            End If

            potentialConcat1 = currentDisjunction
            
            ' Restore variables
            With parseStack
                .length = .length - 4
                tmp = .buffer(.length)
                pendingDisjunction = .buffer(.length + 1)
                currentDisjunction = .buffer(.length + 2)
                potentialConcat2 = .buffer(.length + 3) ' This is correct, potentialConcat1 is the new node!
            End With
            
            If tmp > 0 Then ' capture group
                currentAstNode = ast.length
                AppendThree ast, AST_CAPTURE, potentialConcat1, tmp
                potentialConcat1 = currentAstNode
            ElseIf tmp = -1 Then ' non-capture group
                ' don't do anything
            Else ' lookahead or lookbehind
                currentAstNode = ast.length
                AppendTwo ast, -(tmp + 2) + MIN_AST_CODE, potentialConcat1
                potentialConcat1 = currentAstNode
            End If

        Case RETOK_ATOM_CHAR
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            
            currentAstNode = ast.length
            tmp = currToken.num
            If reFlags And RE_FLAG_IGNORE_CASE Then
                tmp = ReCanonicalizeChar(tmp)
            End If
            AppendTwo ast, AST_CHAR, tmp
                
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_PERIOD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendLong ast, AST_PERIOD
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_BACKREFERENCE
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendTwo ast, AST_BACKREFERENCE, currToken.num
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
        
        Case RETOK_ASSERT_START
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendLong ast, AST_ASSERT_START
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ASSERT_END
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendLong ast, AST_ASSERT_END
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_START_CHARCLASS, RETOK_ATOM_START_CHARCLASS_INVERTED
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            If currToken.t = RETOK_ATOM_START_CHARCLASS Then
                AppendTwo ast, AST_RANGES, 0
            Else
                AppendTwo ast, AST_INVRANGES, 0
            End If
            
            tmp = 0 ' unnecessary, tmp is an output parameter indicating the number of ranges
            ' Todo: Remove that parameter from ParseReRanges -- we can calculate it by comparing
            '   old and new buffer length.
            ParseReRanges lex, ast, tmp, reFlags And RE_FLAG_IGNORE_CASE <> 0
            
            ' patch range count
            ast.buffer(currentAstNode + 1) = tmp
            
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ASSERT_WORD_BOUNDARY
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendLong ast, AST_ASSERT_WORD_BOUNDARY
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ASSERT_NOT_WORD_BOUNDARY
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendLong ast, AST_ASSERT_NOT_WORD_BOUNDARY
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_DIGIT
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendPrefixedPairsArray ast, AST_RANGES, RangeTableDigit
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_NOT_DIGIT
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendPrefixedPairsArray ast, AST_RANGES, RangeTableNotDigit
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_WHITE
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendPrefixedPairsArray ast, AST_RANGES, RangeTableWhite
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_NOT_WHITE
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendPrefixedPairsArray ast, AST_RANGES, RangeTableNotWhite
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
        
        Case RETOK_ATOM_WORD_CHAR
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendPrefixedPairsArray ast, AST_RANGES, RangeTableWordchar
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
        
        Case RETOK_ATOM_NOT_WORD_CHAR
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.length
            AppendPrefixedPairsArray ast, AST_RANGES, RangeTableNotWordChar
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_EOF
            ' Todo: If Not expectEof Then Err.Raise REGEX_ERR_UNEXPECTED_END_OF_PATTERN
            
            ' Close disjunction
            If potentialConcat1 = -1 Then
                currentAstNode = ast.length
                AppendLong ast, AST_EMPTY
                potentialConcat1 = currentAstNode
            Else
                PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            End If
            
            If pendingDisjunction = -1 Then
                currentDisjunction = potentialConcat1
            Else
                ast.buffer(pendingDisjunction + 2) = potentialConcat1
            End If
            
            currentAstNode = ast.length
            AppendThree ast, AST_CAPTURE, currentDisjunction, 0
            ast.buffer(0) = currentAstNode ' patch index of root node into the first word
            
            Exit Do
        Case Else
            Err.Raise REGEX_ERR_UNEXPECTED_REGEXP_TOKEN
        End Select
    Loop
End Sub

Public Function DfsMatch( _
    ByRef outCaptures() As Long, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    Optional ByVal stepsLimit = DEFAULT_STEPS_LIMIT, _
    Optional ByVal reFlags = DFS_FLAGS_NONE _
) As Long
    Dim context As DfsMatcherContext
    DfsMatch = DfsMatchFrom(context, outCaptures, bytecode, inputStr, 0, stepsLimit, reFlags)
End Function

Public Function DfsMatchFrom( _
    ByRef context As DfsMatcherContext, _
    ByRef outCaptures() As Long, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    Optional ByVal stepsLimit = DEFAULT_STEPS_LIMIT, _
    Optional ByVal reFlags = DFS_FLAGS_NONE _
) As Long
    Dim maxCapturePointIdx As Long, res As Long
    
    maxCapturePointIdx = bytecode(0)
    If maxCapturePointIdx >= 0 Then ReDim outCaptures(0 To maxCapturePointIdx) As Long
    
    Do While sp <= Len(inputStr)
        InitializeMatcherContext context, maxCapturePointIdx + 1
        res = DfsRunThreads(outCaptures, context, bytecode, inputStr, sp, stepsLimit, reFlags)
        If res <> -1 Then
            DfsMatchFrom = res
            Exit Function
        End If
        sp = sp + 1
    Loop

    DfsMatchFrom = -1
End Function

Private Function GetBc(ByRef bytecode() As Long, ByRef pc As Long) As Long
    If pc > UBound(bytecode) Then
        GetBc = 0 ' Todo: ??????
        Exit Function
    End If
    GetBc = bytecode(pc)
    pc = pc + 1
End Function

Private Function GetInputCharCode(ByRef inputStr As String, ByRef sp As Long, ByVal spDelta As Long) As Long
    If sp >= Len(inputStr) Then
        GetInputCharCode = -1
    ElseIf sp < 0 Then
        GetInputCharCode = -1
    Else
        GetInputCharCode = AscW(Mid$(inputStr, sp + 1, 1)) ' sp is 0-based and Mid$ is 1-based
        sp = sp + spDelta
    End If
End Function

Private Function PeekInputCharCode(ByRef inputStr As String, ByRef sp As Long) As Long
    If sp >= Len(inputStr) Then
        PeekInputCharCode = -1
    ElseIf sp < 0 Then
        PeekInputCharCode = -1
    Else
        PeekInputCharCode = AscW(Mid$(inputStr, sp + 1, 1)) ' sp is 0-based and Mid$ is 1-based
    End If
End Function


Private Function UnicodeIsLineTerminator(c As Long)
    UnicodeIsLineTerminator = False
End Function

Private Function UnicodeReIsWordchar(c As Long)
    'TODO: Temporary hack
    UnicodeReIsWordchar = ((c >= AscW("A")) And (c <= AscW("Z"))) Or ((c >= AscW("a") And (c <= AscW("z"))))
End Function

Private Sub InitializeMatcherContext(ByRef context As DfsMatcherContext, ByVal nCapturePoints As Long)
    With context
        ' Clear stacks
        .matcherStack.length = 0
        .capturesStack.length = 0
        .qstack.length = 0
        
        .nCapturePoints = nCapturePoints
        AppendFill .capturesStack, nCapturePoints, -1
        
        .master = -1
        .capturesRequireCoW = False
        .qTop = 0
    End With
End Sub

Private Sub PushMatcherStackFrame( _
    ByRef context As DfsMatcherContext, _
    ByVal pc As Long, _
    ByVal sp As Long, _
    ByVal pcLandmark As Long, _
    ByVal spDelta As Long, _
    ByVal q As Long _
)
    With context.matcherStack
        If .length = .capacity Then
            ' Increase capacity
            If .capacity < DFS_MATCHER_STACK_MINIMUM_CAPACITY Then .capacity = DFS_MATCHER_STACK_MINIMUM_CAPACITY Else .capacity = .capacity + .capacity \ 2
            ReDim Preserve .buffer(0 To .capacity - 1) As DfsMatcherStackFrame
        End If
        With .buffer(.length)
           .master = context.master
           .capturesStackState = context.capturesStack.length Or (context.capturesRequireCoW And LONGTYPE_FIRST_BIT)
           .qStackLength = context.qstack.length
           .pc = pc
           .sp = sp
           .pcLandmark = pcLandmark
           .spDelta = spDelta
           .q = q
           .qTop = context.qTop
        End With
        .length = .length + 1
    End With

    context.capturesRequireCoW = True
End Sub

' If current frame is the last remaining frame, returns false
Private Function PopMatcherStackFrame(ByRef context As DfsMatcherContext, ByRef pc As Long, ByRef sp As Long, ByRef pcLandmark As Long, ByRef spDelta As Long, ByRef q As Long) As Boolean
    With context.matcherStack
        If .length = 0 Then
            PopMatcherStackFrame = False
            Exit Function
        End If
    
        .length = .length - 1
        With .buffer(.length)
            context.master = .master
            context.capturesStack.length = .capturesStackState And LONGTYPE_ALL_BUT_FIRST_BIT
            context.qstack.length = .qStackLength
            context.capturesRequireCoW = (.capturesStackState And LONGTYPE_FIRST_BIT) <> 0
            pc = .pc
            sp = .sp
            pcLandmark = .pcLandmark
            spDelta = .spDelta
            q = .q
            context.qTop = .qTop
        End With
    
        PopMatcherStackFrame = True
    End With
End Function

Private Sub ReturnToMasterDiscardCaptures(ByRef context As DfsMatcherContext, ByRef pc As Long, ByRef sp As Long, ByRef pcLandmark As Long, ByRef spDelta As Long, ByRef q As Long)
    With context.matcherStack
        .length = context.master
        With .buffer(.length)
            context.master = .master
            context.capturesStack.length = .capturesStackState And LONGTYPE_ALL_BUT_FIRST_BIT
            context.qstack.length = .qStackLength
            context.capturesRequireCoW = (.capturesStackState And LONGTYPE_FIRST_BIT) <> 0
            pc = .pc
            sp = .sp
            pcLandmark = .pcLandmark
            spDelta = .spDelta
            q = .q
            context.qTop = .qTop
        End With
    End With
End Sub

Private Sub ReturnToMasterPreserveCaptures(ByRef context As DfsMatcherContext, ByRef pc As Long, ByRef sp As Long, ByRef pcLandmark As Long, ByRef spDelta As Long, ByRef q As Long)
    Dim masterCapturesStackLength As Long, i As Long
    
    With context.matcherStack
        .length = context.master
        With .buffer(.length)
            context.master = .master
            masterCapturesStackLength = .capturesStackState And LONGTYPE_ALL_BUT_FIRST_BIT
            context.qstack.length = .qStackLength
            context.capturesRequireCoW = (.capturesStackState And LONGTYPE_FIRST_BIT) <> 0
            pc = .pc
            sp = .sp
            pcLandmark = .pcLandmark
            spDelta = .spDelta
            q = .q
            context.qTop = .qTop
        End With
    End With
    
    With context.capturesStack
        If .length = masterCapturesStackLength Then Exit Sub

        If context.capturesRequireCoW Then
            masterCapturesStackLength = masterCapturesStackLength + context.nCapturePoints
            context.capturesRequireCoW = False
            If .length = masterCapturesStackLength Then Exit Sub
        End If
        
        For i = 1 To context.nCapturePoints
            .buffer(masterCapturesStackLength - i) = .buffer(.length - i)
        Next
        .length = masterCapturesStackLength
    End With
End Sub

Private Sub CopyCaptures(ByRef context As DfsMatcherContext, ByRef target() As Long)
    Dim i As Long, j As Long
    With context
        i = .capturesStack.length - .nCapturePoints
        j = 0
        Do Until i = .capturesStack.length
            target(j) = .capturesStack.buffer(i)
            i = i + 1: j = j + 1
        Loop
    End With
End Sub

Private Sub SetCapturePoint(ByRef context As DfsMatcherContext, ByVal idx As Long, ByVal v As Long)
    With context
        If .capturesRequireCoW Then
            AppendSlice .capturesStack, .capturesStack.length - .nCapturePoints, .nCapturePoints
            .capturesRequireCoW = False
        End If
        .capturesStack.buffer(.capturesStack.length - .nCapturePoints + idx) = v
    End With
End Sub

Private Function GetCapturePoint(ByRef context As DfsMatcherContext, ByVal idx As Long) As Long
    With context
        GetCapturePoint = .capturesStack.buffer( _
            .capturesStack.length - .nCapturePoints + idx _
        )
    End With
End Function

Private Sub PushQCounter(ByRef context As DfsMatcherContext, ByVal q As Long)
    With context
        AppendThree .qstack, .qTop, .matcherStack.length, q
        .qTop = .qstack.length - 1
    End With
End Sub

Private Function PopQCounter(ByRef context As DfsMatcherContext) As Long
    With context
        If .qTop = 0 Then
            PopQCounter = Q_NONE
            Exit Function
        End If
        
        PopQCounter = .qstack.buffer(.qTop)
        If .qstack.buffer(.qTop - 1) = .matcherStack.length Then .qstack.length = .qstack.length - 3
        .qTop = .qstack.buffer(.qTop - 2)
    End With
End Function



Private Function DfsRunThreads( _
    ByRef outCaptures() As Long, _
    ByRef context As DfsMatcherContext, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    ByVal stepsLimit As Long, _
    ByVal reFlags As Long _
) As Long
    Dim pc As Long
    
    ' To avoid infinite loops
    Dim pcLandmark As Long
    
    Dim op As Long
    Dim c1 As Long, c2 As Long
    Dim t As Long
    Dim n As Long
    Dim successfulMatch As Boolean
    Dim r1 As Long, r2 As Long
    Dim b1 As Boolean, b2 As Boolean
    Dim aa As Long, bb As Long, mm As Long
    Dim idx As Long, off As Long
    Dim q As Long, qmin As Long, qmax As Long, qq As Long
    Dim qexact As Long
    Dim stepsCount As Long
    Dim spDelta As Long ' 1 when we walk forwards and -1 when we walk backwards
    
        pc = 1
        pcLandmark = -1
        stepsCount = 0
        spDelta = 1
        q = Q_NONE
       
        GoTo ContinueLoopSuccess

        ' BEGIN LOOP
ContinueLoopFail:
            If Not PopMatcherStackFrame(context, pc, sp, pcLandmark, spDelta, q) Then
                DfsRunThreads = -1
                Exit Function
            End If

ContinueLoopSuccess:

            ' TODO: HACK to prevent infinite loop!!!!
            stepsCount = stepsCount + 1
            If stepsCount >= stepsLimit Then
                DfsRunThreads = -1
                Exit Function
            End If
            
            op = GetBc(bytecode, pc)
    
            ' #if defined(DUK_USE_DEBUG_LEVEL) && (DUK_USE_DEBUG_LEVEL >= 2)
            ' duk__regexp_dump_state(re_ctx);
            ' #End If
            ' DUK_DDD(DUK_DDDPRINT("match: rec=%ld, steps=%ld, pc (after op)=%ld, sp=%ld, op=%ld",
            '                     (long) re_ctx->recursion_depth,
            '                     (long) re_ctx->steps_count,
            '                     (long) (pc - re_ctx->bytecode),
            '                     (long) sp,
            '                     (long) op));
    
            Select Case op
            Case REOP_MATCH
                GoTo Match
            Case REOP_END_LOOKPOS
                ' Reached if pattern inside a positive lookahead matched.
                ReturnToMasterPreserveCaptures context, pc, sp, pcLandmark, spDelta, q
                ' Now we are at the REOP_LOOKPOS opcode.
                pc = pc + 1
                n = GetBc(bytecode, pc)
                pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_END_LOOKNEG
                ' Reached if pattern inside a positive lookahead matched.
                ReturnToMasterDiscardCaptures context, pc, sp, pcLandmark, spDelta, q
                GoTo ContinueLoopFail
            Case REOP_CHAR
                '
                '  Byte-based matching would be possible for case-sensitive
                '  matching but not for case-insensitive matching.  So, we
                '  match by decoding the input and bytecode character normally.
                '
                '  Bytecode characters are assumed to be already canonicalized.
                '  Input characters are canonicalized automatically by
                '  duk__inp_get_cp() if necessary.
                '
                '  There is no opcode for matching multiple characters.  The
                '  regexp compiler has trouble joining strings efficiently
                '  during compilation.  See doc/regexp.rst for more discussion.
    
                pcLandmark = pc - 1
                c1 = GetBc(bytecode, pc)
                c2 = GetInputCharCode(inputStr, sp, spDelta)
                ' DUK_ASSERT(c1 >= 0);
    
                ' DUK_DDD(DUK_DDDPRINT("char match, c1=%ld, c2=%ld", (long) c1, (long) c2));
                If c1 <> c2 Then GoTo ContinueLoopFail
                GoTo ContinueLoopSuccess
            Case REOP_PERIOD
                pcLandmark = pc - 1
                c1 = GetInputCharCode(inputStr, sp, spDelta)
                If c1 < 0 Then
                    GoTo ContinueLoopFail
                ElseIf UnicodeIsLineTerminator(c1) Then
                    GoTo ContinueLoopFail
                End If
                GoTo ContinueLoopSuccess
            Case REOP_RANGES, REOP_INVRANGES
                pcLandmark = pc - 1
                n = GetBc(bytecode, pc) ' assert: >= 1
                c1 = GetInputCharCode(inputStr, sp, spDelta)
                If c1 < 0 Then GoTo ContinueLoopFail
                
                aa = pc - 1
                pc = pc + 2 * n
                bb = pc + 1
                
                ' We are doing a binary search here.
                Do
                    mm = aa + 2 * ((bb - aa) \ 4)
                    If bytecode(mm) >= c1 Then bb = mm Else aa = mm
                    
                    If bb - aa = 2 Then
                        ' bb is the first upper bound index s.t. ary(bb)>=v
                        If bb >= pc Then successfulMatch = False Else successfulMatch = bytecode(bb - 1) <= c1
                        Exit Do
                    End If
                Loop
    
                If (op = REOP_RANGES) <> successfulMatch Then GoTo ContinueLoopFail

                GoTo ContinueLoopSuccess
            Case REOP_ASSERT_START
                If sp <= 0 Then GoTo ContinueLoopSuccess
                If Not (reFlags And RE_FLAG_MULTILINE) Then GoTo ContinueLoopFail
                c1 = PeekInputCharCode(inputStr, sp + spDelta)
                ' E5 Sections 15.10.2.8, 7.3
                If Not UnicodeIsLineTerminator(c1) Then GoTo ContinueLoopFail
                GoTo ContinueLoopSuccess
            Case REOP_ASSERT_END
                c1 = PeekInputCharCode(inputStr, sp)
                If c1 < 0 Then GoTo ContinueLoopSuccess
                If Not (reFlags And RE_FLAG_MULTILINE) Then GoTo ContinueLoopFail
                If Not UnicodeIsLineTerminator(c1) Then GoTo ContinueLoopFail
                GoTo ContinueLoopSuccess
            Case REOP_ASSERT_WORD_BOUNDARY, REOP_ASSERT_NOT_WORD_BOUNDARY
                '
                '  E5 Section 15.10.2.6.  The previous and current character
                '  should -not- be canonicalized as they are now.  However,
                '  canonicalization does not affect the result of IsWordChar()
                '  (which depends on Unicode characters never canonicalizing
                '  into ASCII characters) so this does not matter.
                If sp <= 0 Then
                    b1 = False  ' not a wordchar
                Else
                    c1 = PeekInputCharCode(inputStr, sp - spDelta)
                    b1 = UnicodeReIsWordchar(c1)
                End If
                If sp > Len(inputStr) Then
                    b2 = False ' not a wordchar
                Else
                    c1 = PeekInputCharCode(inputStr, sp)
                    b2 = UnicodeReIsWordchar(c1)
                End If
    
                If (op = REOP_ASSERT_WORD_BOUNDARY) = (b1 = b2) Then GoTo ContinueLoopFail

                GoTo ContinueLoopSuccess
            Case REOP_JUMP
                n = GetBc(bytecode, pc)
                If n > 0 Then
                    ' forward jump (disjunction)
                    pc = pc + n
                    GoTo ContinueLoopSuccess
                Else
                    ' backward jump (end of loop)
                    t = pc + n
                    If pcLandmark <= t Then GoTo ContinueLoopSuccess ' empty match
                                    
                    pc = t: pcLandmark = t
                    GoTo ContinueLoopSuccess
                End If
            Case REOP_SPLIT1
                ' split1: prefer direct execution (no jump)
                n = GetBc(bytecode, pc)
                PushMatcherStackFrame context, pc + n, sp, pcLandmark, spDelta, q
                '.tsStack(.tsLastIndex - 1).pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_SPLIT2
                ' split2: prefer jump execution (not direct)
                n = GetBc(bytecode, pc)
                PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
                pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_EXACTLY_INIT
                PushQCounter context, q
                q = 0
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_EXACTLY_START
                pc = pc + 2 ' skip arguments
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_EXACTLY_END
                qexact = GetBc(bytecode, pc) ' quantity
                n = GetBc(bytecode, pc) ' offset
                q = q + 1
                If q < qexact Then
                    t = pc - n - 3
                    If pcLandmark > t Then pcLandmark = t
                    pc = t
                Else
                    q = PopQCounter(context)
                End If
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_MAX_HUMBLE_INIT, REOP_REPEAT_GREEDY_MAX_INIT
                PushQCounter context, q
                q = -1
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_MAX_HUMBLE_START
                qmax = GetBc(bytecode, pc)
                n = GetBc(bytecode, pc)
                
                q = q + 1
                If q < qmax Then PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
                
                q = PopQCounter(context)
                pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_MAX_HUMBLE_END
                pc = pc + 1  ' skip first argument: quantity
                n = GetBc(bytecode, pc) ' offset
                t = pc - n - 3
                If pcLandmark <= t Then GoTo ContinueLoopFail ' empty match
                
                pc = t: pcLandmark = t
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_GREEDY_MAX_START
                qmax = GetBc(bytecode, pc)
                n = GetBc(bytecode, pc)
                
                q = q + 1
                If q < qmax Then
                    qq = PopQCounter(context)
                    PushMatcherStackFrame context, pc + n, sp, pcLandmark, spDelta, qq
                    PushQCounter context, qq
                Else
                    pc = pc + n
                    q = PopQCounter(context)
                End If
                
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_GREEDY_MAX_END
                pc = pc + 1 ' Skip first argument: quantity
                n = GetBc(bytecode, pc) ' offset
                t = pc - n - 3
                If pcLandmark <= t Then GoTo ContinueLoopSuccess ' empty match
                
                pc = t: pcLandmark = t
                GoTo ContinueLoopSuccess
            Case REOP_SAVE
                idx = GetBc(bytecode, pc)
                If idx >= context.nCapturePoints Then GoTo InternalError
                    ' idx is unsigned, < 0 check is not necessary
                    ' DUK_D(DUK_DPRINT("internal error, regexp save index insane: idx=%ld", (long) idx));
                '.tsStack(.tsLastIndex).saved(idx) = sp
                SetCapturePoint context, idx, sp
                GoTo ContinueLoopSuccess
            Case REOP_CHECK_LOOKAHEAD
                PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
                context.master = context.matcherStack.length - 1
                pc = pc + 2 ' jump over following REOP_LOOKPOS or REOP_LOOKNEG
                ' When we're moving forward, we are at the correct position. When we're moving backward, we have to step one towards the end.
                sp = sp + (1 - spDelta) \ 2
                spDelta = 1
                ' We could set pcLandmark to -1 again here, but we can be sure that pcLandmark < beginning of lookahead, so we can skip that
                GoTo ContinueLoopSuccess
            Case REOP_CHECK_LOOKBEHIND
                PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
                context.master = context.matcherStack.length - 1
                pc = pc + 2 ' jump over following REOP_LOOKPOS or REOP_LOOKNEG
                ' When we're moving backward, we are at the correct position. When we're moving forward, we have to step one towards the beginning.
                sp = sp - (spDelta + 1) \ 2
                spDelta = -1
                ' We could set pcLandmark to -1 again here, but we can be sure that pcLandmark < beginning of lookahead, so we can skip that
                GoTo ContinueLoopSuccess
            Case REOP_LOOKPOS
                ' This point will only be reached if the pattern inside a negative lookahead/back did not match.
                n = GetBc(bytecode, pc)
                pc = pc + n
                GoTo ContinueLoopFail
            Case REOP_LOOKNEG
                ' This point will only be reached if the pattern inside a negative lookahead/back did not match.
                n = GetBc(bytecode, pc)
                pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_BACKREFERENCE
                '
                '  Byte matching for back-references would be OK in case-
                '  sensitive matching.  In case-insensitive matching we need
                '  to canonicalize characters, so back-reference matching needs
                '  to be done with codepoints instead.  So, we just decode
                '  everything normally here, too.
                '
                '  Note: back-reference index which is 0 or higher than
                '  NCapturingParens (= number of capturing parens in the
                '  -entire- regexp) is a compile time error.  However, a
                '  backreference referring to a valid capture which has
                '  not matched anything always succeeds!  See E5 Section
                '  15.10.2.9, step 5, sub-step 3.
    
                pcLandmark = pc - 1
                idx = 2 * GetBc(bytecode, pc) ' backref n -> saved indices [n*2, n*2+1]
                If idx < 2 Then GoTo InternalError
                If idx + 1 >= context.nCapturePoints Then GoTo InternalError
                aa = GetCapturePoint(context, idx)
                bb = GetCapturePoint(context, idx + 1)
                If (aa >= 0) And (bb >= 0) Then
                    If spDelta = 1 Then
                        off = aa
                        Do While off < bb
                            c1 = GetInputCharCode(inputStr, off, 1)
                            c2 = GetInputCharCode(inputStr, sp, 1)
                            ' No need for an explicit c2 < 0 check: because c1 >= 0,
                            ' the comparison will always fail if c2 < 0.
                            If c1 <> c2 Then GoTo ContinueLoopFail
                        Loop
                    Else
                        off = bb - 1
                        Do While off >= aa
                            c1 = GetInputCharCode(inputStr, off, -1)
                            c2 = GetInputCharCode(inputStr, sp, -1)
                            ' No need for an explicit c2 < 0 check: because c1 >= 0,
                            ' the comparison will always fail if c2 < 0.
                            If c1 <> c2 Then GoTo ContinueLoopFail
                        Loop
                    End If
                Else
                    ' capture is 'undefined', always matches!
                End If
                GoTo ContinueLoopSuccess
            Case Else
                'DUK_D(DUK_DPRINT("internal error, regexp opcode error: %ld", (long) op));
                GoTo InternalError
            End Select
        ' END LOOP
    
Match:
        CopyCaptures context, outCaptures
        DfsRunThreads = sp
        Exit Function
        
InternalError:
        ' TODO: Raise correct exception
        Err.Raise 3000
        ' DUK_ERROR_INTERNAL(re_ctx->thr);
        ' DUK_WO_NORETURN(return -1;);
End Function

