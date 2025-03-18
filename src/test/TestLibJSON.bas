Attribute VB_Name = "TestLibJSON"
'''=============================================================================
''' VBA Fast JSON Parser & Serializer - github.com/cristianbuse/VBA-FastJSON
''' ---
''' MIT License - github.com/cristianbuse/VBA-FastJSON/blob/master/LICENSE
''' Copyright (c) 2024 Ion Cristian Buse
'''=============================================================================
''' Additional to own tests, the below test suite includes many tests from:
''' https://github.com/nst/JSONTestSuite
''' ---
''' MIT License - https://github.com/nst/JSONTestSuite/blob/master/LICENSE
''' Copyright (c) 2016 Nicolas Seriot
''' Many thanks to Nicolas! Must-read his article:
''' Parsing JSON is a Minefield - https://seriot.ch/projects/parsing_json.html
'''=============================================================================

Option Explicit
Option Private Module

#Const Windows = (Mac = 0)
Private Const commaA As Byte = &H22

'Test compliance with RFC 8259: https://www.rfc-editor.org/rfc/rfc8259
'Addtional extensions are also tested. See optional parameters for LibJSON.Parse
Public Sub RunAllJSONParseTests()
    TestParseEmptyInvalid
    TestParseLiteralValid
    TestParseLiteralInvalid
    TestParseWhitespaceValid
    TestParseWhitespaceInvalid
    TestParseArrayValid
    TestParseArrayValidLiteral
    TestParseArrayValidNesting
    TestParseArrayValidWithWhitespaces
    TestParseArrayInvalidComma
    TestParseArrayInvalidMisc
    TestParseArrayInvalidUnclosed
    TestParseObjectValid
    TestParseObjectValidNesting
    TestParseObjectInvalidMisc
    TestParseObjectInvalidUnclosed
    TestParseMiscValid
    TestParseMiscInvalid
    TestParseNumberValid
    TestParseNumberValidLargeExponent
    TestParseNumberInvalid
    TestParseStringValid
    TestParseStringValidEscape
    TestParseStringInvalid
    TestParseStringInvalidEscape
    TestParseStringLoneSurrogates
    TestParseStringInvalidUTF8
    Debug.Print "Finished running tests at " & Now()
End Sub

'*******************************************************************************
'Parse Tests
'*******************************************************************************
Private Sub TestParseEmptyInvalid()
    Debug.Assert Not Parse(vbNullString).IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse(" ").IsValid
    Debug.Assert Not Parse(vbNewLine).IsValid
    Debug.Assert Not Parse(vbTab).IsValid
    Debug.Assert Not Parse(vbCr).IsValid
    Debug.Assert Not Parse(vbLf).IsValid
    Debug.Assert Not Parse(vbTab & "  " & vbNewLine & "   ").IsValid
End Sub

Private Sub TestParseLiteralValid()
    Debug.Assert Not Parse("false").Value
    Debug.Assert Parse("true").Value
    Debug.Assert IsNull(Parse("null").Value)
    Debug.Assert Parse("[false,null,true]").Value.Count = 3
End Sub

Private Sub TestParseLiteralInvalid()
    Debug.Assert Not Parse("falsee").IsValid
    Debug.Assert Not Parse("False").IsValid
    Debug.Assert Not Parse("FALSE").IsValid
    Debug.Assert Not Parse("fals").IsValid
    Debug.Assert Not Parse("True").IsValid
    Debug.Assert Not Parse("tRue").IsValid
    Debug.Assert Not Parse("tru").IsValid
    Debug.Assert Not Parse("Null").IsValid
    Debug.Assert Not Parse("nul").IsValid
    Debug.Assert Not Parse("nulll").IsValid
    Debug.Assert Not Parse("undefined").IsValid
    Debug.Assert Not Parse("n/a").IsValid
    Debug.Assert Not Parse("empty").IsValid
    Debug.Assert Not Parse("blank").IsValid
End Sub

Private Sub TestParseWhitespaceValid()
    Debug.Assert Parse("    1").Value = 1
    Debug.Assert Parse("    12").Value = 12
    Debug.Assert Parse(" 123 ").Value = 123
    Debug.Assert Parse("    false").Value = False
    Debug.Assert Parse("    """"").IsValid
    Debug.Assert Parse("    """"").Value = ""
    Debug.Assert Parse("    1  ").Value = 1
    Debug.Assert Parse("    true   ").Value = True
    Debug.Assert Parse("    ""tru""").Value = "tru"
    Debug.Assert Parse(vbNewLine & "1").Value = 1
    Debug.Assert Parse(vbLf & "1").Value = 1
    Debug.Assert Parse(vbCr & "1").Value = 1
    Debug.Assert Parse(vbTab & vbTab & "1").Value = 1
    Debug.Assert Parse("  1  " & vbNewLine & "  " & vbLf & "  " & vbCr & "   " & vbTab).Value = 1
End Sub

Private Sub TestParseWhitespaceInvalid()
    Debug.Assert Not Parse(vbFormFeed & "1").IsValid
    Debug.Assert Not Parse(ChrW$(&H2060) & "0").IsValid 'Word Joiner (WJ)
End Sub

Private Sub TestParseArrayValid()
    Debug.Assert AreEqual(Parse("[]").Value, New Collection)
    Debug.Assert AreEqual(Parse("[""""]").Value, Collection(vbNullString))
    Debug.Assert AreEqual(Parse("[""text"",123,{},[],null]").Value _
                        , Collection("text", 123, New Dictionary, New Collection, Null))
    Debug.Assert AreEqual(Parse("[""" & ChrW$(&H2B7F) & """]").Value _
                        , Collection(ChrW$(&H2B7F))) 'Vertical Tab Key (not same as Line Tabulation / Vertical Tab)
    Debug.Assert AreEqual(Parse("[""abc""]").Value, Collection("abc"))
    Debug.Assert AreEqual(Parse("[ ""abc""]").Value, Collection("abc"))
    Debug.Assert AreEqual(Parse("[ ""abc""]" & vbNewLine).Value, Collection("abc"))
End Sub

Private Sub TestParseArrayValidLiteral()
    Debug.Assert AreEqual(Parse("[false]").Value, Collection(False))
    Debug.Assert AreEqual(Parse("[true]").Value, Collection(True))
    Debug.Assert AreEqual(Parse("[null]").Value, Collection(Null))
    Debug.Assert AreEqual(Parse("[false,null,true]").Value, Collection(False, Null, True))
    Debug.Assert AreEqual(Parse("[null,null,null]").Value, Collection(Null, Null, Null))
End Sub

Private Sub TestParseArrayValidNesting()
    Debug.Assert AreEqual(Parse("[[],[[]]]").Value, Collection(New Collection, Collection(New Collection)))
    '
    Const nestingLevel As Long = 10000
    Dim v As Collection
    Dim c As Collection
    Dim i As Long
    '
    Set v = Parse(String$(nestingLevel, "[") & String$(nestingLevel, "]"), maxNestingDepth:=nestingLevel).Value
    Set c = v
    Do While v.Count = 1
        Set v = v.Item(1)
        i = i + 1
    Loop
    Debug.Assert v.Count = 0
    Debug.Assert i + 1 = nestingLevel
    '
    'The Termination stack frames would fill the stack and lead to 'Out of stack space'
    'We can only fix this if we build a wrapper class around Collection or if we write our own Array class
    Do
        Set c = c(1)
        i = i - 1
    Loop Until i < 5000
End Sub

Private Sub TestParseArrayValidWithWhitespaces()
    Debug.Assert AreEqual(Parse("[ [] ]" & vbNewLine).Value, Collection(New Collection))
    Debug.Assert AreEqual(Parse("[" & vbTab & "]" & vbCr).Value, New Collection)
    Debug.Assert AreEqual(Parse(" [] ").Value, New Collection)
    Debug.Assert AreEqual(Parse("[0" & vbNewLine & "]").Value, Collection(0))
    Debug.Assert AreEqual(Parse("   [  231]").Value, Collection(231))
    Debug.Assert AreEqual(Parse("[1] ").Value, Collection(1))
End Sub

Private Sub TestParseArrayInvalidComma()
    Debug.Assert Not Parse("[0 null]").IsValid
    Debug.Assert Not Parse("[0:1]").IsValid
    Debug.Assert Not Parse("[0;1]").IsValid
    Debug.Assert Not Parse("[0],").IsValid
    Debug.Assert Not Parse("[,123]").IsValid
    Debug.Assert Not Parse("[,1,2,3]").IsValid
    Debug.Assert Not Parse("[1,,2]").IsValid
    Debug.Assert Not Parse("[1,,]").IsValid
    Debug.Assert Not Parse("[null,,]").IsValid
    Debug.Assert Not Parse("[""a"",,]").IsValid
    Debug.Assert Not Parse("[0,]").IsValid
    Debug.Assert Not Parse("["""",]").IsValid
    Debug.Assert Not Parse("[0[1]]").IsValid
    Debug.Assert Not Parse("[0[1],]").IsValid
    Debug.Assert Not Parse("[,]").IsValid
    Debug.Assert Not Parse("[ , 0 ]").IsValid
    Debug.Assert Not Parse("[0,").IsValid
    Debug.Assert Not Parse("[""value"",,""anotherValue""]").IsValid
End Sub

Private Sub TestParseArrayInvalidMisc()
    Debug.Assert Not Parse("[*]").IsValid
    Debug.Assert Not Parse("[-]").IsValid
    Debug.Assert Not Parse("[+]").IsValid
    Debug.Assert Not Parse("[:]").IsValid
    Debug.Assert Not Parse("[\]").IsValid
    Debug.Assert Not Parse("[']").IsValid
    Debug.Assert Not Parse("['").IsValid
    Debug.Assert Not Parse("[,").IsValid
    Debug.Assert Not Parse("[""abc""\f]").IsValid
    Debug.Assert Not Parse("[a]").IsValid
    Debug.Assert Not Parse("[" & ChrW$(1234) & "]").IsValid
    Debug.Assert Not Parse("[" & vbFormFeed & "]").IsValid
    Debug.Assert Not Parse(String$(512, "[")).IsValid
    Debug.Assert Not Parse(RepeatString("[{"""":", 512)).IsValid
    Debug.Assert Not Parse("[false,tru").IsValid
    Debug.Assert Not Parse("[false,nul").IsValid
    Debug.Assert Not Parse("[fals,true").IsValid
    Debug.Assert Not Parse("[true,fals").IsValid
    Debug.Assert Not Parse("[1]a").IsValid
    Debug.Assert Not Parse("[""abc]").IsValid
    Debug.Assert Not Parse("[][]").IsValid
    Debug.Assert Not Parse("[""a"", a]").IsValid
    Debug.Assert Not Parse(BytesToString(&H5B, &HFF, &H5D)).IsValid
    Debug.Assert Not Parse(BytesToString(&H5B, &H61, &HE5, &H5D)).IsValid
End Sub

Private Sub TestParseArrayInvalidUnclosed()
    Debug.Assert Not Parse("[").IsValid
    Debug.Assert Not Parse("[[").IsValid
    Debug.Assert Not Parse("]").IsValid
    Debug.Assert Not Parse("]]").IsValid
    Debug.Assert Not Parse("[}").IsValid
    Debug.Assert Not Parse("{]").IsValid
    Debug.Assert Not Parse("[{").IsValid
    Debug.Assert Not Parse("{[").IsValid
    Debug.Assert Not Parse("[{}").IsValid
    Debug.Assert Not Parse("[[]]]").IsValid
    Debug.Assert Not Parse("[""a""]]").IsValid
    Debug.Assert Not Parse("[""a""").IsValid
    Debug.Assert Not Parse("[a").IsValid
    Debug.Assert Not Parse("[""a""").IsValid
    Debug.Assert Not Parse("[1").IsValid
    Debug.Assert Not Parse("[1]]").IsValid
    Debug.Assert Not Parse("]1]").IsValid
    Debug.Assert Not Parse("1]").IsValid
    Debug.Assert Not Parse("[1," & vbNewLine & "2" & vbNewLine & ",3").IsValid
    Debug.Assert Not Parse("[""1""," & vbNewLine & "2" & vbNewLine & ",3").IsValid
End Sub

Private Sub TestParseObjectValid()
    Debug.Assert AreEqual(Parse("{}").Value, New Dictionary)
    Debug.Assert AreEqual(Parse("{"""":0}").Value, Dictionary(vbNullString, 0))
    Debug.Assert AreEqual(Parse("{""key"":""value""}").Value _
                        , Dictionary("key", "value"))
    Debug.Assert AreEqual(Parse("{""key1"":123,""key2"":true}").Value _
                        , Dictionary("key1", 123, "key2", True))
    Debug.Assert AreEqual(Parse("{""key1"":123,""key2"":true}" & vbNewLine).Value _
                        , Dictionary("key1", 123, "key2", True))
    Debug.Assert AreEqual(Parse("{""nestedObj"":{""key"":123}}").Value _
                        , Dictionary("nestedObj", Dictionary("key", 123)))
    Debug.Assert AreEqual(Parse("{""abc"":""def""}").Value _
                        , Dictionary("abc", "def"))
    Debug.Assert AreEqual(Parse("{""abc"":""def"",""ghi"":""jkl""}").Value _
                        , Dictionary("abc", "def", "ghi", "jkl"))
    Dim d As Dictionary
    If IsFastDict Then 'Fast-Dictionary can handle Duplicate Keys
        Set d = New Dictionary
        d.AllowDuplicateKeys = True
        d.Add "a", "b"
        d.Add "a", "c"
        Debug.Assert AreEqual(Parse("{""a"":""b"",""a"":""c""}", allowDuplicatedKeys:=True).Value, d)
        '
        Set d = New Dictionary
        d.AllowDuplicateKeys = True
        d.Add "a", "b"
        d.Add "a", "b"
        Debug.Assert AreEqual(Parse("{""a"":""b"",""a"":""b""}", allowDuplicatedKeys:=True).Value, d)
    End If
    Debug.Assert AreEqual(Parse("{""ke\u0000y"":null}").Value _
                        , Dictionary("ke" & vbNullChar & "y", Null))
    Debug.Assert AreEqual(Parse("{""min"":-1.0e+308,""max"":1.0E+308}").Value _
                        , Dictionary("min", -1E+308, "max", 1E+308))
    Debug.Assert AreEqual(Parse("{""a"":[{""b"": ""abcdefghijklmnopqrstuvwxyz0123456789@;:~>""}]}").Value _
                        , Dictionary("a", Collection(Dictionary("b", "abcdefghijklmnopqrstuvwxyz0123456789@;:~>"))))
    Debug.Assert AreEqual(Parse("{""key"":[]}").Value _
                        , Dictionary("key", New Collection))
    Debug.Assert AreEqual(Parse("{""key"":""\u0418\u043D\u0442\u0435\u0440\u0435\u0441\u043D\u043E, \u043D\u0430\u043B\u0438?""}").Value _
                        , Dictionary("key", ChrW$(&H418) & ChrW$(&H43D) & ChrW$(&H442) & ChrW$(&H435) & ChrW$(&H440) _
                                          & ChrW$(&H435) & ChrW$(&H441) & ChrW$(&H43D) & ChrW$(&H43E) & ", " _
                                          & ChrW$(&H43D) & ChrW$(&H430) & ChrW$(&H43B) & ChrW$(&H438) & "?"))
    Debug.Assert AreEqual(Parse(vbNewLine & "{""key"":[]}" & vbTab & vbNewLine).Value _
                        , Dictionary("key", New Collection))
    Debug.Assert AreEqual(Parse(ChrW$(&HBBEF) & ChrB$(&HBF) & "{}").Value, New Dictionary) 'UTF8 BOM - ignored
    Debug.Assert AreEqual(Parse(ChrW$(&HFEFF) & "{}").Value, New Dictionary) 'UTF16LE BOM - ignored
    Debug.Assert AreEqual(Parse("{""\uDFAA"":0}").Value, Dictionary(ChrW$(&HDFAA), 0)) 'Single surrogate
End Sub

Private Sub TestParseObjectValidNesting()
    Const nestingLevel As Long = 10000
    Dim v As Dictionary
    Dim i As Long
    Dim h As Dictionary
    '
    Set v = Parse(RepeatString("{""key"":", nestingLevel - 1) & "{" _
                    & String$(nestingLevel, "}"), maxNestingDepth:=nestingLevel).Value
    If IsFastDict() Then Set h = v 'Only Fast-Dictionary can handle deep nesting termination
    Do While v.Count = 1
        Set v = v.Item("key")
        i = i + 1
    Loop
    Debug.Assert v.Count = 0
    Debug.Assert i + 1 = nestingLevel
    '
    Dim w As Object
    Set w = Parse(RepeatString("{""key"":[", nestingLevel) _
                    & RepeatString("]}", nestingLevel), maxNestingDepth:=nestingLevel * 2).Value
    If IsFastDict() Then Set h = w 'Only Fast-Dictionary can handle deep nesting termination
    i = 0
    Do While w.Count = 1
        If TypeOf w Is Collection Then
            Set w = w.Item(1)
        Else
            Set w = w.Item("key")
        End If
        i = i + 1
    Loop
    Debug.Assert w.Count = 0
    Debug.Assert i = nestingLevel * 2 - 1
End Sub

Private Sub TestParseObjectInvalidMisc()
    Debug.Assert Not Parse("{key:""value""}").IsValid
    Debug.Assert Not Parse("{""key"":}").IsValid
    Debug.Assert Not Parse("{""key"":""value"",}").IsValid
    Debug.Assert Not Parse("[{""k1"":1,""k2""}]").IsValid
    Debug.Assert Not Parse("[{""k1"":1 ""k2"":2}]").IsValid
    Debug.Assert Not Parse("[{""k1"":1] ""k2"":2}").IsValid
    Debug.Assert Not Parse("{""a"":""b"",""a"":""c""}", allowDuplicatedKeys:=False).IsValid
    Debug.Assert Not Parse("{""a"":""b"",""a"":""b""}", allowDuplicatedKeys:=False).IsValid
    Debug.Assert Not Parse("{""a"":""b"",""a"":""c""}", allowDuplicatedKeys:=False).IsValid
    Debug.Assert Not Parse("{""a"":""b"",""a"":""b""}", allowDuplicatedKeys:=False).IsValid
    Debug.Assert Not Parse(ChrW$(&HBBEF) & ChrB$(&HBF) & "{}", failIfBOMDetected:=True).IsValid
    Debug.Assert Not Parse(ChrW$(&HFEFF) & "{}", failIfBOMDetected:=True).IsValid
    Debug.Assert Not Parse("{1:0}").IsValid
    Debug.Assert Not Parse("{11111:0}").IsValid
    Debug.Assert Not Parse("{11111e11111:0}").IsValid
    Debug.Assert Not Parse("{[: ""abc""}").IsValid
    Debug.Assert Not Parse("{""key"", null}").IsValid
    Debug.Assert Not Parse("{""key""::null}").IsValid
    Debug.Assert Not Parse(BytesToString(&H7B, &HF0, &H9F, &H87, &HA8, &HF0, &H9F, &H87, &HAD, &H7D)).IsValid '{Emoji}
    Debug.Assert Not Parse("{""a"":""a"" 123}").IsValid
    Debug.Assert Not Parse("{key: 'value'}").IsValid
    Debug.Assert Not Parse(BytesToString(&H7B, &H22, &HB9, &H22, &H3A, &H22, &H30, &H22, &H2C, &H7D)).IsValid
    Debug.Assert Not Parse("{""a"" b}").IsValid
    Debug.Assert Not Parse("{:""b""}").IsValid
    Debug.Assert Not Parse("{""a"" ""b""}").IsValid
    Debug.Assert Not Parse("{1:1}").IsValid
    Debug.Assert Not Parse("{null:null,null:null}").IsValid
    Debug.Assert Not Parse("{""id"":0,,,,,}").IsValid
    Debug.Assert Not Parse("{""id"":0,,,,,}", ignoreTrailingComma:=True).IsValid
    Debug.Assert Not Parse("{'a':0}").IsValid
    Debug.Assert Not Parse("{""id"":0,}").IsValid
    Debug.Assert Not Parse("{""a"":""b""}/**/").IsValid  'Comments
    Debug.Assert Not Parse("{""a"":""b""}/**//").IsValid
    Debug.Assert Not Parse("{""a"":""b""}//").IsValid
    Debug.Assert Not Parse("{""a"":""b""}/").IsValid
    Debug.Assert Not Parse("{""a"":1,,""b"":2}").IsValid
    Debug.Assert Not Parse("{a: ""b""}").IsValid
    Debug.Assert Not Parse("{ ""k"" : ""v"", ""s"" }").IsValid
    Debug.Assert Not Parse("{""a"":""b""}#").IsValid
End Sub

Private Sub TestParseObjectInvalidUnclosed()
    Debug.Assert Not Parse("{").IsValid
    Debug.Assert Not Parse("[}").IsValid
    Debug.Assert Not Parse("{]").IsValid
    Debug.Assert Not Parse("{[").IsValid
    Debug.Assert Not Parse("{"":").IsValid
    Debug.Assert Not Parse("{}}").IsValid
    Debug.Assert Not Parse("[0,{,1]").IsValid
    Debug.Assert Not Parse("{""a"":").IsValid
    Debug.Assert Not Parse("{""a""").IsValid
    Debug.Assert Not Parse("{""a"":""a").IsValid
End Sub

Private Sub TestParseMiscValid()
    Debug.Assert AreEqual(Parse("[{""k1"":true}, {""k2"":null}]").Value _
                        , Collection(Dictionary("k1", True), Dictionary("k2", Null)))
    Debug.Assert AreEqual(Parse("{""arr"":[1,2,3],""obj"":{""k1"":1,""K1"":2}}").Value _
                        , Dictionary("arr", Collection(1, 2, 3), "obj", Dictionary("k1", 1, "K1", 2)))
    Debug.Assert AreEqual(Parse(BytesToString(&H0, &H5B, &H0, &H22, &H0, &HE9, &H0, &H22, &H0, &H5D)).Value _
                        , Collection(ChrW$(&HE9))) 'UTF16BE
    Debug.Assert AreEqual(Parse(BytesToString(&HFF, &HFE, &H5B, &H0, &H22, &H0, &HE9, &H0, &H22, &H0, &H5D, &H0)).Value _
                        , Collection(ChrW$(&HE9))) 'UTF16LE with UTF16LE BOM
    Debug.Assert AreEqual(Parse(BytesToString(&HEF, &HBB, &HBF, &H5B, &H0, &H22, &H0, &HE9, &H0, &H22, &H0, &H5D, &H0)).Value _
                        , Collection(ChrW$(&HE9))) 'UTF16LE with UTF8 BOM
    Debug.Assert AreEqual(Parse("[0,]", ignoreTrailingComma:=True).Value, Collection(0))
    Debug.Assert AreEqual(Parse("["""",]", ignoreTrailingComma:=True).Value, Collection(vbNullString))
    Debug.Assert AreEqual(Parse("{""key"":""value"",}", ignoreTrailingComma:=True).Value, Dictionary("key", "value"))
End Sub

Private Sub TestParseMiscInvalid()
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("<.>").IsValid
    Debug.Assert Not Parse("[<null>]").IsValid
    Debug.Assert Not Parse(BytesToString(&H61, &HC3, &HA5)).IsValid
    Debug.Assert Not Parse("[True]").IsValid
    Debug.Assert Not Parse("{""x"": true,").IsValid
    Debug.Assert Not Parse(BytesToString(&HEF, &HBB, &H7B, &H7D)).IsValid 'Incomplete BOM
    Debug.Assert Not Parse(BytesToString(&HEF, &HBB, &HBF, &H7B, &H7D), failIfBOMDetected:=True).IsValid 'Complete BOM
    Debug.Assert Not Parse(BytesToString(&HEF, &HBB, &HBF)).IsValid 'Complete BOM with no data
    Debug.Assert Not Parse(Chr$(&HE5)).IsValid
    Debug.Assert Not Parse("[").IsValid
    Debug.Assert Not Parse("{").IsValid
    Debug.Assert Not Parse("[}").IsValid
    Debug.Assert Not Parse("{[").IsValid
    Debug.Assert Not Parse("[" & vbNullChar & "]").IsValid
    Debug.Assert Not Parse(StrConv("[" & vbNullChar & "]", vbFromUnicode)).IsValid
    Debug.Assert Not Parse("{}}").IsValid
    Debug.Assert Not Parse("{"""":").IsValid
    Debug.Assert Not Parse("{""a"":/*comment*/""b""}").IsValid
    Debug.Assert Not Parse("{""a"": true} ""x""").IsValid
    Debug.Assert Not Parse("{,").IsValid
    Debug.Assert Not Parse("{""a").IsValid
    Debug.Assert Not Parse("{'a'").IsValid
    Debug.Assert Not Parse("[""\{[""\{[""\{[""\{").IsValid
    Debug.Assert Not Parse(ChrW$(&H17D)).IsValid
    Debug.Assert Not Parse("*").IsValid
    Debug.Assert Not Parse("{""a"":""b""}#{}").IsValid
    Debug.Assert Not Parse(BytesToString(&H5B, &HE2, &H81, &HA0, &H5D)).IsValid
    Debug.Assert Not Parse("[\u000A""""]").IsValid
    Debug.Assert Not Parse("{""abc"":1").IsValid
    Debug.Assert Not Parse(BytesToString(&HC3, &HA5)).IsValid
    Debug.Assert Not Parse(BytesToString(&H5B, &HE2, &H81, &HA0, &H5D)).IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
End Sub

Private Sub TestParseNumberValid()
    Debug.Assert Parse("0").Value = 0
    Debug.Assert Parse("-0").Value = 0
    Debug.Assert Parse("1").Value = 1
    Debug.Assert Parse("-1").Value = -1
    Debug.Assert Parse("12").Value = 12
    Debug.Assert Parse("-12").Value = -12
    Debug.Assert Parse("123").Value = 123
    Debug.Assert Parse(" 3").Value = 3
    Debug.Assert Parse("-1.2").Value = -1.2
    Debug.Assert Parse("0E0").Value = 0
    Debug.Assert Parse("0E1").Value = 0
    Debug.Assert Parse("0E-1").Value = 0
    Debug.Assert Parse("0E+1").Value = 0
    Debug.Assert Parse("12E4").Value = 120000
    Debug.Assert Parse("1E25").Value = 1E+25
    Debug.Assert Parse("1E-2").Value = 0.01
    Debug.Assert Parse("1E+2").Value = 100
    Debug.Assert Parse("123e45").Value = 1.23E+47
    Debug.Assert Parse("123.45e67").Value = 1.2345E+69
    Debug.Assert Parse("123.4567").Value = 123.4567
End Sub

Private Sub TestParseNumberValidLargeExponent()
    Debug.Assert Parse("[1.23456789E-999]").Value(1) = 0
    Debug.Assert Parse("[4.94065645841247E-324]").Value(1) = 4.94065645841247E-324
    Debug.Assert Parse("[1.79769313486231E308]").Value(1) = 1.79769313486231E+308
    Debug.Assert Parse("[-1.79769313486231E308]").Value(1) = -1.79769313486231E+308
    Debug.Assert Parse("[-4.94065645841247E-324]").Value(1) = -4.94065645841247E-324
#If Windows Then
    Debug.Assert Parse("-0.000000000000000000000000000000000000000000000000000001").Value = 0
    Debug.Assert Parse("[-922337203685477.5807]").Value(1) = -922337203685477.5807@
    Debug.Assert Parse("[-79228162514264337593543950335]").Value(1) = CDec("-79228162514264337593543950335")
    Debug.Assert Parse("[79228162514264337593543950335]").Value(1) = CDec("79228162514264337593543950335")
    Debug.Assert Parse("[-7.9228162514264337593543950335]").Value(1) = CDec("-7.9228162514264337593543950335")
    Debug.Assert Parse("[7.9228162514264337593543950335]").Value(1) = CDec("7.9228162514264337593543950335")
#Else
    Debug.Assert Parse("-0.000000000000000000000000000000000000000000000000000001").Value = -1E-54
    Debug.Assert Parse("[-922337203685477]").Value(1) = -922337203685477@
#End If
    Debug.Assert Parse("[-0.0000000000000000000000000001]").Value(1) = -1E-28
    Debug.Assert Parse("[0.1E000100]").Value(1) = 1E+99
    Debug.Assert Parse("[0.1E+00100]").Value(1) = 1E+99
    Debug.Assert Parse("-1234567890123456789012345678901234567890").Value - -1.23456789012346E+39 < 1E+25
    Debug.Assert Parse("1234567890123456789012345678901234567890").Value - 1.23456789012346E+39 < 1E+25
End Sub

Private Sub TestParseNumberInvalidLargeExponent()
    Debug.Assert Not Parse("0.1e0123456789").IsValid
    Debug.Assert Not Parse("-1e+1000").IsValid
    Debug.Assert Not Parse("1.1e+10000000000000").IsValid
    Debug.Assert Not Parse("1.1e-10000000000000").IsValid
    Debug.Assert Not Parse("1e-10000000000000").IsValid
    Debug.Assert Not Parse("-1123456e1000").IsValid
    Debug.Assert Not Parse("1123456e1000").IsValid
End Sub

Private Sub TestParseNumberInvalid()
    Debug.Assert Not Parse("123" & vbNullChar).IsValid
    Debug.Assert Not Parse("0123").IsValid  'Leading zeros are not allowed
    Debug.Assert Not Parse("00123").IsValid
    Debug.Assert Not Parse("0.").IsValid    'Point must be followed by at least a digit
    Debug.Assert Not Parse("1E").IsValid    'Exponent must be followed by at least a digit
    Debug.Assert Not Parse("00E0").IsValid
    Debug.Assert Not Parse("1E-").IsValid
    Debug.Assert Not Parse("1E+").IsValid
    Debug.Assert Not Parse("inf").IsValid
    Debug.Assert Not Parse("+inf").IsValid
    Debug.Assert Not Parse("Inf").IsValid
    Debug.Assert Not Parse("+Inf").IsValid
    Debug.Assert Not Parse("infinity").IsValid
    Debug.Assert Not Parse("Infinity").IsValid
    Debug.Assert Not Parse("-inf").IsValid
    Debug.Assert Not Parse("-Inf").IsValid
    Debug.Assert Not Parse("-infinity").IsValid
    Debug.Assert Not Parse("-Infinity").IsValid
    Debug.Assert Not Parse("NaN").IsValid
    Debug.Assert Not Parse("nan").IsValid
    Debug.Assert Not Parse("-NaN").IsValid
    Debug.Assert Not Parse("-nan").IsValid
    Debug.Assert Not Parse("-nan(ind)").IsValid
    Debug.Assert Not Parse("0.1E+001E00").IsValid
    Debug.Assert Not Parse(".-1").IsValid
    Debug.Assert Not Parse(".1e-2").IsValid
    Debug.Assert Not Parse(".12").IsValid
    Debug.Assert Not Parse("+1").IsValid
    Debug.Assert Not Parse("+12").IsValid
    Debug.Assert Not Parse("++1").IsValid
    Debug.Assert Not Parse("--1").IsValid
    Debug.Assert Not Parse("0.1.2").IsValid
    Debug.Assert Not Parse("0.1.").IsValid
    Debug.Assert Not Parse("0.1E").IsValid
    Debug.Assert Not Parse("0.1e+").IsValid
    Debug.Assert Not Parse("0.1e-").IsValid
    Debug.Assert Not Parse("0.E1").IsValid
    Debug.Assert Not Parse("0.e").IsValid
    Debug.Assert Not Parse("0.e+").IsValid
    Debug.Assert Not Parse("0.E+").IsValid
    Debug.Assert Not Parse("0.e-").IsValid
    Debug.Assert Not Parse("0e").IsValid
    Debug.Assert Not Parse("0e+").IsValid
    Debug.Assert Not Parse("0e-").IsValid
    Debug.Assert Not Parse("-01").IsValid
    Debug.Assert Not Parse("-001").IsValid
    Debug.Assert Not Parse("-00.1").IsValid
    Debug.Assert Not Parse("-1.0.").IsValid
    Debug.Assert Not Parse("1.0e").IsValid
    Debug.Assert Not Parse("-1.0e").IsValid
    Debug.Assert Not Parse("1.0e-").IsValid
    Debug.Assert Not Parse("1.0e+").IsValid
    Debug.Assert Not Parse("1 0.0").IsValid
    Debug.Assert Not Parse("1Ee2").IsValid
    Debug.Assert Not Parse("1.").IsValid
    Debug.Assert Not Parse("-1.").IsValid
    Debug.Assert Not Parse("+1.").IsValid
    Debug.Assert Not Parse("1.e2").IsValid
    Debug.Assert Not Parse("1.e+2").IsValid
    Debug.Assert Not Parse("1.e-2").IsValid
    Debug.Assert Not Parse("2.e+").IsValid
    Debug.Assert Not Parse("0e+-1").IsValid
    Debug.Assert Not Parse("0e-+1").IsValid
    Debug.Assert Not Parse("1+2").IsValid
    Debug.Assert Not Parse("0x0").IsValid
    Debug.Assert Not Parse("0x1").IsValid
    Debug.Assert Not Parse("0x00").IsValid
    Debug.Assert Not Parse("0x12").IsValid
    Debug.Assert Not Parse("1.2a").IsValid
    Debug.Assert Not Parse("-1.2a").IsValid
    Debug.Assert Not Parse("112a").IsValid
    Debug.Assert Not Parse(StrConv("112", vbFromUnicode) & ChrW$(&HE55D)).IsValid
    Debug.Assert Not Parse(StrConv("1E2", vbFromUnicode) & ChrW$(&HE55D)).IsValid
    Debug.Assert Not Parse(StrConv("1", vbFromUnicode) & ChrW$(&HE55D)).IsValid
    Debug.Assert Not Parse("-abc").IsValid
    Debug.Assert Not Parse("- 7").IsValid
    Debug.Assert Not Parse("-07").IsValid
    Debug.Assert Not Parse("-.7").IsValid
    Debug.Assert Not Parse("-0x").IsValid
    Debug.Assert Not Parse("1ex").IsValid
    Debug.Assert Not Parse("1e" & ChrW$(&HE55D)).IsValid
    Debug.Assert Not Parse(ChrW$(-239)).IsValid 'U+FF11 (Fullwidth Digit One Unicode Character)
    Debug.Assert Not Parse("1a3").IsValid
    Debug.Assert Not Parse("1.2a3").IsValid
    Debug.Assert Not Parse("1.2a-3").IsValid
    Debug.Assert Not Parse("1.23456789a-999").IsValid
    Debug.Assert Not Parse("1.23456789x-999").IsValid
    Debug.Assert Not Parse("1.23456789h-999").IsValid
    Debug.Assert Not Parse(BytesToString(&H31, &H65, &HE5)).IsValid
End Sub

Private Sub TestParseStringValid()
    Debug.Assert Parse("""""").Value = vbNullString
    Debug.Assert Parse("""abc""").Value = "abc"
    Debug.Assert Parse("""abc """).Value = "abc "
    Debug.Assert Parse("""abc""" & vbNewLine).Value = "abc"
    Debug.Assert Parse("""abc""" & vbNewLine).Value = "abc"
    Debug.Assert Parse(""" """).Value = " "
    Debug.Assert Parse("""a/*b*/c/*d//e""").Value = "a/*b*/c/*d//e" 'Comments are treated as normal text while inside strings
    Debug.Assert Parse(BytesToString(commaA, &HF4, &H8F, &HBF, &HBF, commaA)).Value = ChrW$(&HDBFF) & ChrW$(&HDFFF) 'Non UTF8 character U+10FFFF
    Debug.Assert Parse(BytesToString(commaA, &HEF, &HBF, &HBF, commaA)).Value = ChrW$(&HFFFF)                       'Non UTF8 character U+FFFF
    Debug.Assert Parse(BytesToString(commaA, &H2C, commaA)).Value = ","
    Debug.Assert Parse(BytesToString(commaA, &HCF, &H80, commaA)).Value = ChrW$(&H3C0) 'Math PI
    Debug.Assert Parse(BytesToString(commaA, &HF0, &H9B, &HBF, &HBF, commaA)).Value = ChrW$(&HD82F) & ChrW$(&HDFFF) 'Reserved character U+1BFFF
    Debug.Assert Parse(BytesToString(commaA, &HE2, &H80, &HA8, commaA)).Value = ChrW$(&H2028) 'Line Separator (LS)
    Debug.Assert Parse(BytesToString(commaA, &HE2, &H80, &HA9, commaA)).Value = ChrW$(&H2029) 'Paragraph Separator (PS)
    Debug.Assert Parse(BytesToString(commaA, &H7F, commaA)).Value = Chr$(&H7F) 'DEL
    Debug.Assert Parse(BytesToString(commaA, &H61, &H7F, &H61, commaA)).Value = "a" & Chr$(&H7F) & "a"
    Debug.Assert Parse(BytesToString(commaA, &HE2, &H8D, &H82, &HE3, &H88, &HB4, &HE2, &H8D, &H82, commaA)).Value = ChrW$(&H2342) & ChrW$(&H3234) & ChrW$(&H2342)
    With New Dictionary
        .AllowDuplicateKeys = True
        .CompareMode = vbTextCompare
        .Add "key", 1
        .Add "KEY", 2
        Debug.Assert AreEqual(Parse("{""key"":1,""KEY"":2}", allowDuplicatedKeys:=True, keyCompareMode:=vbTextCompare).Value, .Self)
    End With
End Sub

Private Sub TestParseStringValidEscape()
    Debug.Assert Parse("""value\nwith\nnewlines""").Value = "value" & vbLf & "with" & vbLf & "newlines"
    Debug.Assert Parse("""value with \""escaped quotes\""""").Value = "value with ""escaped quotes"""
    Debug.Assert Parse("""value\twith\ttabs""").Value = "value" & vbTab & "with" & vbTab & "tabs"
    Debug.Assert Parse("""value with \\backslashes\\""").Value = "value with \backslashes\"
    Debug.Assert Parse("""a\\b""").Value = Parse("""a\u005Cb""").Value
    Debug.Assert Parse("""\u0060\u012a\u12AB""").Value = "`" & ChrW$(&H12A) & ChrW$(&H12AB)
    Debug.Assert Parse("""\uD801\udc37""").Value = ChrW$(&HD801) & ChrW$(&HDC37) 'Surrogate pair
    Debug.Assert Parse("""\ud83d\ude39\ud83d\udc8d""").Value = ChrW$(&HD83D) & ChrW$(&HDE39) & ChrW$(&HD83D) & ChrW$(&HDC8D) 'Surrogate pairs
    Debug.Assert Parse("""\""\\\/\b\f\n\r\t""").Value = """\/" & vbBack & vbFormFeed & vbLf & vbCr & vbTab
    Debug.Assert Parse("""\\u0000""").Value = "\u0000"
    Debug.Assert Parse("""\""""").Value = """"
    Debug.Assert Parse("""\\a""").Value = "\a"
    Debug.Assert Parse("""\\n""").Value = "\n"
    Debug.Assert Parse("""\u0012""").Value = Chr$(&H12) 'Device Control 2
    Debug.Assert Parse(StrConv("""\uFFFF""", vbFromUnicode)).Value = ChrW$(&HFFFF) 'Escaped non-character
    Debug.Assert Parse("""\uDBFF\uDFFF""").Value = ChrW$(&HDBFF) & ChrW$(&HDFFF) 'Last surrogate pair
    Debug.Assert Parse("""new\u00A0line""").Value = "new" & ChrW$(&HA0) & "line" 'No-Break Space (nbsp)
    Debug.Assert Parse("""\u0000""").Value = vbNullChar
    Debug.Assert Parse("""\u002c""").Value = ","
    Debug.Assert Parse("""\uD834\uDd1e""").Value = ChrW$(&HD834) & ChrW$(&HDD1E)
    Debug.Assert Parse("""\u0821""").Value = ChrW$(&H821)
    Debug.Assert Parse("""\u0123""").Value = ChrW$(&H123)
    Debug.Assert Parse("""\u0061\u30af\u30EA\u30b9""").Value = ChrW$(&H61) & ChrW$(&H30AF) & ChrW$(&H30EA) & ChrW$(&H30B9)
    Debug.Assert Parse("""new\u000Aline""").Value = "new" & vbLf & "line"
    Debug.Assert Parse("""\u0022""").Value = """"
    Debug.Assert Parse("""\u005C""").Value = "\"
    Debug.Assert Parse("""\u200B""").Value = ChrW$(&H200B)
    Debug.Assert Parse("""\u2064""").Value = ChrW$(&H2064)
    Debug.Assert Parse("""\uFDD0""").Value = ChrW$(&HFDD0)
    Debug.Assert Parse("""\uFFFE""").Value = ChrW$(&HFFFE)
End Sub

Private Sub TestParseStringInvalid()
    Debug.Assert Not Parse("{""key"":1,""KEY"":2}", keyCompareMode:=vbTextCompare).IsValid
    Debug.Assert Not Parse("""""a").IsValid
    Debug.Assert Not Parse("""\uD800\""").IsValid
    Debug.Assert Not Parse("""\uD800\u""").IsValid
    Debug.Assert Not Parse("""\uD800\u1""").IsValid
    Debug.Assert Not Parse("""\uD800\u1x""").IsValid
    Debug.Assert Not Parse(ChrW$(&H17D)).IsValid
    Debug.Assert Not Parse("""a" & vbNullChar & "a""").IsValid
    Debug.Assert Not Parse("""\" & vbNullChar & """").IsValid
    Debug.Assert Not Parse("""\" & vbTab & """").IsValid
    Debug.Assert Not Parse("""\" & vbLf & """").IsValid
    Debug.Assert Not Parse("""\" & vbCr & """").IsValid
    Debug.Assert Not Parse("""\" & vbFormFeed & """").IsValid
    Debug.Assert Not Parse("""" & vbNewLine & """").IsValid
    Debug.Assert Not Parse("""" & vbTab & """").IsValid
    Debug.Assert Not Parse("""\x00""").IsValid
    Debug.Assert Not Parse("""\\\""").IsValid
    Debug.Assert Not Parse("""\?""").IsValid
    Debug.Assert Not Parse(BytesToString(commaA, &H5C, &HF0, &H9F, &H8C, &H80, commaA)).IsValid '\Emoji
    Debug.Assert Not Parse("""\""").IsValid
    Debug.Assert Not Parse("""\u00a""").IsValid
    Debug.Assert Not Parse("""\uD834\uDd""").IsValid
    Debug.Assert Not Parse("""\uD800\uD800\x""").IsValid
    Debug.Assert Not Parse("""\z""").IsValid
    Debug.Assert Not Parse("""\ugggg""").IsValid
    Debug.Assert Not Parse("""\" & Chr$(&HE5) & """").IsValid
    Debug.Assert Not Parse("""\u" & Chr$(&HE5) & """").IsValid
    Debug.Assert Not Parse("\u0020""asd""").IsValid
    Debug.Assert Not Parse("\n").IsValid 'No quotes
    Debug.Assert Not Parse("'""'").IsValid
    Debug.Assert Not Parse("abc").IsValid
    Debug.Assert Not Parse("""\").IsValid
    Debug.Assert Not Parse("""\UA66D""").IsValid
    Debug.Assert Not Parse(BytesToString(commaA, &H5C, &HE5, commaA)).IsValid
End Sub

'https://datatracker.ietf.org/doc/html/rfc8259#section-7
'\ must be escaped
'" must be escaped
'Control characters U+0000 through U+001F must be escaped
Private Sub TestParseStringInvalidEscape()
    Debug.Assert Not Parse("""\""").IsValid
    Debug.Assert Not Parse("""a\""").IsValid
    Debug.Assert Not Parse("""""""").IsValid
    Debug.Assert Not Parse("""---""---""").IsValid
    Dim i As Long
    For i = &H0 To &H1F
        Debug.Assert Not Parse("""" & Chr$(i) & """").IsValid
    Next i
End Sub

Private Sub TestParseStringLoneSurrogates()
    'Allow
    Debug.Assert Parse("""\uDFAA""", failIfLoneSurrogate:=False).Value = ChrW$(&HDFAA) 'Lone 2nd surrogate
    Debug.Assert Parse("""\uDADA""", failIfLoneSurrogate:=False).Value = ChrW$(&HDADA) 'Lone 1st surrogate
    Debug.Assert Parse("""\uD888\u1234""", failIfLoneSurrogate:=False).Value = ChrW$(&HD888) & ChrW$(&H1234) 'Invalid 2nd surrogate
    Debug.Assert Parse("""\uD800\n""", failIfLoneSurrogate:=False).Value = ChrW$(&HD800) & vbLf
    Debug.Assert Parse("""\uDd1ea""", failIfLoneSurrogate:=False).Value = ChrW$(&HDD1E) & "a"
    Debug.Assert Parse("""\uD800\uD800\n""", failIfLoneSurrogate:=False).Value = ChrW$(&HD800) & ChrW$(&HD800) & vbLf
    Debug.Assert Parse("""\uD800""", failIfLoneSurrogate:=False).Value = ChrW$(&HD800)
    Debug.Assert Parse("""\uD800abc""", failIfLoneSurrogate:=False).Value = ChrW$(&HD800) & "abc"
    Debug.Assert Parse("""\uDd1e\uD834""", failIfLoneSurrogate:=False).Value = ChrW$(&HDD1E) & ChrW$(&HD834) 'Inverted surrogates
    '
    'Fail
    Debug.Assert Not Parse("""\uDFAA""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uDADA""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD888\u1234""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD800\n""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uDd1ea""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD800\uD800\n""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD800""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD800abc""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uDd1e\uD834""", failIfLoneSurrogate:=True).IsValid
End Sub

Private Sub TestParseStringInvalidUTF8()
    'Allow - replace each bad byte with with 0xfffd default character
    Debug.Assert Parse(BytesToString(commaA, &HFF, commaA), failIfInvalidByteSequence:=False).Value = ChrW$(&HFFFD) 'Default character replaces invalid byte sequence
    Debug.Assert Parse(BytesToString(commaA, &HFF, &HFF, &HFF, commaA)).Value = String$(3, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(commaA, &H81, commaA)).Value = ChrW$(&HFFFD)
    Debug.Assert Parse(BytesToString(commaA, &HC0, &HAF, commaA)).Value = String$(2, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(commaA, &HFC, &H83, &HBF, &HBF, &HBF, &HBF, commaA)).Value = String$(6, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(commaA, &HFC, &H80, &H80, &H80, &H80, &H80, commaA)).Value = String$(6, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(commaA, &HE0, &HFF, commaA)).Value = String$(2, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(commaA, &HE6, &H97, &HA5, &HD1, &H88, &HFA, commaA)).Value = ChrW$(&H65E5) & ChrW$(&H448) & ChrW$(&HFFFD)
#If Windows Then
    Debug.Assert Parse(BytesToString(commaA, &HF4, &HBF, &HBF, &HBF, commaA)).Value = String$(3, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(commaA, &HED, &HA0, &H80, commaA)).Value = String$(2, ChrW$(&HFFFD)) 'Single surrogate 0xD800
#Else
    Debug.Assert Parse(BytesToString(commaA, &HF4, &HBF, &HBF, &HBF, commaA)).Value = ChrW$(&HFFFD)
    Debug.Assert Parse(BytesToString(commaA, &HED, &HA0, &H80, commaA)).Value = ChrW$(&HFFFD)
#End If
    '
    'Fail
    Debug.Assert Not Parse(BytesToString(commaA, &HFF, commaA), failIfInvalidByteSequence:=True).IsValid
    
    Debug.Assert Not Parse(BytesToString(commaA, &H5B, &H22, &H81, &H22, &H5D, commaA), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(commaA, &H81, commaA), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(commaA, &HC0, &HAF, commaA), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(commaA, &HFC, &H83, &HBF, &HBF, &HBF, &HBF, commaA), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(commaA, &HFC, &H80, &H80, &H80, &H80, &H80, commaA), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(commaA, &HE0, &HFF, commaA), failIfInvalidByteSequence:=True).IsValid
#If Windows Then
    Debug.Assert Not Parse(BytesToString(commaA, &HF4, &HBF, &HBF, &HBF, commaA), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(commaA, &HED, &HA0, &H80, commaA), failIfInvalidByteSequence:=True).IsValid
#Else
    Debug.Assert Parse(BytesToString(commaA, &HF4, &HBF, &HBF, &HBF, commaA), failIfInvalidByteSequence:=True).Value = ChrW$(&HFFFD)
    Debug.Assert Parse(BytesToString(commaA, &HED, &HA0, &H80, commaA), failIfInvalidByteSequence:=True).Value = ChrW$(&HFFFD)
#End If
End Sub

'*******************************************************************************
'Utilities
'*******************************************************************************
Private Function AreEqual(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean
    On Error GoTo ErrorHandler
    'No need to check for vbDataObject
    If IsObject(v1) Then
        If Not IsObject(v2) Then Exit Function
        '
        If TypeOf v1 Is Collection Then
            If Not TypeOf v2 Is Collection Then Exit Function
            If v1.Count <> v2.Count Then Exit Function
            '
            Dim i As Long
            For i = 1 To v1.Count
                If Not AreEqual(v1(i), v2(i)) Then Exit Function
            Next i
            AreEqual = True
        ElseIf TypeOf v2 Is Dictionary Then
            If Not TypeOf v2 Is Dictionary Then Exit Function
            If v1.Count <> v2.Count Then Exit Function
            '
            Dim k As Variant
            Dim adk As Boolean
            '
            If IsFastDict() Then adk = v1.AllowDuplicateKeys Else adk = False
            If adk Then
                Dim k1() As Variant: k1 = v1.Keys
                Dim k2() As Variant: k2 = v2.Keys
                Dim i1() As Variant: i1 = v1.Items
                Dim i2() As Variant: i2 = v2.Items
                '
                For i = 0 To v1.Count - 1
                    If Not AreEqual(k1(i), k2(i)) Then Exit Function
                    If Not AreEqual(i1(i), i2(i)) Then Exit Function
                Next i
            Else
                For Each k In v1
                    If Not AreEqual(v1(k), v2(k)) Then Exit Function
                Next k
            End If
            AreEqual = True
        Else
            AreEqual = (TypeName(v1) = TypeName(v2))
        End If
        Exit Function
    End If
    If IsObject(v2) Then Exit Function
    '
    Dim vt1 As VbVarType: vt1 = VarType(v1)
    Dim vt2 As VbVarType: vt2 = VarType(v2)
    '
    If vt1 < vt2 Then 'Allow number comparison across types
        #If (x32) Or Mac Then
            Const vbLongLong = 20
        #End If
        '
        If vt1 <= vbNull Or vt1 > vbLongLong Then Exit Function
        If vt2 <= vbNull Or vt2 > vbLongLong Then Exit Function
        If vt1 = vbString Or vt1 = vbError Then Exit Function
        If vt1 = vbString Or vt2 = vbError Then Exit Function
    End If
    If IsNull(v1) Then
        AreEqual = True
    Else
        AreEqual = (v1 = v2)
    End If
ErrorHandler:
End Function
Private Function Collection(ParamArray values() As Variant) As Collection
    Dim v As Variant
    Dim coll As Collection
    '
    Set coll = New Collection
    For Each v In values
        coll.Add v
    Next v
    Set Collection = coll
End Function
Private Function Dictionary(ParamArray values() As Variant) As Dictionary
    Dim i As Long
    Dim dict As Dictionary
    '
    Set dict = New Dictionary
    For i = 0 To UBound(values) Step 2
        dict.Add values(i), values(i + 1)
    Next i
    Set Dictionary = dict
End Function
Private Function RepeatString(ByRef str As String _
                            , ByVal repeatTimes As Long) As String
    If repeatTimes <= 0 Or LenB(str) = 0 Then Exit Function
    '
    Dim newLength As Long: newLength = LenB(str) * repeatTimes
    RepeatString = Space$((newLength + 1) \ 2)
    If newLength Mod 2 = 1 Then RepeatString = MidB$(RepeatString, 2)
    '
    MidB$(RepeatString, 1) = str
    If repeatTimes > 1 Then MidB$(RepeatString, LenB(str) + 1) = RepeatString
End Function
Private Function IsFastDict() As Boolean
    Dim o As Object:     Set o = New Dictionary
    On Error Resume Next
    Dim s As Single:     s = o.LoadFactor
    Dim b As Boolean:    b = o.AllowDuplicateKeys
    Dim d As Dictionary: Set d = o.Self.Factory
    IsFastDict = (Err.Number = 0)
    On Error GoTo 0
End Function
Private Property Get PosInf() As Double 'IEEE754 +inf
    On Error Resume Next
    PosInf = 1 / 0
    On Error GoTo 0
End Property
Private Property Get SNaN() As Double 'IEEE754 signaling NaN (sNaN)
    On Error Resume Next
    SNaN = 0 / 0
    On Error GoTo 0
End Property
Private Property Get NegInf() As Double 'IEEE754 -inf
    NegInf = -PosInf
End Property
Private Property Get QNaN() As Double 'IEEE754 quiet NaN (qNaN)
    QNaN = -SNaN
End Property
'Useful for building UTF8 strings
Private Function BytesToString(ParamArray bytes() As Variant) As String
    Dim i As Long
    Dim b() As Byte: ReDim b(0 To UBound(bytes))
    For i = 0 To UBound(bytes)
        b(i) = bytes(i)
    Next i
    BytesToString = b
End Function
