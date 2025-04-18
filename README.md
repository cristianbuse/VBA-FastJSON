# VBA-FastJSON [![Mentioned in Awesome VBA](https://awesome.re/mentioned-badge.svg)](https://github.com/sancarn/awesome-vba)
Fast Native JSON Parser / Serializer for VBA. Compatible with Windows and Mac.

[RFC 8259](https://datatracker.ietf.org/doc/html/rfc8259) compliant.

This Parser / Serializer is intended for VBA. However, it is compatible with VBA7 / [twinBASIC](https://twinbasic.com/) / VB6 / VBA6.

## Installation

Download the [latest release](https://github.com/cristianbuse/VBA-FastJSON/releases/latest), extract and import the ```LibJSON.bas``` module into your project.

Additionally, a ```Dictionary``` is required. While you can use ```Scripting.Dictionary``` (Microsoft Scripting Runtime reference - scrrun.dll on Windows), it is recommended to use [VBA-FastDictionary](https://github.com/cristianbuse/VBA-FastDictionary) because:
- is Mac compatible
- is faster in almost every way - see [Benchmarking VBA-FastDictionary](https://github.com/cristianbuse/VBA-FastDictionary/blob/master/benchmarking/README.md)
- allows endless nesting
- will still work if ```Scripting.Dictionary``` becomes obsolete

For more information see [Cons of Scripting.Dictionary](https://github.com/cristianbuse/VBA-FastDictionary/blob/master/benchmarking/README.md#scriptingdictionary)

## Parser

For more details see [Parser documentation](https://github.com/cristianbuse/VBA-FastJSON/blob/master/Documentation.md#parser).

```Parse``` method: 
- [RFC 8259](https://datatracker.ietf.org/doc/html/rfc8259) compliant
- memory-efficient, non-recursive implementation - avoids 'Out of stack space' for deep nesting
- fast, for a native implementation 
- automatic encoding detection and conversion. Supports: ```UTF8```, ```UTF16LE```, ```UTF16BE```, ```UTF32LE```, ```UTF32BE```
- various extensions via the available parameters - see [Parser extensions](https://github.com/cristianbuse/VBA-FastJSON/blob/master/Documentation.md#extensions)
- json input can be a ```String``` or a one-dimensional array of ```Byte()``` or ```Integer()``` type
- input is parsed in place without making any copies
- does not throw errors

## Serializer

For more details see [Serializer documentation](https://github.com/cristianbuse/VBA-FastJSON/blob/master/Documentation.md#serializer).

```Serialize``` method: 
- memory-efficient, non-recursive implementation - avoids 'Out of stack space' for deep nesting
- fast, for a native implementation
- supports beautify / minify via the ```indentSpaces``` argument
- by default, cannot fail - see available options via the function parameters
- returns a ```String``` data type
- detects circular object references
- can sort keys
- supports multi-dimensional arrays, row-wise
- supports encoding: ```UTF8```, ```UTF16LE``` (default), ```UTF16BE```, ```UTF32LE```, ```UTF32BE```
- does not throw errors

## Testing

Download the [latest release](https://github.com/cristianbuse/VBA-FastJSON/releases/latest), extract and import the ```TestLibJSON.bas``` module into your project. Run ```RunAllJSONTests``` method. On failure, execution will stop on the first failed ```Assert```.

Many thanks to Nicolas Seriot ([@nst](https://github.com/nst)). This repo includes some of the tests found at [JSONTestSuite](https://github.com/nst/JSONTestSuite). A must-read, see his article: [Parsing JSON is a Minefield](https://seriot.ch/projects/parsing_json.html)!

## Demo

```VBA
Debug.Print Parse("{""key"":[1,2,3,4,5,true]}").Value("key")(6) 'True
'
Debug.Print Parse("[[[[[]]]]]", maxNestingDepth:=4).Error 'Max Depth Hit at char position 5
'
Debug.Print Parse("    false").Value 'False
'
Debug.Print Parse(ChrW$(&HBBEF) & ChrB$(&HBF) & "{}", failIfBOMDetected:=True).IsValid 'False
Debug.Print Parse(ChrW$(&HBBEF) & ChrB$(&HBF) & "{}").Value.Count '0
'
Debug.Print Parse("0E0").Value '0
'
Dim res As Variant
Dim jsonText() As Byte: ReadBytes "myFilePath", jsonText
'
With Parse(jsonText, jpCodeUTF8)
    If .IsValid Then
        If IsObject(.Value) Then Set res = .Value Else res = .Value
    Else
        MsgBox .Error
        Exit Sub
    End If
End With
'
Dim jsonData As New Dictionary
Dim json As String
'
jsonData.Add "d", 1
jsonData.Add 7, 2
jsonData.Add "a", 3
jsonData.Add "b", 4
'
json = Serialize(jsonData)                                        '{"d":1,"a":3,"b":4}
json = Serialize(jsonData, sortKeys:=True)                        '{"a":3,"b":4,"d":1}
json = Serialize(jsonData, sortKeys:=True, forceKeysToText:=True) '{"7":2,"a":3,"b":4,"d":1}
'
json = Serialize(jsonData, indentSpaces:=2)                       '{
                                                                  '  "d": 1,
                                                                  '  "a": 3,
                                                                  '  "b": 4
                                                                  '}
json = Serialize(jsonData, jpCode:=jpCodeUTF8)
```
