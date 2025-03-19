## Table of Contents

- [Parser](#parser)
  - [Extensions](#extensions)
    - [Encoding](#encoding)
    - [Trailing comma](#trailing-comma)
	- [Duplicated keys](#duplicated-keys)
	- [BOM](#bom)
	- [Invalid byte sequence](#invalid-byte-sequence)
	- [Lone surrogates](#lone-surrogates)
	- [Nesting depth](#nesting-depth)

- [Serializer](#serializer)

## Parser

```Parse``` method: 
- [RFC 8259](https://datatracker.ietf.org/doc/html/rfc8259) compliant
- non-recursive implementation - avoids 'Out of stack space' for deep nesting
- fast, for a native implementation 
- automatic encoding detection and conversion. Supports: ```UTF8```, ```UTF16LE```, ```UTF16BE```, ```UTF32LE```, ```UTF32BE```
- various extensions via the available parameters - see [Parser extensions](https://github.com/cristianbuse/VBA-FastJSON/blob/master/Documentation.md#extensions).

Does not raise errors. It returns the following custom structure / type:
```VBA
Public Type ParseResult
    Value As Variant
    IsValid As Boolean
    Error As String
End Type
```

According to [Section 2](https://datatracker.ietf.org/doc/html/rfc8259#section-2):

> A JSON text is a serialized value.  Note that certain previous specifications of JSON constrained a JSON text to be an object or an array.

So, ```Value``` can also be a string, number or literal (null, false, true). VBA users would be familiar with issues when assigning a ```Variant``` returned by a method, as seen [here](https://stackoverflow.com/questions/35750449/how-can-i-assign-a-variant-to-a-variant-in-vba). The use of a custom return type will allow constructs like the following:
```VBA
With Parse(json)
    If .IsValid Then
        If IsObject(.Value) Then
            '...
        Else
            '...
        End If
    Else
        MsgBox .Error
        '...
    End If
End With
```
Which gives the user a way to avoid calling the ```Parse``` method twice or to use an auxiliary method to assign the result ```ByRef```.

However, the following is still possible if the user already trusts the outcome:
```VBA
Dim dict As Dictionary
Set dict = Parse("{...}").Value
``` 
With the minor inconvenience of typing the extra ```.Value```.

If parsing fails, additionally to ```.IsValid``` returning ```False```, the ```.Value``` will be set to the special ```Missing``` value which is a ```vbError``` type of ```Variant``` and can be used with the ```IsMissing``` built-in method. This is purely for convenience.

### Extensions

As per [RFC 8259 section 9](https://datatracker.ietf.org/doc/html/rfc8259#section-9):
> A JSON parser MAY accept non-JSON forms or extensions.

This parser allows extensions via the available arguments. Method signature:
```VBA
Public Function Parse(ByRef jsonText As Variant _
                    , Optional ByVal jpCode As JsonPageCode = jpCodeAutoDetect _
                    , Optional ByVal ignoreTrailingComma As Boolean = False _
                    , Optional ByVal allowDuplicatedKeys As Boolean = False _
                    , Optional ByVal keyCompareMode As VbCompareMethod = vbBinaryCompare _
                    , Optional ByVal failIfBOMDetected As Boolean = False _
                    , Optional ByVal failIfInvalidByteSequence As Boolean = False _
                    , Optional ByVal failIfLoneSurrogate As Boolean = False _
                    , Optional ByVal maxNestingDepth As Long = 128) As ParseResult
```

#### Encoding

[Section 8.1](https://datatracker.ietf.org/doc/html/rfc8259#section-8.1)
> JSON text exchanged between systems that are not part of a closed ecosystem MUST be encoded using UTF-8

Since the intention is to be used also in closed ecosystems, this parsers supports more than just UTF-8.

The optional ```jpCode``` is by default set to ```jpCodeAutoDetect```. This implementation will detect:
- ```UTF8```
- ```UTF16LE```
- ```UTF16BE```
- ```UTF32LE```
- ```UTF32BE```

The input ```jsonText``` will be automatically converted to ```UTF16LE``` and parsed accordingly. This should be convenient for most users as it avoids the need for additional conversion tools. Just a note - if users are in need of such tools for other non-json tasks, then it is recommended to use the excellent [VBA-StringTools](https://github.com/guwidoe/VBA-StringTools) library.

Testing shows that the encoding auto detection in this repository is reliable. However, the user has the ability to force the parser to treat the input ```jsonText``` as a particular encoding e.g. ```Parse json, jpCodeUTF8```, using the available ```JsonPageCode``` enum. This can be useful if the same encoding is always expected under a specific scenario.

#### Trailing comma

By default (```ignoreTrailingComma = False```), trailing commas are not allowed e.g. parsing ```[1,]``` will fail.

If ```ignoreTrailingComma = True``` then a single trailing comma will be allowed e.g. ```[1,]``` or ```{"key":value,}``` allowed but ```[1,,]``` not allowed.

#### Duplicated keys

[Section 4](https://datatracker.ietf.org/doc/html/rfc8259#section-4)

> An object whose names are all unique is interoperable in the sense that all software implementations receiving that object will agree on the name-value mappings.

With the above in mind, by default (```allowDuplicatedKeys = False```), this implementation does not allow duplicated keys e.g. ```{"a":1,"a":2}``` not allowed.

If ```allowDuplicatedKeys = True``` then duplicates are allowed e.g. ```{"a":1,"a":2}```, but note that this will only work if using [VBA-FastDictionary](https://github.com/cristianbuse/VBA-FastDictionary).

If using ```Scripting.Dictionary``` then this argument does nothing i.e. duplicates are not allowed.

#### Key Compare Mode

By default (```keyCompareMode = vbBinaryCompare```), object keys are case-sensitive. For example, ```{"a":1,"A":2}``` is allowed even if ```allowDuplicatedKeys = False```.

The user has the ability to set ```keyCompareMode``` to ```vbTextCompare``` or a specific locale ID (LCID) which would then allow objects like ```{"a":1,"A":2}``` but only if ```allowDuplicatedKeys = True``` and if using VBA-FastDictionary.

#### BOM

[Section 8.1](https://datatracker.ietf.org/doc/html/rfc8259#section-8.1)
> Implementations MUST NOT add a byte order mark (U+FEFF) to the beginning of a networked-transmitted JSON text. In the interests of interoperability, implementations that parse JSON texts MAY ignore the presence of a byte order mark rather than treating it as an error.

By default (```failIfBOMDetected = False```), this parser ignores BOM e.g. byte sequence ```0xFFFE3100``` is treated as ```0x3100``` (where ```0xFFFE``` is the UTF16LE BOM) which results in a number value of ```1``` (ascii 49 dec / 31 hex).

If ```failIfBOMDetected = True``` then parsing fails if BOM is detected at the beginning of the json input.

#### Invalid byte sequence

Please note that ```failIfInvalidByteSequence``` only has effect if conversion is needed (e.g. UTF8 to UTF16LE). For UTF16LE inputs, this does nothing. This is because the APIs used to convert between encodings (```iconv``` on Mac and ```MultiByteToWideChar``` on Windows) will either error out or replace the invalid byte sequences with a replacement character.

By default (```failIfInvalidByteSequence = False```), each invalid byte sequence / unit is replaced with ```U+FFFD``` replacement character e.g. ```0x22FF22``` (UTF8) will return a ```0xFFFD``` string. See approach [3](https://unicode.org/review/pr-121.html).

Parsing fails if ```failIfInvalidByteSequence = True``` and an invalid sequence is detected.

Please also note that in specific scenarios, the number of ```0xFFFD``` characters replacing invalid sequences may differ between Mac and Windows as seen in testing [here](https://github.com/cristianbuse/VBA-FastJSON/blob/02833b183a06c2234c3104ecf2af7447f6267559/src/test/TestLibJSON.bas#L740-L746).

#### Lone surrogates

According to RFC, an escaped invalid lone UTF-16 surrogate is perfectly valid grammar e.g. ```"\uD800"```, and should be parsed.

By default (```failIfLoneSurrogate = False```), will allow lone surrogates (U+D800 to U+DFFF).

Parsing will fail if  ```failIfLoneSurrogate = True``` and a lone surrogate is detected.

#### Nesting depth

In real-world scenarios, a depth of more than 16 levels is rarely encountered or good practice.

Regardless, the default maximum depth for this parser is ```maxNestingDepth = 128```. This works well with both ```Collection``` and ```Scripting.Dictionary``` and avoids 'Out of stack space' issues.

Note that ```VBA-FastDictionary``` has no nesting limit unlike ```Scripting.Dictionary```. As seen in [this](https://github.com/cristianbuse/VBA-FastJSON/blob/02833b183a06c2234c3104ecf2af7447f6267559/src/test/TestLibJSON.bas#L300) test, it can easily manage 10000 nesting levels and more.

This option is in place to avoid application crashes and allow full user control.

## Serializer

Work in progress...
