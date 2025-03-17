Attribute VB_Name = "LibJSON"
'''=============================================================================
''' VBA Fast JSON Parser & Serializer - github.com/cristianbuse/VBA-FastJSON
''' ---
''' MIT License - github.com/cristianbuse/VBA-FastJSON/blob/master/LICENSE
''' Copyright (c) 2024 Ion Cristian Buse
'''=============================================================================

''==============================================================================
'' Description:
''  - RFC 8259 compliant: https://datatracker.ietf.org/doc/html/rfc8259
''                        https://www.rfc-editor.org/rfc/rfc8259
''  - Mac OS compatible
''  - Performant, for a native VBA implementation
''  - Parser:
''     * Non-Recursive - avoids 'Out of stack space' for deep nesting
''     * Comments not supported, trailing or inline
''       They are simply treated as normal text if inside a json string
''     * Supports 'extensions' via the available arguments - see repository docs
''==============================================================================

'===============================================================================
' For convenience, the following are extracted from the above-mentioned RFC:
'  - A JSON text is a sequence of tokens: six structural characters, strings,
'    numbers, and three literal names: false, null, true
'  - The literal names MUST be lowercase.  No other literal names are allowed
'  - Structural characters are [ { ] } : ,
'  - Whitespace: &H09 (Hor. tab), &H0A (Lf or NewLine), &H0D (Cr), &H20 (Space)
'  - A JSON value MUST be an object, array, number, string, or literal
'  - An object has zero or more name/value pairs. A name is a string
'  - The names within an object SHOULD be unique, for better interoperability
'  - There is no requirement that the values in an array be of the same type
'  - Numbers:
'     * Leading zeros are not allowed
'     * A fraction part is a decimal point followed by one or more digits
'     * An exponent part begins with the letter E in uppercase or lowercase,
'       which may be followed by a plus or minus sign. The E and optional
'       sign are followed by one or more digits
'     * Infinity and NaN are not permitted
'     * Hex numbers are not allowed
'  - Strings:
'     * A string begins and ends with quotation marks
'     * All Unicode characters may be placed within the quotation marks, except
'       for the characters that MUST be escaped: question mark, reverse solidus
'       and control characters (U+0000 through U+001F)
'     * Any character may be escaped
'     * If the character is in the BMP (U+0000 through U+FFFF), then it may be
'       represented as a six-character sequence: \u followed by four hex digits
'       that encode the character's code point. The hex letters A through F can
'       be uppercase or lowercase
'     * If the character is outside the BMP, the character is represented as a
'       12-character sequence, encoding the UTF-16 surrogate pair
'       E.g. (U+1D11E) may be represented as "\uD834\uDD1E"
'===============================================================================

Option Explicit
Option Private Module

#If Mac Then
    #If VBA7 Then 'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv.3.html
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr, ByRef inBuf As LongPtr, ByRef inBytesLeft As LongPtr, ByRef outBuf As LongPtr, ByRef outBytesLeft As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As LongPtr, ByVal fromCode As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr) As Long
        Private Declare PtrSafe Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
        Private Declare Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long, ByRef inBuf As Long, ByRef inBytesLeft As Long, ByRef outBuf As Long, ByRef outBytesLeft As Long) As Long
        Private Declare Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As Long, ByVal fromCode As Long) As Long
        Private Declare Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long) As Long
        Private Declare Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
        Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
        Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
        Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    #End If
#End If

#Const Windows = (Mac = 0)
#Const x64 = Win64

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Private Enum DataTypeSize
    byteSize = 1
    intSize = 2
    longSize = 4
#If x64 Then
    ptrSize = 8
#Else
    ptrSize = 4
#End If
End Enum

#If x64 Then
    Private Const NullPtr As LongLong = 0^
#Else
    Private Const NullPtr As Long = 0&
#End If
Private Const VT_BYREF As Long = &H4000

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type
Private Enum SAFEARRAY_OFFSETS
    cDimsOffset = 0
    fFeaturesOffset = cDimsOffset + intSize
    cbElementsOffset = fFeaturesOffset + intSize
    cLocksOffset = cbElementsOffset + longSize
    pvDataOffset = cLocksOffset + ptrSize
    rgsaboundOffset = pvDataOffset + ptrSize
    rgsabound0_cElementsOffset = rgsaboundOffset
    rgsabound0_lLboundOffset = rgsabound0_cElementsOffset + longSize
End Enum

Private Type ByteAccessor
    arr() As Byte
    sa As SAFEARRAY_1D
End Type
Private Type IntegerAccessor
    arr() As Integer
    sa As SAFEARRAY_1D
End Type
Private Type LongAccessor
    arr() As Long
    sa As SAFEARRAY_1D
End Type
Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

Private Type FourByteTemplate
    b(0 To 3) As Byte
End Type
Private Type LongTemplate
    l As Long
End Type

Private Enum CharCode
    ccNull = 0          '0x00 'nullChar'
    ccBack = 8          '0x08 \b
    ccTab = 9           '0x09 \t
    ccLf = 10           '0x0A \n
    ccFormFeed = 12     '0x0C \f
    ccCr = 13           '0x0D \r
    ccSpace = 32        '0x20 'space'
    ccBang = 33         '0x21 !
    ccDoubleQuote = 34  '0x22 "
    ccPlus = 43         '0x2B +
    ccComma = 44        '0x2C ,
    ccMinus = 45        '0x2D -
    ccDot = 46          '0x2E .
    ccSlash = 47        '0x2F /
    ccZero = 48         '0x30 0
    ccNine = 57         '0x39 9
    ccColon = 58        '0x3A :
    ccArrayStart = 91   '0x5B [
    ccBackslash = 92    '0x5C \
    ccArrayEnd = 93     '0x5D ]
    ccBacktick = 96     '0x60 `
    ccLowB = 98         '0x62 b
    ccLowF = 102        '0x66 f
    ccLowN = 110        '0x6E n
    ccLowR = 114        '0x72 r
    ccLowT = 116        '0x74 t
    ccLowU = 117        '0x75 u
    ccObjectStart = 123 '0x7B {
    ccObjectEnd = 125   '0x7D }
End Enum

Private Enum CharType
    whitespace = 1
    numDigit = 2
    numSign = 3
    numExp = 4
    numDot = 5
End Enum

Private Type CharMap
    toType(ccTab To ccObjectEnd) As CharType
    nibs(ccZero To ccLowF) As Integer 'Nibble: 0 to F. Byte: 00 to FF
    nib1(0 To 15) As Integer
    nib2(0 To 15) As Integer
    nib3(0 To 15) As Integer
    nib4(0 To 15) As Integer
End Type

Private Enum AllowedToken
    allowNone = 0
    allowColon = 1
    allowComma = 2
    allowRBrace = 4
    allowRBracket = 8
    allowString = 16
    allowValue = 32
End Enum

Private Type ContextInfo
    coll As Collection
    dict As Dictionary
    tAllow As AllowedToken
    isDict As Boolean
    pendingKey As String
    pendingKeyPos As Long
End Type

Private Type JSONOptions
    ignoreTrailingComma As Boolean
    allowDuplicatedKeys As Boolean
    compMode As VbCompareMethod
    failIfLoneSurrogate As Boolean
    maxDepth As Long
End Type

Public Enum JsonPageCode
    jpCodeAutoDetect = -1
    [_jpcNone] = 0
    jpCodeUTF8 = 65001
    jpCodeUTF16LE = 1200
    jpCodeUTF16BE = 1201
    jpCodeUTF32LE = 12000
    jpCodeUTF32BE = 12001
    [_jpcCount] = 5
End Enum

Private Type BOM
    b() As Byte
    jpCode As JsonPageCode
    sizeB As Long
End Type

Public Type ParseResult
    Value As Variant
    IsValid As Boolean
    Error As String
End Type

'*******************************************************************************
'Parses a json string or byte/integer array
' - RFC 8259 compliant
' - Does not throw
' - Returns a convenient custom Type and it's .Value can be an object or not
' - Supports UTF8, UTF16 (LE and BE) and UTF32 (LE and BE on Mac only)
' - Returs UTF16LE texts only
' - Accepts Default Class Members so no need to check for IsObject
' - Numbers are rounded to the nearest Double and if possible Decimal is used
'Parameters:
' * jsonText: String or Byte / Integer 1D array
' * jpCode: json page code:
'     - jpCodeAutoDetect (default)
'     - force specific cp e.g. UTF8
' * ignoreTrailingComma:
'     - False (Default): e.g. [1,] not allowed
'     - True:            e.g. [1,] allowed, [1,,] not allowed
' * allowDuplicatedKeys:
'     - False (Default): e.g. {"a":1,"a":2} not allowed
'     - True:            e.g. {"a":1,"a":2} allowed
'                        Only works with VBA-FastDictionary
' * keyCompareMode:
'     - vbBinaryCompare (Default): e.g. {"a":1,"A":2} allowed
'                                  even if 'allowDuplicatedKeys' = False
'     - vbTextCompare (or LCID):   e.g. {"a":1,"A":2} allowed
'                                  only if 'allowDuplicatedKeys' = True
' * failIfBOMDetected:
'     - False (Default): e.g. 0xFFFE3100 same as 0x3100 => Number: 1
'     - True:            fails if any BOM detected
' * failIfInvalidByteSequence:
'     - False (Default): replaces each byte/unit with U+FFFD. See approach 3:
'                        https://unicode.org/review/pr-121.html
'                        e.g. 0x22FF22 (UTF8) => String: 0xFFFD
'     - True:            fails if invalid sequence detected
' * failIfLoneSurrogate:
'     - False (Default): allows lone U+D800 to U+DFFF e.g. "\uDFAA" allowed
'     - True:            fails if any lone surrogate detected
' * maxNestingDepth (default 128). Note that VBA-FastDictionary has no nesting
'                                  limit unlike Scripting.Dictionary
'*******************************************************************************
Public Function Parse(ByRef jsonText As Variant _
                    , Optional ByVal jpCode As JsonPageCode = jpCodeAutoDetect _
                    , Optional ByVal ignoreTrailingComma As Boolean = False _
                    , Optional ByVal allowDuplicatedKeys As Boolean = False _
                    , Optional ByVal keyCompareMode As VbCompareMethod = vbBinaryCompare _
                    , Optional ByVal failIfBOMDetected As Boolean = False _
                    , Optional ByVal failIfInvalidByteSequence As Boolean = False _
                    , Optional ByVal failIfLoneSurrogate As Boolean = False _
                    , Optional ByVal maxNestingDepth As Long = 128) As ParseResult
    Const pArrayOffset As Long = 8
    Static chars As IntegerAccessor
    Static bytes As ByteAccessor
    Static ptrs As PointerAccessor
    Dim jOptions As JSONOptions
    Dim bomCode As JsonPageCode
    Dim sizeB As Long
    Dim buff As String
    Dim vt As VbVarType: vt = VarType(jsonText)
    '
    If chars.sa.cDims = 0 Then 'Init memory accessors
        InitAccessor VarPtr(bytes), bytes.sa, byteSize
        InitAccessor VarPtr(chars), chars.sa, intSize
        InitAccessor VarPtr(ptrs), ptrs.sa, ptrSize
    End If
    If vt = vbString Then
        chars.sa.pvData = StrPtr(jsonText)
        sizeB = LenB(jsonText)
    ElseIf vt = vbArray + vbByte Or vt = vbArray + vbInteger Then
        ptrs.sa.pvData = VarPtr(jsonText)
        ptrs.sa.rgsabound0.cElements = 2 'Need 2 for reading 'sizeB'
        '
        vt = CLng(ptrs.arr(0) And &HFFFF&) 'VarType - Little Endian so fine
        ptrs.sa.pvData = ptrs.sa.pvData + pArrayOffset 'Read pointer in Variant
        If vt And VT_BYREF Then ptrs.sa.pvData = ptrs.arr(0)
        If ptrs.arr(0) = NullPtr Then GoTo UnexpectedInput
        ptrs.sa.pvData = ptrs.arr(0) 'SAFEARRAY address i.e. ArrPtr
        '
        'Check for One-Dimensional
        If CLng(ptrs.arr(0) And &HFF&) <> 1& Then GoTo UnexpectedInput
        '
        ptrs.sa.pvData = ptrs.sa.pvData + pvDataOffset
        chars.sa.pvData = ptrs.arr(0) 'Data address
        sizeB = CLng(ptrs.arr(1) And &H7FFFFFFF)  '# of array elements
        If vt = vbArray + vbInteger Then sizeB = sizeB * 2
    Else
        GoTo UnexpectedInput
    End If
    '
    If (sizeB >= 2) Then
        bytes.sa.pvData = chars.sa.pvData
        bytes.sa.rgsabound0.cElements = sizeB
        bomCode = DetectBOM(bytes.arr, sizeB, chars.sa.pvData)
        If (bomCode <> [_jpcNone]) And failIfBOMDetected Then
            Parse.Error = "BOM not allowed"
            GoTo Clean
        End If
    End If
    If sizeB = 0 Then
        Parse.Error = "Expected at least one character"
        GoTo Clean
    End If
    If jpCode = jpCodeAutoDetect Then
        bytes.sa.pvData = chars.sa.pvData
        bytes.sa.rgsabound0.cElements = sizeB
        jpCode = DetectCodePage(bytes.arr, sizeB)
        If jpCode = [_jpcNone] Then
            If bomCode = [_jpcNone] Then
                Parse.Error = "Could not determine encoding"
                GoTo Clean
            End If
            jpCode = bomCode
        End If
    End If
    '
    If jpCode = jpCodeUTF16LE Then
        chars.sa.rgsabound0.cElements = sizeB \ 2
    ElseIf jpCode = jpCodeUTF16BE Then
        ReverseBytes chars.sa.pvData, sizeB, buff _
                   , chars.sa.rgsabound0.cElements, chars.sa.pvData
    Else
        If Not Decode(chars.sa.pvData, sizeB, jpCode, buff _
                    , chars.sa.rgsabound0.cElements, chars.sa.pvData _
                    , Parse.Error, failIfInvalidByteSequence) Then GoTo Clean
    End If
    '
    jOptions.ignoreTrailingComma = ignoreTrailingComma
    jOptions.allowDuplicatedKeys = allowDuplicatedKeys
    jOptions.compMode = keyCompareMode
    jOptions.failIfLoneSurrogate = failIfLoneSurrogate
    jOptions.maxDepth = maxNestingDepth 'Negative numbers will allow 0 depth
    '
    Parse.IsValid = ParseChars(chars.arr, jOptions, Parse.Value, Parse.Error)
Clean:
    chars.sa.rgsabound0.cElements = 0: chars.sa.pvData = NullPtr
    bytes.sa.rgsabound0.cElements = 0: bytes.sa.pvData = NullPtr
    ptrs.sa.rgsabound0.cElements = 0:  ptrs.sa.pvData = NullPtr
Exit Function
UnexpectedInput:
    Parse.Error = "Expected JSON String or Byte/Integer 1D Array"
    GoTo Clean
End Function

Private Function DetectBOM(ByRef bytes() As Byte _
                         , ByRef sizeB As Long _
                         , ByRef ptr As LongPtr) As JsonPageCode
    Static boms(0 To [_jpcCount] - 1) As BOM
    Dim i As Long
    Dim j As Long
    Dim wasFound As Boolean
    '
    If boms(0).sizeB = 0 Then 'https://en.wikipedia.org/wiki/Byte_order_mark
        InitBOM boms(0), jpCodeUTF8, &HEF, &HBB, &HBF
        InitBOM boms(1), jpCodeUTF16LE, &HFF, &HFE
        InitBOM boms(2), jpCodeUTF16BE, &HFE, &HFF
        InitBOM boms(3), jpCodeUTF32LE, &HFF, &HFE, &H0, &H0
        InitBOM boms(4), jpCodeUTF32BE, &H0, &H0, &HFE, &HFF
    End If
    For i = 0 To [_jpcCount] - 1
        With boms(i)
            If sizeB >= .sizeB Then
                wasFound = True
                For j = 0 To .sizeB - 1
                    If bytes(j) <> .b(j) Then
                        wasFound = False
                        Exit For
                    End If
                Next j
                If wasFound Then
                    DetectBOM = .jpCode
                    sizeB = sizeB - .sizeB
                    ptr = ptr + .sizeB
                    Exit Function
                End If
            End If
        End With
    Next i
End Function
Private Sub InitBOM(ByRef b As BOM _
                  , ByVal jpCode As JsonPageCode _
                  , ParamArray bomBytes() As Variant)
    Dim i As Long
    b.jpCode = jpCode
    b.sizeB = UBound(bomBytes) + 1
    ReDim b.b(0 To b.sizeB - 1)
    For i = 0 To b.sizeB - 1
        b.b(i) = bomBytes(i)
    Next i
End Sub
Private Function DetectCodePage(ByRef bytes() As Byte _
                              , ByVal sizeB As Long) As JsonPageCode
    'We assume first character must be ASCII
    Dim fbt As FourByteTemplate
    Dim lt As LongTemplate
    Dim i As Long
    '
    For i = 0 To 3
        If i < sizeB Then
            If bytes(i) = 0 Then 'Null
            ElseIf bytes(i) < &H80 Then 'ASCII
                fbt.b(i) = 1
            Else
                fbt.b(i) = 2
            End If
        Else
            fbt.b(i) = &H88
        End If
    Next i
    '
    LSet lt = fbt
    If lt.l = &H1 Then
        DetectCodePage = jpCodeUTF32LE
    ElseIf lt.l = &H1000000 Then
        DetectCodePage = jpCodeUTF32BE
    Else
        lt.l = lt.l And &HFFFF&
        If lt.l = &H1 Then
            DetectCodePage = jpCodeUTF16LE
        ElseIf lt.l = &H100 Then
            DetectCodePage = jpCodeUTF16BE
        ElseIf (lt.l And &HFF) = &H1 Then
            DetectCodePage = jpCodeUTF8
        End If
    End If
End Function

'Converts from 'jpCode' to VBA's internal UTF-16LE
Private Function Decode(ByVal jsonPtr As LongPtr _
                      , ByVal sizeB As Long _
                      , ByVal jpCode As JsonPageCode _
                      , ByRef outBuff As String _
                      , ByRef outBuffSize As Long _
                      , ByRef outBuffPtr As LongPtr _
                      , ByRef outErrDesc As String _
                      , ByVal failIfInvalidByteSequence As Boolean) As Boolean
    #If Mac Then
        outBuff = Space$(sizeB * 2)
        outBuffSize = sizeB * 4
        outBuffPtr = StrPtr(outBuff)
        '
        Dim inBytesLeft As LongPtr:  inBytesLeft = sizeB
        Dim outBytesLeft As LongPtr: outBytesLeft = outBuffSize
        Static collDescriptors As New Collection
        Dim cd As LongPtr
        Dim defaultChar As String: defaultChar = ChrW$(&HFFFD)
        Dim defPtr As LongPtr:     defPtr = StrPtr(defaultChar)
        Dim nonRev As LongPtr
        Dim inPrevLeft As LongPtr
        Dim outPrevLeft As LongPtr
        '
        On Error Resume Next
        cd = collDescriptors(CStr(jpCode))
        On Error GoTo 0
        '
        If cd = NullPtr Then
            Static descTo As String
            Dim descFrom As String: descFrom = PageCodeDesc(jpCode)
            If LenB(descTo) = 0 Then descTo = PageCodeDesc(jpCodeUTF16LE)
            cd = iconv_open(StrPtr(descTo), StrPtr(descFrom))
            If cd = -1 Then
                outErrDesc = "Unsupported page code conversion"
                Exit Function
            End If
            collDescriptors.Add cd, CStr(jpCode)
        End If
        Do
            CopyMemory ByVal errno_location, 0&, longSize
            inPrevLeft = inBytesLeft
            outPrevLeft = outBytesLeft
            nonRev = iconv(cd, jsonPtr, inBytesLeft, outBuffPtr, outBytesLeft)
            If nonRev >= 0 Then Exit Do
            Const EILSEQ As Long = 92
            Const EINVAL As Long = 22
            Dim errNo As Long: CopyMemory errNo, ByVal errno_location, longSize
            '
            If (errNo = EILSEQ Eqv errNo = EINVAL) Or failIfInvalidByteSequence _
            Then
                Select Case errNo
                    Case EILSEQ: outErrDesc = "Invalid byte sequence: "
                    Case EINVAL: outErrDesc = "Incomplete byte sequence: "
                    Case Else:   outErrDesc = "Failed conversion: "
                End Select
                outErrDesc = outErrDesc & " from CP" & jpCode
                Exit Function
            End If
            CopyMemory ByVal outBuffPtr, ByVal defPtr, intSize
            outBytesLeft = outBytesLeft - intSize
            outBuffPtr = outBuffPtr + intSize
            jsonPtr = jsonPtr + byteSize
            inBytesLeft = inBytesLeft - byteSize
        Loop
        If nonRev > 0 Then
            'Maybe there is a way to mimic Windows behaviour - repeat 0xFFFD
            '...
        End If
        outBuffSize = (outBuffSize - CLng(outBytesLeft)) \ 2
        outBuffPtr = StrPtr(outBuff)
        Decode = True
    #Else
        Const MB_ERR_INVALID_CHARS As Long = 8
        Dim charCount As Long
        Dim dwFlags As Long
        '
        If failIfInvalidByteSequence Then
            Select Case jpCode
            Case jpCodeUTF8, jpCodeUTF32LE, jpCodeUTF32BE
                dwFlags = MB_ERR_INVALID_CHARS
            End Select
        End If
        charCount = MultiByteToWideChar(jpCode, dwFlags, jsonPtr, sizeB, 0, 0)
        If charCount = 0 Then
            Const ERROR_INVALID_PARAMETER      As Long = 87
            Const ERROR_NO_UNICODE_TRANSLATION As Long = 1113
            '
            Select Case Err.LastDllError
            Case ERROR_NO_UNICODE_TRANSLATION
                outErrDesc = "Invalid CP" & jpCode & " byte sequence"
            Case ERROR_INVALID_PARAMETER
                outErrDesc = "Code Page: " & jpCode & " not supported"
            Case Else
                outErrDesc = "Unicode conversion failed"
            End Select
            Exit Function
        End If
        outBuff = Space$(charCount)
        outBuffPtr = StrPtr(outBuff)
        outBuffSize = charCount
        '
        MultiByteToWideChar jpCode, dwFlags, jsonPtr _
                          , sizeB, outBuffPtr, charCount
        Decode = (charCount = outBuffSize)
        If Not Decode Then outErrDesc = "Unicode conversion failed"
    #End If
End Function
#If Mac Then
Private Function PageCodeDesc(ByVal jpCode As JsonPageCode) As String
    Dim result As String
    Select Case jpCode
        Case jpCodeUTF8:    PageCodeDesc = "UTF-8"
        Case jpCodeUTF16LE: PageCodeDesc = "UTF-16LE"
        Case jpCodeUTF16BE: PageCodeDesc = "UTF-16BE"
        Case jpCodeUTF32LE: PageCodeDesc = "UTF-32LE"
        Case jpCodeUTF32BE: PageCodeDesc = "UTF-32BE"
    End Select
    PageCodeDesc = StrConv(PageCodeDesc, vbFromUnicode)
End Function
#End If

Private Sub ReverseBytes(ByVal jsonPtr As LongPtr _
                       , ByVal sizeB As Long _
                       , ByRef outBuff As String _
                       , ByRef outBuffSize As Long _
                       , ByRef outBuffPtr As LongPtr)
    Static bytesSrc As ByteAccessor
    Static bytesDest As ByteAccessor
    Dim i As Long
    Dim j As Long
    '
    If bytesSrc.sa.cDims = 0 Then 'Init memory accessors
        InitAccessor VarPtr(bytesSrc), bytesSrc.sa, byteSize
        InitAccessor VarPtr(bytesDest), bytesDest.sa, byteSize
    End If
    '
    outBuffSize = sizeB \ 2
    outBuff = Space$(outBuffSize)
    outBuffPtr = StrPtr(outBuff)
    '
    bytesSrc.sa.pvData = jsonPtr
    bytesSrc.sa.rgsabound0.cElements = sizeB
    bytesDest.sa.pvData = outBuffPtr
    bytesDest.sa.rgsabound0.cElements = outBuffSize * 2 'Ignore odd byte if any
    '
    For i = 0 To bytesDest.sa.rgsabound0.cElements - 1 Step 2
        j = i + 1
        bytesDest.arr(i) = bytesSrc.arr(j)
        bytesDest.arr(j) = bytesSrc.arr(i)
    Next i
    '
    bytesSrc.sa.rgsabound0.cElements = 0
    bytesSrc.sa.pvData = NullPtr
    bytesDest.sa.rgsabound0.cElements = 0
    bytesDest.sa.pvData = NullPtr
End Sub

Private Sub InitAccessor(ByVal accPtr As LongPtr _
                       , ByRef sa As SAFEARRAY_1D _
                       , ByVal elemSize As DataTypeSize)
    InitSafeArray sa, elemSize
    MemLongPtr(accPtr) = VarPtr(sa)
End Sub

Private Sub InitSafeArray(ByRef sa As SAFEARRAY_1D, ByVal elemSize As Long)
    Const FADF_AUTO As Long = &H1
    Const FADF_FIXEDSIZE As Long = &H10
    Const FADF_COMBINED As Long = FADF_AUTO Or FADF_FIXEDSIZE
    With sa
        .cDims = 1
        .fFeatures = FADF_COMBINED
        .cbElements = elemSize
        .cLocks = 1
    End With
End Sub

Private Property Let MemLongPtr(ByVal memAddress As LongPtr, ByVal newValue As LongPtr)
    #If Mac Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, ptrSize
    #ElseIf TWINBASIC Then
        PutMemPtr memAddress, newValue
    #Else
        Static pa As PointerAccessor
        If pa.sa.cDims = 0 Then
            InitSafeArray pa.sa, ptrSize
            CopyMemory pa, VarPtr(pa.sa), ptrSize 'Only API call
        End If
        '
        pa.sa.pvData = memAddress
        pa.sa.rgsabound0.cElements = 1
        pa.arr(0) = newValue
        pa.sa.rgsabound0.cElements = 0
        pa.sa.pvData = NullPtr
    #End If
End Property
Private Property Get MemLongPtr(ByVal memAddress As LongPtr) As LongPtr
    #If Mac Or (VBA7 = 0) Then
        CopyMemory MemLongPtr, ByVal memAddress, ptrSize
    #ElseIf TWINBASIC Then
        GetMemPtr memAddress, MemLongPtr
    #Else
        Static pa As PointerAccessor
        '
        If pa.sa.cDims = 0 Then
            InitSafeArray pa.sa, ptrSize
            MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
        End If
        '
        pa.sa.pvData = memAddress
        pa.sa.rgsabound0.cElements = 1
        MemLongPtr = pa.arr(0)
        pa.sa.rgsabound0.cElements = 0
        pa.sa.pvData = NullPtr
    #End If
End Property

'Non-recursive parser
Private Function ParseChars(ByRef inChars() As Integer _
                          , ByRef inOptions As JSONOptions _
                          , ByRef v As Variant _
                          , ByRef outError As String _
                          , Optional ByVal vMissing As Variant) As Boolean
    Static cm As CharMap
    Static buff As IntegerAccessor
    Static longs As LongAccessor
    Dim i As Long
    Dim j As Long
    Dim afterCommaArr As AllowedToken
    Dim afterCommaDict As AllowedToken
    '
    If buff.sa.cDims = 0 Then
        InitAccessor VarPtr(buff), buff.sa, intSize
        InitAccessor VarPtr(longs), longs.sa, longSize
        InitCharMap cm
    End If
    '
    If inOptions.ignoreTrailingComma Then
        afterCommaArr = allowValue Or allowRBracket
        afterCommaDict = allowString Or allowRBrace
    Else
        afterCommaArr = allowValue
        afterCommaDict = allowString
    End If
    '
    On Error GoTo ErrorHandler
    '
    Dim cInfo As ContextInfo
    Dim depth As Long
    Dim ch As Integer
    Dim wasValue As Boolean
    Dim parents() As ContextInfo: ReDim parents(0 To 0)
    Dim buffSize As Long: buffSize = 16
    Dim sBuff As String:  sBuff = Space$(buffSize)
    Dim UB As Long:       UB = UBound(inChars)
    '
    i = 0
    cInfo.tAllow = allowValue
    buff.sa.pvData = StrPtr(sBuff)
    buff.sa.rgsabound0.cElements = buffSize
    '
    Do While i <= UB
        ch = inChars(i)
        wasValue = False
        If ch < ccTab Or ch > ccObjectEnd Then
            GoTo Unexpected
        ElseIf cm.toType(ch) = whitespace Then 'Skip
        ElseIf ch = ccArrayStart Or ch = ccObjectStart Then
            If (cInfo.tAllow And allowValue) = 0 Then GoTo Unexpected
            depth = depth + 1
            If depth > inOptions.maxDepth Then Err.Raise 5, , "Max Depth Hit"
            If depth > UBound(parents) Then ReDim Preserve parents(0 To depth)
            parents(depth) = cInfo
            '
            cInfo = parents(0) 'Clears members
            cInfo.isDict = (ch = ccObjectStart)
            If cInfo.isDict Then
                Set cInfo.dict = New Dictionary
                If inOptions.allowDuplicatedKeys Then
                    cInfo.dict.AllowDuplicateKeys = True
                End If
                cInfo.dict.CompareMode = inOptions.compMode
                cInfo.tAllow = allowString Or allowRBrace
            Else
                Set cInfo.coll = New Collection
                cInfo.tAllow = allowValue Or allowRBracket
            End If
        ElseIf ch = ccArrayEnd Then
            If (cInfo.tAllow And allowRBracket) = 0 Then GoTo Unexpected
            If Not IsEmpty(v) Then cInfo.coll.Add v
            Set v = cInfo.coll
            cInfo = parents(depth)
            depth = depth - 1
            wasValue = True
        ElseIf ch = ccObjectEnd Then
            If (cInfo.tAllow And allowRBrace) = 0 Then GoTo Unexpected
            If Not IsEmpty(v) Then cInfo.dict.Add cInfo.pendingKey, v
            Set v = cInfo.dict
            cInfo = parents(depth)
            depth = depth - 1
            wasValue = True
        ElseIf ch = ccComma Then
            If (cInfo.tAllow And allowComma) = 0 Then GoTo Unexpected
            If cInfo.isDict Then
                cInfo.dict.Add cInfo.pendingKey, v
                cInfo.tAllow = afterCommaDict
            Else
                cInfo.coll.Add v
                cInfo.tAllow = afterCommaArr
            End If
            v = Empty
        ElseIf ch = ccColon Then
            If (cInfo.tAllow And allowColon) = 0 Then GoTo Unexpected
            cInfo.tAllow = allowValue
        ElseIf ch = ccDoubleQuote Then
            If cInfo.tAllow < allowString Then GoTo Unexpected
            Dim wasHighSurrogate As Boolean: wasHighSurrogate = False
            Dim isLowSurrogate As Boolean
            Dim endFound As Boolean: endFound = False
            '
            j = 0
            For i = i + 1 To UB
                ch = inChars(i)
                If ch = ccDoubleQuote Then
                    endFound = True
                    Exit For
                ElseIf ch = ccBackslash Then
                    i = i + 1
                    Select Case inChars(i)
                        Case ccDoubleQuote, ccSlash, ccBackslash: ch = inChars(i)
                        Case ccLowB: ch = ccBack
                        Case ccLowF: ch = ccFormFeed
                        Case ccLowN: ch = ccLf
                        Case ccLowR: ch = ccCr
                        Case ccLowT: ch = ccTab
                        Case ccLowU 'u followed by 4 hex digits (nibbles)
                            ch = cm.nib1(cm.nibs(inChars(i + 1))) _
                               + cm.nib2(cm.nibs(inChars(i + 2))) _
                               + cm.nib3(cm.nibs(inChars(i + 3))) _
                               + cm.nib4(cm.nibs(inChars(i + 4)))
                            i = i + 4
                        Case Else: Err.Raise 5, , "Invalid escape"
                    End Select
                ElseIf ch >= ccNull And ch < ccSpace Then
                    Err.Raise 5, , "U+0000 through U+001F must be escaped"
                End If
                '
                'https://www.unicode.org/versions/Unicode16.0.0/core-spec/chapter-5/#G11318
                If inOptions.failIfLoneSurrogate Then
                    isLowSurrogate = (ch >= &HDC00 And ch < &HE000)
                    If wasHighSurrogate Xor isLowSurrogate Then
                        Err.Raise 5, , "Lone surrogate not allowed"
                    ElseIf isLowSurrogate Then
                        wasHighSurrogate = False
                    Else
                        wasHighSurrogate = (ch >= &HD800 And ch < &HDC00)
                    End If
                End If
                '
                buff.arr(j) = ch
                j = j + 1
                If j = buffSize Then
                    buff.sa.rgsabound0.cElements = 0
                    buff.sa.pvData = NullPtr
                    sBuff = sBuff & Space$(buffSize)
                    buffSize = buffSize * 2
                    buff.sa.pvData = StrPtr(sBuff)
                    buff.sa.rgsabound0.cElements = buffSize
                End If
            Next i
            If Not endFound Then
                Err.Raise 5, , "Incomplete string"
            ElseIf inOptions.failIfLoneSurrogate And wasHighSurrogate Then
                Err.Raise 5, , "Lone surrogate not allowed"
            End If
            '
            If cInfo.tAllow And allowString Then
                cInfo.pendingKey = Left$(sBuff, j)
                cInfo.pendingKeyPos = i - j
                cInfo.tAllow = allowColon
            Else
                v = Left$(sBuff, j)
                wasValue = True
            End If
        ElseIf (cInfo.tAllow And allowValue) = 0 Then
            GoTo Unexpected
        ElseIf cm.toType(ch) = numDigit Or ch = ccMinus Then
            Dim hasLeadZero As Boolean: hasLeadZero = False
            Dim hasDot As Boolean:      hasDot = False
            Dim hasExp As Boolean:      hasExp = False
            Dim digitsCount As Long
            Dim ct As CharType
            Dim prevCT As CharType
            '
            j = 0
            buff.arr(j) = ch
            hasLeadZero = (ch = ccZero)
            ct = cm.toType(ch)
            digitsCount = -CLng(ct = numDigit)
            For i = i + 1 To UB
                prevCT = ct
                ch = inChars(i)
                ct = cm.toType(ch)
                If ct = numDigit Then
                    If hasLeadZero Then
                        i = i - 1
                        Err.Raise 5, , "Leading zeroes are now allowed"
                    End If
                    hasLeadZero = (digitsCount = 0) And (ch = ccZero)
                    digitsCount = digitsCount + 1
                ElseIf ch = ccDot Then
                    If prevCT <> numDigit Then Err.Raise 5, , "Expected digit not ."
                    If hasDot Or hasExp Then Err.Raise 5, , "Unexpected ."
                    hasDot = True
                    hasLeadZero = False
                ElseIf ct = numExp Then
                    If prevCT <> numDigit Then Err.Raise 5, , "Expected digit not E"
                    If hasExp Then Err.Raise 5, , "Unexpected E-notation"
                    hasExp = True
                    hasLeadZero = False
                ElseIf ct = numSign Then
                    If prevCT <> numExp Then Err.Raise 5, , "Unexpected " & Chr$(ch)
                Else
                    Exit For
                End If
                '
                j = j + 1
                If j = buffSize Then
                    buff.sa.rgsabound0.cElements = 0
                    buff.sa.pvData = NullPtr
                    sBuff = sBuff & Space$(buffSize)
                    buffSize = buffSize * 2
                    buff.sa.pvData = StrPtr(sBuff)
                    buff.sa.rgsabound0.cElements = buffSize
                End If
                buff.arr(j) = ch
            Next i
            If ct > numDigit Then Err.Raise 5, , "Expected digit"
            '
            If buff.arr(j) = ccZero And hasDot And Not hasExp Then
                'Remove trailing zeroes
                digitsCount = digitsCount - j
                Do While j > 0
                    If buff.arr(j - 1) <> ccZero Then Exit Do
                    j = j - 1
                Loop
                If cm.toType(buff.arr(j - 1)) = numDigit Then j = j - 1
                digitsCount = digitsCount + j
            End If
            '
            v = Left$(sBuff, j + 1)
            #If Mac Then
                v = CDbl(v)
            #Else
                Const maxDigits As Long = 15 'Double supports 15 digits
                If digitsCount > maxDigits Then
                    On Error Resume Next
                    v = CDec(v)
                    On Error GoTo ErrorHandler
                    If VarType(v) = vbString Then v = CDbl(v)
                Else
                    v = CDbl(v)
                End If
            #End If
            wasValue = True
            i = i - 1
        ElseIf ch = ccLowF Then
            If inChars(i + 1) <> 97 Or inChars(i + 2) <> 108 _
            Or inChars(i + 3) <> 115 Or inChars(i + 4) <> 101 Then Err.Raise 9
            v = False
            i = i + 4
            wasValue = True
        ElseIf ch = ccLowN Then
            If inChars(i + 1) <> 117 Or inChars(i + 2) <> 108 _
                                     Or inChars(i + 3) <> 108 Then Err.Raise 9
            v = Null
            i = i + 3
            wasValue = True
        ElseIf ch = ccLowT Then
            If inChars(i + 1) <> 114 Or inChars(i + 2) <> 117 _
                                     Or inChars(i + 3) <> 101 Then Err.Raise 9
            v = True
            i = i + 3
            wasValue = True
        Else
            GoTo Unexpected
        End If
        If wasValue Then
            If cInfo.isDict Then
                cInfo.tAllow = allowComma Or allowRBrace
            Else
                cInfo.tAllow = (allowComma Or allowRBracket) * Sgn(depth)
            End If
        End If
        i = i + 1
    Loop
    If depth > 0 Then GoTo Unexpected
    '
    If IsEmpty(v) Then
        outError = "Expected more than just whitespace"
        v = vMissing
    Else
        ParseChars = True
    End If
    '
    buff.sa.rgsabound0.cElements = 0
    buff.sa.pvData = NullPtr
Exit Function
Unexpected:
    If i <= UB Then
        If ch < ccBang Or ch > ccObjectEnd Then
            v = "\u" & Right$("000" & Hex$(ch), 4)
        Else
            v = ChrW$(ch)
        End If
        If cInfo.tAllow = allowNone Then Err.Raise 5, , "Extra " & v
        If cInfo.tAllow And allowValue Then Err.Raise 5, , "Unexpected " & v
    End If
    Err.Raise 5, , "Expected " & AllowedChars(cInfo.tAllow)
ErrorHandler:
    buff.sa.rgsabound0.cElements = 0
    buff.sa.pvData = NullPtr
    If Err.Number = 9 Then
        Select Case ch
            Case ccBackslash: If i > UB Then outError = "Incomplete escape" _
                                        Else outError = "Invalid hex"
            Case ccLowF:  outError = "Expected 'false'"
            Case ccLowN:  outError = "Expected null'"
            Case ccLowT:  outError = "Expected 'true'"
            Case Else: outError = "Invalid literal"
        End Select
    ElseIf Err.Number = 457 Then
        outError = "Duplicated key"
        i = cInfo.pendingKeyPos
    Else
        outError = Err.Description
    End If
    If i > UB Then
        outError = outError & " at end of JSON input"
    Else
        outError = outError & " at char position " & i + 1
    End If
    v = vMissing
End Function

Private Function AllowedChars(ByVal ta As AllowedToken) As String
    If ta And allowString Then AllowedChars = """"
    If ta And allowColon Then AllowedChars = AllowedChars & ":"
    If ta And allowComma Then AllowedChars = AllowedChars & ","
    If ta And allowRBrace Then AllowedChars = AllowedChars & "}"
    If ta And allowRBracket Then AllowedChars = AllowedChars & "]"
    If ta And allowValue Then AllowedChars = AllowedChars & "%"
    If Len(AllowedChars) = 2 Then
        AllowedChars = Left$(AllowedChars, 1) & " or " & Right$(AllowedChars, 1)
    End If
    AllowedChars = Replace(AllowedChars, "%", "Value")
End Function

Private Sub InitCharMap(ByRef cm As CharMap)
    Dim i As Long
    '
    'Map ascii character codes to specific json tokens
    'Avoids the use of Select Case
    cm.toType(ccTab) = whitespace
    cm.toType(ccLf) = whitespace
    cm.toType(ccCr) = whitespace
    cm.toType(ccSpace) = whitespace 'Space
    For i = ccZero To ccNine
        cm.toType(i) = numDigit
    Next i
    cm.toType(ccPlus) = numSign
    cm.toType(ccMinus) = numSign
    cm.toType(ccDot) = numDot
    cm.toType(69) = numExp  'e
    cm.toType(101) = numExp 'E
    '
    'Map nibbles in escaped 4-hex Unicode chracters e.g. \u2713 (check mark)
    'Avoids the use of ChrW by precomputing all hex digits and their position
    For i = ccColon To ccBacktick
        cm.nibs(i) = &H8000 'Force 'Subscript out of range' when used with nib#
    Next i
    For i = 0 To 9
        cm.nibs(i + ccZero) = i
    Next i
    For i = 10 To 15
        cm.nibs(i + 55) = i 'A to F
        cm.nibs(i + 87) = i 'a to f
    Next i
    'All hex digits 0 to F have been mapped in 'nibs'. Now map by position
    'E.g. resCharCode = nib1(nibs(hexDigit1)) + nib2(nibs(hexDigit2) + ...
    'Also avoids using CLng("&H" & ... by directly shifting the nibbles
    For i = 0 To 15
        cm.nib1(i) = (i + 16 * (i > 7)) * &H1000 'Account for sign bit
        cm.nib2(i) = i * &H100
        cm.nib3(i) = i * &H10
        cm.nib4(i) = i 'Only needed to raise error if not 0 to 15 / F
    Next i
End Sub
