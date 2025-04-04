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
        Private Declare PtrSafe Function memmove Lib "/usr/lib/libc.dylib" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr, ByRef inBuf As LongPtr, ByRef inBytesLeft As LongPtr, ByRef outBuf As LongPtr, ByRef outBytesLeft As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As LongPtr, ByVal fromCode As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr) As Long
        Private Declare PtrSafe Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As LongPtr
    #Else
        Private Declare Function memmove Lib "/usr/lib/libc.dylib" (Destination As Any, Source As Any, ByVal Length As Long) As Long
        Private Declare Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long, ByRef inBuf As Long, ByRef inBytesLeft As Long, ByRef outBuf As Long, ByRef outBytesLeft As Long) As Long
        Private Declare Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As Long, ByVal fromCode As Long) As Long
        Private Declare Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long) As Long
        Private Declare Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByVal pvargSrc As LongPtr)
        Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
        Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
    #Else
        Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByVal pvargSrc As Long)
        Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
        Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    #End If
#End If

#Const Windows = (Mac = 0)
#Const x64 = Win64
#If (x64 = 0) Or Mac Then
    Const vbLongLong = 20
#End If

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
    currSize = 8
#If x64 Then
    variantSize = 24
#Else
    variantSize = 16
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
Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type
Private Type CurrencyAccessor
    arr() As Currency
    sa As SAFEARRAY_1D
End Type
Private Type VariantAccessor
    arr() As Variant
    vt() As Integer
    sa As SAFEARRAY_1D
End Type
Private Type SABoundAccessor
    rgsabound() As SAFEARRAYBOUND
    sa As SAFEARRAY_1D
End Type

Private Type FakeVariant
    vt As Integer
    wReserved(0 To 2) As Integer
    ptrs(0 To 1) As LongPtr
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
    ccCapitalE = 69     '0x45 E
    ccArrayStart = 91   '0x5B [
    ccBackslash = 92    '0x5C \
    ccArrayEnd = 93     '0x5D ]
    ccBacktick = 96     '0x60 `
    ccLowB = 98         '0x62 b
    ccLowE = 101        '0x65 e
    ccLowF = 102        '0x66 f
    ccLowN = 110        '0x6E n
    ccLowR = 114        '0x72 r
    ccLowT = 116        '0x74 t
    ccLowU = 117        '0x75 u
    ccObjectStart = 123 '0x7B {
    ccObjectEnd = 125   '0x7D }
    ccDel = 127         '0x7F DEL - last ASCII
End Enum

Private Enum CharType
    whitespace = 1
    numDigit = 2
    numSign = 3
    numExp = 4
    numDot = 5
End Enum

Private Type CharacterMap
    toType(ccTab To ccObjectEnd) As CharType
    nibs(ccZero To ccLowF) As Integer 'Nibble: 0 to F. Byte: 00 to FF
    nib1(0 To 15) As Integer
    nib2(0 To 15) As Integer
    nib3(0 To 15) As Integer
    nib4(0 To 15) As Integer
    literal(0 To 2) As Currency
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

Private Type ParseContextInfo
    coll As Collection
    dict As Dictionary
    tAllow As AllowedToken
    isDict As Boolean
    pendingKey As String
    pendingKeyPos As Long
End Type

Private Type SerializeContextInfo
    arrKeys() As Variant
    arrItems() As Variant
    isDict As Boolean
    incIndex As Long
    currIndex As Long
    ub As Long
    iUnkPtr As LongPtr
    firstItemPtr As LongPtr
    firstKeyPtr As LongPtr
    epClose As EncodedPosition
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

Private Type TextBuffer
    Text As String
    Index As Long
    Limit As Long
End Type

Private Enum EncodedPosition
    epTrue = -1
    [_epMin] = epTrue
    epFalse = 0
    epNull
    epText
    epArray
    epDict
    epLBrace
    epRBrace
    epLBracket
    epRBracket
    epComma
    [_epMax] = epComma
End Enum

Private Type InfOrNaNInfo
    Offset As LongPtr
    Mask As Integer
End Type

Private Type EncodedString
    s As String
    sLen As Long
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
' - Returns UTF16LE texts only
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
'     Only applicable if conversion is needed (e.g. UTF8 to UTF16LE)
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
    Static isFDict As Boolean
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
        isFDict = IsFastDict()
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
    jOptions.allowDuplicatedKeys = allowDuplicatedKeys And isFDict
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
            memmove ByVal errno_location, 0&, longSize
            inPrevLeft = inBytesLeft
            outPrevLeft = outBytesLeft
            nonRev = iconv(cd, jsonPtr, inBytesLeft, outBuffPtr, outBytesLeft)
            If nonRev >= 0 Then Exit Do
            Const EILSEQ As Long = 92
            Const EINVAL As Long = 22
            Dim errNo As Long: memmove errNo, ByVal errno_location, longSize
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
            memmove ByVal outBuffPtr, ByVal defPtr, intSize
            outBytesLeft = outBytesLeft - intSize
            outBuffPtr = outBuffPtr + intSize
            jsonPtr = jsonPtr + byteSize
            inBytesLeft = inBytesLeft - byteSize
        Loop
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
                       , Optional ByRef outBuffSize As Long _
                       , Optional ByRef outBuffPtr As LongPtr)
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
    If elemSize = variantSize Then 'Init auxiliary VarType
        MemLongPtr(accPtr + ptrSize) = VarPtr(sa)
    End If
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

Private Property Let MemLongPtr(ByVal memAddress As LongPtr _
                              , ByVal newValue As LongPtr)
    #If Mac Then
        memmove ByVal memAddress, newValue, ptrSize
    #ElseIf TWINBASIC Then
        PutMemPtr memAddress, newValue
    #Else
        Static pa As PointerAccessor
        If pa.sa.cDims = 0 Then
            InitSafeArray pa.sa, ptrSize
            MemLongPtrRef(VarPtr(pa)) = VarPtr(pa.sa)
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
    #If Mac Then
        memmove MemLongPtr, ByVal memAddress, ptrSize
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

#If Windows Then
Private Property Let MemLongPtrRef(ByVal memAddress As LongPtr _
                                 , ByVal newValue As LongPtr)
    Const VT_BYREF As Long = &H4000
    Dim memValue As Variant
    Dim remoteVT As Variant
    Dim fv As FakeVariant
    '
    fv.ptrs(0) = VarPtr(memValue)
    fv.vt = vbInteger + VT_BYREF
    VariantCopy remoteVT, VarPtr(fv) 'Init VarType ByRef
    '
#If Win64 Then 'Cannot assign LongLong ByRef
    Dim c As Currency
    RemoteAssign memValue, VarPtr(newValue), remoteVT, vbCurrency + VT_BYREF, c, memValue
    RemoteAssign memValue, memAddress, remoteVT, vbCurrency + VT_BYREF, memValue, c
#Else 'Can assign Long ByRef
    RemoteAssign memValue, memAddress, remoteVT, vbLong + VT_BYREF, memValue, newValue
#End If
End Property
Private Sub RemoteAssign(ByRef memValue As Variant _
                       , ByRef memAddress As LongPtr _
                       , ByRef remoteVT As Variant _
                       , ByVal newVT As VbVarType _
                       , ByRef targetVariable As Variant _
                       , ByRef newValue As Variant)
    memValue = memAddress
    remoteVT = newVT
    targetVariable = newValue
    remoteVT = vbEmpty
End Sub
#End If

'Non-recursive parser
Private Function ParseChars(ByRef inChars() As Integer _
                          , ByRef inOptions As JSONOptions _
                          , ByRef v As Variant _
                          , ByRef outError As String _
                          , Optional ByVal vMissing As Variant) As Boolean
    Static cm As CharacterMap
    Static buff As IntegerAccessor
    Static curr As CurrencyAccessor
    Dim i As Long
    Dim j As Long
    Dim afterCommaArr As AllowedToken
    Dim afterCommaDict As AllowedToken
    '
    If buff.sa.cDims = 0 Then
        InitAccessor VarPtr(buff), buff.sa, intSize
        InitAccessor VarPtr(curr), curr.sa, currSize
        curr.sa.pvData = VarPtr(curr)
        curr.sa.rgsabound0.cElements = 1
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
    Dim cInfo As ParseContextInfo
    Dim depth As Long
    Dim ch As Integer
    Dim wasValue As Boolean
    Dim parents() As ParseContextInfo
    Dim ubParents As Long
    Dim buffSize As Long: buffSize = 16
    Dim sBuff As String:  sBuff = Space$(buffSize)
    Dim ub As Long:       ub = UBound(inChars)
    Dim chDot As Integer: chDot = AscW(Mid$(CStr(1.1), 2, 1)) 'Locale Dot
    '
    i = 0
    cInfo.tAllow = allowValue
    buff.sa.pvData = StrPtr(sBuff)
    buff.sa.rgsabound0.cElements = buffSize
    '
    Do While i <= ub
        ch = inChars(i)
        wasValue = False
        If ch < ccTab Or ch > ccObjectEnd Then
            GoTo Unexpected
        ElseIf cm.toType(ch) = whitespace Then 'Skip
        ElseIf ch = ccArrayStart Or ch = ccObjectStart Then
            If (cInfo.tAllow And allowValue) = 0 Then GoTo Unexpected
            depth = depth + 1
            If depth > inOptions.maxDepth Then Err.Raise 5, , "Max Depth Hit"
            If depth > ubParents Then
                ReDim Preserve parents(0 To depth)
                ubParents = depth
            End If
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
            For i = i + 1 To ub
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
            For i = i + 1 To ub
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
                    If prevCT <> numDigit Or hasDot _
                                          Or hasExp Then GoTo Unexpected
                    hasDot = True
                    hasLeadZero = False
                    ch = chDot 'We avoid Val() because we also want CDec()
                ElseIf ct = numExp Then
                    If prevCT <> numDigit Or hasExp Then GoTo Unexpected
                    hasExp = True
                    hasLeadZero = False
                ElseIf ct = numSign Then
                    If prevCT <> numExp Then GoTo Unexpected
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
        Else 'Check for literal: false, null, true
            If ch = ccLowF Then
                i = i + 4
                v = False
            ElseIf ch = ccLowN Then
                i = i + 3
                v = Null
            ElseIf ch = ccLowT Then
                i = i + 3
                v = True
            Else
                GoTo Unexpected
            End If
            If i > ub Then Err.Raise 9
            curr.sa.pvData = VarPtr(inChars(i - 3))
            If cm.literal((ch And &H18) \ &H8) <> curr.arr(0) Then Err.Raise 9
            wasValue = True
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
    curr.sa.pvData = VarPtr(curr)
Exit Function
Unexpected:
    If i <= ub Then
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
            Case ccBackslash: If i > ub Then outError = "Incomplete escape" _
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
    If i > ub Then
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

Private Sub InitCharMap(ByRef cm As CharacterMap)
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
    cm.toType(ccCapitalE) = numExp
    cm.toType(ccLowE) = numExp
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
    '
    'Map false, null, true to 8-byte values
    cm.literal(0) = 2842946657609.3281@ 'alse
    cm.literal(1) = 3039976134888.6638@ 'null
    cm.literal(2) = 2842947516642.1108@ 'true
End Sub

Private Function IsFastDict() As Boolean
    Dim o As Object:     Set o = New Dictionary
    On Error Resume Next
    Dim s As Single:     s = o.LoadFactor
    Dim b As Boolean:    b = o.AllowDuplicateKeys
    Dim d As Dictionary: Set d = o.Self.Factory
    IsFastDict = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Serializes the input data to a json string
' - Does not throw
' - By default, this method cannot fail.
'   It can only fail in the following scenarios:
'     1) A non-text key is found and 'failIfNonTextKeys' is set to 'True'
'     2) A circular reference is found and 'failIfCircularRef' is set to 'True'
'     3) Conversion failed and 'jpCode' was not UTF16LE
'   Read more on these parameters below.
'   On failure, it returns a null String and an 'outError' ByRef (String)
' - By default, returns UTF16LE json string - see 'jpCode' below
' - Invalid data types (direct or nested) are replaced with Null:
'     - Empty
'     - User Defined Type
'     - Nothing or an interface not implementing 'ToDictionary'
'     - Uninitialized Array
'     - Special Single/Double values: +Inf, -Inf, SNaN, QNaN
'     - Circular references (by default - see below)
'Parameters:
' * jsonData - can be any of the following:
'     - primitive (String, Number, Boolean, Null)
'     - Array or Collection
'     - Dictionary
'     - Any class that implements a 'ToDictionary() As Dictionary' method
' * indentSpaces:
'     - Default is 0 - no indentation
'     - For positive values, indenting will be added up to 'maxIndent' (16)
' * escapeNonASCII (character codes less than 0 and higher than 127):
'     - False:          codes outside 0-127 are not escaped
'     - True (Default): codes outsde 0-126 are escaped using the \u0000 notation
'                       Note DEL (127) is escaped even though is ASCII
' * sortKeys:
'     - False (Default): does nothing
'     - True:            sorts dictionary keys, ascending.
'                        Useful for debugging or comparing JSON data
' * forceKeysToText:
'     - False (Default): non-text keys are skipped
'     - True:            converts vbError, vbDate and number data types to text
' * failIfNonTextKeys:
'     - False (Default): does nothing - non-text keys are skipped,
'                        unless 'failIfNonTextKeys' is 'True' and converted
'     - True:            fails if a non-text key is found.
'                        Note that if 'forceKeysToText' is 'True' then only
'                        fails if there are keys that cannot be converted
' * failIfCircularRef:
'     - False (Default): inserts Null if circular reference is found
'     - True:            fails if circular reference is found
'     An object can reference itself, either directly or indirectly
' * formatDateISO
'     - False (Default): yyyy-mm-dd hh:nn:ss
'     - True:            yyyy-mm-ddThh:nn:ss.sssZ (see FormatISOExt method)
' * jpCode:
'     - Default is UTF16LE. Supports UTF8, UTF16BE and UTF32 (Mac only)
' * outError:
'     Returns error message (ByRef) on failure
'*******************************************************************************
Public Function Serialize(ByRef jsonData As Variant _
                        , Optional ByVal indentSpaces As Long = 0 _
                        , Optional ByVal escapeNonASCII As Boolean = True _
                        , Optional ByVal sortKeys As Boolean = False _
                        , Optional ByVal forceKeysToText As Boolean = False _
                        , Optional ByVal failIfNonTextKeys As Boolean = False _
                        , Optional ByVal failIfCircularRef As Boolean = False _
                        , Optional ByVal formatDateISO As Boolean = False _
                        , Optional ByVal jpCode As JsonPageCode = jpCodeUTF16LE _
                        , Optional ByRef outError As String) As String
    Const dateF As String = "yyyy-mm-dd hh:nn:ss"
    Const maxBit As Long = &H40000000
    Const maxBuf As Long = &H7FFFFFFF
    Const initLevels As Long = 16
    Const maxIndent As Long = 16
    Const pOffset As Long = 8
    Static encoded([_epMin] To [_epMax]) As EncodedString
    Static escaped(ccNull To ccSpace - 1) As EncodedString
    Static map(0 To 255) As String
    Static ints As IntegerAccessor
    Static chars As IntegerAccessor
    Static ptrs As PointerAccessor
    Static vars As VariantAccessor
    Static bounds As SABoundAccessor
    Static infOrNaN(vbSingle To vbDouble) As InfOrNaNInfo
    Static numChar(ccSpace To ccBacktick) As Byte
    Dim fv As FakeVariant
    Dim colonLen As Long
    Dim commaLen As Long
    Dim nLineLen As Long
    Dim ep As EncodedPosition
    Dim beautify As Boolean
    Dim depth As Long
    Dim ub As Long
    Dim levels() As SerializeContextInfo
    Dim currentIndent As Long
    Dim tempIndent As Long
    Dim spaces As String
    Dim spCount As Long
    Dim buff As TextBuffer
    Dim i As Long
    Dim j As Long
    '
    'Map for escaped control characters?
    '
    If indentSpaces > 0 Then
        beautify = True
        If indentSpaces > maxIndent Then indentSpaces = maxIndent
        colonLen = 2
        nLineLen = Len(vbNewLine)
        spCount = indentSpaces * initLevels 'Sufficient for most cases
        spaces = Space$(spCount)
    Else
        colonLen = 1
        indentSpaces = 0 'In case negative
    End If
    '
    buff.Index = 1
    ub = initLevels - 1
    ReDim levels(0 To ub)
    levels(0).incIndex = 1
    '
    If vars.sa.cDims = 0 Then
        'Init memory accessors
        InitAccessor VarPtr(ints), ints.sa, intSize
        InitAccessor VarPtr(chars), chars.sa, intSize
        InitAccessor VarPtr(ptrs), ptrs.sa, ptrSize
        InitAccessor VarPtr(vars), vars.sa, variantSize
        InitAccessor VarPtr(bounds), bounds.sa, longSize * 2
        'Init maps
        InitEncoded encoded
        InitEscaped escaped, map
        InitNumChar numChar
        InitInfOrNaN infOrNaN
    End If
    '
    commaLen = 1 + nLineLen
    encoded(epLBrace).sLen = commaLen
    encoded(epLBracket).sLen = commaLen
    '
    'Point accesors to input data
    vars.sa.pvData = VarPtr(jsonData)
    ptrs.sa.pvData = vars.sa.pvData + pOffset
    ints.sa.pvData = vars.sa.pvData
    '
    vars.sa.rgsabound0.cElements = 1
    ptrs.sa.rgsabound0.cElements = 1
    ints.sa.rgsabound0.cElements = 1
    '
    If vars.vt(0) And VT_BYREF Then
        ptrs.sa.pvData = ptrs.arr(0)
        fv.ptrs(0) = ptrs.arr(0)
        fv.vt = vars.vt(0) Xor VT_BYREF
        vars.sa.pvData = VarPtr(fv)
    End If
    '
    Do
        Dim vt As VbVarType: vt = vars.vt(0)
        If vt = vbObject Or vt = vbDataObject Then
            If vars.arr(0) Is Nothing Then GoTo InsertNull
            '
            Dim iUnk As IUnknown: Set iUnk = vars.arr(0)
            Dim iPtr As LongPtr:  iPtr = ObjPtr(iUnk)
            '
            'Check for circular references
            'For deep nesting might be worth implementing a Dictionary
            For i = depth To 0 Step -1
                If levels(i).iUnkPtr = iPtr Then
                    If failIfCircularRef Then
                        outError = "Circular reference detected"
                        GoTo Clean
                    End If
                    GoTo InsertNull
                End If
            Next i
            '
            Dim coll As Collection
            Dim dict As Dictionary
            Dim isDict As Boolean
            '
            isDict = (TypeOf vars.arr(0) Is Dictionary)
            If isDict Then
                Set dict = vars.arr(0)
            ElseIf TypeOf vars.arr(0) Is Collection Then
                Set coll = vars.arr(0)
            Else 'Try 'ToDictionary' method via late-binding / IDispatch::Invoke
                On Error Resume Next
                Set dict = Nothing
                Set dict = vars.arr(0).ToDictionary()
                On Error GoTo 0
                If dict Is Nothing Then GoTo InsertNull
                isDict = True
            End If
            '
            depth = depth + 1
            If depth > ub Then
                ub = ub * 2 - 1
                ReDim Preserve levels(0 To ub)
            End If
            '
            With levels(depth)
                .isDict = isDict
                If isDict Then
                    .ub = dict.Count - 1
                    If .ub = -1 Then
                        ep = epDict
                        .epClose = epFalse
                    Else
                        .arrKeys = dict.Keys()
                        '
                        Dim vtKey As VbVarType
                        Dim lastFound As Long: lastFound = -1
                        For i = 0 To .ub
                            If IsObject(.arrKeys(i)) Then
                                vtKey = vbDataObject
                            Else
                                vtKey = VarType(.arrKeys(i))
                            End If
                            If vtKey = vbString Then
                                lastFound = i
                            ElseIf vtKey = vbNull Or vtKey = vbDataObject _
                                                  Or vtKey > vbLongLong Then
                                If failIfNonTextKeys Then
                                    outError = "Invalid key data type"
                                    GoTo Clean
                                End If
                                .arrKeys(i) = Empty
                            ElseIf forceKeysToText Then
                                If vtKey = vbDate Then
                                    If formatDateISO Then
                                        .arrKeys(i) = FormatISOExt(.arrKeys(i))
                                    Else
                                        .arrKeys(i) = Format$(.arrKeys(i), dateF)
                                    End If
                                Else
                                    .arrKeys(i) = CStr(.arrKeys(i))
                                End If
                                lastFound = i
                            ElseIf failIfNonTextKeys Then
                                outError = TypeName(.arrKeys(i)) & " key found"
                                GoTo Clean
                            Else
                                .arrKeys(i) = Empty
                            End If
                        Next i
                        If lastFound < 0 Then
                            .currIndex = .ub
                            ep = epDict
                            .epClose = epFalse
                        Else
                            .ub = lastFound
                            .currIndex = -1
                            ep = epLBrace
                            .epClose = epRBrace
                            .iUnkPtr = iPtr
                            .arrItems = dict.Items()
                            If sortKeys Then
                                QuickSortKeys .arrKeys, .arrItems, 0, .ub
                            End If
                            .firstItemPtr = VarPtr(.arrItems(0))
                            .firstKeyPtr = VarPtr(.arrKeys(0))
                            vars.sa.pvData = .firstKeyPtr
                        End If
                    End If
                Else
                    .ub = coll.Count - 1
                    If .ub < 0 Then
                        ep = epArray
                        .epClose = epFalse
                    Else
                        ep = epLBracket
                        .epClose = epRBracket
                        .iUnkPtr = iPtr
                        '
                        'Iterate collection with dynamic iterator
                        ReDim .arrItems(0 To .ub + 1) 'Extra member for safety
                        .firstKeyPtr = VarPtr(.arrItems(0))
                        '
                        vars.sa.pvData = .firstKeyPtr
                        '
                        For Each vars.arr(0) In coll 'This actually works
                            vars.sa.pvData = vars.sa.pvData + variantSize
                        Next vars.arr(0)
                        '
                        vars.sa.pvData = .firstKeyPtr
                        .currIndex = -1
                    End If
                End If
                .incIndex = 1
            End With
        ElseIf vt >= vbArray Then
            ptrs.sa.pvData = vars.sa.pvData + pOffset
            If ptrs.arr(0) = NullPtr Then GoTo InsertNull 'Uninitialized
            ptrs.sa.pvData = ptrs.arr(0) 'SAFEARRAY address
            '
            depth = depth + 1
            If depth > ub Then
                ub = ub * 2 - 1
                ReDim Preserve levels(0 To ub)
            End If
            '
            bounds.sa.pvData = ptrs.sa.pvData + rgsabound0_cElementsOffset
            bounds.sa.rgsabound0.cElements = CLng(ptrs.arr(0) And &HFF&)
            '
            Dim totalElem As Long: totalElem = 1
            '
            For i = 0 To bounds.sa.rgsabound0.cElements - 1
                totalElem = totalElem * bounds.rgsabound(i).cElements
                If totalElem = 0 Then Exit For
            Next i
            '
            With levels(depth)
                .incIndex = 1
                If totalElem = 0 Then
                    .ub = -1
                    ep = epArray
                    .epClose = epFalse
                Else
                    ep = epLBracket
                    .epClose = epRBracket
                    .currIndex = -1
                    If bounds.sa.rgsabound0.cElements = 1 Then
                        .ub = totalElem - 1
                        If vt = vbArray + vbVariant Then 'Just access data
                            ptrs.sa.pvData = ptrs.sa.pvData + pvDataOffset
                            .firstKeyPtr = ptrs.arr(0)
                        Else 'Copy
                            ReDim .arrItems(0 To .ub)
                            .firstKeyPtr = VarPtr(.arrItems(0))
                            j = bounds.rgsabound(0).lLbound
                            bounds.rgsabound(0).lLbound = 0
                            If vt = vbArray + vbObject Then
                                For i = 0 To .ub
                                    Set .arrItems(i) = vars.arr(0)(i)
                                Next i
                            Else
                                For i = 0 To .ub
                                    .arrItems(i) = vars.arr(0)(i)
                                Next i
                            End If
                            bounds.rgsabound(0).lLbound = j
                        End If
                        vars.sa.pvData = .firstKeyPtr
                    Else 'Multi-Dimensional
                        .ub = bounds.sa.rgsabound0.cElements - 1
                        .ub = bounds.rgsabound(.ub).cElements - 1
                        ReDim .arrItems(0 To .ub)
                        NDArrayToCollections .arrItems, vars.arr(0), bounds _
                                                      , totalElem, ptrs.arr(0)
                        .firstKeyPtr = VarPtr(.arrItems(0))
                        vars.sa.pvData = .firstKeyPtr
                    End If
                End If
            End With
            bounds.sa.rgsabound0.cElements = 0
            bounds.sa.pvData = NullPtr
        ElseIf vt <= vbNull Or vt >= vbUserDefinedType Then 'Empty, Null or UDT
'Or by jump: Nothing, Unsupported Interface, Uninitialized Array, +-Inf, [SQ]NaN
InsertNull: ep = epNull
        ElseIf vt = vbBoolean Then
            ep = CLng(vars.arr(0))
        Else
            If vt = vbString Then
                ints.sa.pvData = StrPtr(vars.arr(0))
                ints.sa.rgsabound0.cElements = Len(vars.arr(0))
                '
                If ints.sa.rgsabound0.cElements = 0 Then
                    encoded(epText).s = """"""
                    encoded(epText).sLen = 2
                Else
                    'Just in case each character needs to be escaped
                    i = buff.Index + ints.sa.rgsabound0.cElements * 6 _
                                   + currentIndent + 1
                    If i > buff.Limit Then
                        j = buff.Limit
                        If i And maxBit Then buff.Limit = maxBuf _
                                        Else buff.Limit = i * 2
                        buff.Text = buff.Text & Space$(buff.Limit - j)
                    End If
                    If currentIndent > 0 Then
                        Mid$(buff.Text, buff.Index, currentIndent) = spaces
                        buff.Index = buff.Index + currentIndent
                        tempIndent = currentIndent
                        currentIndent = 0
                    End If
                    '
                    chars.sa.pvData = StrPtr(buff.Text) + buff.Index * 2
                    chars.sa.rgsabound0.cElements = i - buff.Index
                    Mid$(buff.Text, buff.Index) = """"
                    buff.Index = buff.Index + 1
                    '
                    Dim ch As Integer
                    i = 0
                    j = 0
                    For i = 0 To ints.sa.rgsabound0.cElements - 1
                        ch = ints.arr(i)
                        If ch >= ccSpace And ch < ccDel Then
                            If ch = ccDoubleQuote Or ch = ccBackslash Then
                                chars.arr(j) = ccBackslash
                                j = j + 1
                            End If
                            chars.arr(j) = ch
                            j = j + 1
                        ElseIf (ch And &HFFE0) = 0 Then 'Escape U+0000 to U+001F
                            Mid$(buff.Text, buff.Index + j) = escaped(ch).s
                            j = j + escaped(ch).sLen
                        ElseIf escapeNonASCII Then
                            'Note we included non-printable ASCII DEL (127)
                            Dim k As Long: k = buff.Index + j
                            Mid$(buff.Text, k) = "\u00"
                            Mid$(buff.Text, k + 2) = map((ch + &H10000) \ &H100& And &HFF&)
                            Mid$(buff.Text, k + 4) = map(ch And &HFF&)
                            j = j + 6
                        Else
                            chars.arr(j) = ch
                            j = j + 1
                        End If
                    Next i
                    '
                    chars.sa.rgsabound0.cElements = 0
                    chars.sa.pvData = NullPtr
                    buff.Index = buff.Index + j
                    '
                    encoded(epText).s = """"
                    encoded(epText).sLen = 1
                End If
                ints.sa.rgsabound0.cElements = 1
            ElseIf vt = vbError Then
                encoded(epText).s = """" & CStr(vars.arr(0)) & """"
                encoded(epText).sLen = Len(encoded(epText).s)
            ElseIf vt = vbDate Then 'Quotes already included in formatting
                If formatDateISO Then
                    encoded(epText).s = FormatISOExt(vars.arr(0))
                Else
                    encoded(epText).s = Format$(vars.arr(0), dateF)
                End If
                encoded(epText).sLen = Len(encoded(epText).s)
            ElseIf (vt And &H14) <> &H4 Then 'Byte, Integer, Long, LongLong
                encoded(epText).s = CStr(vars.arr(0))
                encoded(epText).sLen = Len(encoded(epText).s)
            Else
                If vt < vbCurrency Then 'Float - check for +-Inf, [SQ]NaN
                    With infOrNaN(vt)
                        ints.sa.pvData = vars.sa.pvData + .Offset
                        If (ints.arr(0) And .Mask) = .Mask Then GoTo InsertNull
                    End With
                End If
                encoded(epText).s = CStr(vars.arr(0))
                encoded(epText).sLen = Len(encoded(epText).s)
                '
                'Fix dot if it's something else e.g. comma, quote, space etc.
                ints.sa.pvData = StrPtr(encoded(epText).s)
                ints.sa.rgsabound0.cElements = encoded(epText).sLen
                For i = 0 To ints.sa.rgsabound0.cElements - 1
                    If numChar(ints.arr(i)) = 0 Then
                        ints.arr(i) = ccDot
                        Exit For
                    End If
                Next i
                ints.sa.rgsabound0.cElements = 1
            End If
            ep = epText
        End If
        '
        Do
            i = buff.Index + encoded(ep).sLen + currentIndent
            If i + commaLen > buff.Limit Then 'Increase buffer capacity
                j = buff.Limit
                If i And maxBit Then buff.Limit = maxBuf Else buff.Limit = i * 2
                buff.Text = buff.Text & Space$(buff.Limit - j)
            End If
            If beautify Then
                If currentIndent = 0 Then
                    currentIndent = tempIndent
                    tempIndent = 0
                Else
                    Mid$(buff.Text, buff.Index, currentIndent) = spaces
                    buff.Index = buff.Index + currentIndent
                End If
            End If
            Mid$(buff.Text, buff.Index) = encoded(ep).s
            buff.Index = i
            '
            Dim isNotLastItem As Boolean
            Do
                With levels(depth)
                    isNotLastItem = (.currIndex < .ub) Or (.incIndex = 0)
                End With
                If isNotLastItem Then Exit Do
                '
                ep = levels(depth).epClose
                depth = depth - 1
                If depth < 0 Then Exit Do
            Loop While ep = epFalse
            If isNotLastItem Or depth < 0 Then Exit Do
            '
            If beautify Then
                Mid$(buff.Text, buff.Index) = vbNewLine
                buff.Index = buff.Index + nLineLen
                currentIndent = currentIndent - indentSpaces
            End If
        Loop
        If depth < 0 Then Exit Do
        '
        With levels(depth)
            .currIndex = .currIndex + .incIndex
            If .currIndex = 0 Then
                If beautify And (.incIndex > 0) Then
                    currentIndent = currentIndent + indentSpaces

                    If currentIndent > spCount Then
                        spCount = currentIndent * 2
                        spaces = Space$(spCount)
                    End If
                End If
            ElseIf .incIndex > 0 Then
                Mid$(buff.Text, buff.Index, commaLen) = encoded(epComma).s
                buff.Index = buff.Index + commaLen
                vars.sa.pvData = .firstKeyPtr + variantSize * .currIndex
            End If
            If .isDict Then
                If .incIndex = 0 Then
                    Const colonSpace As String = ": "
                    Mid$(buff.Text, buff.Index, colonLen) = colonSpace
                    buff.Index = buff.Index + colonLen
                    vars.sa.pvData = .firstItemPtr + variantSize * .currIndex
                    tempIndent = currentIndent
                    currentIndent = 0
                    .incIndex = 1
                Else
                    Do While vars.vt(0) = vbEmpty 'Skip key
                        .currIndex = .currIndex + 1
                        vars.sa.pvData = vars.sa.pvData + variantSize
                    Loop
                    .incIndex = 0
                End If
            End If
        End With
    Loop
    If jpCode = jpCodeUTF16LE Then
        Serialize = Left$(buff.Text, buff.Index - 1)
    ElseIf jpCode = jpCodeUTF16BE Then
        ReverseBytes StrPtr(buff.Text), (buff.Index - 1) * 2, Serialize
    Else
        outError = "Conversion not yet supported"
    End If
Clean:
    ints.sa.rgsabound0.cElements = 0: ints.sa.pvData = NullPtr
    ptrs.sa.rgsabound0.cElements = 0: ptrs.sa.pvData = NullPtr
    vars.sa.rgsabound0.cElements = 0: vars.sa.pvData = NullPtr
End Function

Private Sub InitEncoded(ByRef encoded() As EncodedString)
    encoded(epTrue).s = "true"
    encoded(epFalse).s = "false"
    encoded(epNull).s = "null"
    encoded(epArray).s = "[]"
    encoded(epDict).s = "{}"
    encoded(epLBrace).s = "{" & vbNewLine
    encoded(epRBrace).s = "}"
    encoded(epLBracket).s = "[" & vbNewLine
    encoded(epRBracket).s = "]"
    encoded(epComma).s = "," & vbNewLine
    Dim i As Long
    For i = [_epMin] To [_epMax]
        encoded(i).sLen = Len(encoded(i).s)
    Next i
End Sub
Private Sub InitEscaped(ByRef escapedControls() As EncodedString _
                      , ByRef hexMax() As String)
    Dim i As Long
    '
    For i = &H0 To &H9
        hexMax(i) = Chr$(i + 48)
    Next i
    For i = &HA To &HF
        hexMax(i) = Chr$(i + 55)
    Next i
    For i = &H10 To &HFF
        hexMax(i) = hexMax(i \ &H10) & hexMax(i And &HF)
    Next i
    For i = &H0 To &HF
         hexMax(i) = "0" & hexMax(i)
    Next i
    '
    escapedControls(ccBack).s = "\b"
    escapedControls(ccTab).s = "\t"
    escapedControls(ccLf).s = "\n"
    escapedControls(ccFormFeed).s = "\f"
    escapedControls(ccCr).s = "\r"
    '
    For i = ccBack To ccCr
        escapedControls(i).sLen = 2
    Next i
    escapedControls(11).sLen = 0
    '
    For i = ccNull To ccSpace - 1
        If escapedControls(i).sLen = 0 Then
            escapedControls(i).sLen = 6
            escapedControls(i).s = "\u00" & hexMax(i)
        End If
    Next i
End Sub

Private Sub InitNumChar(ByRef numChar() As Byte)
    Dim i As Long
    For i = ccZero To ccNine
        numChar(i) = 1
    Next i
    numChar(ccPlus) = 1
    numChar(ccMinus) = 1
    numChar(ccCapitalE) = 1
End Sub

Private Sub InitInfOrNaN(ByRef infOrNaN() As InfOrNaNInfo)
    infOrNaN(vbSingle).Mask = &H7F80
    infOrNaN(vbDouble).Mask = &H7FF0
    infOrNaN(vbSingle).Offset = 10
    infOrNaN(vbDouble).Offset = 14
End Sub

'Adds values, row-wise, from a multi-dimensional array into nested collections
'Based on method with same name at:
'https://github.com/cristianbuse/VBA-ArrayTools/blob/master/src/LibArrayTools.bas
Private Sub NDArrayToCollections(ByRef arr() As Variant _
                               , ByRef src As Variant _
                               , ByRef bounds As SABoundAccessor _
                               , ByVal totalElem As Long _
                               , ByRef cDimsPtr As LongPtr)
#If x64 Then
    Const dimMask As LongLong = &HFFFFFFFFFFFFFF00^
#Else
    Const dimMask As Long = &HFFFFFF00^
#End If
    Dim rgsabound0 As SAFEARRAYBOUND
    Dim rowMajorIndexes() As Long
    Dim depth() As Long
    Dim i As Long
    Dim ub As Long: ub = bounds.sa.rgsabound0.cElements - 1
    '
    ReDim depth(0 To ub)
    depth(0) = 1
    For i = 1 To ub
        depth(i) = depth(i - 1) * bounds.rgsabound(i - 1).cElements
    Next i
    '
    ReDim rowMajorIndexes(0 To totalElem - 1)
    AddRowMajorIndexes rowMajorIndexes, bounds.rgsabound, depth, 0, ub, 0, 0
    '
    'Make source one-dimensional
    rgsabound0 = bounds.rgsabound(0)
    cDimsPtr = (cDimsPtr And dimMask) Or 1
    bounds.rgsabound(0).cElements = totalElem
    bounds.rgsabound(0).lLbound = 0
    '
    'Populate
    Dim j As Long 'Acts as an index and needs to be incremented ByRef
    For i = 0 To bounds.rgsabound(ub).cElements - 1
        Set arr(i) = AddCollections(rowMajorIndexes, src, bounds.rgsabound _
                                  , ub - 1, rgsabound0.cElements, j)
    Next i
    '
    'Restore source to multi-dimensional
    bounds.rgsabound(0) = rgsabound0
    cDimsPtr = (cDimsPtr And dimMask) Or bounds.sa.rgsabound0.cElements
End Sub
Private Sub AddRowMajorIndexes(ByRef arr() As Long _
                             , ByRef rgsabound() As SAFEARRAYBOUND _
                             , ByRef depth() As Long _
                             , ByVal currDim As Long _
                             , ByVal lastDim As Long _
                             , ByVal colWiseIndex As Long _
                             , ByRef rowWiseIndex As Long)
    Dim i As Long
    If currDim = lastDim Then
        For i = 0 To rgsabound(currDim).cElements - 1
            arr(colWiseIndex + i * depth(currDim)) = rowWiseIndex
            rowWiseIndex = rowWiseIndex + 1
        Next i
    Else
        Dim nextDim As Long: nextDim = currDim + 1
        For i = 0 To rgsabound(currDim).cElements - 1
            AddRowMajorIndexes arr, rgsabound, depth, nextDim, lastDim _
                             , colWiseIndex + i * depth(currDim), rowWiseIndex
        Next i
    End If
End Sub
Private Function AddCollections(ByRef arrIndex() As Long _
                              , ByRef arrData As Variant _
                              , ByRef rgsabound() As SAFEARRAYBOUND _
                              , ByVal currDim As Long _
                              , ByVal firstSize As Long _
                              , ByRef elemIndex As Long) As Collection
    Dim collResult As New Collection
    Dim i As Long
    If currDim = 0 Then
        For i = 1 To firstSize
            collResult.Add arrData(arrIndex(elemIndex))
            elemIndex = elemIndex + 1
        Next i
    Else
        Dim prevDim As Long: prevDim = currDim - 1
        For i = 1 To rgsabound(currDim).cElements
            collResult.Add AddCollections(arrIndex, arrData, rgsabound _
                                        , prevDim, firstSize, elemIndex)
        Next i
    End If
    Set AddCollections = collResult
End Function

'Uses the yyyy-mm-ddThh:nn:ss.sZ extended format
Private Function FormatISOExt(ByRef vDate As Variant) As String
    Const secondsPerDay As Double = 86400
    Const usPerSecond As Double = 1000000
    Dim days As Double: days = vDate
    Dim frac As Double: frac = (days - Int(days)) * secondsPerDay
    '
    frac = CLng((frac - Int(frac)) * usPerSecond) / usPerSecond
    If frac = 0 Or frac = 1 Then
        FormatISOExt = Format$(days, "yyyy-mm-ddThh:nn:ssZ")
    Else
        days = days - frac / secondsPerDay 'Avoid rounding
        FormatISOExt = Format$(days, "yyyy-mm-ddThh:nn:ss") _
                     & Format$(frac, ".0##Z")
    End If
End Function

'This is not a 'Stable' Sort i.e. equal values do not preserve their order
'https://en.wikipedia.org/wiki/Quicksort
Private Sub QuickSortKeys(ByRef arrKeys() As Variant _
                        , ByRef arrItems() As Variant _
                        , ByVal lb As Long _
                        , ByVal ub As Long)
    If lb >= ub Then Exit Sub
    '
    Dim newLB As Long: newLB = lb
    Dim newUB As Long: newUB = ub
    Dim p As Long:     p = (lb + ub) \ 2
    '
    Do While newLB <= newUB
        Do While newLB < ub
            If arrKeys(newLB) >= arrKeys(p) Then Exit Do
            newLB = newLB + 1
        Loop
        Do While newUB > lb
            If arrKeys(p) >= arrKeys(newUB) Then Exit Do
            newUB = newUB - 1
        Loop
        If newLB <= newUB Then
            If newLB <> newUB Then 'Swap
                'Can be more efficient if using memory accessors
                Dim setL As Boolean: setL = IsObject(arrItems(newLB))
                Dim setU As Boolean: setU = IsObject(arrItems(newUB))
                Dim v As Variant:    v = arrKeys(newLB) 'Keys are Text or Empty
                '
                arrKeys(newLB) = arrKeys(newUB)
                arrKeys(newUB) = v
                If setL Then Set v = arrItems(newLB) Else v = arrItems(newLB)
                If setU Then Set arrItems(newLB) = arrItems(newUB) _
                             Else arrItems(newLB) = arrItems(newUB)
                If setL Then Set arrItems(newUB) = v Else arrItems(newUB) = v
            End If
            newLB = newLB + 1
            newUB = newUB - 1
        End If
    Loop
    QuickSortKeys arrKeys, arrItems, lb, newUB
    QuickSortKeys arrKeys, arrItems, newLB, ub
End Sub
