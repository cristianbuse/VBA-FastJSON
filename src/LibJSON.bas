Attribute VB_Name = "LibJSON"
'''=============================================================================
''' VBA Fast JSON Parser / Serializer
''' --------------------------------------------
''' https://github.com/cristianbuse/VBA-FastJSON
''' --------------------------------------------
''' MIT License
'''
''' Copyright (c) 2024 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

''==============================================================================
'' Description:
''    * Mac OS compatible
''    * Performant
''    * Non-Recursive Parser
''==============================================================================

Option Explicit
Option Private Module

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If

#Const Windows = (Mac = 0)
#Const x64 = Win64

#If x64 Then
    Private Const PTR_SIZE As Long = 8
    Private Const NULL_PTR As LongLong = 0^
#Else
    Private Const PTR_SIZE As Long = 4
    Private Const NULL_PTR As Long = 0&
#End If
Private Const INT_SIZE As Long = 2

Public Type SAFEARRAYBOUND
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

Private Type IntegerArray
    arr() As Integer
End Type

Private Enum CharType
    whitespace = 1
    numDigit = 2
    numSign = 3
    numExp = 4
    numDot = 5
End Enum

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
End Type

'*******************************************************************************
'Parses a json string into a dictionary or an array
'Expects a UTF16LE string. To convert from UTF8 or other Code Pages use:
'   https://github.com/guwidoe/VBA-StringTools - 'Decode' method
'   OR
'   https://github.com/cristianbuse/VBA-FileTools - 'ConvertText' method
'   before calling this method
'*******************************************************************************
Public Function Parse(ByRef jsonText As Variant) As Variant
    Const methodName As String = "Parse"
    '
    If VarType(jsonText) <> vbString Then
        Err.Raise 5, methodName, "Expected JSON String"
    End If
    '
    Static chars As IntegerArray
    Static sa As SAFEARRAY_1D
    Dim v As Variant
    Dim errMsg As String
    Dim res As Boolean
    '
    If sa.cDims = 0 Then
        InitSafeArray sa, INT_SIZE
        CopyMemory ByVal VarPtr(chars), VarPtr(sa), PTR_SIZE
    End If
    sa.pvData = StrPtr(jsonText)
    sa.rgsabound0.cElements = Len(jsonText)
    '
    res = ParseChars(chars.arr, Parse, errMsg)
    '
    sa.rgsabound0.cElements = 0
    sa.pvData = NULL_PTR
    '
    If Not res Then Err.Raise 5, methodName, errMsg
End Function

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

'Non-recursive parse
Private Function ParseChars(ByRef inChars() As Integer _
                          , ByRef outResult As Variant _
                          , ByRef outErrMsg As String) As Boolean
    Static charMap(9 To 125) As CharType
    Static nibs(48 To 102) As Integer 'Nibble: 0 to F. Byte: 00 to FF
    Static nib1(0 To 15) As Integer
    Static nib2(0 To 15) As Integer
    Static nib3(0 To 15) As Integer
    Static nib4(0 To 15) As Integer
    Static buff As IntegerArray
    Static sa As SAFEARRAY_1D
    Dim i As Long
    Dim j As Long
    Dim v As Variant
    '
    If sa.cDims = 0 Then 'Init static variables
        InitSafeArray sa, INT_SIZE
        CopyMemory ByVal VarPtr(buff), VarPtr(sa), PTR_SIZE
        '
        charMap(9) = whitespace  'Tab
        charMap(10) = whitespace 'Lf
        charMap(13) = whitespace 'Cr
        charMap(32) = whitespace 'Space
        For i = 48 To 57
            charMap(i) = numDigit '0 to 9
        Next i
        charMap(43) = numSign '+
        charMap(45) = numSign '-
        charMap(46) = numDot  '.
        charMap(69) = numExp  'e
        charMap(101) = numExp 'E
        '
        For i = 58 To 96
            nibs(i) = &H8000 'Force 'Subscript out of range' when used with nib#
        Next i
        For i = 0 To 9
            nibs(i + 48) = i '0 to 9
        Next i
        For i = 10 To 15
            nibs(i + 55) = i 'A to F
            nibs(i + 87) = i 'a to f
        Next i
        For i = 0 To 15
            nib1(i) = (i + 16 * (i > 7)) * &H1000
            nib2(i) = i * &H100
            nib3(i) = i * &H10
            nib4(i) = i 'Only needed to raise error if not 0 to 15
        Next i
    End If
    '
    On Error GoTo ErrorHandler
    '
    Dim cInfo As ContextInfo
    Dim depth As Long
    Dim ch As Integer
    Dim UB As Long: UB = UBound(inChars)
    Dim parents() As ContextInfo: ReDim parents(0 To 0)
    Dim buffSize As Long: buffSize = 16
    Dim sBuff As String:  sBuff = Space$(buffSize)
    Dim wasValue As Boolean
    '
    i = 0
With cInfo       'Not indented intentionally
    .tAllow = allowValue
Do While i <= UB 'Not indented intentionally
    ch = inChars(i)
    wasValue = False
    If ch < 9 Or ch > 125 Then
        GoTo Unexpected
    ElseIf charMap(ch) = whitespace Then 'Skip
    ElseIf ch = 91 Or ch = 123 Then '[ or {
        If (.tAllow And allowValue) = 0 Then GoTo Unexpected
        depth = depth + 1
        If depth > UBound(parents) Then ReDim Preserve parents(0 To depth)
        parents(depth) = cInfo
        '
        cInfo = parents(0) 'Clears members. Does not affect With block
        .isDict = (ch = 123)
        If .isDict Then
            Set .dict = New Dictionary
            .tAllow = allowString Or allowRBrace
        Else
            Set .coll = New Collection
            .tAllow = allowValue Or allowRBracket
        End If
    ElseIf ch = 93 Then ']
        If (.tAllow And allowRBracket) = 0 Then GoTo Unexpected
        If Not IsEmpty(v) Then .coll.Add v
        Set v = .coll
        cInfo = parents(depth)
        depth = depth - 1
        wasValue = True
    ElseIf ch = 125 Then '}
        If (.tAllow And allowRBrace) = 0 Then GoTo Unexpected
        If Not IsEmpty(v) Then .dict.Add .pendingKey, v
        Set v = .dict
        cInfo = parents(depth)
        depth = depth - 1
        wasValue = True
    ElseIf ch = 44 Then ',
        If (.tAllow And allowComma) = 0 Then GoTo Unexpected
        If .isDict Then
            .dict.Add .pendingKey, v
            .tAllow = allowString
        Else
            .coll.Add v
            .tAllow = allowValue
        End If
        v = Empty
    ElseIf ch = 58 Then ':
        If .tAllow And allowColon Then .tAllow = allowValue Else GoTo Unexpected
    ElseIf ch = 34 Then '"
        If .tAllow < allowString Then GoTo Unexpected
        Dim endFound As Boolean: endFound = False
        '
        j = 0
        sa.pvData = StrPtr(sBuff)
        sa.rgsabound0.cElements = buffSize
        For i = i + 1 To UB
            ch = inChars(i)
            If ch = 34 Then '"
                endFound = True
                Exit For
            ElseIf ch = 92 Then '\
                i = i + 1
                Select Case inChars(i)
                Case 34, 47, 92: buff.arr(j) = inChars(i) '" / \
                Case 98:         buff.arr(j) = 8          'b >> vbBack
                Case 102:        buff.arr(j) = 12         'f >> vbFormFeed
                Case 110:        buff.arr(j) = 10         'n >> vbLf
                Case 114:        buff.arr(j) = 13         'r >> vbCr
                Case 116:        buff.arr(j) = 9          't >> vbTab
                Case 117 'u followed by 4 hex digits (nibbles)
                    buff.arr(j) = nib1(nibs(inChars(i + 1))) _
                                + nib2(nibs(inChars(i + 2))) _
                                + nib3(nibs(inChars(i + 3))) _
                                + nib4(nibs(inChars(i + 4)))
                    i = i + 4
                Case Else
                    Err.Raise 5, , "Invalid escape"
                End Select
            Else
                buff.arr(j) = ch
            End If
            j = j + 1
            If j = buffSize Then
                sBuff = sBuff & Space$(buffSize)
                buffSize = buffSize * 2
                sa.pvData = StrPtr(sBuff)
                sa.rgsabound0.cElements = buffSize
            End If
        Next i
        If Not endFound Then Err.Raise 5, , "Incomplete string"
        sa.rgsabound0.cElements = 0
        sa.pvData = NULL_PTR
        '
        If .tAllow And allowString Then
            .pendingKey = Left$(sBuff, j)
            .tAllow = allowColon
        Else
            v = Left$(sBuff, j)
            wasValue = True
        End If
    ElseIf (.tAllow And allowValue) = 0 Then
        GoTo Unexpected
    ElseIf charMap(ch) = numDigit Or ch = 45 Then
        Dim hasDot As Boolean: hasDot = False
        Dim hasExp As Boolean: hasExp = False
        Dim digitsCount As Long
        Dim ct As CharType
        Dim prevCT As CharType
        '
        j = 0
        sa.pvData = StrPtr(sBuff)
        sa.rgsabound0.cElements = buffSize
        buff.arr(j) = ch
        ct = charMap(ch)
        digitsCount = -CLng(ct = numDigit)
        For i = i + 1 To UB
            prevCT = ct
            ch = inChars(i)
            ct = charMap(ch)
            If ct = numDigit Then
                digitsCount = digitsCount + 1
            ElseIf ct = numDot Then
                If prevCT <> numDigit Then Err.Raise 5, , "Expected digit not ."
                If hasDot Or hasExp Then Err.Raise 5, , "Unexpected ."
                hasDot = True
            ElseIf ct = numExp Then
                If prevCT <> numDigit Then Err.Raise 5, , "Expected digit not E"
                If hasExp Then Err.Raise 5, , "Unexpected E-notation"
                hasExp = True
            ElseIf ct = numSign Then
                If prevCT <> numExp Then Err.Raise 5, , "Unexpected " & Chr$(ch)
            Else
                Exit For
            End If
            '
            j = j + 1
            If j = buffSize Then
                sBuff = sBuff & Space$(buffSize)
                buffSize = buffSize * 2
                sa.pvData = StrPtr(sBuff)
                sa.rgsabound0.cElements = buffSize
            End If
            buff.arr(j) = ch
        Next i
        sa.rgsabound0.cElements = 0
        sa.pvData = NULL_PTR
        If prevCT <> numDigit Then Err.Raise 5, , "Expected digit"
        '
        v = Left$(sBuff, j + 1)
        #If Mac Then
            v = CDbl(v)
        #Else
            Const maxDigits As Long = 15 'Double supports 15 digits
            If digitsCount > maxDigits Then
                v = CDec(v)
            Else
                v = CDbl(v)
            End If
        #End If
        wasValue = True
        i = i - 1
    ElseIf ch = 102 Then 'f
        If inChars(i + 1) <> 97 Or inChars(i + 2) <> 108 _
        Or inChars(i + 3) <> 115 Or inChars(i + 4) <> 101 Then Err.Raise 9
        v = False
        i = i + 4
        wasValue = True
    ElseIf ch = 110 Then 'n
        If inChars(i + 1) <> 117 Or inChars(i + 2) <> 108 _
                                 Or inChars(i + 3) <> 108 Then Err.Raise 9
        v = Null
        i = i + 3
        wasValue = True
    ElseIf ch = 116 Then 't
        If inChars(i + 1) <> 114 Or inChars(i + 2) <> 117 _
                                 Or inChars(i + 3) <> 101 Then Err.Raise 9
        v = True
        i = i + 3
        wasValue = True
    Else
        GoTo Unexpected
    End If
    If wasValue Then
        If .isDict Then
            .tAllow = allowComma Or allowRBrace
        Else
            .tAllow = (allowComma Or allowRBracket) * Sgn(depth)
        End If
    End If
    i = i + 1
Loop     'Not indented intentionally
End With 'Not indented intentionally
    If depth > 0 Then GoTo Unexpected
    If IsObject(v) Then Set outResult = v Else outResult = v
    ParseChars = True
    '
Exit Function
Unexpected:
    If i <= UB Then
        If ch < 33 Or ch > 125 Then v = Format(ch, "\\u0000") Else v = ChrW$(ch)
        If cInfo.tAllow = allowNone Then Err.Raise 5, , "Extra " & v
        If cInfo.tAllow And allowValue Then Err.Raise 5, , "Unexpected " & v
    End If
    Err.Raise 5, , "Expected " & AllowedChars(cInfo.tAllow)
ErrorHandler:
    sa.rgsabound0.cElements = 0
    sa.pvData = NULL_PTR
    If Err.Number = 9 Then
        Select Case ch
            Case 92: If i > UB Then outErrMsg = "Incomplete escape" _
                               Else outErrMsg = "Invalid hex"
            Case 102:  outErrMsg = "Expected 'false'"
            Case 110:  outErrMsg = "Expected null'"
            Case 116:  outErrMsg = "Expected 'true'"
            Case Else: outErrMsg = "Invalid literal"
        End Select
    ElseIf Err.Number = 457 Then
        outErrMsg = "Duplicated key"
    Else
        outErrMsg = Err.Description
    End If
    If i > UB Then
        outErrMsg = outErrMsg & " at end of JSON input"
    Else
        outErrMsg = outErrMsg & " at index " & i + 1
    End If
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
