Attribute VB_Name = "VB6_Funcs"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''©Rd'
'  VB6 intrinsic functions not available in VB5
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Const LOCALE_SDECIMAL = &HE&           ' Character(s) used as the decimal separator.
Private Const LOCALE_STHOUSAND = &HF&          ' Character(s) used to separate groups of digits to the left of the decimal.
Private Const LOCALE_SGROUPING = &H10&         ' Sizes for each group of digits to the left of the decimal.
Private Const LOCALE_IDIGITS = &H11&           ' Number of fractional digits.
Private Const LOCALE_ILZERO = &H12&            ' Leading zero for decimal.

Private Const LOCALE_SCURRENCY = &H14&         ' String used as the local monetary symbol.
Private Const LOCALE_SMONDECIMALSEP = &H16&    ' Character(s) used as the monetary decimal separator.
Private Const LOCALE_SMONTHOUSANDSEP = &H17&   ' Character(s) used as the monetary separator between groups of digits to the left of the decimal.
Private Const LOCALE_SMONGROUPING = &H18&      ' Sizes for each group of monetary digits to the left of the decimal.
Private Const LOCALE_ICURRDIGITS = &H19&       ' Number of fractional digits for the local monetary format.
Private Const LOCALE_INEGCURR = &H1C&          ' Negative currency mode.

Private Const LOCALE_SNEGATIVESIGN = &H51&     ' String value for the negative sign.

Private Const LOCALE_USER_DEFAULT = &H400&     ' Default user locale.
Private Const LOCALE_SYSTEM_DEFAULT = &H800&   ' Default system locale.
Private Const LOCALE_IFIRSTDAYOFWEEK = &H100C& ' First day of week specifier.
Private Const LOCALE_INEGNUMBER = &H1010&      ' Negative number mode.

Private Type tLocaleInfo
    sNegChr As String
    sGroups As String
    sDecSep As String
    sZero As String
    iGroups As Long
    iNegNum As Long
    iNumDec As Long
End Type
Private mLI As tLocaleInfo

Private Type tLocaleCurrencyInfo
    sSymbol As String
    sNegChr As String
    sGroups As String
    sDecSep As String
    sZero As String
    iGroups As Long
    iNegNum As Long
    iNumDec As Long
End Type
Private mLCI As tLocaleCurrencyInfo

Public Enum eDayOfWeek
    UseSystem = 0  ' Use National Language Support (NLS) API setting.
    Sunday = 1     ' Sunday (default)
    Monday = 2     ' Monday
    Tuesday = 3    ' Tuesday
    Wednesday = 4  ' Wednesday
    Thursday = 5   ' Thursday
    Friday = 6     ' Friday
    Saturday = 7   ' Saturday
    #If False Then
      Dim UseSystem, Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday
    #End If
End Enum

Public Enum eTristate
    UseDefault = -2  ' Use the setting from the computer's regional settings.
    etTrue = -1      ' True
    etFalse = 0      ' False
    #If False Then
      Dim UseDefault, etTrue, etFalse
    #End If
End Enum

Public Enum eCompareMethod
    UseOptionCompare = -1  ' Performs a comparison using the setting of the Option Compare statement.
    BinaryCompare = 0      ' Performs a binary comparison.
    TextCompare = 1        ' Performs a textual comparison.
    #If False Then
      Dim UseOptionCompare, BinaryCompare, TextCompare
    #End If
End Enum

Public Enum eNamedFormat
    GeneralDate = 0  ' Display a date and/or time. If there is a date part, display it as a short date. If there is a time part, display it as a long time. If present, both parts are displayed.
    LongDate = 1     ' Display a date using the long date format specified in your computer's regional settings.
    ShortDate = 2    ' Display a date using the short date format specified in your computer's regional settings.
    LongTime = 3     ' Display a time using the time format specified in your computer's regional settings.
    ShortTime = 4    ' Display a time using the 24-hour format (hh:mm).
    #If False Then
      Dim GeneralDate, LongDate, ShortDate, LongTime, ShortTime
    #End If
End Enum

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' InStrRev
'
' Returns the position of an occurrence of one string within another,
' from the end of string. Similar to InStr but searches from the end.
'
' VBA.Strings.InStrRev(StringCheck As String, _
'                      StringMatch As String, _
'                     [Start As Long = -1], _
'                     [Compare As VbCompareMethod = vbBinaryCompare]) As Long
'
' StringCheck:  The string to be searched.
' StringMatch:  The substring to search for.
' Start:        Specifies the search starting position. If omitted searches
'               entire string, else search back from specified position.
' Compare:      Comparison to use when evaluating substrings.
'
' Returns:      The last start position of StringMatch within StringCheck.
'
' If StringCheck is "" or vbNullString:        Returns 0 (zero)
' If StringMatch is "" or vbNullString:        Returns Start or last character position of StringCheck if Start omitted
' If StringMatch is not found:                 Returns 0 (zero)
' If StringMatch is found within StringCheck:  Returns position at which match is found
' If Start > Len(StringCheck):                 Returns 0 (zero)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function InStrRev(sString As String, sMatch As String, _
                Optional ByVal lRightStart As Long = -1, _
                Optional ByVal eCompare As eCompareMethod = BinaryCompare, _
                Optional ByVal lLeftLimit As Long = 1, _
                Optional ByVal fStrictCompat As Boolean) As Long
    Dim lPos As Long
    If lRightStart = -1 Then lRightStart = Len(sString)

    If fStrictCompat And lRightStart < 1 Then Err.Raise 5
    If fStrictCompat Then lRightStart = lRightStart - Len(sMatch) + 1

    lPos = InStr(lLeftLimit, sString, sMatch, eCompare)

    Do Until lPos = 0 Or lPos > lRightStart
        InStrRev = lPos
        lPos = InStr(InStrRev + 1, sString, sMatch, eCompare)
    Loop
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filter
'
' Returns a zero-based array containing a subset of the specified
' string array InputStrings based on the specified filter criteria.
'
' VBA.Strings.Filter(InputStrings, _
'                    Match As String, _
'                   [Include As Boolean = True], _
'                   [Compare As VbCompareMethod = vbBinaryCompare])
'
' InputStrings:  One-dimensional array of strings to be searched.
' Match:         String to search for.
' Include:       Specifies whether to return substrings that include or exclude Match.
'                If True, Filter returns the subset of the array that contains Match as a substring.
'                If False, Filter returns the subset of the array that does not contain Match as a substring.
' Compare:       Comparison to use when evaluating substrings.
'
' If no matches of Match are found within InputStrings, Filter returns an empty array.
' An error occurs if InputStrings is Null or is not a one-dimensional array.
'
' The array returned by the Filter function contains only enough elements to contain
' the number of matched items.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Filter(InputStrings() As String, sMatch As String, _
                Optional ByVal fInclude As Boolean = True, _
                Optional ByVal eCompare As eCompareMethod = BinaryCompare) As Variant ' BMc

    Dim asRet() As String, c As Long, i As Long, s As String
    Const cChunk As Long = 20
    ReDim asRet(0 To cChunk) As String

    On Error GoTo FilterResize

    For i = LBound(InputStrings) To UBound(InputStrings)
        s = InputStrings(i)
        If InStr(1, s, sMatch, eCompare) Then
            If fInclude Then
                asRet(c) = s
                c = c + 1
            End If
        Else
            If Not fInclude Then
                asRet(c) = s
                c = c + 1
            End If
        End If
    Next
    If c = 0 Then
        Erase asRet
    Else
        ReDim Preserve asRet(0 To c - 1)
    End If
    Filter = asRet
    Exit Function

FilterResize:

    If Err.Number = 9 Then
        ReDim Preserve asRet(0 To c + cChunk) As String
        Resume              ' Try again
    End If
    Err.Raise Err.Number    ' Other VB error for client
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Join
'
' Returns a string created by joining the substrings contained in an array.
'
' VBA.Strings.Join(InputStrings() As String, [Delimiter As String]) As String
'
' InputStrings:  One-dimensional array containing substrings to be joined.
'
' Delimiter:     Character(s) used to separate the substrings in the returned string.
'                If omitted, the space character (" ") is used as the separater.
'                If delimiter is "", all items in the list are concatenated with no delimiters.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Join(InputStrings() As String, Optional Delimiter As String = " ") As String
    Dim iLen As Long
    Dim iPos As Long
    Dim idx As Long
    Dim iLb As Long
    Dim iUb As Long

    iLb = LBound(InputStrings)
    iUb = UBound(InputStrings)

    iLen = (iUb - iLb) * Len(Delimiter)
    For idx = iLb To iUb
       iLen = iLen + Len(InputStrings(idx))
    Next

    Join = Space$(iLen)
    iPos = 1

    If LenB(Delimiter) Then
       For idx = iLb To iUb - 1
           MidB$(Join, iPos) = InputStrings(idx)
           iPos = iPos + LenB(InputStrings(idx))
           MidB$(Join, iPos) = Delimiter
           iPos = iPos + LenB(Delimiter)
       Next
       MidB$(Join, iPos) = InputStrings(idx)
    Else
       For idx = iLb To iUb
           MidB$(Join, iPos) = InputStrings(idx)
           iPos = iPos + LenB(InputStrings(idx))
       Next
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Split - Fast substitue Split function
'
' Returns a one-dimensional array of substrings.
' Optionally specify/limit the number of substrings.
' Optionally specify the lower boundary of the array, or zero-based if omitted.
'
' VBA.Strings.Split(Expression As String,[Delimiter],[Count],[Compare],[LBound])
'
' Expression:  Required. String expression containing substrings and delimiters.
'              If expression is a zero-length string(""), Split returns an empty
'              array, that is, an array with no elements and no data!
'
' Delimiter:   Optional. String character(s) used to identify substring limits.
'              If omitted, the space character (" ") is assumed to be the delimiter.
'              If delimiter is a zero-length string, a single-element array containing
'              the entire expression string is returned. If specifed, the delimiter
'              character is removed from the substrings during the split process.
'
' Count:       Optional. Number of substrings to be returned.
'              If omitted or <= zero indicates that all substrings are returned.
'
' Compare:     Optional. Long indicating the kind of comparison to use when evaluating the delimiter.
'              The compare argument can have the following values:
'              vbUseCompareOption –1 Performs a comparison using the setting of the Option Compare statement.
'              vbBinaryCompare     0 Performs a binary comparison.
'              vbTextCompare       1 Performs a textual comparison.
'
' LBound:      Optional. Specifies the lower boundary of the array, or zero-based if omitted.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Function Split(InputString As String, Optional Delimiter As String = " ", Optional ByVal iCntLimit As Long, Optional ByVal eCompare As eCompareMethod = BinaryCompare, Optional ByVal iLb As Long) As String() ' Variant for VB5

    Dim aSplit() As String
    Dim sLCtext As String
    Dim sLCdelim As String
    Dim aHits() As Long
    Dim iHitPos As Long
    Dim iPos As Long
    Dim iHits As Long
    Dim iHit As Long
    Dim iLen As Long
    Dim iDelim As Long
    Dim iOffset As Long

    On Error GoTo FreakOut

    iLen = LenB(InputString)
    If iLen Then Else GoTo ExitFunc ' No text

    iDelim = LenB(Delimiter)
    If iDelim Then Else GoTo ExitOneItem ' Nothing to find

    ReDim aHits(0 To (iLen \ iDelim)) As Long ' Allow max possible

    If (eCompare = TextCompare) Then
        ' Better to convert once to lowercase than on every call to InStr
        sLCdelim = LCase$(Delimiter): sLCtext = LCase$(InputString)
        iPos = InStrB(1, sLCtext, sLCdelim, vbBinaryCompare)
    Else                      ' Do first search
        iPos = InStrB(1, InputString, Delimiter, vbBinaryCompare)
    End If

    Do While (iPos)
        If iPos And 1& Then
            aHits(iHits) = iPos
            iHits = iHits + 1
            If iHits = iCntLimit - 1 Then Exit Do
            iOffset = iDelim  ' Offset next start pos
        Else
            iOffset = 1       ' Byte offset start pos
        End If

        If (eCompare = TextCompare) Then
            iPos = InStrB(iPos + iOffset, sLCtext, sLCdelim)
        Else
            iPos = InStrB(iPos + iOffset, InputString, Delimiter)
        End If
    Loop

    If iHits Then Else GoTo ExitOneItem ' No hits

    aHits(iHits) = iLen + 1
    ReDim aSplit(iLb To iLb + iHits) As String

    iPos = 1
    For iHit = 0 To iHits
        iHitPos = aHits(iHit)
        aSplit(iHit + iLb) = MidB$(InputString, iPos, iHitPos - iPos)
        iPos = iHitPos + iDelim
    Next

ExitFunc:
    Split = aSplit ' Compiler will optimize
    Exit Function

ExitOneItem:
    ReDim aSplit(iLb To iLb) As String ' One item
    aSplit(iLb) = InputString
    Split = aSplit ' Compiler will optimize
    Exit Function

FreakOut:
    MsgBox "Error - " & Err.Number & ": " & Err.Description

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' SplitVb5 - Even faster Split sub-routine
'
' Assigns byref to a one-dimensional array of substrings.
' Optionally specify/limit the number of substrings.
' Optionally specify the lower boundary of the array, or zero-based if omitted.
' Further details as per Split above.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub SplitVb5(InputString As String, asSplit() As String, Optional Delimiter As String = " ", Optional ByVal iCntLimit As Long, Optional ByVal eCompare As eCompareMethod = BinaryCompare, Optional ByVal iLb As Long)

    Dim sLCtext As String
    Dim sLCdelim As String
    Dim aHits() As Long
    Dim iHitPos As Long
    Dim iPos As Long
    Dim iHits As Long
    Dim iHit As Long
    Dim iLen As Long
    Dim iDelim As Long
    Dim iOffset As Long

    On Error GoTo FreakOut

    iLen = LenB(InputString)
    If iLen Then Else Erase asSplit: Exit Sub ' No text

    iDelim = LenB(Delimiter)
    If iDelim Then Else GoTo ExitOneItem ' Nothing to find

    ReDim aHits(0 To (iLen \ iDelim)) As Long ' Allow max possible

    If (eCompare = TextCompare) Then
        ' Better to convert once to lowercase than on every call to InStr
        sLCdelim = LCase$(Delimiter): sLCtext = LCase$(InputString)
        iPos = InStrB(1, sLCtext, sLCdelim, vbBinaryCompare)
    Else                      ' Do first search
        iPos = InStrB(1, InputString, Delimiter, vbBinaryCompare)
    End If

    Do While (iPos)
        If iPos And 1& Then
            aHits(iHits) = iPos
            iHits = iHits + 1
            If iHits = iCntLimit - 1 Then Exit Do
            iOffset = iDelim  ' Offset next start pos
        Else
            iOffset = 1       ' Byte offset start pos
        End If

        If (eCompare = TextCompare) Then
            iPos = InStrB(iPos + iOffset, sLCtext, sLCdelim)
        Else
            iPos = InStrB(iPos + iOffset, InputString, Delimiter)
        End If
    Loop

    If iHits Then Else GoTo ExitOneItem ' No hits

    aHits(iHits) = iLen + 1
    ReDim asSplit(iLb To iLb + iHits) As String

    iPos = 1
    For iHit = 0 To iHits
        iHitPos = aHits(iHit)
        asSplit(iHit + iLb) = MidB$(InputString, iPos, iHitPos - iPos)
        iPos = iHitPos + iDelim
    Next

    Exit Sub

ExitOneItem:
    ReDim asSplit(iLb To iLb) As String ' One item
    asSplit(iLb) = InputString
    Exit Sub

FreakOut:
    MsgBox "Error - " & Err.Number & ": " & Err.Description

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Replace - Fast substitute Replace function in pure VB.
'
' Returns a string in which a specified substring has been replaced
' with another substring the specified number of times.
'
' VBA.Strings.Replace(Expression, Find, With, [Start], [Count], [Compare])
'
' Expression:  String expression containing substrings to replace.
' Find:        Substring being searched for.
' With:        Replacement substring.
' Start:       Position within expression where substring search is to begin. If omitted, 1 is assumed.
' Count:       Number of substring substitutions to perform. If omitted, makes all possible substitutions.
' Compare:     Comparison to use when evaluating substrings.
'
' Replace returns the following values:
'
' If Expression is zero-length ("") returns a zero-length string ("")
' If Find is zero-length ("") returns a copy of Expression.
' If With is zero-length ("") returns a copy of Expression with all occurences of Find removed.
' If Start > Len(Expression) returns a copy of Expression.
' If Start < 1 the search begins at 1.
' If Count <= 0 all possible substitutions are performed.
'
' Removed functionality:
'
' The return value of VB's Replace function is a string, with substitutions made,
' that begins at the position specified by Start and concludes at the end of the
' expression string. It is not a copy of the original string from the beginning
' of the string, with substitutions made, if Start is > 1.
'
' By contrast, this implementation returns a full copy of the expression string
' from the beginning of the string to the end of the string, with substitutions
' made from the Start position until the Count number of substitutions are made
' or until the end of the string.
'
' Extra functionality:
'
' This Replace implementation returns information about replacements made
' through the return value of the Count parameter (zero if no replacements),
' and the start position in the returned string of the last replacement
' through the return value of the Start parameter (zero if none found).
'
' Using this feature you can limit the number of replacements (as you would
' Using VB 's Replace through the Count parameter), but could make subsequent
' calls to Replace by passing Start + 1 (or Start + Len(With)) to step through
' the replacement process as needed, and stop when Count returns with zero.
'
' Or simply just to display how many replacements were made.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Replace(sString As String, sTerm As String, sNewTerm As String, _
                        Optional lStart As Long = 1, Optional lHitCnt As Long, _
                        Optional ByVal eCompare As eCompareMethod = BinaryCompare) As String

    Dim lLenOld As Long, lLenNew As Long, lCnt As Long
    Dim lHit As Long, lLenOrig As Long, lOffset As Long
    Dim lSize As Long, lOffStart As Long, lHitPos As Long
    Dim sTermL As String, sStrL As String, lProg As Long
    Dim alHits() As Long, lPos As Long, fSkip As Boolean

    On Error GoTo FreakOut

    If lStart < 1 Then lPos = 1 Else lPos = lStart + lStart - 1 ' Validate start pos
    lStart = 0

    lLenOrig = LenB(sString)
    If (lLenOrig = 0) Then Exit Function ' No text

    lLenOld = LenB(sTerm)
    If (lLenOld = 0) Then GoTo ShortCirc ' Nothing to find
    lLenNew = LenB(sNewTerm)

    lOffset = lLenNew - lLenOld
    lSize = 500 ' lSize = Arr chunk size
    ReDim alHits(0 To lSize) As Long

    If (eCompare = TextCompare) Then
        ' Better to convert once to lowercase than on every call to InStr
        sTermL = LCase$(sTerm): sStrL = LCase$(sString)
        lHit = InStrB(lPos, sStrL, sTermL, vbBinaryCompare)
    Else                        ' Do first search
        lHit = InStrB(lPos, sString, sTerm, vbBinaryCompare)
    End If

    Do While (lHit)             ' Do until no more hits
        If (lHit And 1&) Then
            alHits(lCnt) = lHit ' Record hits
            lCnt = lCnt + 1

            If (lCnt = lHitCnt) Then Exit Do
            If (lCnt = lSize) Then
                lSize = lSize + 5000
                ReDim Preserve alHits(0 To lSize) As Long
            End If

            lOffStart = lLenOld ' Offset next start pos
        Else
           lOffStart = 1        ' Byte offset start pos
        End If
        If (eCompare = TextCompare) Then
            lHit = InStrB(lHit + lOffStart, sStrL, sTermL)
        Else
            lHit = InStrB(lHit + lOffStart, sString, sTerm)
        End If
    Loop

    lHitCnt = lCnt
    If (lCnt = 0) Then GoTo ShortCirc   ' No hits

    lSize = lLenOrig + (lOffset * lCnt) ' lSize = result chr count
    If (lSize = 0) Then Exit Function   ' Result is an empty string
    Replace = Space$(lSize * 0.5)       ' Pre-allocate memory

    lOffStart = 1: lPos = 1
    If (lLenNew) Then
       For lHit = 0 To lCnt - 1
           lHitPos = alHits(lHit)
           lProg = lHitPos - lPos
           If (lProg) Then          ' Build new string
               MidB$(Replace, lOffStart) = MidB$(sString, lPos, lProg)
               lOffStart = lOffStart + lProg
           End If                   ' ©Rd
           MidB$(Replace, lOffStart) = sNewTerm
           lOffStart = lOffStart + lLenNew
           lPos = lHitPos + lLenOld ' No offset orig str
       Next
    Else
       For lHit = 0 To lCnt - 1
           lHitPos = alHits(lHit)
           lProg = lHitPos - lPos   ' Build new string
           If (lProg) Then
               MidB$(Replace, lOffStart) = MidB$(sString, lPos, lProg)
               lOffStart = lOffStart + lProg
           End If
           lPos = lHitPos + lLenOld ' No offset orig str
       Next
    End If

    If lOffStart <= lSize Then MidB$(Replace, lOffStart) = MidB$(sString, lPos)
    lStart = (lOffStart + 1 - lLenNew) * 0.5 ' Last hit pos in returned string
FreakOut:
    Exit Function

ShortCirc: ' If nothing to do
    Replace = sString

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' StrReverse
'
' VBA.Strings.StrReverse(Expression As String) As String
'
' Reverses the specified string.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function StrReverse(InputString As String) As String
    Dim iLen As Long, iIdx As Long
    iLen = Len(InputString)
    If iLen Then
        StrReverse = Space$(iLen)
        Do Until iIdx = iLen
            Mid$(StrReverse, iIdx + 1) = Mid$(InputString, iLen - iIdx)
            iIdx = iIdx + 1
        Loop
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Round
'
' Returns a number rounded to a specified number of decimal places.
'
' VBA.Math.Round(Expression, [NumDecimalPlaces]) As Double
'
' Expression:        Numeric expression being rounded.
' NumDecimalPlaces:  Number indicating how many places to the right
'                    of the decimal are included in the rounding.
'                    If omitted, integers (whole numbers) are returned.
'                    Negatives round to tens (-1), hundreds (-2), etc.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Round(ByVal Expression As Double, Optional ByVal NumDecimalPlaces As Long) As Double
    Round = Int(Expression * (10# ^ NumDecimalPlaces) + 0.5) / (10# ^ NumDecimalPlaces)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' FormatNumber
'
' VBA.Strings.FormatNumber(Expression, [NumDecimalPlaces], [IncludeLeadingZero], _
'                                      [UseParensForNegNumbers], [GroupDigits])
'
' Returns an expression formatted as a number.
'
' Rounds the result to the specified decimal places.
'
' Expression:              Expression to be formatted.
'
' NumDecimalPlaces:        Indicates how many places to the right of the decimal are displayed.
'                          Default (–1) indicates that the computer's regional settings are used.
'
' IncludeLeadingZero:      Tristate constant that indicates whether or not a leading zero is
'                          displayed for fractional values.
'
' UseParensForNegNumbers:  Tristate constant that indicates whether or not to place negative
'                          values within parentheses.
'
' GroupDigits:             Tristate constant that indicates whether or not numbers are grouped
'                          using the group delimiter specified in the computer's regional settings.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FormatNumber(ByVal Expression As Double, Optional ByVal NumDecimalPlaces As Long = -1, Optional ByVal IncludeLeadingZero As eTristate = UseDefault, Optional ByVal UseParensForNegNumbers As eTristate = UseDefault, Optional ByVal GroupDigits As eTristate = UseDefault) As String
    Dim sFormat As String
    sFormat = GetLocaleFormat(NumDecimalPlaces, IncludeLeadingZero, GroupDigits)
    Select Case NumDecimalPlaces
        Case Is = 0
            FormatNumber = Format$(Round(Expression, NumDecimalPlaces), sFormat)
        Case Else
            FormatNumber = Format$(Round(Expression, NumDecimalPlaces), sFormat & mLI.sDecSep & String$(NumDecimalPlaces, "0"))
    End Select
    If (Expression < 0) And UseParensForNegNumbers Then
        FormatNumber = FormatNegativeNumbers(FormatNumber, UseParensForNegNumbers)
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' FormatPercent Function
'
' VBA.Strings.FormatPercent(Expression, [NumDecimalPlaces], [IncludeLeadingZero], _
'                                       [UseParensForNegNumbers], [GroupDigits])
'
' Returns an expression formatted as a percentage, (multipled by 100)
' with a trailing % character.
'
' Rounds the result to the specified decimal places.
'
' Expression:              Expression to be formatted.
'
' NumDecimalPlaces:        Indicates how many places to the right of the decimal are displayed.
'                          Default (–1) indicates that the computer's regional settings are used.
'
' IncludeLeadingZero:      Tristate constant that indicates whether or not a leading zero is
'                          displayed for fractional values.
'
' UseParensForNegNumbers:  Tristate constant that indicates whether or not to place negative
'                          values within parentheses.
'
' GroupDigits:             Tristate constant that indicates whether or not numbers are grouped
'                          using the group delimiter specified in the computer's regional settings.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FormatPercent(ByVal Expression As Double, Optional ByVal NumDecimalPlaces As Long = -1, Optional ByVal IncludeLeadingZero As eTristate = UseDefault, Optional ByVal UseParensForNegNumbers As eTristate = UseDefault, Optional ByVal GroupDigits As eTristate = UseDefault) As String
    Dim sFormat As String
    sFormat = GetLocaleFormat(NumDecimalPlaces, IncludeLeadingZero, GroupDigits)
    Select Case NumDecimalPlaces
        Case Is = 0
            FormatPercent = Format$(Round(Expression * 100#, NumDecimalPlaces), sFormat) & "%"
        Case Else
            FormatPercent = Format$(Round(Expression * 100#, NumDecimalPlaces), sFormat & mLI.sDecSep & String$(NumDecimalPlaces, "0")) & "%"
    End Select
    If (Expression < 0) And UseParensForNegNumbers Then
        FormatPercent = FormatNegativeNumbers(FormatPercent, UseParensForNegNumbers)
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' FormatCurrency Function
'
' VBA.Strings.FormatCurrency(Expression, [NumDecimalPlaces], [IncludeLeadingZero], _
'                                        [UseParensForNegNumbers], [GroupDigits])
'
' Returns an expression formatted as a currency value using
' the currency symbol defined in the system control panel.
'
' Expression:              Expression to be formatted.
'
' NumDecimalPlaces:        Indicates how many places to the right of the decimal are displayed.
'                          Default (–1) indicates that the computer's regional settings are used.
'
' IncludeLeadingZero:      Tristate constant that indicates whether or not a leading zero is
'                          displayed for fractional values.
'
' UseParensForNegNumbers:  Tristate constant that indicates whether or not to place negative
'                          values within parentheses.
'
' GroupDigits:             Tristate constant that indicates whether or not numbers are grouped
'                          using the group delimiter specified in the computer's regional settings.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FormatCurrency(ByVal Expression As Double, Optional ByVal NumDecimalPlaces As Long = -1, Optional ByVal IncludeLeadingZero As eTristate = UseDefault, Optional ByVal UseParensForNegNumbers As eTristate = UseDefault, Optional ByVal GroupDigits As eTristate = UseDefault) As String
    Dim sFormat As String
    sFormat = GetLocaleCurrencyFormat(NumDecimalPlaces, IncludeLeadingZero, GroupDigits)
    Select Case NumDecimalPlaces
        Case Is = 0
            FormatCurrency = Format$(Round(Expression, NumDecimalPlaces), mLCI.sSymbol & sFormat)
        Case Else
            FormatCurrency = Format$(Round(Expression, NumDecimalPlaces), mLCI.sSymbol & sFormat & mLCI.sDecSep & String$(NumDecimalPlaces, "0"))
    End Select
    If (Expression < 0) And UseParensForNegNumbers Then
        FormatCurrency = FormatNegativeCurrency(FormatCurrency, UseParensForNegNumbers)
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' FormatDateTime Function
'
' VBA.Strings.FormatDateTime(Expression, [NamedFormat])
'
' Returns an expression formatted as a date or time.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FormatDateTime(ByVal Expression, Optional ByVal NamedFormat As eNamedFormat = GeneralDate) As String
    Select Case NamedFormat
        Case GeneralDate:  FormatDateTime = Format$(Expression, "General Date") ' Shows date and time if expression contains both. If expression is only a date or a time, the missing information is not displayed. Date display is determined by user's system settings.
        Case LongDate:     FormatDateTime = Format$(Expression, "Long Date")    ' Uses the Long Date format specified by user's system settings.
        Case ShortDate:    FormatDateTime = Format$(Expression, "Short Date")   ' Uses the Short Date format specified by user's system settings.
        Case LongTime:     FormatDateTime = Format$(Expression, "Long Time")    ' Displays a time using user's system's long-time format; includes hours, minutes, seconds.
        Case ShortTime:    FormatDateTime = Format$(Expression, "Short Time")   ' Shows the hour and minute using the 24-hour hh:mm format.
    End Select
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' MonthName Function
'
' VBA.Strings.MonthName(Month As Long, [Abbreviate])
'
' Returns a string indicating the specified month.
'
' Month:      The numeric designation of the month.
'             For example, January is 1, February is 2, and so on.
'
' Abbreviate: Indicates if the month name is to be abbreviated.
'             If omitted the month name is not abbreviated.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function MonthName(ByVal Month As Long, Optional ByVal Abbreviate As Boolean) As String
    Dim sAbbrev As String
    Dim sDate As String
    If Abbreviate Then sAbbrev = "mmm" Else sAbbrev = "mmmm"
    sDate = Month & "/" & Month & "/29" ' mmm   Abreviated month names (Jan, Feb, Mar, etc).
    MonthName = Format$(sDate, sAbbrev) ' mmmm  Full month names (January, February, etc).
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSystemFirstDayOfWeek (Custom Support Function)
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetSystemFirstDayOfWeek() As eDayOfWeek
    Dim sFDOW As String
    Dim lRet As Long
    sFDOW = vbNullChar & vbNullChar
    ' NLS API value:  0 Mon, 1 Tues, 2 Wed, 3 Thurs, 4 Fri, 5 Sat, 6 Sun
    ' VB enumeration: 2 Mon, 3 Tues, 4 Wed, 5 Thurs, 6 Fri, 7 Sat, 1 Sun
    lRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_IFIRSTDAYOFWEEK, sFDOW, 2&)
    ' Convert API value to VB DayOfWeek enumeration (+ 2, mod sunday to 1)
    GetSystemFirstDayOfWeek = (CLng(Left$(sFDOW, 1)) + 1) Mod 7 + 1
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' WeekdayName Function
'
' VBA.Strings.WeekdayName(Weekday, [Abbreviate], [FirstDayOfWeek])
'
' Returns a string indicating the specified day of the week.
'
' Weekday:        The numeric designation for the day of the week.
'                 This value depends on the setting of FirstDayOfWeek.
'
' Abbreviate:     Indicates if the weekday name is to be abbreviated.
'                 If omitted the weekday name is not abbreviated.
'
' FirstDayOfWeek: Enumeration indicating the first day of the week.
'                 If omitted defaults to Sunday as first day of the
'                 week, so passing Weekday as 1 returns "Sunday".
'                 UseSystem - Uses National Language Support API
'                 setting returned for the LOCALE_SYSTEM_DEFAULT.
'
' IMPORTANT:
'
' The MSDN documentation states that FirstDayOfWeek defaults to
' Sunday if omitted, which corresponds to the Weekday function.
'
' However, the VB6 WeekdayName function actually defaults to
' UseSystem instead, while the VB6 Weekday function defaults
' to Sunday. This can cause errors if this parameter is not
' explicitly passed to both functions in VB6.
'
' This WeekdayName impementation defaults to Sunday for more
' predictable behavior. Code written for VB6 that explicitly
' passes this parameter to ensure correct behavior will work
' as expected using this function, while code that doesn't
' will work as expected only with this implementation.
'
' If today is a Tuesday:
'
' Weekday(Now) == 3
' VB6_Funcs.WeekdayName(Weekday(Now)) = "Tuesday"
' VBA.Strings.WeekdayName(Weekday(Now)) = "Wednesday"
'
' Weekday(Now, UseSystem) == 2
' VB6_Funcs.WeekdayName(2, UseSystem) = "Tuesday"
' VBA.Strings.WeekdayName(2, UseSystem) = "Tuesday"
'
' Weekday(Now, Thursday) == 6
' VB6_Funcs.WeekdayName(6, True, Thursday) = "Tue"
' VBA.Strings.WeekdayName(6, True, Thursday) = "Tue"
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function WeekdayName(ByVal Weekday As Long, Optional ByVal Abbreviate As Boolean, Optional ByVal FirstDayOfWeek As eDayOfWeek = Sunday) As String
    Dim sAbbrev As String
    If Abbreviate Then sAbbrev = "ddd" Else sAbbrev = "dddd"
    If FirstDayOfWeek = UseSystem Then               ' Strangely, Format$ ignores the FirstDayOfWeek
        FirstDayOfWeek = GetSystemFirstDayOfWeek     ' parameter when formatting "ddd" or "dddd" and
    End If                                           ' uses it's default -> Sunday.
    Weekday = Weekday + FirstDayOfWeek - 1           ' So we modify Weekday instead.
    WeekdayName = Format$(Weekday, sAbbrev, Sunday)  ' And (ineffectually) enforce Sunday.
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Weekday  (This is an existing VB5 Function)
'
' VBA.DateTime.Weekday(Date, [FirstDayOfWeek])
'
' Returns a whole number representing the day of the week.
'
' Weekday returns Sun = 1 to Sat = 7, unless modified by
' the FirstDayOfWeek parameter.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Weekday(ByVal vDate, Optional ByVal FirstDayOfWeek As eDayOfWeek = Sunday) As Long
    Weekday = Format$(vDate, "w", FirstDayOfWeek)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' FormatNegativeNumbers (Custom Function)
'
' Uses the computer's regional settings to format negative numbers.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FormatNegativeNumbers(NegativeNumber As String, Optional ByVal UseParensForNegNumbers As eTristate = UseDefault) As String
    Dim lNeg As Long
    If UseParensForNegNumbers = etFalse Then FormatNegativeNumbers = NegativeNumber: Exit Function
    If LenB(mLI.sDecSep) = 0 Then InitLocaleInfo
    lNeg = Len(mLI.sNegChr)
    Select Case IIf(UseParensForNegNumbers = etTrue, 0, mLI.iNegNum)
      Case 0: FormatNegativeNumbers = "(" & Mid$(NegativeNumber, 1 + lNeg) & ")"
      Case 1: FormatNegativeNumbers = NegativeNumber
      Case 2: FormatNegativeNumbers = mLI.sNegChr & " " & Mid$(NegativeNumber, 1 + lNeg)
      Case 3: FormatNegativeNumbers = Mid$(NegativeNumber, 1 + lNeg) & mLI.sNegChr
      Case 4: FormatNegativeNumbers = Mid$(NegativeNumber, 1 + lNeg) & " " & mLI.sNegChr
    End Select
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' FormatNegativeCurrency (Custom Function)
'
' Uses the computer's regional settings to format negative currency.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FormatNegativeCurrency(NegativeCurrency As String, Optional ByVal UseParensForNegNumbers As eTristate = UseDefault) As String
    Dim lDol As Long
    Dim lNeg As Long
    If UseParensForNegNumbers = etFalse Then FormatNegativeCurrency = NegativeCurrency: Exit Function
    If LenB(mLCI.sDecSep) = 0 Then InitLocaleCurrencyInfo
    lDol = Len(mLCI.sSymbol)
    lNeg = Len(mLCI.sNegChr)
    Select Case IIf(UseParensForNegNumbers = etTrue, 0, mLCI.iNegNum)
      Case 0:  FormatNegativeCurrency = "(" & Mid$(NegativeCurrency, 1 + lNeg) & ")"
      Case 1:  FormatNegativeCurrency = NegativeCurrency
      Case 2:  FormatNegativeCurrency = mLCI.sSymbol & mLCI.sNegChr & Mid$(NegativeCurrency, 1 + lNeg + lDol)
      Case 3:  FormatNegativeCurrency = Mid$(NegativeCurrency, 1 + lNeg) & mLCI.sNegChr
      Case 4:  FormatNegativeCurrency = "(" & Mid$(NegativeCurrency, 1 + lNeg + lDol) & mLCI.sSymbol & ")"
      Case 5:  FormatNegativeCurrency = mLCI.sNegChr & Mid$(NegativeCurrency, 1 + lNeg + lDol) & mLCI.sSymbol
      Case 6:  FormatNegativeCurrency = Mid$(NegativeCurrency, 1 + lNeg + lDol) & mLCI.sNegChr & mLCI.sSymbol
      Case 7:  FormatNegativeCurrency = Mid$(NegativeCurrency, 1 + lNeg + lDol) & mLCI.sSymbol & mLCI.sNegChr
      Case 8:  FormatNegativeCurrency = mLCI.sNegChr & Mid$(NegativeCurrency, 1 + lNeg + lDol) & " " & mLCI.sSymbol
      Case 9:  FormatNegativeCurrency = mLCI.sNegChr & mLCI.sSymbol & " " & Mid$(NegativeCurrency, 1 + lNeg + lDol)
      Case 10: FormatNegativeCurrency = Mid$(NegativeCurrency, 1 + lNeg + lDol) & " " & mLCI.sSymbol & mLCI.sNegChr
      Case 11: FormatNegativeCurrency = mLCI.sSymbol & " " & Mid$(NegativeCurrency, 1 + lNeg + lDol) & mLCI.sNegChr
      Case 12: FormatNegativeCurrency = mLCI.sSymbol & " " & mLCI.sNegChr & Mid$(NegativeCurrency, 1 + lNeg + lDol)
      Case 13: FormatNegativeCurrency = Mid$(NegativeCurrency, 1 + lNeg + lDol) & mLCI.sNegChr & " " & mLCI.sSymbol
      Case 14: FormatNegativeCurrency = "(" & mLCI.sSymbol & " " & Mid$(NegativeCurrency, 1 + lNeg + lDol) & ")"
      Case 15: FormatNegativeCurrency = "(" & Mid$(NegativeCurrency, 1 + lNeg + lDol) & " " & mLCI.sSymbol & ")"
    End Select
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Custom Functions
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

' Remove double quotes from start and/or end if present
Public Function TrimQuotes(sString As String) As String
   Dim s As String, i As Long, j As Long
   s = Trim$(sString): j = Len(s)
   If (j <> 0) Then
      If (Left$(s, 1) = Chr$(34)) Then i = 2 Else i = 1
      If (Right$(s, 1) = Chr$(34)) Then j = j - i
      If (j - i > 0) Then TrimQuotes = Mid$(s, i, j)
   End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FormatBytes(ByVal ByteSize As Double) As String 'Original code Ben White
    Dim x As Long, y As Long, z As Double
    Dim sBytes As String, sUnits As String
    Do: x = x + 1
       z = 2# ^ (x * 10#)
       For y = 1 To 3
          If ByteSize < z * (10# ^ y) Then
             sBytes = FormatNumber(ByteSize / z, 3 - y)
             Exit Do
    End If: Next: Loop
    sUnits = Choose(x, "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB", "??", "??", "??", "??")
    FormatBytes = Space$(5 - Len(sBytes)) & sBytes & Space$(2) & sUnits & Space$(1)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''

' Returns 'the' before 'their' when in the same location
Public Function MultiInStr(sSrc As String, sTerms() As String, _
                           Optional ByVal lStart As Long = 1, _
                           Optional ByVal eCompare As VbCompareMethod = vbBinaryCompare, _
                           Optional ByVal lRightLimit As Long = -1, _
                           Optional ByRef lHitItemIndex As Long) As Long ' Kenneth Buckmaster
    Dim iPos As Long
    Dim iHit As Long
    Dim iIdx As Long
    Dim bHit As Boolean
    If lRightLimit = -1 Then lRightLimit = Len(sSrc)
    iHit = Len(sSrc) + 1
    For iIdx = LBound(sTerms) To UBound(sTerms)
       iPos = InStr(lStart, sSrc, sTerms(iIdx), eCompare)
       If iPos Then
          If iPos < iHit Then
              bHit = True
          ElseIf iPos = iHit Then
              bHit = LenB(sTerms(iIdx)) < LenB(sTerms(lHitItemIndex))
          End If
          If bHit Then
              iHit = iPos
              lHitItemIndex = iIdx
              bHit = False
          End If
       End If
    Next
    If iHit < Len(sSrc) + 1 Then MultiInStr = iHit
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''

' Always returns 'heir' before 'their' for reverse search
' Always returns 'their' before 'the' for reverse search
Public Function MultiInStrR(sSrc As String, sTerms() As String, _
                            Optional ByVal lRightStart As Long = -1, _
                            Optional ByVal eCompare As VbCompareMethod = vbBinaryCompare, _
                            Optional ByVal lLeftLimit As Long = 1, _
                            Optional ByRef lHitItemIndex As Long) As Long ' Kenneth Buckmaster
    Dim iLast As Long
    Dim iPos As Long
    Dim iHit As Long
    Dim iIdx As Long
    Dim bHit As Boolean
    If lRightStart = -1 Then lRightStart = Len(sSrc)
    For iIdx = LBound(sTerms) To UBound(sTerms)
       iPos = InStr(lLeftLimit, sSrc, sTerms(iIdx), eCompare)
       Do Until iPos = 0 Or iPos > lRightStart
          iLast = iPos
          iPos = InStr(iLast + 1, sSrc, sTerms(iIdx), eCompare)
       Loop
       If iLast > iHit Then
          bHit = True
       ElseIf iLast = iHit Then
          bHit = LenB(sTerms(iIdx)) > LenB(sTerms(lHitItemIndex))
       End If
       If bHit Then
          iHit = iLast
          lHitItemIndex = iIdx
          lLeftLimit = iLast
          iLast = 0
          bHit = False
       End If
    Next
    If iHit Then MultiInStrR = iHit
End Function

'-BuildStr-----------------------------------------------
'  This function can replace vb's string & concatenation.
'  The speed is very similar for simple appends:
'     sResult = sResult & "text"
'     sResult = BuildStr(sResult, "text")
'  But for more substrings this function is much faster
'  because vb's multiple appending is very slow:
'     sResult = sResult & "some" & "more" & "text"
'     sResult = BuildStr(sResult, "some", "more", "text")
'  Notice you can safely pass as an argument the variable
'  that the function is assigning back to (compiler safe).
'  You can also specify the delimiter character(s) to
'  insert between the appended substrings, and will work
'  correctly if an argument is omitted or passed empty:
'     sMsg = BuildStr("s1", , "s2", "s3", vbCrLf)
'     MsgBox BuildStr("", sMsg, , "s4", vbCrLf)
'--------------------------------------------------------
Public Function BuildStr(Str1 As String, Optional Str2 As String, Optional Str3 As String, _
                                         Optional Str4 As String, Optional Delim As String) As String
    Dim LenWrk As Long, LenAll As Long
    Dim LenDlm As Long, TotDlm As Long
    Dim Len1 As Long, Len2 As Long
    Dim Len3 As Long, Len4 As Long
    Dim iPos As Long
    Len1 = LenB(Str1): Len2 = LenB(Str2)
    Len3 = LenB(Str3): Len4 = LenB(Str4)
    LenDlm = LenB(Delim)
    If (LenDlm) Then
        TotDlm = -LenDlm
        If (Len1) Then TotDlm = 0
        If (Len2) Then TotDlm = TotDlm + LenDlm
        If (Len3) Then TotDlm = TotDlm + LenDlm
        If (Len4) Then TotDlm = TotDlm + LenDlm
    End If
    LenAll = Len1 + Len2 + Len3 + Len4 + TotDlm
    If (LenAll > 0) Then
        BuildStr = Space$(LenAll * 0.5)
        iPos = 1
        If (Len1) Then
            MidB$(BuildStr, iPos) = Str1
            LenWrk = Len1
        End If
        If (Len2) Then
            If (LenDlm) Then If (LenWrk) Then GoSub InsDelim
            MidB$(BuildStr, iPos + LenWrk) = Str2
            LenWrk = LenWrk + Len2
        End If
        If (Len3) Then
            If (LenDlm) Then If (LenWrk) Then GoSub InsDelim
            MidB$(BuildStr, iPos + LenWrk) = Str3
            LenWrk = LenWrk + Len3
        End If
        If (Len4) Then
            If (LenDlm) Then If (LenWrk) Then GoSub InsDelim
            MidB$(BuildStr, iPos + LenWrk) = Str4
        End If
    End If
    Exit Function
InsDelim:
    MidB$(BuildStr, iPos + LenWrk) = Delim
    LenWrk = LenWrk + LenDlm
    Return
End Function

' ===========================================================================
' The RTrimChr function removes from sStr the first occurrence from the right
' of the specified character(s) and everything following it, and returns just
' the start of the string up to but not including the specified character(s).
' It always searches from right to left starting at the end of sStr. If the
' character(s) does not exist in sStr then the whole of sStr is returned and
' lRetPos is set to Len(sStr) + 1. sChar defaults to a backslash if omitted.
' ===========================================================================
Public Function RTrimChr(sStr As String, Optional sChar As String = "\", Optional ByRef lRetPos As Long, _
                         Optional ByVal eCompare As eCompareMethod = BinaryCompare) As String
    Dim lPos As Long
    lRetPos = Len(sStr) + 1
    If LenB(sChar) Then
        lPos = InStr(1, sStr, sChar, eCompare)
        Do Until lPos = 0
            lRetPos = lPos
            lPos = InStr(lRetPos + 1, sStr, sChar, eCompare)
        Loop
    End If
    ' Return sStr w/o sChar and any following substring
    RTrimChr = LeftB$(sStr, lRetPos + lRetPos - 2)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Private Support Routines
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub InitLocaleInfo()
    Dim sBuff As String
    Dim iLenRet As Long
    sBuff = String$(16, vbNullChar)

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_ILZERO, sBuff, 2)
    mLI.sZero = String$(CLng(Left$(sBuff, 1)), "0")  ' Leading zero for decimal.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_INEGNUMBER, sBuff, 2)
    mLI.iNegNum = CLng(Left$(sBuff, 1))              ' Negative number mode.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SNEGATIVESIGN, sBuff, 3)
    mLI.sNegChr = Left$(sBuff, iLenRet - 1)          ' String value for the negative sign.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SDECIMAL, sBuff, 3)
    mLI.sDecSep = Left$(sBuff, iLenRet - 1)          ' Character(s) used as the decimal separator.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_IDIGITS, sBuff, 3)
    mLI.iNumDec = CLng(Left$(sBuff, iLenRet - 1))    ' Number of fractional digits.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_STHOUSAND, sBuff, 8)
    mLI.sGroups = Left$(sBuff, iLenRet - 1)          ' Character(s) used to separate groups of digits to the left of the decimal.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SGROUPING, sBuff, 16)
    mLI.iGroups = CLng(Left$(sBuff, 1))              ' Sizes for each group of digits to the left of the decimal.
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetLocaleFormat(NumDecimalPlaces As Long, ByVal IncludeLeadingZero As eTristate, ByVal GroupDigits As eTristate) As String
    Dim sZero As String
    If LenB(mLI.sDecSep) = 0 Then InitLocaleInfo

    If NumDecimalPlaces < 0 Then
        NumDecimalPlaces = mLI.iNumDec
    End If
    If IncludeLeadingZero = UseDefault Then
        sZero = mLI.sZero
    ElseIf IncludeLeadingZero Then
        sZero = "0"
    End If
    If GroupDigits = UseDefault Then
        If mLI.iGroups Then
            GetLocaleFormat = "#" & mLI.sGroups & String$(mLI.iGroups - Len(sZero), "#") & sZero
        Else
            GetLocaleFormat = "#" & sZero
        End If
    ElseIf GroupDigits Then
        GetLocaleFormat = "#" & mLI.sGroups & String$(3 - Len(sZero), "#") & sZero
    Else
        GetLocaleFormat = "#" & sZero
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub InitLocaleCurrencyInfo()
    Dim sBuff As String
    Dim iLenRet As Long
    sBuff = String$(16, vbNullChar)

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_ILZERO, sBuff, 2)
    mLCI.sZero = String$(CLng(Left$(sBuff, 1)), "0")  ' Leading zero for decimal.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_INEGCURR, sBuff, 2)
    mLCI.iNegNum = CLng(Left$(sBuff, 1))              ' Negative currency mode.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SNEGATIVESIGN, sBuff, 3)
    mLCI.sNegChr = Left$(sBuff, iLenRet - 1)          ' String value for the negative sign.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SCURRENCY, sBuff, 3)
    mLCI.sSymbol = Left$(sBuff, iLenRet - 1)          ' String used as the local monetary symbol.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SMONDECIMALSEP, sBuff, 3)
    mLCI.sDecSep = Left$(sBuff, iLenRet - 1)          ' Character(s) used as the monetary decimal separator.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_ICURRDIGITS, sBuff, 3)
    mLCI.iNumDec = CLng(Left$(sBuff, iLenRet - 1))    ' Number of fractional digits for the local monetary format.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SMONTHOUSANDSEP, sBuff, 8)
    mLCI.sGroups = Left$(sBuff, iLenRet - 1)          ' Character(s) used as the monetary separator between groups to the left of the decimal.

    iLenRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SMONGROUPING, sBuff, 16)
    mLCI.iGroups = CLng(Left$(sBuff, 1))              ' Sizes for each group of monetary digits to the left of the decimal.
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetLocaleCurrencyFormat(NumDecimalPlaces As Long, ByVal IncludeLeadingZero As eTristate, ByVal GroupDigits As eTristate) As String
    Dim sZero As String
    If LenB(mLCI.sDecSep) = 0 Then InitLocaleCurrencyInfo

    If NumDecimalPlaces < 0 Then
        NumDecimalPlaces = mLCI.iNumDec
    End If
    If IncludeLeadingZero = UseDefault Then
        sZero = mLCI.sZero
    ElseIf IncludeLeadingZero Then
        sZero = "0"
    End If
    If GroupDigits = UseDefault Then
        If mLCI.iGroups Then
            GetLocaleCurrencyFormat = "#" & mLCI.sGroups & String$(mLCI.iGroups - Len(sZero), "#") & sZero
        Else
            GetLocaleCurrencyFormat = "#" & sZero
        End If
    ElseIf GroupDigits Then
        GetLocaleCurrencyFormat = "#" & mLCI.sGroups & String$(3 - Len(sZero), "#") & sZero
    Else
        GetLocaleCurrencyFormat = "#" & sZero
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
