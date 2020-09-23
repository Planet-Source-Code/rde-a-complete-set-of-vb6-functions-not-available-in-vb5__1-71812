Attribute VB_Name = "PerfTimer"
Option Explicit                                              '-©Rd-

' Performance Counter API's
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public Enum ePerf
    epSeconds   ' Seconds
    epMillisecs ' Milliseconds - one thousandth of a second
    epMicrosecs ' Microseconds - one millionth of a second
End Enum

Private mDblFreq As Double

' ===========================================================================
' Precise Timer - ProfileStart and ProfileStop
' ===========================================================================
'
' ProfileStart returns the current value of the high-resolution performance
' counter as a Double data type. This value is passed on to ProfileStop which
' subtracts it from an ending value and returns the difference (elapsed time).
'
' In the case of no high-resolution timer, ProfileStart returns zero.
' Multiple performance timers can be run concurrently if required.
'
' The result is returned in the time scale specified, and is accurate to
' the maximum decimal places returned by QueryPerformanceCounter.
'
' In the case of no high-resolution timer, ProfileStop returns zero.
' ===========================================================================
Public Function ProfileStart() As Double
    Dim curStart As Currency
    Dim curFreq As Currency
    If (mDblFreq = 0) Then
        QueryPerformanceFrequency curFreq
        mDblFreq = CDbl(curFreq)
    End If
    If (mDblFreq) Then
        QueryPerformanceCounter curStart
        ProfileStart = CDbl(curStart)
    End If
End Function

' Dim d As Double
' d = ProfileStart
' txtDisplay = FormatElapsed(ProfileStop(d, epMillisecs))

Public Function ProfileStop(ByVal dblStart As Double, Optional ByVal eTimeScale As ePerf) As Double
    Dim curStop As Currency
    If (mDblFreq) Then
        QueryPerformanceCounter curStop ' cpu tick accurate
        Select Case eTimeScale
            Case epSeconds:   ProfileStop = CDbl(curStop - dblStart) / mDblFreq
            Case epMillisecs: ProfileStop = (CDbl(curStop - dblStart) / mDblFreq) * 1000#
            Case epMicrosecs: ProfileStop = (CDbl(curStop - dblStart) / mDblFreq) * 1000000#
            ' Multiply by 1000 to convert to milliseconds or 1000000 for microseconds µs
        End Select
    End If
End Function

' ===========================================================================
' Elapsed Time Formatting - FormatElapsed
' ===========================================================================
'
' The return value from ProfileStop can be passed to this function to format
' the elapsed time to a string representation of the decimal value.
'
' The result returned is accurate to the maximum decimal places contained in
' the dblElapsed argument if the optional argument is omitted or set to < 0.
'
' Otherwise, it returns the elapsed time rounded to the number of decimal
' places specified by lngDecPlaces.
' ===========================================================================
Public Function FormatElapsed(ByVal dblElapsed As Double, Optional ByVal lngDecPlaces As Long = -1) As String
    Select Case lngDecPlaces
        Case Is < 0
            FormatElapsed = CStr(CDec(dblElapsed))
        Case Is = 0
            FormatElapsed = Format$(CStr(CDec(dblElapsed)), "#0")
        Case Else
            FormatElapsed = Format$(CStr(CDec(dblElapsed)), "#0." & String$(lngDecPlaces, "0"))
    End Select
End Function
' ===========================================================================
