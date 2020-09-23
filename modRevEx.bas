Attribute VB_Name = "modRevEx"
Option Explicit

'.------------------------------------------------------------------------------
'. RevEx - creates a random alphanumeric string based on a pattern.
'.
'. Think of it as regular expression pattern matching in reverse.
'.
'. Useful if you need to create a random combination of values in a
'. certain format. Letters (lower/upper case) and decimal numbers or
'. any combination or part thereof: boolean, hex, octal values; dates, times,
'. vowels, consonants, phone numbers, zip codes, GUIDs; IP, social security,
'. credit card numbers ...
'.
'. Usage:
'. .RevEx ("[pattern][pattern][pattern]")
'.
'. Examples:
'. [a][b][c]                    = specific lower case characters
'. [A][B][C]                    = specific upper case characters
'. [#][#][#]                    = numbers (decimal)
'. [123][abc][ABC][9zZ]         = sets (choose *only* from set)
'. [!123][!abc][!ABC][!0aA]     = negative sets (choose *not* from set)
'. [0-9][a-z][A-Z][a-Z]         = ranges (choose *only* from range)
'. [!0-9][!a-z][!A-Z][!a-Z]     = negative ranges (choose *not* from range)
'.
'. Note:
'. Depending on what kind of data you'd like to create, be aware that some
'. results will need postprocessing to make them valid. For example: If you'd like
'. to get the days in a month (00-31) - there's currently no pattern to match this.
'. So better check, parse, pad, replace and/or combine to make them valid...
'.
'.
'. Hi and Bye from Germany...
'.
'. Send comments, bug reports, etc... to Roman.Kehr@gmx.de
'.------------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' DECLARATIONS - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim strPossibleValues(3) As String
Dim strPatternStartTag As String
Dim strPatternEndTag As String
Dim strPatternDelimiter As String
Dim strNumberTag As String
Dim strNotTag As String
Dim strRangeTag As String

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' DECLARATIONS - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'-------------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' FUNCTIONS (PUBLIC) - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'.------------------------------------------------------------------------------
'.  Function  : Public Function RevEx
'.
'.
'.  Prereq.    : -
'.
'.  Parameters: strPattern As String
'.
'.  Returns   : Variant (NULL = Error)
'.
'.  Comments  : Creates and returns a random alphanumeric string based on a
'.              supplied pattern. Prepares the arguments and calls the actual
'.              pattern matching routine.
'.
'.  Author    : Roman Kehr - 14.04.2002
'.  Changed   : -
'.------------------------------------------------------------------------------

Public Function RevEx(strPattern) As Variant
    Dim arrPattern() As String
    Dim vntRetval As Variant
    Dim vntDummy As Variant
    
    ' change this, if you don't like the way the parameters are passed
    strPatternStartTag = "["
    strPatternEndTag = "]"
    strPatternDelimiter = "]["
    strNumberTag = "#" ' < change to s.th. else, if you need to create the "#" character
    strNotTag = "!"
    strRangeTag = "-"
    
    ' the set of allowed values
    ' this is only important for numbers and ranges
    ' it's no problem to use different characters, as long as you use a set
    ' .RevEx ("[><*]")
    '
    strPossibleValues(0) = "0123456789" ' numbers only
    strPossibleValues(1) = "abcdefghijklmnopqrstuvwxyz" ' lower case only
    strPossibleValues(2) = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" ' upper case only
    
    On Error GoTo Errorhandler
    
    ' clean up passed parameter, remove start / end tags if necessary
    strPattern = Trim(strPattern)
    If Left(strPattern, 1) = strPatternStartTag Then strPattern = Right(strPattern, Len(strPattern) - 1)
    If Right(strPattern, 1) = strPatternEndTag Then strPattern = Left(strPattern, Len(strPattern) - 1)
    
    ' string -> array
    arrPattern = Split(strPattern, strPatternDelimiter)
    
    ' iterate through the array and call actual pattern matching function
    Dim i As Integer
    For i = LBound(arrPattern) To UBound(arrPattern)
        vntRetval = vntRetval & MatchPattern(arrPattern(i))
    Next i
    
    RevEx = vntRetval
    
    Exit Function
Errorhandler:
    RevEx = Null
End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' FUNCTIONS (PUBLIC) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'-------------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' FUNCTIONS (PRIVATE) - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'.------------------------------------------------------------------------------
'.  Function   : Private MatchPattern
'.
'.
'.  Prereq.    : -
'.
'.  Parameters : strPatternFragment As String
'.
'.  Returns    : Variant (NULL = Error)
'.
'.  Comments   : This function does the actual pattern matching
'.               Returns Variant, so that Errors can be returned to calling
'.               function
'.
'.  Author     : Roman Kehr 14.04.2002
'.  Changed    :
'.------------------------------------------------------------------------------

Private Function MatchPattern(strPatternFragment As String) As Variant
    Dim bolInclusive As Boolean
    Dim strDummy As String
    Dim vntRetval As Variant
    Dim i As Integer
    
    On Error GoTo Errorhandler
    
    ' let's check simple cases first
    ' this order is important, to catch cases in which only
    ' the NOT tag ("!") or the RANGE tag ("-") or space (" ") has been passed
    If Len(strPatternFragment) = 1 Then
        
        If strPatternFragment = strNumberTag Then
            ' return single number
            ' could have returned the random number right away,
            ' but this way it's more consistent and easier to maintain
            vntRetval = Mid(strPossibleValues(0), GetRandomNumber(1, Len(strPossibleValues(0))), 1)
        Else
            ' do nothing, return parameter
            vntRetval = strPatternFragment
        End If
    
    Else ' now on to the more complex cases
        
        ' clean up
        strPatternFragment = Trim(strPatternFragment)
        
        ' is it a *negative* set or range ?
        If Left(strPatternFragment, 1) = strNotTag Then ' yes
            bolInclusive = False ' set flag accordingly
            strPatternFragment = Right(strPatternFragment, Len(strPatternFragment) - 1) ' remove NOT tag
        Else ' nope
            bolInclusive = True ' set flag accordingly
        End If
        
        ' is it a set or a range ?
        If InStr(1, strPatternFragment, strRangeTag) = False Then ' it's a set
            
            If bolInclusive = True Then ' it's a *positive* set
                vntRetval = Mid(strPatternFragment, GetRandomNumber(1, Len(strPatternFragment)), 1)
            Else ' it's a *negative* set
            
                strDummy = strPossibleValues(0) + strPossibleValues(1) + strPossibleValues(2)
                For i = 1 To Len(strPatternFragment)
                    strDummy = Replace(strDummy, Mid(strPatternFragment, i, 1), "")
                Next i
                
                vntRetval = Mid(strDummy, GetRandomNumber(1, Len(strDummy)), 1)
                
            End If
        
        Else ' it's a range
           
            Dim strStart As String
            Dim strEnd As String
            Dim strReplace As String
            Dim intStart As Integer
            Dim intEnd As Integer
            
            ' we now need to take a look at *all* the possible values, bring them together: array -> string
            For i = LBound(strPossibleValues) To UBound(strPossibleValues)
                strDummy = strDummy & strPossibleValues(i)
            Next i
            
            ' get the characters marking start and end of the range
            strStart = Left(strPatternFragment, 1)
            strEnd = Right(strPatternFragment, 1)
            
            ' now get the character's position
            intStart = InStr(1, strDummy, strStart)
            intEnd = InStr(1, strDummy, strEnd)
            
            ' get the range
            strReplace = Mid(strDummy, intStart, intEnd - intStart + 1)
            
            ' is it a positive or a negative range?
            If bolInclusive = True Then ' positive
                strDummy = strReplace
            Else ' negative
                strDummy = Replace(strDummy, strReplace, "") ' remove range from possible values
            End If
            
            ' get value from what's left
            vntRetval = Mid(strDummy, GetRandomNumber(1, Len(strDummy)), 1)
            
        End If
    End If
    
    MatchPattern = vntRetval
    
    Exit Function
Errorhandler:
    MatchPattern = Null
End Function

'.------------------------------------------------------------------------------
'.  Function   : Private Function GetRandomNumber
'.
'.
'.  Prereq.    : -
'.
'.  Parameters : intLowBound As Integer
'.               intHighBound As Integer
'.
'.  Returns    : Integer
'.
'.  Comments   : Returns a random integer in the supplied range
'.
'.  Author     : Roman Kehr 14.04.2002
'.  Changed    :
'.------------------------------------------------------------------------------

Private Function GetRandomNumber(intLowBound As Integer, intHighBound As Integer) As Integer
    Randomize
    GetRandomNumber = Int(((intHighBound - intLowBound + 1) * Rnd + intLowBound))
End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' FUNCTIONS (PRIVATE) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
