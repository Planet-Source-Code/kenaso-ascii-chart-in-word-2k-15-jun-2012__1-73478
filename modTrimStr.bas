Attribute VB_Name = "modTrimStr"
' ***************************************************************************
' Routine:       modTrimStr
'
' Description:   This is Greg Millers TrimStr function modified for speed.
'                Now much faster but has the same bullet proof behavior.
'
'                TrimStr trims all control and other non-alphanumeric
'                characters from the passed string. Through the process
'                of elimination, it skips characters not recognized as
'                valid characters from before and after the string.
'                There are many examples that trim the null characters
'                from the end of the string but not all examples seem
'                to work correctly on all strings.
'
' References:    
'   21-Apr-1998  Greg Miller - Original code
'   TrimNullStr
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=854&lngWId=1
'
'   22-Feb-2011  RD Edwards - Refined and optimized for maximum speed
'   Functional TrimNull functions
'   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=71593&lngWId=1
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Apr-1998  Greg Miller - Original code
'              TrimNullStr
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=854&lngWId=1
' 22-Feb-2011  RD Edwards - Refined and optimized for maximum speed
'              Functional TrimNull functions
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=71593&lngWId=1
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const FADF_NO_REDIM As Long = &H11

' ***************************************************************************
' Type Structures
' ***************************************************************************
  ' SAFEARRAY Header, used in place of
  ' the real one to trick VB into allowing
  ' access to string data in-place
  Private Type SAFEARRAY1D
      cDims      As Integer   ' Count of dimensions in this array
      fFeatures  As Integer   ' Bitfield flags indicating attributes of array
      cbElements As Long      ' Byte size of each element of the array
      cLocks     As Long      ' Number of times the array has been locked without
                              ' corresponding unlock. The cLocks field is a
                              ' reference count that indicates how many times the
                              ' array has been locked. When there is no lock, you
                              ' are not supposed to access the array data, which
                              ' is located in pvData.
      pvData     As Long      ' Pointer to start of array data (use only if cLocks > 0)
      cElements  As Long      ' Count of elements in this dimension
      lLbound    As Long      ' The lower-bounding index of this dimension
      lUbound    As Long      ' The upper-bounding index of this dimension
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' ZeroMemory is used for clearing contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       TrimStr
'
' Description:   This is Greg Millers TrimStr function modified for speed.
'                Now much faster but has the same bullet proof behavior.
'
'                TrimStr trims all control and other non-alphanumeric
'                characters from the passed string. Through the process
'                of elimination, it skips characters not recognized as
'                valid characters from before and after the string.
'                There are many examples that trim the null characters
'                from the end of the string but not all examples seem
'                to work correctly on all strings.
'
' Parameters:    strData - Data string to be evaluated
'
' Returns:       Modified data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Apr-1998  Greg Miller - Original code
'              TrimNullStr
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=854&lngWId=1
' 22-Feb-2011  RD Edwards - Refined and optimized for maximum speed
'              Functional TrimNull functions
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=71593&lngWId=1
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 05-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code
' ***************************************************************************
Public Function TrimStr(ByVal strData As String) As String

    Dim lngEnd     As Long
    Dim lngStart   As Long
    Dim lngLength  As Long
    Dim aintChrs() As Integer
    Dim typSA      As SAFEARRAY1D

    On Error GoTo TrimStr_CleanUp
    
    ' if no data then leave
    If Len(Trim$(strData)) = 0 Then
        TrimStr = vbNullString
        Exit Function
    End If
    
    ZeroMemory typSA, Len(typSA)         ' Empty type structure
    
    ' Load type structure
    With typSA
        If .cDims = 0 Then
            
            .cDims = 1                   ' 1 Dimensional
            .fFeatures = FADF_NO_REDIM   ' Cannot REDIM the array
            .cbElements = 2&             ' This is an integer array
            .lLbound = 1&                ' Set lower-bound to one
            
            ' Convert data string into numeric
            ' equivalents into temp array
            CopyMemory ByVal ArrayPtr(aintChrs), VarPtr(typSA), 4&
        End If
        
        .pvData = StrPtr(strData)   ' Point at source string
        .cElements = Len(strData)   ' Set string length
        .cLocks = 1&                ' Lock the array
        lngLength = .cElements      ' Get length of string
    End With

    ' See if we have some data
    If (lngLength <> 0) Then
            
        ' Find first valid character by
        ' parsing forwards thru data string
        For lngStart = 1 To lngLength
            
            Select Case aintChrs(lngStart)
                   Case 33 To 126, 160 To 223
                        Exit For   ' Found valid character
            End Select
            
        Next lngStart
    
        ' Find last valid character by parsing
        ' backwards thru data string
        For lngEnd = lngLength To lngStart Step -1
            
            Select Case aintChrs(lngEnd)
                   Case 33 To 126, 160 To 223
                        Exit For   ' Found valid character
            End Select
            
        Next lngEnd
    
        lngLength = (lngEnd - lngStart) + 1   ' Calc data length
        
    End If
    
    Select Case lngLength
           Case Is > 0: TrimStr = Mid$(strData, lngStart, lngLength)  ' Format data to be returned
           Case Else:   TrimStr = vbNullString                        ' Return empty string
    End Select

TrimStr_CleanUp:
    ZeroMemory typSA, Len(typSA)   ' Empty type structure
    On Error GoTo 0                ' Nullify this error trap
    
End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

' ***************************************************************************
' Routine:       ArrayPtr
'
' Description:   This function returns a pointer to the SAFEARRAY header of
'                any Visual Basic array, including a Visual Basic string
'                array. Substitutes both ArrPtr and StrArrPtr. This function
'                will work with vb5 or vb6 without modification.
'
' Parameters:    vntData - Data to be evaluated
'
' Returns:       Zero     - Bad data (FALSE)
'                Non-zero - Good data (TRUE)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Feb-2011  RD Edwards - Refined and optimized for maximum speed
'              Functional TrimNull functions
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=71593&lngWId=1
' ***************************************************************************
Private Function ArrayPtr(ByRef vntData As Variant) As Long

    Dim lngDataType As Long   ' Must be a long integer

    On Error GoTo ArrayPtr_Exit
 
    ' Get the real VarType of the argument,
    ' this is similar to VarType(), but
    ' returns also the VT_BYREF bit
    CopyMemory lngDataType, vntData, 2&

    ' if a valid array was passed
    If (lngDataType And vbArray) = vbArray Then
      
        ' get the address of the SAFEARRAY descriptor
        ' stored in the second half of the Variant
        ' parameter that has received the array.
        ' Thanks to Francesco Balena.
        CopyMemory ArrayPtr, ByVal VarPtr(vntData) + 8&, 4&
        
    End If

ArrayPtr_Exit:
    On Error GoTo 0   ' Nullify this error trap
    
End Function

