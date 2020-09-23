Attribute VB_Name = "modProcesses"

' ***************************************************************************
' Routine:   modProcesses
'
' Purpose:   Stops designated processes completely.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote module
' 05-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added ProcessTerminate() routine.  
'              - Removed obsolete code. 
' ***************************************************************************
Option Explicit
  
' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "modProcesses"
  Private Const MAX_SIZE    As Long = 260

' ****************************************************************************
' Type Structures
' ****************************************************************************
  Private Type HANDLE_COLLECTION
      Handle    As Long
      ProcessID As Long
      Caption   As String   ' Used for debugging only
  End Type
  
  Private Type LUID
      LowPart   As Long
      HighPart  As Long
  End Type

  Private Type TOKEN_PRIVILEGES
      PrivilegeCount As Long
      TheLuid        As LUID
      Attributes     As Long
  End Type

' ****************************************************************************
' API Declares
' ****************************************************************************
  ' The EnumWindows() function enumerates all top-level windows on the screen
  ' by passing the handle of each window, in turn, to an application-defined
  ' callback function. EnumWindows() continues until the last top-level window
  ' is enumerated or the callback function returns FALSE.
  Private Declare Function EnumWindows Lib "user32" _
          (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

  ' GetParent returns the handle of the parent window of another window.
  ' If successful, the function returns a handle to the parent window.
  Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
  
  ' Always close an objects handle if it is not being used.
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long
        
  ' The TerminateProcess() function is used to unconditionally cause a
  ' process to exit and not save anything.
  Private Declare Function TerminateProcess Lib "kernel32" _
          (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

  ' The GetWindowText() function copies the text of the specified windows
  ' (Parent) title bar (if it has one) into a buffer. If the specified window
  ' is a control, the text of the control is copied.
  Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
          (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
  
  ' The GetWindowThreadProcessId() function retrieves the identifier of the
  ' thread that created the specified window and, optionally, the identifier
  ' of the process that created the window.
  Private Declare Function GetWindowThreadProcessId Lib "user32" _
          (ByVal hwnd As Long, lpdwProcessId As Long) As Long

  ' OpenProcess function returns a handle of an existing process object.
  ' If the function succeeds, the return value is an open handle of the
  ' specified process.
  Private Declare Function OpenProcess Lib "kernel32" _
          (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
          ByVal dwProcessId As Long) As Long
  
  ' OpenProcessToken function opens the access token associated with a
  ' process. If the function succeeds, the return value is nonzero.
  Private Declare Function OpenProcessToken Lib "advapi32.dll" _
          (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
          TokenHandle As Long) As Long
  
  ' AdjustTokenPrivileges function enables or disables privileges in
  ' the specified access token. Enabling or disabling privileges in an
  ' access token requires TOKEN_ADJUST_PRIVILEGES access.  If the function
  ' fails, the return value is zero.
  Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
          (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
          NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
          PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
  
  ' LookupPrivilegeValue function retrieves the locally unique identifier
  ' (LUID) used on a specified system to locally represent the specified
  ' privilege name.  If the function succeeds, the return value is nonzero.
  Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" _
          Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
          ByVal lpName As String, lpLuid As LUID) As Long
  
  ' GetCurrentProcess function returns a pseudohandle for the current
  ' process.  The return value is a pseudohandle to the current process.
  Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

  ' ZeroMemory is used for clearing contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Type Structure)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a typ Handles
' Variable name:     matypHandles
'
' ***************************************************************************
  Private mblnFoundit      As Boolean
  Private mlngHwnd         As Long
  Private mlngCount        As Long
  Private mlngProcessID    As Long
  Private mstrPartialTitle As String
  Private matypHandles()   As HANDLE_COLLECTION

' ***************************************************************************
' Routine:       StopStubbornPgms
'
' Description:   This routine is used to stop processes that are known for
'                hanging when shutting down a PC.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 09-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Optimized routine
' ***************************************************************************
Public Sub StopStubbornPgms()

    Dim lngIndex   As Long
    Dim astrData() As String
            
    ' Partial names of applications.
    ' Allow room for growth.
    Const APPL_CNT As Long = 20
    
    ' Verify most of the active applications
    ' are deactivated.  Makes for an easier
    ' shutdown process.
    
    Erase astrData()          ' Always start with empty arrays
    ReDim astrData(APPL_CNT)  ' Size temp array
    
    ' Preload array with one blank space
    For lngIndex = 0 To APPL_CNT - 1
        astrData(lngIndex) = Chr$(32)
    Next lngIndex
    
    ' Load array with partial names of known
    ' applications that sometimes hang during
    ' shutdown process
    astrData(0) = "cd"         ' CD burner software (Partial caption title)
    astrData(1) = "burner"     ' CD burner software (Partial caption title)
    astrData(2) = "ccapp"      ' CD burner software (Partial caption title)
    astrData(3) = "media"      ' multimedia software (Partial caption title)
    astrData(4) = "inbox"      ' Mail applications (Partial caption title)
    astrData(5) = "fax"        ' Fax monitoring (Partial caption title)
    astrData(6) = "mshow"      ' MS web cast application (Partial caption title)
    astrData(7) = "hpsysdrv"   ' HP system driver (Task mgr name)
    astrData(8) = "ca iss"     ' CA anti-virus pgms (Partial caption title)
    astrData(9) = "ccevtmgr"   ' CA anti-virus event mgr (Task mgr name)
    
    ' Loop thru array
    For lngIndex = 0 To UBound(astrData) - 1
            
        ' See if any more data in array
        DoEvents
        If Len(Trim$(astrData(lngIndex))) = 0 Then
            Exit For    ' exit For..Next loop
        End If
             
        ' Search for process by caption title
        StopProcessByName astrData(lngIndex)
        DoEvents
        
    Next lngIndex
    
    Erase astrData()  ' Always empty arrays when not needed
    
End Sub

' ***************************************************************************
' Routine:       StopProcessByHandle
'
' Description:   This routine is used to stop a specific processes if window
'                handle is known.
'
' Parameters:    lngHwnd - Unique handle designating process to be closed
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub StopProcessByHandle(ByVal lngHwnd As Long)

    ' Stop a single process
    
    If Val(lngHwnd) < 1 Then
        InfoMsg "Incoming process handle is missing." & _
                vbNewLine & vbNewLine & MODULE_NAME & "StopProcessByHandle"
        Exit Sub
    End If

    ' Search for a specific process handle
    lngHwnd = FindProcessByHandle(lngHwnd)
    
    DoEvents
    If lngHwnd > 0 Then
            
        ProcessTerminate mlngProcessID   ' Close process ID handle (Task mgr name)
        ProcessTerminate lngHwnd         ' Close process window handle
    
    End If
    
End Sub

' ***************************************************************************
' Routine:       StopProcessByName
'
' Description:   This routine is used to stop a specific processes.  All
'                occurances of processes with the same data in the caption
'                will be identified and closed.
'
' Parameters:    strCaption - Full or partial caption data to search for
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 07-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Sub StopProcessByName(ByVal strCaption As String)

    ' Stop a single process by caption name

    Dim lngHwnd  As Long
    Dim lngIndex As Long

    Erase matypHandles()   ' Always start with empty arrays

    If Len(Trim$(strCaption)) > 0 Then

        ' Search for process by caption title
        lngHwnd = FindProcessByCaption(strCaption)
        
        If lngHwnd > 0 Then
        
            For lngIndex = 0 To UBound(matypHandles) - 1
            
                ' If multiple occurances of
                ' this application are active
                ' then close all of them
                With matypHandles(lngIndex)
                    lngHwnd = .Handle
                    mlngProcessID = .ProcessID
                End With
                
                ProcessTerminate mlngProcessID   ' Close process ID handle (Task mgr name)
                ProcessTerminate lngHwnd         ' Close process window handle
                                                    
            Next lngIndex
            
        End If
    End If

    Erase matypHandles()   ' Always empty arrays when not needed

End Sub

' ***************************************************************************
' Routine:       FindProcessByCaption
'
' Description:   This routine will set an external search flag to FALSE and
'                perform an enumeration of all active programs, either hidden,
'                minimized, or displayed.
'
' Parameters:    strCaption - partial/full name of caption title
'
' Returns:       Found it  - Returns first process handle
'                Not found - Returns zero
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 07-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function FindProcessByCaption(ByVal strCaption As String) As Long

    mlngCount = 0
    mlngProcessID = 0
    mblnFoundit = False
    
    Erase matypHandles()   ' Always start with empty arrays

    If Len(Trim$(strCaption)) > 0 Then

        ReDim matypHandles(1)   ' Size array to one entry
        
        mstrPartialTitle = strCaption
        
        ' Search all active applications
        EnumWindows AddressOf FindProcessCaption, &H0
        
        If mblnFoundit Then
            FindProcessByCaption = matypHandles(0).Handle   ' Return first process handle
        Else
            Erase matypHandles()       ' Empty array
            FindProcessByCaption = 0   ' Process not found
        End If
    
    End If

End Function

' ***************************************************************************
' Routine:       FindProcessByHandle
'
' Description:   This routine will set an external search flag to FALSE and
'                perform an enumeration of all active programs, either hidden,
'                minimized, or displayed.
'
' Parameters:    lngHwnd - Specific process handle to search for
'
' Returns:       Found it  - Returns first process handle
'                Not found - Returns zero
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function FindProcessByHandle(ByVal lngHwnd As Long) As Long

    mlngHwnd = lngHwnd   ' Init variables
    mlngProcessID = 0
    mblnFoundit = False
    
    ' Search all active applications
    EnumWindows AddressOf FindProcessHandle, &H0
    
    If mblnFoundit Then
        FindProcessByHandle = mlngHwnd  ' Return process handle
    Else
        FindProcessByHandle = 0         ' Process not found
    End If
    
End Function

' ***************************************************************************
' Routine:       ProcessTerminate
'
' Description:   The following Function terminates a process, given a
'                process ID Or the windows handle of a form owned by the
'                process. This has the same effect As pressing End Task
'                In Task Mananger.
'
'                In WIN NT, click the "Processes" tab in the "Task Manager"
'                to see the process ID (or PID) for an application.
'                Must specify either lngHwnd or lngProcessID.  Equivalent
'                to pressing Alt+Ctrl+Del then "End Task"

' Parameters:   lngProcessID - Optional - process ID (or PID) to terminate
'               lngHwnd - Optional - application handle designating parent
'                         process
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Apr-2001  Andrew Baker
'              http://en.allexperts.com/q/Visual-Basic-1048/Kill-Process-VB-its-1.htm
' 05-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function ProcessTerminate(Optional ByVal lngProcessID As Long = 0, _
                                 Optional ByVal lngHwnd As Long = 0) As Boolean
                                 
    Dim lngRetVal       As Long
    Dim lngExitCode     As Long
    Dim lngTokenHwnd    As Long
    Dim lngProcessHwnd  As Long
    Dim lngThisProcHwnd As Long
    Dim lngBufferNeeded As Long
    Dim typLUID         As LUID
    Dim typTokenPriv    As TOKEN_PRIVILEGES
    Dim typTokenPrivNew As TOKEN_PRIVILEGES
    
    Const SE_DEBUG_NAME           As String = "SeDebugPrivilege"
    Const TOKEN_QUERY             As Long = &H8
    Const PROCESS_TERMINATE       As Long = &H1
    Const SE_PRIVILEGE_ENABLED    As Long = &H2
    Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
    
    On Error Resume Next
    
    ' Empty type structures
    ZeroMemory typLUID, Len(typLUID)
    ZeroMemory typTokenPriv, Len(typTokenPriv)
    ZeroMemory typTokenPrivNew, Len(typTokenPrivNew)
        
    If lngHwnd > 0 Then
        
        ' Get process ID from the window handle
        lngRetVal = GetWindowThreadProcessId(lngHwnd, lngProcessID)
        
    End If
   
    If lngProcessID > 0 Then
        
        ' Give Kill permissions to this process
        lngThisProcHwnd = GetCurrentProcess
        
        OpenProcessToken lngThisProcHwnd, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lngTokenHwnd
        LookupPrivilegeValue "", SE_DEBUG_NAME, typLUID
        
        ' Set the number of privileges to be change
        typTokenPriv.PrivilegeCount = 1
        typTokenPriv.TheLuid = typLUID
        typTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        
        ' Enable the kill privilege in the access token of this process
        AdjustTokenPrivileges lngTokenHwnd, False, typTokenPriv, Len(typTokenPrivNew), typTokenPrivNew, lngBufferNeeded
        
        ' Open the process to kill
        lngProcessHwnd = OpenProcess(PROCESS_TERMINATE, 0, lngProcessID)
        
        If lngProcessHwnd Then
            
            ' Obtained process handle, kill the process
            ProcessTerminate = CBool(TerminateProcess(lngProcessHwnd, lngExitCode))
            
            Call CloseHandle(lngProcessHwnd)  ' Always close handles when not needed
            
        End If
    
    End If
    
    ' Empty type structures
    ZeroMemory typLUID, Len(typLUID)
    ZeroMemory typTokenPriv, Len(typTokenPriv)
    ZeroMemory typTokenPrivNew, Len(typTokenPrivNew)
        
    On Error GoTo 0

End Function


' ***************************************************************************
' ****               Internal procedures & functions                     ****
' ***************************************************************************

' ***************************************************************************
' Routine:       FindProcessHandle
'
' Description:   This routine will search ALL active programs running under
'                Windows, including the hidden and minimized.  It will
'                look for a specific parent handle.
'
' Parameters:    lngHwnd - Generic application handle to check all
'                          active programs
'                lngNotUsed - Not used (but required for callbacks)
'
' Returns:       Sets an external flag to TRUE/FALSE based on the findings.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Function FindProcessHandle(ByVal lngHwnd As Long, _
                          Optional ByVal lngNotUsed As Long = 0) As Long

    Dim lngLength    As Long
    Dim lngRetCode   As Long
    Dim strFullTitle As String
    
    strFullTitle = Space$(MAX_SIZE)                              ' Preload with spaces not nulls
    lngLength = GetWindowText(lngHwnd, strFullTitle, MAX_SIZE)   ' Capture process title
        
    ' Is this our process handle?
    If mlngHwnd = lngHwnd Then
            
        ' Save complete title.  May be referenced by future application.
        If lngLength > 0 Then
            strFullTitle = Left$(strFullTitle, lngLength)    ' Format process title
        End If
        
        lngRetCode = GetParent(lngHwnd)  ' Get parent application handle
        
        ' If this is a child process, the return
        ' value will be a parent process handle
        If lngRetCode > 0 Then
            lngHwnd = lngRetCode  ' Found parent window
        End If
            
        ' Get process identifier
        lngRetCode = GetWindowThreadProcessId(lngHwnd, mlngProcessID)
        
        mblnFoundit = True     ' Set boolean flag
        FindProcessHandle = 0  ' No more searching
    Else
        FindProcessHandle = 1  ' Set flag for another interation
    End If

    CloseHandle lngHwnd   ' Close current handle

End Function

' ***************************************************************************
' Routine:       FindProcessCaption
'
' Description:   This routine will search ALL active programs running under
'                Windows, including the hidden and minimized.  It will
'                look for all parent names with the same partial/full caption
'                title.
'
' Parameters:    lngHwnd - Generic application handle to check all
'                          active programs
'                lngNotUsed - Not used (but required for callbacks)
'
' Returns:       Sets an external flag to TRUE/FALSE based on the findings.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Function FindProcessCaption(ByVal lngHwnd As Long, _
                           Optional ByVal lngNotUsed As Long = 0) As Long

    Dim lngLength    As Long
    Dim lngRetCode   As Long
    Dim strFullTitle As String
    
    strFullTitle = Space$(MAX_SIZE)   ' Preload with spaces not nulls
    
    lngLength = GetWindowText(lngHwnd, strFullTitle, MAX_SIZE)   ' Capture process title
        
    If lngLength > 0 Then
    
        strFullTitle = Left$(strFullTitle, lngLength)   ' Format process title
        
        ' See if this is our process title. Since we may only have a
        ' partial title, then we have to do an INSTR comparison.
        If InStr(1, strFullTitle, mstrPartialTitle, vbTextCompare) > 0 Then
                        
            lngRetCode = GetParent(lngHwnd)  ' Get parent application handle
            
            ' If this is a child process, the return
            ' value will be a parent process handle
            If lngRetCode > 0 Then
                lngHwnd = lngRetCode   ' Found parent window
                lngRetCode = 0         ' Reset to zero
            End If
            
            ' Get Process ID and Thread ID
            lngRetCode = GetWindowThreadProcessId(lngHwnd, mlngProcessID)
            
            ' Save pertinent data
            With matypHandles(mlngCount)
                .Handle = lngHwnd
                .ProcessID = mlngProcessID
                .Caption = mstrPartialTitle   ' This entry used for debugging only
            End With
            
            mlngCount = mlngCount + 1               ' Increment index counter
            ReDim Preserve matypHandles(mlngCount)  ' Increase array size by one
            mblnFoundit = True                      ' Set boolean flag
        
        End If
    End If

    FindProcessCaption = 1  ' Set flag for another interation
    
    CloseHandle lngHwnd     ' Always close current handle

End Function

