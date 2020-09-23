Attribute VB_Name = "modManifest"
' ***************************************************************************
'  Module:      modManifest.bas
'
'  Purpose:     This module contains routines designed to provide standard
'               formatting for message boxes.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added IsPathValid() routine.
' 28-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic in IsWinXPorNewer() routine
' 31-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated error trap by clearing error number in InitComctl32()
'              routine.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const VER_PLATFORM_WIN32_NT  As Long = 2
  Private Const FILE_ATTRIBUTE_HIDDEN  As Long = &H2&
  Private Const FILE_ATTRIBUTE_NORMAL  As Long = &H80&
  
  ' Set of bit flags that indicate which common
  ' control classes will be loaded from the DLL.
  ' The dwICC value of INIT_COMMON_CTRLS can be
  ' a combination of the following:
  Private Const ICC_ANIMATE_CLASS      As Long = &H80&     ' Load animate control class
  Private Const ICC_BAR_CLASSES        As Long = &H4&      ' Load toolbar, status bar, trackbar, tooltip control classes
  Private Const ICC_COOL_CLASSES       As Long = &H400&    ' Load rebar control class
  Private Const ICC_DATE_CLASSES       As Long = &H100&    ' Load date and time picker control class
  Private Const ICC_HOTKEY_CLASS       As Long = &H40&     ' Load hot key control class
  Private Const ICC_INTERNET_CLASSES   As Long = &H800&    ' Load IP address class
  Private Const ICC_LINK_CLASS         As Long = &H8000&   ' Load a hyperlink control class. Must have trailing ampersand.
  Private Const ICC_LISTVIEW_CLASSES   As Long = &H1&      ' Load list-view and header control classes
  Private Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&   ' Load a native font control class
  Private Const ICC_PAGESCROLLER_CLASS As Long = &H1000&   ' Load pager control class
  Private Const ICC_PROGRESS_CLASS     As Long = &H20&     ' Load progress bar control class
  Private Const ICC_STANDARD_CLASSES   As Long = &H4000&   ' Load user controls that include button, edit, static, listbox,
                                                           '      combobox, scrollbar
  Private Const ICC_TREEVIEW_CLASSES   As Long = &H2&      ' Load tree-view and tooltip control classes
  Private Const ICC_TAB_CLASSES        As Long = &H8&      ' Load tab and tooltip control classes
  Private Const ICC_UPDOWN_CLASS       As Long = &H10&     ' Load up-down control class
  Private Const ICC_USEREX_CLASSES     As Long = &H200&    ' Load ComboBoxEx class
  Private Const ICC_WIN95_CLASSES      As Long = &HFF&     ' Load animate control, header, hot key, list-view, progress bar,
                                                           '      status bar, tab, tooltip, toolbar, trackbar, tree-view,
                                                           '      and up-down control classes

  ' All bit flags combined. Total value = &HFFFF& (65535)
  Private Const ALL_FLAGS As Long = ICC_ANIMATE_CLASS Or ICC_BAR_CLASSES Or ICC_COOL_CLASSES Or _
                                    ICC_DATE_CLASSES Or ICC_HOTKEY_CLASS Or ICC_INTERNET_CLASSES Or _
                                    ICC_LINK_CLASS Or ICC_LISTVIEW_CLASSES Or ICC_NATIVEFNTCTL_CLASS Or _
                                    ICC_PAGESCROLLER_CLASS Or ICC_PROGRESS_CLASS Or ICC_STANDARD_CLASSES Or _
                                    ICC_TREEVIEW_CLASSES Or ICC_TAB_CLASSES Or ICC_UPDOWN_CLASS Or _
                                    ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES
                                                    
' ***************************************************************************
' Type structures
' ***************************************************************************
  ' The OSVERSIONINFOEX data structure contains operating system version
  ' information. The information includes major and minor version numbers,
  ' a build number, a platform identifier, and information about product
  ' suites and the latest Service Pack installed on the system. This structure
  ' is used with the GetVersionEx and VerifyVersionInfo functions.
  Private Type OSVERSIONINFOEX
      OSVSize           As Long           ' size of this data structure (in bytes)
      dwVerMajor        As Long           ' ex: 5
      dwVerMinor        As Long           ' ex: 01
      dwBuildNumber     As Long           ' ex: 2600
      PlatformID        As Long           ' Identifies operating system platform
      szCSDVersion      As String * 128   ' ex: "Service Pack 3"
      wServicePackMajor As Integer
      wServicePackMinor As Integer
      wSuiteMask        As Integer
      wProductType      As Byte
      wReserved         As Byte
  End Type

  ' Used with manifest files
  Private Type INIT_COMMON_CTRLS
      dwSize As Long   ' size of this structure
      dwICC  As Long   ' flags indicating which classes to be initialized
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long
  
  ' Initializes the entire common control dynamic-link library. Exported by
  ' all versions of Comctl32.dll.
  Private Declare Sub InitCommonControls Lib "comctl32" ()
  
  ' Initializes specific common controls classes from the common control
  ' dynamic-link library. Returns TRUE (non-zero) if successful, or FALSE
  ' otherwise. Began being exported with Comctl32.dll version 4.7
  ' (IE3.0 & later).
  Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
          (iccex As INIT_COMMON_CTRLS) As Boolean

  ' This function obtains extended information about the version of the
  ' operating system that is currently running.
  Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
          (LpVersionInformation As Any) As Long

  ' SetFileAttributes Function sets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Private Declare Function SetFileAttributes Lib "kernel32" _
          Alias "SetFileAttributesA" _
          (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

  ' ZeroMemory is used for clearing contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)


' ***************************************************************************
' Routine:       InitComctl32
'
' Description:   This will create the XP Manifest file and utilize it. You
'                will only see the results when the exe (not in the IDE)
'                is run.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jan-2006  Randy Birch
'              http://vbnet.mvps.org/
' 03-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 31-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated error trap by clearing error number
' ***************************************************************************
Public Sub InitComctl32()

    Dim typICC As INIT_COMMON_CTRLS
    
    CreateManifestFile
         
    On Error GoTo Use_Old_Version
         
    With typICC
        .dwSize = LenB(typICC)
        .dwICC = ALL_FLAGS
    End With
    
    ' VB will generate error 453 "Specified DLL function not found"
    ' if InitCommonControlsEx can't be located in the library. The
    ' error is trapped and the original InitCommonControls is called
    ' instead below.
    If InitCommonControlsEx(typICC) = 0 Then
        InitCommonControls
    End If
    
    On Error GoTo 0
    Exit Sub
    
Use_Old_Version:
    Err.Clear
    InitCommonControls
    On Error GoTo 0
    
End Sub


' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CreateManifestFile
'
' Description:   If this is Windows XP and the manifest file does not exist
'                then one will be created.  If this is not Windows XP and
'                the manifest file exist, it will be deleted.
'
' Parameters:    None.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jan-2006  Randy Birch
'              http://vbnet.mvps.org/
' 03-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub CreateManifestFile()
    
    Dim hFile       As Long
    Dim strXML      As String
    Dim strCompany  As String
    Dim strExeName  As String
    Dim strFilename As String
    
    On Error Resume Next

    strXML = vbNullString
    strCompany = "Kens.Software."   ' Enter unique name here. Periods are delimiters.
    strExeName = App.EXEName        ' EXE name without an extension
    strFilename = QualifyPath(App.Path) & strExeName & ".exe.manifest"
        
    ' If this is Windows XP or newer and
    ' if the manifest file does not exist
    ' then create it and shutdown this
    ' application.
    If IsWinXPorNewer Then
        
        ' Checks if the manifest has already been created
        If IsPathValid(strFilename) Then
            Exit Sub
        Else
            ' Create the manifest file
            strXML = strXML & "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & _
                              "UTF-8" & Chr$(34) & " standalone=" & Chr$(34) & "yes" & Chr$(34) & "?>"
            strXML = strXML & vbNewLine & "<assembly xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v1" & _
                              Chr$(34) & " manifestVersion=" & Chr$(34) & "1.0" & Chr$(34) & ">"
            strXML = strXML & vbNewLine & "  <assemblyIdentity"
            strXML = strXML & vbNewLine & "    version=" & Chr$(34) & "1.0.0.0" & Chr$(34)
            strXML = strXML & vbNewLine & "    processorArchitecture=" & Chr$(34) & "X86" & Chr$(34)
            strXML = strXML & vbNewLine & "    name=" & Chr$(34) & strCompany & strExeName & Chr$(34)
            strXML = strXML & vbNewLine & "    type=" & Chr$(34) & "win32" & Chr$(34)
            strXML = strXML & vbNewLine & "  />"
            strXML = strXML & vbNewLine & "  <description>" & strCompany & strExeName & "</description>"
            strXML = strXML & vbNewLine & "  <dependency>"
            strXML = strXML & vbNewLine & "    <dependentAssembly>"
            strXML = strXML & vbNewLine & "      <assemblyIdentity"
            strXML = strXML & vbNewLine & "        type=" & Chr$(34) & "win32" & Chr$(34)
            strXML = strXML & vbNewLine & "        name=" & Chr$(34) & "Microsoft.Windows.Common-Controls" & Chr$(34)
            strXML = strXML & vbNewLine & "        version=" & Chr$(34) & "6.0.0.0" & Chr$(34)
            strXML = strXML & vbNewLine & "        processorArchitecture=" & Chr$(34) & "X86" & Chr$(34)
            strXML = strXML & vbNewLine & "        publicKeyToken=" & Chr$(34) & "6595b64144ccf1df" & Chr$(34)
            strXML = strXML & vbNewLine & "        language=" & Chr$(34) & "*" & Chr$(34)
            strXML = strXML & vbNewLine & "      />"
            strXML = strXML & vbNewLine & "    </dependentAssembly>"
            strXML = strXML & vbNewLine & "  </dependency>"
            strXML = strXML & vbNewLine & "</assembly>"
            
            hFile = FreeFile
            Open strFilename For Output As #hFile
            Print #hFile, strXML
            Close #hFile
                        
            ' reset file attributes to hidden
            SetFileAttributes strFilename, FILE_ATTRIBUTE_HIDDEN
                    
            ' Display shutdown message
            InfoMsg "Manifest file has been re-initialized." & vbNewLine & vbNewLine & _
                    "This application must be restarted."
            TerminateProgram   ' shutdown this application
        End If
        
    Else
        ' If this is not Windows XP or newer and
        ' if the manifest file does exist then
        ' delete the file because it is not needed.
        If IsPathValid(strFilename) Then
            
            ' reset file attributes
            SetFileAttributes strFilename, FILE_ATTRIBUTE_NORMAL
            Kill strFilename
            
        End If
    End If
    
    On Error GoTo 0
    
End Sub

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Determines whether a path to a file system object such as
'                a file or directory is valid. This function tests the
'                validity of the path. A path specified by Universal Naming
'                Convention (UNC) is limited to a file only; that is,
'                \\server\share\file is permitted. A UNC path to a server
'                or server share is not permitted; that is, \\server or
'                \\server\share. This function returns FALSE if a mounted
'                remote drive is out of service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'
' Reference:     http://msdn.microsoft.com/en-us/library/bb773584(v=vs.85).aspx
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))
   
End Function
  
' ***************************************************************************
' Routine:       IsWinXPorNewer
'
' Description:   Test to see if operating system is Windows XP or newer.
'
' Returns:       TRUE - Operating system is Windows XP or newer
'                FALSE - Earlier version of Windows
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jan-2006  Randy Birch
'              http://vbnet.mvps.org/
' 28-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified logic selection
' ***************************************************************************
Private Function IsWinXPorNewer() As Boolean

    Dim typOSVIEX As OSVERSIONINFOEX
    
    IsWinXPorNewer = False                ' Preset flag to FALSE
    ZeroMemory typOSVIEX, Len(typOSVIEX)  ' Clear type structure
    typOSVIEX.OSVSize = Len(typOSVIEX)    ' Size of type structure (Required)
    
    ' Capture type of operating system
    If GetVersionEx(typOSVIEX) <> 0 Then
        
        With typOSVIEX
            If .PlatformID = VER_PLATFORM_WIN32_NT Then
                If (.dwVerMajor = 5 And .dwVerMinor >= 1) Or _
                   (.dwVerMajor > 5) Then
                   
                    IsWinXPorNewer = True
                End If
            End If
        End With
        
    End If

    ZeroMemory typOSVIEX, Len(typOSVIEX)  ' Clear type structure

End Function

