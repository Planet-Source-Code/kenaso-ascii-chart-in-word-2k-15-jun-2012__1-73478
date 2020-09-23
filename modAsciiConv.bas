Attribute VB_Name = "modAsciiConv"
' ***************************************************************************
' Routine:       modAsciiConv
'
' Description:   Create an ASCII conversion chart to be inserted into
'                MS Word 2000.  The modics are here for an easy transition
'                to other versions of MS Word.
'
' ASCII conversion chart created for Microsoft Word 2000
' by Kenneth Ives  kenaso@tx.rr.com
'
' *** IMPORTANT ***
' Must make a reference to "Microsoft Word n.0 object Library"
' Where "n" represents the version number.
' In the VB IDE, select Project>>References...
'
'      or
'
' In the VB IDE, select Project>>References...
' Select Browse... button
' Navigate to MS Office folder:
' Ex:   Office 2000   C:\Program Files\Microsoft Office\Office\MSWord9.olb
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const FMT_03      As String = "@@@"
  Private Const FMT_04      As String = "@@@@"
  Private Const FMT_05      As String = "@@@@@"
  Private Const FMT_11      As String = "@@@@@@@@@@@"
  Private Const FMT_12      As String = "@@@@@@@@@@@@"
  Private Const MAX_BYTE    As Long = 256

' ***************************************************************************
' Type Structures
' ***************************************************************************
  Private Type CONVERTED_DATA
      Decimal As String
      Hex     As String
      Binary  As String
      Symbol  As String
  End Type
  
' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Object)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m obj Word
' Variable name:     mobjWord
'
' ***************************************************************************
  Private mobjWord      As Word.Application
  Private mobjWordDoc   As Word.Document
  Private mstrTimeStamp As String
    
Public Function BuildChart(ByVal strCustomData As String, _
                           ByVal blnOnePage As Boolean) As String

    ' Called by frmChart.StartProcessing()
    
    Dim avntSymbol   As Variant
    Dim lngIndex     As Long
    Dim strOutput    As String
    Dim astrBinary() As String
    Dim atypCD()     As CONVERTED_DATA
  
    avntSymbol = Empty   ' always start with empty variants
    Erase atypCD()       ' always start with empty arrays
    Erase astrBinary()
    
    ReDim atypCD(MAX_BYTE)   ' Size type structure array
    strOutput = vbNullString           ' Verify string is empty
    mstrTimeStamp = vbNullString
    
    ' See if user wants time stamp on report
    If gblnDateOnChart Then
        mstrTimeStamp = Format$(Now(), "dd-MMM-yyyy") & "  " & _
                        Format$(Now(), "Long Time")
    End If
    
    If Len(Trim$(strCustomData)) = 0 Then
        strCustomData = vbNullString   ' Verify string is empty
    End If
    
    astrBinary() = LoadBinaryArray()   ' preload binary string array
    
    ' Fill array with first thirty-two multi-char symbols
    avntSymbol = Array("NUL", "SOH", "STX", "ETX", "EOT", "ENQ", "ACK", "BEL", _
                       "BS", "TAB", "LF", "VT", "FF", "CR", "SO", "SI", "DLE", _
                       "DC1", "DC2", "DC3", "DC4", "NAK", "SYN", "ETB", "CAN", _
                       "EM", "SUB", "ESC", "FS", "GS", "RS", "US")
    
    ' load conversion array
    For lngIndex = 0 To MAX_BYTE - 1
        
        With atypCD(lngIndex)
             .Decimal = CStr(lngIndex)                ' Numeric value
             .Hex = Right$("00" & Hex$(lngIndex), 2)  ' Hex string
             .Binary = astrBinary(lngIndex)           ' Binary string
             
             Select Case lngIndex
                    Case 0 To 31: .Symbol = Format$(avntSymbol(lngIndex), FMT_03)  ' Multi-char symbols
                    Case Else:    .Symbol = Format$(Chr$(lngIndex), FMT_03)        ' Single symbol
             End Select
        End With
        
    Next lngIndex
    
    avntSymbol = Empty   ' always empty variants when not needed
    Erase astrBinary()   ' always empty arrays when not needed
    
    ' Format and load data into MS Word
    If blnOnePage Then
        
        strOutput = OnePage(atypCD())            ' Format output string
        FormatOnePage strOutput, strCustomData   ' Portrait mode - 1 page
    
    Else
        
        strOutput = TwoPages(atypCD(), 0, True)                 ' Format page one
        strOutput = strOutput & TwoPages(atypCD(), 128, False)  ' Format page two
        FormatTwoPages strOutput, strCustomData                 ' Landscape mode - 2 pages
    
    End If
    
    Erase atypCD()     ' always empty arrays when not needed
    strOutput = vbNullString     ' Verify string is empty
    TerminateProgram   ' Terminate this application
    
End Function



' ***************************************************************************
' ****               Internal procedures & functions                     ****
' ***************************************************************************

Private Function LoadBinaryArray() As String()

    ' Called by BuildChart()
    
    Dim lngIndex     As Long     ' Index counter
    Dim avntBinary   As Variant  ' 0-F in binary
    Dim astrOutput() As String   ' hold formatted data

    avntBinary = Empty           ' Start with empty variants
    Erase astrOutput()           ' Start with empty arrays
    ReDim astrOutput(MAX_BYTE)   ' Size temp string array
    
    ' Load temp binary array
    avntBinary = Array("0000", "0001", "0010", "0011", _
                       "0100", "0101", "0110", "0111", _
                       "1000", "1001", "1010", "1011", _
                       "1100", "1101", "1110", "1111")

    ' Preload 8-bit binary string array
    ' Ex:  astrOutput(0) = "0000 0000"
    '         ...
    '      astrOutput(255) = "1111 1111"
    For lngIndex = 0 To MAX_BYTE - 1
        astrOutput(lngIndex) = avntBinary(lngIndex \ &H10) & " " & _
                               avntBinary(lngIndex And &HF)
    Next lngIndex
        
    LoadBinaryArray = astrOutput()   ' Return binary array
    
    avntBinary = Empty  ' always empty variants when not needed
    Erase astrOutput()  ' always empty arrays when not needed
    
End Function

Private Function OnePage(ByRef atypCD() As CONVERTED_DATA) As String

    ' Called by BuildChart()
    
    Dim lngLoop    As Long
    Dim lngIndex   As Long
    Dim lngStart   As Long
    Dim strTemp    As String
    Dim strOutput  As String
    Dim strHeading As String
    
    Const HEADING1  As String = "DEC  HEX   BINARY    SYM"
    Const HEADING2  As String = "DEC  HEX   BINARY   SYM"
    
    strOutput = vbNullString   ' Verify strings are empty
    strHeading = vbNullString
    
    ' Format column headings
    strHeading = HEADING1 & Space$(5) & HEADING1 & Space$(5) & _
                 HEADING2 & Space$(6) & HEADING2 & vbNewLine
    
    ' prepare the column headings
    strOutput = String$(110, 45) & vbNewLine                ' dashed line
    strOutput = strOutput & strHeading                   ' heading titles
    strOutput = strOutput & String$(110, 45) & vbNewLine    ' dashed line
              
    lngStart = 0   ' Initialize the array starting position
              
    ' load the output string for all four columns
    For lngLoop = 1 To 4
    
        For lngIndex = lngStart To (lngStart + 15)
            
            strTemp = vbNullString

            ' load first column
            With atypCD(lngIndex)
                 strTemp = Format$(.Decimal, FMT_03)
                 strTemp = strTemp & Format$(.Hex, FMT_05)
                 strTemp = strTemp & Format$(.Binary, FMT_11)
                                                  
                 Select Case lngIndex
                        Case 0 To 15: strTemp = strTemp & Format$(.Symbol, FMT_05) & Space$(5)
                        Case Else:    strTemp = strTemp & Format$(.Symbol, FMT_04) & Space$(6)
                 End Select
            End With
            
            ' load second column
            With atypCD(lngIndex + 16)
                 
                 strTemp = strTemp & Format$(.Decimal, FMT_03)
                 strTemp = strTemp & Format$(.Hex, FMT_05)
                 strTemp = strTemp & Format$(.Binary, FMT_11)
                 
                 Select Case (lngIndex + 16)
                        Case 16 To 31: strTemp = strTemp & Format$(.Symbol, FMT_05) & Space$(5)
                        Case Else:     strTemp = strTemp & Format$(.Symbol, FMT_04) & Space$(6)
                 End Select
            End With
            
            ' load third column
            With atypCD(lngIndex + 32)
                 strTemp = strTemp & Format$(.Decimal, FMT_03)
                 strTemp = strTemp & Format$(.Hex, FMT_05)
                 strTemp = strTemp & Format$(.Binary, FMT_11)
                 strTemp = strTemp & Format$(.Symbol, FMT_03) & Space$(7)
            End With
            
            ' load fourth column
            With atypCD(lngIndex + 48)
                 strTemp = strTemp & Format$(.Decimal, FMT_03)
                 strTemp = strTemp & Format$(.Hex, FMT_05)
                 strTemp = strTemp & Format$(.Binary, FMT_11)
                 strTemp = strTemp & Format$(.Symbol, FMT_03) & vbNewLine
            End With
                      
            ' insert blank line between groups of 16
            If lngIndex > 0 And lngIndex < 255 Then
                Select Case lngIndex
                       Case 15, 79, 143
                            strTemp = strTemp & vbNewLine & vbNewLine
                End Select
            End If
            
            ' append formatted data to output string
            strOutput = strOutput & strTemp
            
        Next lngIndex
            
        lngStart = lngStart + 64   ' set new starting position
            
    Next lngLoop
    
    OnePage = strOutput   ' return formatted data
    
    strTemp = vbNullString    ' Verify strings are empty
    strOutput = vbNullString
    strHeading = vbNullString
    
End Function

Private Function TwoPages(ByRef atypCD() As CONVERTED_DATA, _
                          ByVal lngStart As Long, _
                          ByVal blnPageOne As Boolean) As String

    ' Called by BuildChart()
    
    Dim lngLoop      As Long
    Dim lngIndex     As Long
    Dim strTemp      As String
    Dim strOutput    As String
    Dim strHeading1  As String
    Dim strHeading2  As String
    Dim blnFirstHalf As Boolean
    
    Const HEADING1 As String = "DEC  HEX    BINARY    SYM"
    Const HEADING2 As String = "DEC  HEX    BINARY   SYM"
    
    blnFirstHalf = True
    strOutput = vbNullString   ' Verify strings are empty
    strOutput = String$(117, 45) & vbNewLine   ' dashed line
    
    ' Format column headings
    strHeading1 = HEADING1 & Space$(6) & _
                  HEADING1 & Space$(6) & _
                  HEADING2 & Space$(7) & _
                  HEADING2 & vbNewLine
                  
    strHeading2 = HEADING2 & Space$(7) & _
                  HEADING2 & Space$(7) & _
                  HEADING2 & Space$(7) & _
                  HEADING2 & vbNewLine
    
    ' Append column titles
    If blnPageOne Then
        strOutput = strOutput & strHeading1
    Else
        strOutput = strOutput & strHeading2
    End If
        
    strOutput = strOutput & String$(117, 45) & vbNewLine  ' dashed line
        
    ' load the output string for all four columns
    For lngLoop = 1 To 2
                          
        For lngIndex = lngStart To (lngStart + 15)
        
            strTemp = vbNullString
        
            ' load first column
            With atypCD(lngIndex)
                 strTemp = Format$(.Decimal, FMT_03)
                 strTemp = strTemp & Format$(.Hex, FMT_05)
                 
                 If lngIndex < 32 Then
                     strTemp = strTemp & Format$(.Binary, FMT_12)
                     strTemp = strTemp & Format$(.Symbol, FMT_05) & Space$(6)
                 Else
                     strTemp = strTemp & Format$(.Binary, FMT_12)
                     strTemp = strTemp & Format$(.Symbol, FMT_03) & Space$(8)
                 End If
            End With
            
            ' load second column
            With atypCD(lngIndex + 16)
                 strTemp = strTemp & Format$(.Decimal, FMT_03)
                 strTemp = strTemp & Format$(.Hex, FMT_05)
                 
                 If lngIndex < 32 Then
                     strTemp = strTemp & Format$(.Binary, FMT_12)
                     strTemp = strTemp & Format$(.Symbol, FMT_05) & Space$(6)
                 Else
                     strTemp = strTemp & Format$(.Binary, FMT_12)
                     strTemp = strTemp & Format$(.Symbol, FMT_03) & Space$(8)
                 End If
            End With
            
            ' load third column
            With atypCD(lngIndex + 32)
                 strTemp = strTemp & Format$(.Decimal, FMT_03)
                 strTemp = strTemp & Format$(.Hex, FMT_05)
                 strTemp = strTemp & Format$(.Binary, FMT_12)
                 strTemp = strTemp & Format$(.Symbol, FMT_03) & Space$(8)
            End With
            
            ' load fourth column
            With atypCD(lngIndex + 48)
                 strTemp = strTemp & Format$(.Decimal, FMT_03)
                 strTemp = strTemp & Format$(.Hex, FMT_05)
                 strTemp = strTemp & Format$(.Binary, FMT_12)
                 strTemp = strTemp & Format$(.Symbol, FMT_03) & vbNewLine
            End With
                      
            strOutput = strOutput & strTemp
            
        Next lngIndex
        
        ' Prepare second half of the page
        If blnFirstHalf Then
            strTemp = String$(117, 45) & vbNewLine                ' Dashed line
            strTemp = strTemp & vbNewLine & vbNewLine & vbNewLine       ' three blank lines
            strOutput = strOutput & strTemp                    ' Append to output string
            strOutput = strOutput & String$(117, 45) & vbNewLine  ' Dashed line
            strOutput = strOutput & strHeading2                ' Append column heading
            strOutput = strOutput & String$(117, 45) & vbNewLine  ' dashed line
        End If
        
        lngStart = lngStart + 64
        blnFirstHalf = False
        
    Next lngLoop
    
    If blnPageOne Then
        strOutput = strOutput & Chr$(12)  ' insert a page break
    End If
    
    TwoPages = strOutput  ' return formatted data
    
    strTemp = vbNullString    ' Verify strings are empty
    strOutput = vbNullString

End Function

Private Sub FormatOnePage(ByVal strData As String, _
                 Optional ByVal strCustomData As String = vbNullString)

    ' Called by BuildChart()
    
    On Error GoTo FormatOnePage_CleanUp
    
    Set mobjWord = New Word.Application   ' start the WORD application
    
    ' if Word fails to start then leave
    If mobjWord Is Nothing Then
        GoTo FormatOnePage_CleanUp
    End If
           
    PageDesign strData, wdOrientPortrait  ' Design page layout
    AddHeader                             ' Insert page header
    AddFooter strCustomData               ' Insert custom data in footer
    
    ' Do not close Word.  Let user do that manually.
    mobjWordDoc.SaveAs "ASCII_Chart_1.doc"
    mobjWord.Application.Visible = True   ' Show word document
    
FormatOnePage_CleanUp:
    Set mobjWordDoc = Nothing   ' free all objects from memory
    Set mobjWord = Nothing
    
End Sub

Private Sub FormatTwoPages(ByVal strData As String, _
                  Optional ByVal strCustomData As String = vbNullString)

    ' Called by BuildChart()
    
    On Error GoTo FormatTwoPages_CleanUp
        
    Set mobjWord = New Word.Application   ' start the WORD application
    
    ' if Word fails to start then leave
    If mobjWord Is Nothing Then
        GoTo FormatTwoPages_CleanUp
    End If
    
    PageDesign strData, wdOrientLandscape  ' Design page layout
    AddHeader                              ' Add page header
    AddFooter strCustomData                ' Insert custom data in footer
    
    ' Do not close Word.  Let user do that manually.
    mobjWordDoc.SaveAs "ASCII_Chart_2.doc"
    mobjWord.Application.Visible = True   ' Show word document
    
FormatTwoPages_CleanUp:
    Set mobjWordDoc = Nothing   ' free all objects from memory
    Set mobjWord = Nothing
    
End Sub

'************************************************************************
' Most of the code in this routine was obtained by recording a macro in
' MS Word and then copying the code to here with some minor tweaking.
'************************************************************************
Private Sub PageDesign(ByVal strData As String, _
                       ByVal lngLayout As Long)
    
    ' Called by FormatOnePage()
    '           FormatTwoPages()
    
    Dim sngWidth    As Single
    Dim sngHeight   As Single
    Dim sngFontSize As Single
    
    ' Page layout
    Select Case lngLayout
    
           Case wdOrientPortrait    ' 0 - Portrait mode
                sngWidth = 8.5!
                sngHeight = 11!
                sngFontSize = 8!
           
           Case wdOrientLandscape   ' 1 - Landscape mode
                sngWidth = 11!
                sngHeight = 8.5!
                sngFontSize = 10!
    End Select
    
    mobjWord.Application.Visible = False       ' hide Word application
    Set mobjWordDoc = mobjWord.Documents.Add   ' create new document
            
    ' Format main body of the document
    With mobjWord
        With .Selection
            .WholeStory                  ' select whole document (CTRL+A)
            .Font.Name = "Courier New"   ' font name
            .Font.Size = sngFontSize     ' font size
            .Font.Bold = True            ' bold output
        End With
        
        ' determine page setup
        With .ActiveDocument.PageSetup
            .LineNumbering.Active = False
            .Orientation = lngLayout                    ' Page layout
            .TopMargin = InchesToPoints(0.3)
            .BottomMargin = InchesToPoints(0.6)
            .LeftMargin = InchesToPoints(0.5)
            .RightMargin = InchesToPoints(0.5)
            .Gutter = InchesToPoints(0)
            .HeaderDistance = InchesToPoints(0.3)
            .FooterDistance = InchesToPoints(0.6)
            .PageWidth = InchesToPoints(sngWidth)       ' Page width
            .PageHeight = InchesToPoints(sngHeight)     ' Page height
            .FirstPageTray = wdPrinterDefaultBin
            .OtherPagesTray = wdPrinterDefaultBin
            .SectionStart = wdSectionNewPage
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .VerticalAlignment = wdAlignVerticalCenter  ' Center top to bottom on page
            .SuppressEndnotes = False
            .MirrorMargins = False
            .TwoPagesOnOne = False
            .GutterPos = wdGutterPosLeft
        End With
        
        With .ActiveWindow
            ' Force to "Print View" window style
            If .View.SplitSpecial = wdPaneNone Then
                .ActivePane.View.Type = wdPrintView
            Else
                .View.Type = wdPrintView
            End If
                
            ' return to main document window
            .ActivePane.View.SeekView = wdSeekMainDocument
        End With
        
        With .Selection
            .Collapse                                          ' Release CTRL+A
            .ParagraphFormat.Alignment = wdAlignParagraphLeft  ' Left align all data for this document
            .Text = strData                                    ' Insert data into main document
            .HomeKey Unit:=wdStory                             ' Jump to top of page
        End With
    End With
    
End Sub

'************************************************************************
' VBA macro examples to insert text into a Word 2000 document
' http://support.microsoft.com/kb/212682
'************************************************************************
Private Sub AddHeader()

    ' Called by FormatOnePage()
    '           FormatTwoPages()
    
    With mobjWord
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        
        With .Selection
            .HeaderFooter.Range.Text = "ASCII Conversion Chart"   ' Insert header text
            .HomeKey Unit:=wdLine                                 ' Draw a line in header area
            .EndKey Unit:=wdLine, Extend:=wdExtend                ' Extend line to width of page
            .Font.Name = "Times New Roman"                        ' Set font attributes
            .Font.Size = 20!
            .Font.Bold = wdToggle
            .ParagraphFormat.Alignment = wdAlignParagraphCenter   ' Center data on page
        End With
        
        With .ActiveWindow
            ' Force to "Print View" window style
            If .View.SplitSpecial = wdPaneNone Then
                .ActivePane.View.Type = wdPrintView
            Else
                .View.Type = wdPrintView
            End If
                
            ' return to main document window
            .ActivePane.View.SeekView = wdSeekMainDocument
        End With
    End With

End Sub

'************************************************************************
' VBA macro examples to insert text into a Word 2000 document
' http://support.microsoft.com/kb/212682
'************************************************************************
Private Sub AddFooter(Optional ByVal strCustomData As String = vbNullString)

    ' Called by FormatOnePage()
    '           FormatTwoPages()
    
    If Len(Trim$(strCustomData)) = 0 Then
        
        ' Need time stamp on report
        If gblnDateOnChart Then
            strCustomData = mstrTimeStamp   ' Time stamp only
        End If
    Else
    
        ' There is data on the optional footer
        ' line and user also wants a time stamp
        If gblnDateOnChart Then
            strCustomData = strCustomData & Chr$(13) & mstrTimeStamp
        End If
    End If
    
    ' Format footer
    With mobjWord
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
        
        With .Selection
            .HeaderFooter.Range.Text = strCustomData             ' Insert optional footer data
            .Font.Name = "Times New Roman"                       ' Set font attributes
            .Font.Size = 10!
            .Font.Bold = wdToggle
            .HomeKey Unit:=wdLine                                ' Draw a line
            .EndKey Unit:=wdLine, Extend:=wdExtend               ' Extend line to width of page
            .ParagraphFormat.Alignment = wdAlignParagraphCenter  ' Center data on page
            .Borders(wdBorderTop).LineStyle = Options.DefaultBorderLineStyle
            .Borders(wdBorderTop).LineWidth = Options.DefaultBorderLineWidth
            .Borders(wdBorderTop).Color = Options.DefaultBorderColor
        End With
        
        With .ActiveWindow
            ' Force to "Print View" window style
            If .View.SplitSpecial = wdPaneNone Then
                .ActivePane.View.Type = wdPrintView
            Else
                .View.Type = wdPrintView
            End If
                
            ' return to main document window
            .ActivePane.View.SeekView = wdSeekMainDocument
        End With
    End With

End Sub


