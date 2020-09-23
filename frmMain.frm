VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   6345
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6345
   Begin VB.TextBox txtCustomData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      MaxLength       =   55
      TabIndex        =   2
      Text            =   "txtCustomData"
      Top             =   1710
      Width           =   6135
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   650
      Index           =   1
      Left            =   5580
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2130
      Width           =   650
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   650
      Index           =   0
      Left            =   4800
      Picture         =   "frmMain.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2130
      Width           =   650
   End
   Begin VB.PictureBox picManifest 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   765
      ScaleHeight     =   615
      ScaleWidth      =   5100
      TabIndex        =   5
      Top             =   720
      Width           =   5100
      Begin VB.CheckBox chkDate 
         Caption         =   "Show date at bottom of page"
         Height          =   465
         Left            =   2700
         TabIndex        =   11
         Top             =   45
         Width           =   2400
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Create a single page chart"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   0
         Top             =   45
         Value           =   -1  'True
         Width           =   2580
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Create a two page chart"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   45
         TabIndex        =   1
         Top             =   315
         Width           =   2580
      End
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   135
         TabIndex        =   6
         Top             =   855
         Width           =   3075
      End
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   180
      TabIndex        =   10
      Top             =   2340
      Width           =   3480
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter data to be inserted in footer of document (Max 55 characters)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   9
      Top             =   1440
      Width           =   5595
   End
   Begin VB.Label lblAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2565
      TabIndex        =   8
      Top             =   435
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ASCII Conversion Chart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1590
      TabIndex        =   7
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************
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
'     Office 2000   C:\Program Files\Microsoft Office\Office\MSWord9.olb
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' *************************************************************************
Option Explicit

' *************************************************************************
' Module variables
' *************************************************************************
  Private mblnOnePage As Boolean
  
Private Sub chkDate_Click()

    gblnDateOnChart = CBool(chkDate.Value)
    
End Sub

Private Sub cmdChoice_Click(Index As Integer)
    
    Select Case Index
           
           Case 0   ' Start processing
                DoEvents
                gstrOptTitle = txtCustomData.Text
                frmMain.Hide

                DoEvents
                frmChart.StartProcessing mblnOnePage, Trim$(Me.txtCustomData.Text)
           
           Case Else
                TerminateProgram
    End Select
    
End Sub

Private Sub Form_Load()
          
    optPrint_Click 0  ' Preset to single page
    
    With frmMain
        .Caption = PGM_NAME & gstrVersion
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                 "No warranties or guarantees implied or intended."
                                 
        .txtCustomData.Text = gstrOptTitle
        .chkDate.Value = IIf(gblnDateOnChart, vbChecked, vbUnchecked)
        
        ' Center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show
    End With
    
    txtCustomData.SetFocus
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        TerminateProgram
    End If
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub optPrint_Click(Index As Integer)
    mblnOnePage = Not CBool(Index)
End Sub

Private Sub txtCustomData_GotFocus()

    ' Highlight contents in text box
    With txtCustomData
         .SelStart = 0             ' start with first char in string
         .SelLength = Len(.Text)   ' to end of data
    End With
  
End Sub
