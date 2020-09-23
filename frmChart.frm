VERSION 5.00
Begin VB.Form frmChart 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2820
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   3900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmChart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00FF0000&
      Height          =   2940
      Left            =   -45
      ScaleHeight     =   2880
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   -45
      Width           =   3975
      Begin VB.Label lblAuthor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freeware by Kenneth Ives"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   840
         TabIndex        =   3
         Top             =   2295
         Width           =   2130
      End
      Begin VB.Label lblFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   720
         TabIndex        =   2
         Top             =   1305
         Width           =   2490
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Building ASCII Conversion Chart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2415
         Left            =   255
         TabIndex        =   1
         Top             =   203
         Width           =   3390
      End
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************
' ASCII conversion chart created for Microsoft Word 2000
' by Kenneth Ives  kenaso@tx.rr.com
'
' *** IMPORTANT ***
' In the VB IDE, select Project>>References...
' Must make a reference to "Microsoft Word n.0 object Library"
' Where "n" represents the version number.
'
'      or
'
' In the VB IDE, select Project >> References...
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

Private Sub Form_Load()
          
    With frmChart
        .Hide
        .lblFile.Caption = "File will be saved to " & vbNewLine & _
                           Chr$(34) & "My Documents" & Chr$(34) & " folder"
                           
        ' Center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Hide
    End With
    
    DoEvents
    
End Sub

Public Sub StartProcessing(ByVal blnOnePage As Boolean, _
                  Optional ByVal strCustomData As String = vbNullString)

    DoEvents
    frmChart.Show
    
    DoEvents
    If Len(Trim$(strCustomData)) = 0 Then
        strCustomData = vbNullString   ' verify string is empty
    End If
        
    BuildChart strCustomData, blnOnePage
    
End Sub

