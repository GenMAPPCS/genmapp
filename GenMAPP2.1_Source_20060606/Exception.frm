VERSION 5.00
Begin VB.Form frmException_old 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fix Exception Line"
   ClientHeight    =   2844
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   9552
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   796
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7740
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Save all entries to this point and leave rest in exception file"
      Top             =   2400
      Width           =   1752
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7740
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Process the change in the exception line"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1752
   End
   Begin VB.CommandButton cmdDont 
      Caption         =   "Don't Add Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5760
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Ignore this line. Don't add it to the Expression Dataset"
      Top             =   2400
      Width           =   1812
   End
   Begin VB.CommandButton cmdThis 
      Caption         =   "Add This To Other"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5760
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Add this gene to the Other category"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.TextBox txtTitles 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Exception.frx":0000
      Top             =   360
      Width           =   9372
   End
   Begin VB.CheckBox chkDont 
      BackColor       =   &H0080FFFF&
      Caption         =   "Don't add line to dataset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   4980
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   2952
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Add All To Other"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5760
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Add all the remaining genes to the Other category without confirmation"
      Top             =   1980
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.OptionButton optSwissNo 
      BackColor       =   &H0080FFFF&
      Caption         =   "SwissProt/TrEMBL Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4980
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "SwissProt or TrEMBL gene accession number"
      Top             =   2100
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.TextBox txtLine 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   552
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   900
      Width           =   9432
   End
   Begin VB.OptionButton optGenBank 
      BackColor       =   &H0080FFFF&
      Caption         =   "GenBank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4980
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "GenBank accession number"
      Top             =   1620
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.OptionButton optSwissProt 
      BackColor       =   &H0080FFFF&
      Caption         =   "SwissProt/TrEMBL Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4980
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "SwissProt or TrEMBL gene name"
      Top             =   1860
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.OptionButton optOther 
      BackColor       =   &H0080FFFF&
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4980
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Something other than SwissProt or TrEMBL"
      Top             =   2340
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.Label lblException 
      Height          =   252
      Left            =   2160
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   72
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type of gene identification:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2460
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   2364
   End
   Begin VB.Label lblPrimaryType 
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1932
   End
End
Attribute VB_Name = "frmException_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  The type of gene, basically the checked option button, is carried in lblPrimaryType
'     Also returned in lblPrimaryType. If exception not for identification, returned unchanged
'  Returning Tag:
'     Empty    Normal process
'     Quit     Interrupt processing of exception file at this point
'     AddAll   Add all remaining rows of the exception file to Other
'              is checked (also the value of lblPrimaryType).
'     AddThis  Add current row of the exception file to Other
'     DontAdd  Don't add row to Exception Dataset

Private Sub cmdDont_Click()
   Tag = "DontAdd"
   Hide
End Sub

Private Sub cmdOK_Click()
'   If optGenBank Then
'      lblPrimaryType = "G"
'   ElseIf optSwissProt Then
'      lblPrimaryType = "S"
'   ElseIf optSwissNo Then
'      lblPrimaryType = "N"
'   ElseIf optOther Then
'      lblPrimaryType = "O"
'   ElseIf chkDont = vbChecked Then                                             'Simply pass through
'      lblPrimaryType = "X"                                                          'Skip this line
'   Else
'      lblPrimaryType = ""
''      MsgBox "Must choose the type of gene identification (GenBank, etc.).", _
'             vbExclamation + vbOKOnly, "Exiting Exception Handler"
''      Exit Sub                               'No primary type >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   End If
   Hide
End Sub

Private Sub cmdProcess_Click()
   Hide
End Sub

Private Sub cmdQuit_Click()
   Tag = "Quit"
   Hide
End Sub

Private Sub cmdAll_Click()
   Tag = "AddAll"
   lblPrimaryType = "O"
   cmdProcess_Click
End Sub

Private Sub cmdThis_Click()
   Tag = "AddThis"
   lblPrimaryType = "O"
   cmdProcess_Click
End Sub

Private Sub Form_Activate()
'   chkDont = vbUnchecked
   If lblException = "unidentified" Then                                'Gene type to be identified
'      optGenBank.Visible = True
'      optSwissProt.Visible = True
'      optSwissNo.Visible = True
'      optOther.Visible = True
'      lblType.Visible = True
      cmdThis.Visible = True
      cmdAll.Visible = True
      cmdProcess.Visible = True
'      cmdOK.Visible = False
   Else
'      optGenBank.Visible = False
'      optSwissProt.Visible = False
'      optSwissNo.Visible = False
'      optOther.Visible = False
'      lblType.Visible = False
      cmdThis.Visible = False
      cmdAll.Visible = False
      cmdProcess.Visible = True
'      cmdOK.Visible = True
   End If
   DoEvents
   Tag = ""
   Select Case lblPrimaryType
   Case "G"
      optGenBank = True
   Case "S"
      optSwissProt = True
   Case "N"
      optSwissNo = True
   Case "O"
      optOther = True
   Case "X"
      chkDont = vbChecked
   Case Else
      optGenBank = False
      optSwissProt = False
      optSwissNo = False
      optOther = False
   End Select
End Sub

