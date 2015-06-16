VERSION 5.00
Begin VB.Form frmMAPPInfo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MAPP Information"
   ClientHeight    =   3852
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5352
   Icon            =   "MAPPInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3852
   ScaleWidth      =   5352
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   3300
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3420
      Width           =   972
   End
   Begin VB.TextBox txtNotes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   672
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      ToolTipText     =   "These notes will only appear in this window."
      Top             =   2640
      Width           =   3912
   End
   Begin VB.TextBox txtRemarks 
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
      Left            =   1380
      MaxLength       =   50
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Appears in the Information area on the MAPP graphic."
      Top             =   1920
      Width           =   3912
   End
   Begin VB.TextBox txtModify 
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
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   4
      ToolTipText     =   "Enter date MAPP was last modified."
      Top             =   1560
      Width           =   3912
   End
   Begin VB.TextBox txtCopyright 
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
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   6
      ToolTipText     =   "Copyright date and entity. E.g. 2000, Gladstone Institutes"
      Top             =   2280
      Width           =   3912
   End
   Begin VB.TextBox txtEMail 
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
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Contact this address for more information about this MAPP"
      Top             =   1200
      Width           =   3912
   End
   Begin VB.TextBox txtMaint 
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
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   "Person who maintains the published MAPP."
      Top             =   840
      Width           =   3912
   End
   Begin VB.TextBox txtAuthor 
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
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Creator of the MAPP."
      Top             =   480
      Width           =   3912
   End
   Begin VB.TextBox txtTitle 
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
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   0
      ToolTipText     =   "The title that appears at the top of the MAPP graphic."
      Top             =   120
      Width           =   3912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   4320
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3420
      Width           =   972
   End
   Begin VB.Label lblMAPP 
      Height          =   312
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   2412
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
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
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   2700
      Width           =   528
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   1980
      Width           =   804
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last modified"
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
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
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
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2340
      Width           =   852
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
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
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1260
      Width           =   576
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maintained by"
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
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   900
      Width           =   1224
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
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
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   540
      Width           =   588
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
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
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   384
   End
End
Attribute VB_Name = "frmMAPPInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MAPPInfoDirty As Boolean
Dim title As String, author As String, maint As String, email As String, modify As String
Dim remarks As String, copyright As String, notes As String


Private Sub Form_Activate()
   MAPPInfoDirty = False
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++ Save Original Data In Case User Cancels
   title = txtTitle
   author = txtAuthor
   maint = txtMaint
   email = txtEMail
   modify = txtModify
   remarks = txtRemarks
   copyright = txtCopyright
   notes = txtNotes
End Sub
Private Sub cmdOK_Click()
   Dim invalidChrsFound As Boolean
   
   txtTitle = TextToSql(Dat(txtTitle))
      If InvalidChr(txtTitle, "Title") Then
         invalidChrsFound = True
         txtTitle.SetFocus
      End If
   txtAuthor = TextToSql(Dat(txtAuthor))
      If InvalidChr(txtAuthor, "Author") Then
         invalidChrsFound = True
         txtAuthor.SetFocus
      End If
   txtMaint = TextToSql(Dat(txtMaint))
      If InvalidChr(txtMaint, "Maintained by") Then
         invalidChrsFound = True
         txtMaint.SetFocus
      End If
   txtEMail = TextToSql(Dat(txtEMail))
      If InvalidChr(txtEMail, "E-mail") Then
         invalidChrsFound = True
         txtEMail.SetFocus
      End If
   txtModify = TextToSql(Dat(txtModify))
      If InvalidChr(txtModify, "Last modified") Then
         invalidChrsFound = True
         txtModify.SetFocus
      End If
   txtRemarks = TextToSql(Dat(txtRemarks))
      If InvalidChr(txtRemarks, "Remarks") Then
         invalidChrsFound = True
         txtRemarks.SetFocus
      End If
   txtCopyright = TextToSql(Dat(txtCopyright))
      If InvalidChr(txtCopyright, "Copyright") Then
         invalidChrsFound = True
         txtCopyright.SetFocus
      End If
   txtNotes = TextToSql(Dat(txtNotes))
      If InvalidChr(txtNotes, "Notes") Then
         invalidChrsFound = True
         txtNotes.SetFocus
      End If
   
   If invalidChrsFound Then Exit Sub                       '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   Hide
End Sub
Private Sub cmdCancel_Click()
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Restore Original Data
   txtTitle = title
   txtAuthor = author
   txtMaint = maint
   txtEMail = email
   txtModify = modify
   txtRemarks = remarks
   txtCopyright = copyright
   txtNotes = notes
   Hide
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Cancel = True
   If MAPPInfoDirty Then
      If MsgBox("Save changes?", vbQuestion + vbYesNo, "Closing MAPP Information") = vbYes Then
         cmdOK_Click
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   cmdCancel_Click
End Sub
Private Sub txtAuthor_Change()
   MAPPInfoDirty = True
End Sub
Private Sub txtCopyright_Change()
   MAPPInfoDirty = True
End Sub
Private Sub txtRemarks_Change()
   MAPPInfoDirty = True
End Sub
Private Sub txtEMail_Change()
   MAPPInfoDirty = True
End Sub
Private Sub txtMaint_Change()
   MAPPInfoDirty = True
End Sub
Private Sub txtModify_Change()
   MAPPInfoDirty = True
End Sub
Private Sub txtNotes_Change()
   MAPPInfoDirty = True
End Sub
Private Sub txtTitle_Change()
   MAPPInfoDirty = True
End Sub
Sub Clear()
   txtAuthor = ""
   txtCopyright = ""
   txtRemarks = ""
   txtEMail = ""
   txtMaint = ""
   txtModify = ""
   txtNotes = ""
   txtTitle = ""
   lblMAPP = ""
End Sub
