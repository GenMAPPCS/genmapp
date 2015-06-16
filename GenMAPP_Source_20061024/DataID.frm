VERSION 5.00
Begin VB.Form frmDataID 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Type Specification"
   ClientHeight    =   3036
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6480
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3036
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   4260
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1032
   End
   Begin VB.ListBox lstSystemCodes 
      Height          =   288
      ItemData        =   "DataID.frx":0000
      Left            =   2100
      List            =   "DataID.frx":0002
      TabIndex        =   12
      Top             =   4140
      Visible         =   0   'False
      Width           =   144
   End
   Begin VB.ListBox lstSystems 
      Height          =   1248
      ItemData        =   "DataID.frx":0004
      Left            =   1320
      List            =   "DataID.frx":0006
      TabIndex        =   8
      Top             =   2580
      Visible         =   0   'False
      Width           =   2832
   End
   Begin VB.ListBox lstTitles 
      Columns         =   3
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      IntegralHeight  =   0   'False
      ItemData        =   "DataID.frx":0008
      Left            =   60
      List            =   "DataID.frx":000A
      Style           =   1  'Checkbox
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   300
      Width           =   6372
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1032
   End
   Begin VB.CheckBox chkTitle 
      BackColor       =   &H0080FFFF&
      Height          =   192
      Index           =   0
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System"
      Height          =   240
      Index           =   2
      Left            =   60
      TabIndex        =   11
      Top             =   2940
      Visible         =   0   'False
      Width           =   672
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cataloging"
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   948
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Primary"
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   2580
      Visible         =   0   'False
      Width           =   696
   End
   Begin VB.Label lblInstructions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check the box if the column contains character data."
      Height          =   240
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   4608
   End
   Begin VB.Label lblDataHead 
      BackStyle       =   0  'Transparent
      Caption         =   "Data in First Row"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   2400
      TabIndex        =   3
      Top             =   5640
      Width           =   1872
   End
   Begin VB.Label lblTitleHead 
      BackStyle       =   0  'Transparent
      Caption         =   "Column"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   5640
      Width           =   972
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   240
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   5700
      Visible         =   0   'False
      Width           =   48
   End
End
Attribute VB_Name = "frmDataID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cancelExit As Boolean                        'Sends message to QueryUnload if user cancels exit

Private Sub cmdCancel_Click()
   Tag = "Cancel"
   Hide
End Sub

Private Sub cmdOK_Click()
   Tag = ""
   Hide
'   If lstSystems.SelCount Then
'      Hide
'   Else
'      MsgBox "Must click on a Primary Cataloging System.", vbExclamation + vbOKOnly
'   End If
End Sub

Private Sub Form_Activate()
   Tag = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   cancelExit = False
   cmdOK_Click
   If cancelExit Then Cancel = True
End Sub
