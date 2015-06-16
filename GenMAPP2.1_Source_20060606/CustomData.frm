VERSION 5.00
Begin VB.Form frmCustomData_Old 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Data"
   ClientHeight    =   3972
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   3444
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CustomData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3972
   ScaleWidth      =   3444
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLink 
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Top             =   480
      Width           =   2508
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   2400
      TabIndex        =   4
      Top             =   3540
      Width           =   972
   End
   Begin VB.TextBox txtData 
      Height          =   2160
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "CustomData.frx":08CA
      Top             =   1200
      Width           =   3228
   End
   Begin VB.TextBox txtLabel 
      Height          =   300
      Left            =   840
      TabIndex        =   0
      Top             =   60
      Width           =   2508
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link to"
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   540
      Width           =   588
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   528
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   492
   End
End
Attribute VB_Name = "frmCustomData_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   Hide
End Sub

