VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "About GenMAPP"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAbout 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   0
      Picture         =   "About.frx":08CA
      ScaleHeight     =   7695
      ScaleWidth      =   7905
      TabIndex        =   1
      Top             =   0
      Width           =   7905
      Begin VB.Label lblBuild 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Build"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   420
         Width           =   2835
      End
      Begin VB.Label lblProgramTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Program Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   2835
      End
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Build"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   420
      Width           =   48
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
   lblProgramTitle.visible = False
   lblBuild.visible = False
'   lblProgramTitle = PROGRAM_TITLE
'   lblBuild = BUILD
   CurrentX = 10
   CurrentY = picAbout.Height + 50
   Print "Build: " & BUILD
   Width = picAbout.Width + Width - ScaleWidth
   Height = CurrentY + 50 + TextHeight("A") + Height - ScaleHeight '300 '200
   Left = (Screen.Width - Width) / 2
   Top = (Screen.Height - Height) / 2
End Sub

