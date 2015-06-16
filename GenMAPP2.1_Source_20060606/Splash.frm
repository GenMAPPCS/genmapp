VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6195
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   6195
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSplash 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4875
      Left            =   0
      Picture         =   "Splash.frx":2926E
      ScaleHeight     =   4815
      ScaleWidth      =   7905
      TabIndex        =   0
      Top             =   0
      Width           =   7965
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   With picSplash
   Height = picSplash.Height
   Width = picSplash.Width
   Top = (Screen.Height - Height) / 2
   Left = (Screen.Width - Width) / 2
   .CurrentX = 20
   .CurrentY = picSplash.Height - 300
   .foreColor = vbWhite
   picSplash.Print BUILD
   End With
End Sub

