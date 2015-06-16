VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About MAPPFinder 2.0"
   ClientHeight    =   10635
   ClientLeft      =   5325
   ClientTop       =   3150
   ClientWidth     =   8940
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Build: 20050220"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   10320
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   10695
      Left            =   0
      Picture         =   "frmAbout.frx":08CA
      Top             =   0
      Width           =   8955
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
