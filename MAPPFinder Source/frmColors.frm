VERSION 5.00
Begin VB.Form frmColors 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Color key"
   ClientHeight    =   4785
   ClientLeft      =   14025
   ClientTop       =   435
   ClientWidth     =   1665
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   1665
   Begin VB.Image Image1 
      Height          =   3345
      Left            =   0
      Picture         =   "frmColors.frx":08CA
      Top             =   1320
      Width           =   1680
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The GO terms and Local MAPPs are colored based on the nested percent of genes changed."
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
   TreeForm.Show
   frmColors.Hide
End Sub
