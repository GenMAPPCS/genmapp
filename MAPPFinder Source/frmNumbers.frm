VERSION 5.00
Begin VB.Form frmNumbers 
   BackColor       =   &H00C0FFFF&
   Caption         =   "What do the numbers mean?"
   ClientHeight    =   3660
   ClientLeft      =   5565
   ClientTop       =   7455
   ClientWidth     =   14925
   Icon            =   "frmNumbers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   14925
   Begin VB.Image Image1 
      Height          =   3405
      Left            =   120
      Picture         =   "frmNumbers.frx":08CA
      Top             =   120
      Width           =   14670
   End
End
Attribute VB_Name = "frmNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
   TreeForm.Show
   frmNumbers.Hide
End Sub
