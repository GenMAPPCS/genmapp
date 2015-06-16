VERSION 5.00
Begin VB.Form frmNodeClickResponse 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Click Options"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   Icon            =   "frmNodeClickResponse.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Export the list of genes in my Expression Dataset that have been linked to that GO term or MAPP."
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   4815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Open the corresponding GenMAPP MAPP file."
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clicking on a term in the GO hierarchy or the Local MAPPs will:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmNodeClickResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   If TreeForm.OpenMAPPWhenClicked Then
      Option1(0).Value = True
   Else
      Option1(1).Value = True
   End If
   
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 0 Then
      TreeForm.OpenMAPPWhenClicked = True
   Else
      TreeForm.OpenMAPPWhenClicked = False
   End If
End Sub
