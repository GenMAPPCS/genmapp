VERSION 5.00
Begin VB.Form frmrankedMAPPs 
   Caption         =   "Form1"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   6240
      Width           =   10575
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   10575
   End
   Begin VB.Label Label3 
      Caption         =   "Gene Ontology Terms"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   5880
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Local MAPPs"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "The list below shows those Local MAPPs and GO terms that meet the your criteria of"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frmrankedMAPPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub
