VERSION 5.00
Begin VB.Form frmCalculation 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Calculation Summary"
   ClientHeight    =   7665
   ClientLeft      =   450
   ClientTop       =   795
   ClientWidth     =   9810
   Icon            =   "frmCalculation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9810
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmCalculation.frx":08CA
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   6840
      Width           =   9495
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Number meeting criterion"
      Height          =   375
      Left            =   3720
      TabIndex        =   26
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Number in Expression Dataset"
      Height          =   375
      Left            =   6840
      TabIndex        =   25
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Probes"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Probes found in Cluster System"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Genes found on a MAPP"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblLocalProbeC 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblGenesOnMAPPE 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   20
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lbinClusterLocalE 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   19
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblgenesonMAPPC 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblinClusterLocalC 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblLocalProbeE 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblprobeE 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   15
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblnoClusterC 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblGenesinGOC 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblnoClusterE 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblGenesinGOE 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblprobeC 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblcriterion 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   9615
   End
   Begin VB.Label lblgeneingo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Genes found in GO"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Probes found in Cluster System"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Probes"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Number in Expression Dataset"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Number meeting criterion"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "For the z score, the number of genes in GO or on all of the local MAPPs are used for the calculations."
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   9375
   End
   Begin VB.Label lblLocalCriteria 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gene Ontology Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Local MAPPs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   Me.Hide
   Cancel = 1
   
End Sub

Private Sub lblLocalStat_Click()

End Sub

Public Sub Clear_Form()
   lblLocalProbeC.Caption = ""
   lblLocalProbeE.Caption = ""
   lblinClusterLocalC.Caption = ""
   lbinClusterLocalE.Caption = ""
   lblgenesonMAPPC.Caption = ""
   lblGenesOnMAPPE.Caption = ""
   lblLocalCriteria.Caption = ""
   lblLocalProbeC.Caption = ""
   lblLocalProbeE.Caption = ""
   lblnoClusterC.Caption = ""
   lblnoClusterE.Caption = ""
   'lblGenesC.Caption = ""
'   lblGenesE.Caption = ""
   lblGenesinGOC.Caption = ""
   lblGenesinGOE.Caption = ""
   lblcriterion.Caption = ""
   

End Sub

