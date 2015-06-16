VERSION 5.00
Begin VB.Form frmLocalMAPPsCFG 
   Caption         =   "Load Local MAPPs"
   ClientHeight    =   2775
   ClientLeft      =   4365
   ClientTop       =   3240
   ClientWidth     =   5775
   Icon            =   "frmLocalMAPPsCFG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   5775
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Local MAPPs"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000000&
      Caption         =   "Human"
      Height          =   615
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000000&
      Caption         =   "Mouse "
      Height          =   615
      Index           =   2
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000000&
      Caption         =   "Yeast"
      Height          =   615
      Index           =   4
      Left            =   3960
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"frmLocalMAPPsCFG.frx":208E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "Select the species"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmLocalMAPPsCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim species As String
Dim speciesselected As Boolean
   
Private Sub Command1_Click()
   On Error GoTo error:
   Dim dbPath As Database
   If speciesselected = False Then
      MsgBox "You must select a species before proceeding.", vbOKOnly
      GoTo nospecies
   End If
   
   'load the local mapps
   Select Case species
      Case "human"
         Set dbPath = OpenDatabase(programpath & "MAPPFinder Human.mdb")
      Case "mouse"
         Set dbPath = OpenDatabase(programpath & "MAPPFinder Mouse.mdb")
      Case "yeast"
         Set dbPath = OpenDatabase(programpath & "MAPPFinder Yeast.mdb")
   End Select
   
   frmLocalMAPPs.Load dbPath
   frmLocalMAPPsCFG.Hide
   frmLocalMAPPs.Show vbModal
   
error:
   Select Case Err.Number
      Case 3024 'the error for not having the database
         MsgBox "The database MAPPFinder " & species & ".mdb was not found in the folder" _
         & " containing this application. Please move it to this folder, or downloaded from GenMAPP.org.", vbOKOnly
   End Select
nospecies:
End Sub

Private Sub Command2_Click()
   frmLocalMAPPsCFG.Hide
   frmStart.Show
End Sub

Private Sub Form_Load()

   If mammalOK = False Then
      Option1(0).Enabled = False
      Option1(2).Enabled = False
   End If
   If yeastOK = False Then
      Option1(4).Enabled = False
   End If
   Option1(0).Value = False
   Option1(2).Value = False
   Option1(4).Value = False
End Sub

Private Sub Help_Click()
   frmHelp.Show
End Sub

Private Sub Option1_Click(Index As Integer)
   speciesselected = True
   Select Case Index
      Case 0
         species = "human"
      Case 2
         species = "mouse"
      Case 4
         species = "yeast"
   End Select
   frmLocalMAPPs.setSpecies (species)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormCode Then                                            'Exiting from mnuExit
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   
 Exit_Click

End Sub
Private Sub Exit_Click()
   End
End Sub


