VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStart 
   BackColor       =   &H00C0FFFF&
   Caption         =   "MAPPFinder 2.0 Main Menu"
   ClientHeight    =   4230
   ClientLeft      =   5385
   ClientTop       =   3660
   ClientWidth     =   5940
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5940
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Local MAPPs"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calculate New Results"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Existing Results"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblDB 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Current gene database: "
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "To load existing results, click here."
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "To calculate new results, click here. "
      Height          =   735
      Left            =   2520
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "If you would like to use local MAPPs as part of your analysis, or want to change which local MAPPs are loaded, click here."
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu changeGeneDatabase 
         Caption         =   "Choose Gene Database"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu MAPPFinderhelp 
         Caption         =   "MAPPFinder Help"
      End
      Begin VB.Menu about 
         Caption         =   "About MAPPFinder"
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
   frmAbout.Show
End Sub

Private Sub changeGeneDatabase_Click()
   Dim Fsys As New FileSystemObject
   Dim newfile As TextStream, oldfile As TextStream
   Dim line As String
   Dim dbMAPPfinder As Database
   Dim rsdate As Recordset
   
   CommonDialog1.FileName = databaseloc
   CommonDialog1.Filter = "GenMAPP Gene Database|*.gdb"
   CommonDialog1.ShowOpen
   databaseloc = CommonDialog1.FileName
   UpdateDBlabel 'updates the DB label on all forms
   MousePointer = vbHourglass
   Set dbMAPPfinder = OpenDatabase(databaseloc)
   Set rsdate = dbMAPPfinder.OpenRecordset("SELECT version FROM info")
   If dbDate <> rsdate!Version Then
      dbDate = rsdate!Version
      'TreeForm.FormLoad 'need to reload the treeform with the correct ontology files
   End If
   
   dbMAPPfinder.Close
   
   Set newfile = Fsys.CreateTextFile(programpath & "mftemp.$tm")
   Set oldfile = Fsys.OpenTextFile(programpath & "MAPPFinder.cfg")
   
   newfile.WriteLine (oldfile.ReadLine)
   newfile.WriteLine (oldfile.ReadLine)
   newfile.WriteLine (databaseloc)
   oldfile.ReadLine
   newfile.WriteLine (oldfile.ReadLine)
   newfile.Close
   oldfile.Close
   Kill programpath & "MAPPFinder.cfg"
   Name programpath & "mftemp.$tm" As programpath & "MAPPFinder.cfg"
   MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
   frmStart.Hide
   frmLoadFiles.FormLoad
   
   
   
End Sub

Private Sub Command2_Click()
   frmInput.Show
   frmStart.Hide
End Sub
Private Sub Close_Click()
    End
End Sub

Private Sub Command3_Click()
   frmStart.Hide
   frmLocalMAPPs.LoadSpecies
   frmLocalMAPPs.Show
   
End Sub

Private Sub Form_Load()
   'lblDB.Caption = databaseloc
   If lblDB.Caption = "" Then
      MsgBox "You must select a gene database. Go to the File menu.", vbOKOnly
   End If
   
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

Private Sub MAPPFinderhelp_Click()
   Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, programpath & "\GenMAPP.chm::/MAPPFinder.htm", HH_DISPLAY_TOPIC, 0)

End Sub
