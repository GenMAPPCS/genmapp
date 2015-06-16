VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmInput 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Select the Expression Dataset"
   ClientHeight    =   3975
   ClientLeft      =   3465
   ClientTop       =   3990
   ClientWidth     =   7680
   Icon            =   "frmInputNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   120
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "OK"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFindFile 
      Caption         =   "Find File"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtFileName 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Text            =   "Load GenMAPP .gex File"
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Label lblDB 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   3480
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmInputNew.frx":08CA
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1680
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmInputNew.frx":0989
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   6975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuChangeGeneDB 
         Caption         =   "Choose Gene Database"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Mappfinderhelp 
         Caption         =   "MAPPFinder Help"
      End
      Begin VB.Menu about 
         Caption         =   "About MAPPFinder"
      End
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is a test MAPPFinder 2.0
'Written by Scott Doniger
'Completed 6/30/2003
'I'm changing this in the copy


Private Sub about_Click()
   frmAbout.Show
End Sub

Private Sub cmdCancel_Click()
   txtFileName.Text = ""
   frmInput.Hide
   frmStart.Show
End Sub

Private Sub cmdFindFile_Click()
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "GenMAPP Expression Datasets|*.gex"
   CommonDialog1.ShowOpen
   txtFileName.Text = CommonDialog1.FileName
End Sub



Private Sub cmdNext_Click()
   On Error GoTo error
   Dim rscolorsets As DAO.Recordset
   
   If txtFileName.Text = "" Then
      MsgBox "You have not selected an Expression Dataset. Please do so.", vbOKOnly
   Else
      Set dbExpressionData = OpenDatabase(txtFileName.Text)
      Set rscolorsets = dbExpressionData.OpenRecordset("Select * from DISPLAY")
      Set rscolorsets = dbExpressionData.OpenRecordset("SELECT ColorSet, Criteria FROM [ColorSet]")
      frmCriteria.Load (txtFileName.Text)
      txtFileName.Text = ""
      frmInput.Hide
      
   End If
error:
   Select Case Err.Number
      Case 3024
         MsgBox "MAPPFinder cannot find the file. Please select a valid file.", vbOKOnly
      Case 3051
         MsgBox "The file you have selected has been set to read-only or it is locked by another program" _
            & ", please make sure it is closed, or not set to read-only.", vbOKOnly
      Case 3078
         MsgBox "You have selected a GenMAPP version 1.0 Expression Dataset. To use MAPPFinder 2.0, you" _
         & " will need to install GenMAPP version 2.0 and convert the version 1.0 Expression Dataset to version 2.0.", vbOKOnly
         frmStart.Show
         Me.Hide
   End Select
End Sub

Private Sub Close_Click()
    End
End Sub

Private Sub Exit_Click()
   End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormCode Then                                            'Exiting from mnuExit
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   Exit_Click

End Sub

Private Sub MAPPFinderhelp_Click()
  Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, programpath & "\GenMAPP.chm::/MAPPFinder.htm", HH_DISPLAY_TOPIC, 0)

End Sub

Private Sub mnuChangeGeneDB_Click()
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
