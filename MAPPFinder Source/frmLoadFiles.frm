VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLoadFiles 
   BackColor       =   &H00C0FFFF&
   Caption         =   "MAPPFinder 2.0"
   ClientHeight    =   5460
   ClientLeft      =   3465
   ClientTop       =   3600
   ClientWidth     =   7815
   Icon            =   "frmLoadFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7815
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoadFiles 
      Caption         =   "Load Files"
      Height          =   735
      Left            =   4080
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdLocal 
      Caption         =   "Local MAPP Results"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtLocal 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   5415
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "Gene Ontology Results"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtGO 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label lblspecies 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "If this isn't the correct species, you must change the Gene Database."
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label lblDB 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Species Selected:"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Please select the Gene Ontology and Local MAPP results files that you would like to load."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   7215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuChangeGeneDB 
         Caption         =   "Choose Gene Database"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu mappfinderhelp 
         Caption         =   "MAPPFinder Help"
      End
      Begin VB.Menu about 
         Caption         =   "About MAPPFinder"
      End
   End
End
Attribute VB_Name = "frmLoadFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public speciesselected As Boolean


Private Sub about_Click()
   frmAbout.Show
End Sub

Private Sub cmdCancel_Click()
   txtLocal.Text = ""
   txtGO.Text = ""
   frmStart.Show
   frmLoadFiles.Hide
End Sub



Private Sub cmdLocal_Click()
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "MAPPFinder Results|*-Local.txt"
   CommonDialog1.ShowOpen
   txtLocal.Text = CommonDialog1.FileName
End Sub

Private Sub cmdGO_Click()
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "MAPPFinder Results|*-GO.txt"
   CommonDialog1.ShowOpen
   txtGO.Text = CommonDialog1.FileName
   
End Sub

Public Sub cmdLoadFiles_Click()
   On Error GoTo error
   MousePointer = vbHourglass
   
   If txtGO.Text = "" And txtLocal.Text = "" Then
      MsgBox "You have not selected a file, please select the files you would like to load.", vbOKOnly
      GoTo nospecies
   End If
   
   
   If TreeForm.TviewCount = 0 Then 'the tree form isn't loaded need to load the ontology files
      TreeForm.setDatabase (frmStart.lblDB.Caption)
      TreeForm.setSpecies (lblspecies.Caption)
      TreeForm.FormLoad
   End If
   
   TreeForm.LoadFiles txtGO.Text, txtLocal.Text
   
   
   frmLoadFiles.Hide
   TreeForm.Show
   'frmColors.Show
   'frmNumbers.Show
   
   txtGO.Text = ""
   txtLocal.Text = ""
   
error:
   Select Case Err.Number
   
   End Select
   
nospecies:
   MousePointer = vbDefault
   
End Sub
Public Sub setSpecies(s As String)
   lblspecies.Caption = s
End Sub

Private Sub Close_Click()
    End
End Sub


Public Sub FormLoad()

   Dim dbMAPPfinder As Database
   Dim rsSpecies As Recordset
   Set dbMAPPfinder = OpenDatabase(databaseloc)
   Set rsSpecies = dbMAPPfinder.OpenRecordset("SELECT Species FROM INFO")
      'the database should be species specific, so this will in most cases by 1, but SwissProt shows up as a MOD
      'and is also in SwissProt.
   If rsSpecies.RecordCount = 1 Then
      lblspecies.Caption = Replace(rsSpecies![species], "|", "")
   Else
      MsgBox "This database has multiple species. MAPPFinder needs you to use a species specific database.", vbOKOnly
   End If
   Me.Show
   dbMAPPfinder.Close
error:
   Select Case Err.Number
      Case 3078
         MsgBox "The database you loaded does not appear to be a gene database. Please select " _
            & "another database file and make sure it's a gene database.", vbOKOnly
         dbMAPPfinder.Close
         frmLoadFiles.Hide
         frmStart.Show
   End Select
         
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
   LoadSpecies
   MousePointer = vbDefault
   
End Sub

Public Sub LoadSpecies()
   Dim dbMAPPfinder As Database
   Dim rsSpecies As Recordset
   Set dbMAPPfinder = OpenDatabase(databaseloc)
   Set rsSpecies = dbMAPPfinder.OpenRecordset("SELECT Species FROM INFO")
      'the database should be species specific, so this will in most cases by 1, but SwissProt shows up as a MOD
      'and is also in SwissProt.
   If rsSpecies.RecordCount = 1 Then
      lblspecies.Caption = Replace(rsSpecies![species], "|", "")
   Else
      MsgBox "This database has multiple species. MAPPFinder needs you to use a species specific database.", vbOKOnly
   End If
   dbMAPPfinder.Close
End Sub

