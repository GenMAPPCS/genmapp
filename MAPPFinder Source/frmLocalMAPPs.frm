VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLocalMAPPs 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Choose Local MAPPs Folder"
   ClientHeight    =   5640
   ClientLeft      =   4305
   ClientTop       =   3600
   ClientWidth     =   5865
   Icon            =   "frmLocalMAPPs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5865
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdLocalMAPPs 
      Caption         =   "Load MAPPs"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtLocalMAPPs 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton cmdChooseFolder 
      Caption         =   "Choose Folder"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmLocalMAPPs.frx":08CA
      Height          =   855
      Left            =   360
      TabIndex        =   9
      Top             =   4680
      Width           =   5175
   End
   Begin VB.Label Lblspecies 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "If this isn't the correct species, you must change the Gene Database."
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Species Selected:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmLocalMAPPs.frx":098C
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   3600
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmLocalMAPPs.frx":0A7F
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   5295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuChooseGeneDB 
         Caption         =   "Choose Gene Database"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu MAPPfinderhelp 
         Caption         =   "MAPPFinder Help"
      End
      Begin VB.Menu about 
         Caption         =   "About MAPPFinder"
      End
   End
End
Attribute VB_Name = "frmLocalMAPPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAX_GENES = 100
Dim MyPath As String
Dim dbMAPPfinder As Database
Dim dbLocalMAPPs As Database

Dim Fsys As New FileSystemObject
Dim localMAPPstext As TextStream
Dim species As String
Dim mappnum As Integer
Dim continue As Boolean

Public Sub Load(MAPPFinderDB As Database)
   Set dbMAPPfinder = MAPPFinderDB
   
End Sub

Private Sub about_Click()
   frmAbout.Show
End Sub

Private Sub cmbSpecies_Change()
   species = cmbSpecies.Text
End Sub

Private Sub cmdCancel_Click()
   txtLocalMAPPs.Text = ""
   frmLocalMAPPs.Hide
   frmStart.Show
End Sub

Private Sub cmdChooseFolder_Click()
   frmFindFolder.setSpecies (species)
   frmFindFolder.Load
   frmFindFolder.Show vbModal
End Sub
Public Sub setSpecies(newspecies As String)
   species = newspecies
End Sub
Private Sub cmdLocalMAPPs_Click()
startover:
    MousePointer = vbHourglass
   Dim basefolder As String
   Dim i As Integer
   Dim dbname As String
   Dim slash As Integer
   continue = True
   
   
   If txtLocalMAPPs.Text = "" Then
      MsgBox "You have not selected a local MAPPs folder. Please do so."
   ElseIf (Len(txtLocalMAPPs.Text) = 3) Then 'ie C:\ or D:\
      MsgBox "You cannot load your entire hard drive. Please select a more specific folder.", vbOKOnly
   Else

   
   species = lblspecies.Caption
   Set localMAPPstext = Fsys.CreateTextFile(programpath & "LocalMAPPs_" & species & ".txt")
      
   
   Set dbMAPPfinder = OpenDatabase(databaseloc)
     
  
    Fsys.CopyFile programpath & "LocalMAPPTmpl.gtp", programpath & "LocalMAPPs_" & species & ".gmf", True
    CompactDatabase programpath & "LocalMAPPs_" & species & ".gmf", programpath & "temp.$tm"
    Kill programpath & "LocalMAPPs_" & species & ".gmf"
    Name programpath & "temp.$tm" As programpath & "LocalMAPPs_" & species & ".gmf"
    Set dbLocalMAPPs = OpenDatabase(programpath & "LocalMAPPs_" & species & ".gmf")
   
   mappnum = 0
   localMAPPstext.WriteLine ("!autogenerated-by:     MetaMAPP program")
   localMAPPstext.WriteLine ("!saved-by:             MAPPFinder")
   localMAPPstext.WriteLine ("!date:                 " & Time & Date)
   localMAPPstext.WriteLine ("!Version: ")
   localMAPPstext.WriteLine ("!note:         DO NOT ALTER THIS FILE OR YOU WILL SCREW UP MAPPFinder!!!!")
   localMAPPstext.WriteLine ("LocalPath: " & txtLocalMAPPs.Text)
   dbLocalMAPPs.Execute ("DELETE * FROM GeneToMAPP")
   'dbLocalMAPPs.Execute ("DELETE * FROM GeneToMAPPCount")
   
   MyPath = txtLocalMAPPs.Text ' Set the path.
   TreeForm.setLocalPath (MyPath)
   i = InStrRev(MyPath, "\")
   If i = Len(MyPath) Then  'you want to extract the highest folder, so go one slash further
      basefolder = Left(MyPath, Len(MyPath) - 1)
      i = InStrRev(basefolder, "\")
      basefolder = Mid(basefolder, i + 1, Len(basefolder) - i)
      localMAPPstext.WriteLine ("<" & basefolder & " ; " & basefolder)
      BuildMetaMAPP MyPath, 0
   Else
      basefolder = Mid(MyPath, i + 1, Len(MyPath) - i)
      localMAPPstext.WriteLine ("<" & basefolder & " ; " & basefolder)
      BuildMetaMAPP MyPath & "\", 1
   End If
   'dbLocalMAPPs.Close
   'CompactDatabase programpath & "LocalMAPPs_" & species & ".gmf", "lm.$tm"
   'Kill nprogrampath & "LocalMAPPs_" & species & ".gmf"
   'Name "lm.$tm" As programpath & "LocalMAPPs_" & species & ".gmf"
   'Set dbLocalMAPPs = OpenDatabase(programpath & "LocalMAPPs_" & species & ".gmf")
   dbLocalMAPPs.TableDefs.Delete "GeneToMAPPCount"
   dbLocalMAPPs.Execute ("SELECT First(MAPP) as MappName, Count(Mapp) " _
                  & "AS MAPPCount INTO [GeneToMAPPCount] From GeneToMAPP " _
                  & "GROUP BY Mapp")
   
   localMAPPstext.Close
   dbMAPPfinder.Close
   dbLocalMAPPs.Close
   'TreeForm.FormLoad don't need to do this until they load files or calculate results.
   txtLocalMAPPs.Text = ""
   frmStart.Show
   frmLocalMAPPs.Hide
   End If
nospecies:

error:
    Select Case Err.Number
        Case 3343
            'something is wrong with the database.
            Kill programpath & "LocalMAPPs_" & species & ".gmf"
            Resume startover
    End Select
   MousePointer = vbDefault
End Sub

Private Sub BuildMetaMAPP(Path As String, indent As Integer)
   Dim MyName As String
   Dim Paths As String
   Dim steps As Integer, i As Integer, j As Integer
   MyName = Dir(Path, vbDirectory)    ' Retrieve the first entry.
   steps = 1
   While MyName <> "" And continue                    ' Start the loop.
    ' Ignore the current directory and the encompassing directory.
      If MyName <> "." And MyName <> ".." Then
        ' Use bitwise comparison to see if MyName is a directory.
        
            'Debug.Print "<" & Mid(MyName, 1, Len(MyName) - 5)
         If (GetAttr(Path & MyName) = vbDirectory) Then ' it represents a directory.
            'Debug.Print MyName ' Display entry only if it
             For j = 0 To indent
               localMAPPstext.Write (" ")
            Next j
            localMAPPstext.WriteLine ("<" & MyName & " ; " & MyName)
            BuildMetaMAPP Path & MyName & "\", indent + 1
            'now we need to take steps to get dir back to where it was before the recursion
            MyName = Dir(Path, vbDirectory)
            For i = 2 To steps
               MyName = Dir()
            Next i
        ElseIf InStr(1, UCase(MyName), ".MAPP") Then 'a MAPP file
            MAPPOK = checkMAPPName(MyName)
            If MAPPOK Then
               For j = 0 To indent
                  localMAPPstext.Write (" ")
               Next j
               localMAPPstext.WriteLine ("<" & Mid(MyName, 1, Len(MyName) - 5) & " ; " & Mid(MyName, 1, Len(MyName) - 5))
               continue = AddMapp(Path & MyName)
            Else 'the mapp name was no good
               MsgBox "The MAPP, " & MyName & " contains an illegal charachter " _
                  & "(the character ' or ; are not allowed in MAPP file names). This MAPP will be left out of the" _
                  & " MAPPFinder analysis. Fix it's name and reload this folder to include it.", vbOKOnly
            End If
            
        End If
      End If
      DoEvents
      MyName = Dir() ' Get next entry
      steps = steps + 1
   Wend
End Sub

Private Function AddMapp(MappPath As String) As Boolean
   On Error GoTo error
   Dim dbMAPP As Database
   Dim rsGenes As DAO.Recordset
   Dim rsMasterGenes As DAO.Recordset
   Dim MAPPName As String
   Dim genbanks As String, swissprots As String
   Dim rsGenbanks As DAO.Recordset, rsSwissName As DAO.Recordset, rsSwissNums As DAO.Recordset
   Dim MAPPGenes As New Collection
   Dim maint As String, author As String
   Dim genes As Boolean, id As String
   Dim primaryIDs(MAX_GENES) As String                          'Gene IDs to search relationals for
      '  Primary IDs are those IDs in the systemIn that are used to search relational tables.
      '  A Primary ID is added to Genes() returned if it is found in the systemIn. (And if the
      '  systemIn is in the SystemsList, i.e. represented in the Expression Dataset.)
      '  The search for Primary IDs is in both the ID and Secondary ID columns of the systemIn.
      '  It is possible that more than one Primary ID may be found if the secondary ID shows
      '  up in more than one row.
   
   AddMapp = True
   MAPPName = Mid(MappPath, InStrRev(MappPath, "\") + 1, Len(MappPath) - 5 - InStrRev(MappPath, "\"))
   mappnum = mappnum + 1
   If mappnum >= 4999 Then
      MsgBox "The maximum number of local MAPPs you can load is 5000. You have exceeded this." _
         & " If you meant to load this many MAPPs, contact genmapp@gladstone.ucsf.edu, otherwise" _
         & " select a more specific folder. The first 5000 MAPPs will be used in MAPPFinder.", vbOKOnly
      AddMapp = False
   Else
   'Debug.Print MappName
   genes = False
   Set dbMAPP = OpenDatabase(MappPath)
   
   Set rsGenes = dbMAPP.OpenRecordset("Select DISTINCT ID, SystemCode FROM [Objects] Where Type = 'Gene'")
    
   While rsGenes.EOF = False
      If rsGenes!id <> "" Then
        genes = True
      
      
        id = getPrimary(rsGenes!id, rsGenes!systemcode)
      
      
        If addMAPPgene(MAPPGenes, id) Then
            dbLocalMAPPs.Execute ("INSERT INTO GenetoMAPP(ID, SystemCode, MAPP)" _
                  & " VALUES ('" & id & "', '" & rsGenes![systemcode] & "', '" _
                  & TextToSql(MAPPName) & "')")
            dbLocalMAPPs.Execute ("INSERT INTO GenesOnMAPP(ID) VALUES ('" & id & "')")
        End If
      End If
      rsGenes.MoveNext
   Wend
   dbMAPP.Close
   End If
error:
   Select Case Err.Number
      Case 3051
         MsgBox "The MAPP " & MappPath & " is read-only, or open somewhere else. It will be" _
            & " left out of the Local MAPPs analysis. If you want to include this MAPP, you need to " _
            & "close it, or make sure it isn't read-only.", vbOKOnly
      Case 3061
         MsgBox "The MAPP " & MappPath & " appears to be a version 1.0 MAPP File. You are running MAPPFinder 2.0" _
               & " and need to use GenMAPP 2.0 MAPPs. Please check this. You can open version 1.0 MAPPs in version 2.0" _
               & " and they will be converted for you.", vbOKOnly
   End Select

End Function
Private Function addMAPPgene(MAPPGenes As Collection, id As String) As Boolean
On Error GoTo error
   MAPPGenes.Add id, id
   addMAPPgene = True
   Exit Function
error:
   addMAPPgene = False
   
End Function
Private Function getPrimary(idIn As String, systemIn As String) As String
   Dim Index As Integer, lastIndex As Integer
   Dim primaryID As String                          'Gene IDs to search relationals for
      '  Primary IDs are those IDs in the systemIn that are used to search relational tables.
      '  A Primary ID is added to Genes() returned if it is found in the systemIn. (And if the
      '  systemIn is in the SystemsList, i.e. represented in the Expression Dataset.)
      '  The search for Primary IDs is in both the ID and Secondary ID columns of the systemIn.
      '  It is possible that more than one Primary ID may be found if the secondary ID shows
      '  up in more than one row.
   Dim system As String                            'Cataloging system code currently being examined
   Dim column As String                                                       'Any secondary column
   Dim rsSystems As Recordset                                                    'The Systems table
   Dim rsSystem As Recordset
   Dim pipe As Integer, slash As Integer
   Dim sql As String, found As Boolean
   
   
   Set rsSystems = dbMAPPfinder.OpenRecordset( _
                   "SELECT * FROM Systems WHERE SystemCode = '" & systemIn & "'", _
                   dbOpenForwardOnly)                                    'Get the system table name
                                                 'Eg: SELECT * FROM Systems WHERE SystemCode = 'Rm'
   
   
   
   
   If rsSystems.EOF Then '-----------------------------------------------------System Doesn't Exist
      getPrimary = idIn 'have to leave it alone
                                            'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   Else
       If rsSystems![Date] <> "" Or systemIn = "O" Then           'Date or Other, supported system
      '  Other always supported. See comments at beginning of sub
      supportedSystem = True
      End If
      If supportedSystem Then
         '=================================================================Check ID Column For Gene In
         Set rsSystem = dbMAPPfinder.OpenRecordset( _
                        "SELECT * FROM " & rsSystems!system & _
                        "   WHERE ID = '" & idIn & "'", _
                        dbOpenForwardOnly)        'Eg: SELECT * FROM SwissProt WHERE ID = 'CALM_HUMAN'
         If Not rsSystem.EOF Then '____________________________________________Found idIn in ID column
      '         lastPrimaryIndex = lastPrimaryIndex + 1
      '         primaryIDs(lastPrimaryIndex) = idIn
            getPrimary = idIn
            
         Else '________________________________________________________________Check Secondary Columns
            '  If the received idIn is not in the ID column of the cataloging system table (systemIn)
            '  go to the secondary columns and search
            '  Typical column listing in Systems!Columns
            '     ID|Accession\SMBF|Nicknames\sF|Protein|Functions\B|
            found = False
            pipe = 3                                                                     'End of "ID|"
            slash = InStr(pipe + 1, rsSystems!columns, "\")                                'Next slash
            Do While slash And Not found '-----------------------------------------------------Each Secondary Column
               pipe = InStrRev(rsSystems!columns, "|", slash)            'Next pipe with slash in unit
               '  If the Columns column had
               '     ID|Whatever|Accession\SMBF|Nicknames\sF|Protein|Functions\B|
               '  then the pipe would be the one beginning the |Accession\SMBF|
               '  not the |Whatever| unit.
               If UCase(Mid(rsSystems!columns, slash + 1, 1)) = "S" Then      'This is a search column
                  column = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)             'Column name
                  If Mid(rsSystems!columns, slash + 1, 1) = "S" Then 'Multiple IDs surrounded by pipes
                     sql = "SELECT ID" & _
                           "   FROM " & rsSystems!system & _
                           "   WHERE [" & column & "] LIKE '*|" & idIn & "|*'"               'Use LIKE
                     'Eg: SELECT ID FROM SwissProt WHERE Accession LIKE '*|A1234|*'
                  Else                                                        'Single ID without pipes
                     sql = "SELECT ID" & _
                           "   FROM " & rsSystems!system & _
                           "   WHERE [" & column & "] = '" & idIn & "'"                         'Use =
                          'Eg: SELECT ID FROM SwissProt WHERE Accession = 'A1234'
                          'Doesn't happen with SwissProt!
                  End If
                  Set rsSystem = dbMAPPfinder.OpenRecordset(sql, dbOpenForwardOnly)         'Get all records
                  If rsSystem.EOF = False Then                            'you found a Primary ID
                   'unlike in GenMAPP, we're only storing the first Primary ID encountered. It complicates
                   'things too much to look at a secondary that is in multiple primarys.
                   'This is such a rare event (only a few SwissProt Accession numbers) so I'm not going
                   'to worry about it. -sd 3/26/03
                     found = True
                     getPrimary = rsSystem!id
                  End If
               End If
               slash = InStr(slash + 1, rsSystems!columns, "\")                            'Next slash
            Loop
            If found = False Then 'this ID wasn't a primary or a secondary. we have to leave it alone
               getPrimary = idIn
            End If
         End If
      Else 'not a supported system so leave it alone
         getPrimary = idIn
      End If
   End If
End Function


'need to build the table that counts the number of genes in each mapp
Private Sub buildCountTable()
   
   dbLocalMAPPs.TableDefs.Delete "GeneToMAPPCount"
   dbLocalMAPPs.Execute ("SELECT First(MAPP) as MappName, Count(Mapp) " _
                  & "AS MAPPCount INTO [GeneToMAPPCount] From GeneToMAPP " _
                  & "GROUP BY Mapp")

End Sub

Private Sub Form_Load()
  On Error GoTo error
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
   
error:
   Select Case Err.Number
    Case 3024
        MsgBox "You must load a database.", vbOKOnly
        frmStart.Show
End Select
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

Private Sub MAPPFinderhelp_Click()
 Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, programpath & "\GenMAPP.chm::/MAPPFinder.htm", HH_DISPLAY_TOPIC, 0)
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

Function TextToSql(txt As String) As String '**************************** Makes Text SQL Compatible
    Dim Index As Integer                     'copied from GenMAPP 1.0 Source code
    Dim sql As String
   
    sql = txt
    For Index = 1 To Len(txt)
      Select Case Mid(txt, Index, 1)
      Case "'"                            'Convert single quote to typographer's close single quote
         Mid(sql, Index, 1) = Chr(146)
      Case Else
      End Select
    Next Index
    For Index = 1 To Len(txt)
      Select Case Mid(txt, Index, 1)
      Case "!"                            'Convert single quote to typographer's close single quote
         Mid(sql, Index, 1) = Chr(32)
      Case Else
      End Select
   Next Index
   TextToSql = sql
End Function


Public Function checkMAPPName(MAPPName As String) As Boolean
   If (InStr(1, MAPPName, "'") <> 0) Or (InStr(1, MAPPName, ";") <> 0) Then
      checkMAPPName = False
   Else
      checkMAPPName = True
   End If
End Function

Private Sub mnuChooseGeneDB_Click()
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
