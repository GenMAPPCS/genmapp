VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCriteria 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Calculate New Results"
   ClientHeight    =   8790
   ClientLeft      =   6405
   ClientTop       =   1500
   ClientWidth     =   6390
   Icon            =   "frmCriteriaNewtry3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   6390
   Begin VB.CheckBox chkStats 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Click here to calculate p values (calculating p values will add between 5-10 minutes to the runtime per criterion)."
      Height          =   495
      Left            =   1080
      TabIndex        =   20
      Top             =   4440
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check2"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Dataset"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -120
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   5520
      Width           =   4575
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Browse"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ListBox lstcriteria 
      Height          =   1230
      Left            =   3120
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.ListBox lstColorSet 
      Height          =   1230
      ItemData        =   "frmCriteriaNewtry3.frx":08CA
      Left            =   360
      List            =   "frmCriteriaNewtry3.frx":08CC
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdRunMAPPFinder 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Run MAPPFinder"
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lblspecies 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "If this isn't the correct species, you must change the Gene Database."
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   2880
      Width           =   4935
   End
   Begin VB.Label lblDB 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   7920
      Width           =   6135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmCriteriaNewtry3.frx":08CE
      Height          =   615
      Left            =   360
      TabIndex        =   19
      Top             =   1800
      Width           =   5655
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "(-criterion# and -GO or -Local will be added to the file name. Do not remove -GO or -Local from the file name.)"
      Height          =   495
      Left            =   1440
      TabIndex        =   18
      Top             =   6000
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gray = No Local MAPPs loaded"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Local MAPPs"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gene Ontology"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select the type of analysis you would like to run."
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save Results as: "
      Height          =   495
      Left            =   1080
      TabIndex        =   12
      Top             =   5160
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Species Selected:"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select Criteria to filter by:"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select Color Set:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblProgress 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   7320
      Width           =   5415
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mnuChangeGeneDB 
         Caption         =   "Choose Gene Database"
      End
      Begin VB.Menu localMAPPs 
         Caption         =   "Load Local MAPPs"
      End
      Begin VB.Menu newresults 
         Caption         =   "Calculate New Results"
      End
      Begin VB.Menu loadExisting 
         Caption         =   "Load Existing Results"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
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
Attribute VB_Name = "frmCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MAPPFinder 1.0
'Written by Scott Doniger
'Completed 12/11/2001
'This file contains the MAPPFinder algorithm. It links Expression Data to GO terms and then
'calculates the percentage of genes in each GO term that changed based on the criteria selected
'by the user.

'Input: a .Gex file (GenMAPP expression Data file)
'Output: 3 text files, Component, Function, Process with the results.
'The first time MAPPFinder is run on a .gex file, a .gdb file is created storing
'the GO annotations for that entire chip. This total chip data is used for subsequent
'criteria.
 
 
Const MAX_CRITERIA = 30
Const GOCount = 46000 'the highest number assigned to a go term. This is pretty large, but oh well.
Const MAPPCount = 5000 'the highest number of mapps that can be loaded as local
Const GOLENGTH = 10
Const TRIALS = 2000
Const IDSIZE = 30 'the field size for a gene ID (needs to match GenMAPP)

 Dim rscolorsets As DAO.Recordset  'stores the colorsets of the .gex file
 Dim dbExpressionData As Database 'stores the expression table of the .gex file
 Dim sql(MAX_CRITERIA) As String
 Dim criteriaSelected(MAX_CRITERIA) As Boolean
 Public species As String, GODate As String
 Dim fullname As String
 Dim newfilename As String
 Dim dbMAPPfinder As Database 'the MAPPFinder database MAPPFinder 1.0.mdb
 Dim dbLocalMAPPs As Database
 Dim dbChipData As Database, dbChipDataLocal As Database 'the entire chips annotations.
 Dim gotable As String, FSO As Object
 Dim expressionName As String
 Dim filelocation As String, colorset As String
 Dim chipName As String
 Public clustersystem As String, clustercode As String
 Dim speciesselected As Boolean, geneontology As Boolean
 Public LocalMAPPsLoaded As Boolean
 Dim colorsetclicked As Boolean, criteriaclicked As Boolean
 Dim GOrelation As String
 Dim Statistics As Boolean
 Dim localR As Long, bigR As Long, noCluster As Long
 Dim localN As Long, bigN As Long
 Dim distinctgenes As Long, nomapp As Long
 Dim nolocalCluster As Long, GEXsize As Long, currentcriterion As Integer
 
 Dim relations(MAX_RELATIONS, 3) As String
 ' 0 = relation
 ' 1 = P or R or S(is the Expression data ID the primary or related field of the relation)
                  ' or are the Expression ID and the Cluster ID the same type
 ' 2 = systemcode of gex
 ' 3 = any secondary IDs that need to be searched
 Dim PrimaryGenes As New Collection
   'key is GEX ID, item is a primarygene object
   '1 = Cluster ID
   'redim to size of GEX and index by orderno
Dim ClusterGenes As New Collection
   'key is ClusterID, value is a ClusterGene Object for the ID
Dim GOterms As New Collection
   'key is GOID (string)
   'value is a GOterm object with all of the numbers
Dim LocalPrimaryGenes As New Collection
   'key is ID
   'value is a PrimaryGene Object
Dim LocalClusterGenes As New Collection
   'key is an ID of a systemcode supported by the GenesToMAPP table
   'value is a clustergene object for that gene on a mapp
Dim localMAPPsCollection As New Collection
   'key is a mappname
   'object is a GOterm object (no nested values)
Dim GEXgenes As New Collection ' for the Local MAPPs
   'a collection of cluster genes storing the GEX->MAPP relation
 
         
Public Sub Load(FileName As String)
   On Error GoTo error
   Dim colorset As String
   Dim slash As Integer
   Dim dbTemp As Database
   Dim rsSpecies As Recordset
   speciesselected = False
   chipbuiltOK = False
   filelocation = FileName
   Statistics = False
   slash = InStrRev(FileName, "\")
   newfilename = Mid(FileName, 1, Len(FileName) - 4) 'everything but .gex
   Set dbExpressionData = OpenDatabase(FileName)
  
   Set rscolorsets = dbExpressionData.OpenRecordset("SELECT ColorSet FROM [ColorSet]")
   If rscolorsets.EOF = True Then
      MsgBox "There are no colorsets in this Expression Dataset File. Please return to GenMAPP" _
           & " and define at least one color set.", vbOKOnly
   Else
      While rscolorsets.EOF = False
         lstColorSet.AddItem rscolorsets![colorset]
         rscolorsets.MoveNext
      Wend
   End If
   lstcriteria.Enabled = False
   lstcriteria.Visible = False
   DisplayALL = True
   Check2.Enabled = False
   colorsetclicked = False
   criteriaclicked = False
   geneontology = False
   LocalMAPPsLoaded = False
   
   Set dbTemp = OpenDatabase(databaseloc)
   Set rsSpecies = dbTemp.OpenRecordset("SELECT [MOD] FROM Systems WHERE [MOD] <> Null" _
                                       & " AND [Date] <> Null ORDER BY [MOD]")
   'the database should be species specific, so this will in most cases by 1, but SwissProt shows up as a MOD
   'and is also in SwissProt.
   If rsSpecies.RecordCount = 1 Then
      lblspecies.Caption = rsSpecies![Mod]
   ElseIf rsSpecies.RecordCount = 2 Then
      'SwissProt and a MOD
      While rsSpecies.EOF = False
         If rsSpecies![Mod] <> "Homo sapiens" Then
            lblspecies.Caption = rsSpecies![Mod]
         End If
         rsSpecies.MoveNext
      Wend
   Else
      MsgBox "MAPPFinder requires a species specific database. Please change your gene database.", vbOKOnly
   End If
   
   species = lblspecies.Caption
   If Dir(Module1.programpath & "LocalMAPPs_" & species & ".txt") = ("LocalMAPPs_" & species & ".txt") Then
      Check2.Enabled = True
   Else
      Check2.Enabled = False
   End If
   
   dbTemp.Close
   frmInput.Hide
   Me.Show
   
   
   Exit Sub
error:
   Select Case Err.Number
      Case 3078
         MsgBox "The database you loaded does not appear to be a gene database. Please" _
            & " make sure you select a gene database and then try again.", vbOKOnly
         dbTemp.Close
         frmInput.Hide
         'frmCriteria.Hide
         frmStart.Show
   End Select
End Sub

Private Sub Frame2_DragDrop(Source As Control, x As Single, y As Single)

End Sub



Private Sub about_Click()
   frmAbout.Show
End Sub

Private Sub Check1_Click()
   If geneontology = True Then
      geneontology = False
   Else
      geneontology = True
   End If
End Sub

Private Sub Check2_Click()
   If LocalMAPPsLoaded = True Then
      LocalMAPPsLoaded = False
   Else
      LocalMAPPsLoaded = True
   End If
End Sub

Private Sub resetcheck2()
   Check2.Value = 0
End Sub

Private Sub chkStats_Click()
   Statistics = chkStats.Value
End Sub



Private Sub cmdFile_Click()
   Dim FileName As String
   Dim criteria As String, criterion As Integer
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "Text Files|*.txt"
   CommonDialog1.ShowSave
   txtFile.Text = CommonDialog1.FileName
   
   If txtFile.Text <> "" Then
   
   FileName = Left(txtFile.Text, Len(txtFile.Text) - 4)
   If invalidFileName(FileName) Then
      MsgBox "A filename cannot contain any of the following characters: /\:*?" & Chr(34) & "<>| are not", vbOKOnly
      txtFile.Text = ""
      Exit Sub
   End If
   For criterion = 0 To lstcriteria.ListCount - 1
      If lstcriteria.Selected(criterion) Then
         
         If Dir(FileName & "-criterion" & criterion & "-GO.txt") <> "" Or Dir(FileName & "-" & criterion & "-Local.txt") <> "" Then
            If MsgBox("Overwrite the existing " & txtFile.Text & "-criterion" & criterion & "-GO and -Local?", vbOKCancel) = vbCancel Then
               txtFile.Text = ""
               Exit For
            End If
         End If
      End If
   Next criterion
   
   End If
End Sub
 
Private Sub cmdRunMAPPFinder_Click()
'On Error GoTo error

   Dim rsFilter As DAO.Recordset, rsType As DAO.Recordset, rsRelation As Recordset
   Dim tblGOAll As TableDef, rstemp2 As DAO.Recordset
   Dim tblgo As TableDef, tblresults As TableDef, rschip As DAO.Recordset, tblnestedresults As TableDef
   Dim rsFunction As DAO.Recordset, rsProcess As DAO.Recordset, rsComponent As DAO.Recordset
   Dim percentage As Single, metFilter As Integer, noGO As Integer 'no GeneOntology available
   Dim others As Integer, noSwissProt As Integer 'no swissprot counts the number of genes that can't be converted
   Dim present As Single
   Dim output As TextStream
   Dim criteria As String, trembl As Boolean
   Dim genmappID As String, GOID As String, clusterID As String
   Dim MGIsAdded As New Collection, GenMAPPsAdded As New Collection
   Dim genecounter() As String, GOarray(GOCount) As Integer, mapparray() As Integer
   Dim progress As Integer, rsinGONested As DAO.Recordset, indata As Integer
   Dim i As Long, numofsystems As Integer, rsGenes As Recordset
   Dim criterion As Integer, relationNotFound As Boolean
   Dim r As Long, rsGOcount As Recordset
   Dim N As Long, GOIDs As New Collection, goterm As New goterm
   Dim teststat As Double, Cluster As New ClusterGene, filtergenes As Integer
   Dim results() As Double, noClusterC As Integer, clusterC As Integer
   Dim pgene As PrimaryGene, GONode As Node, cg As ClusterGene
   Dim GOresultsExist As Boolean, localResultsExist As Boolean
   Dim currentlocalcriterion As Integer
   
   GOresultsExist = False
   localResultsExist = False
   
   
   
   If geneontology = False And LocalMAPPsLoaded = False Then
      MsgBox "You have not selected the type of MAPPFinder analysis you would like to run." _
            & " Please select Gene Ontology, Local MAPPs, or both.", vbOKOnly
      Exit Sub
   End If
   
   If txtFile.Text = "" Then
      MsgBox "You have not selected a file to save the results to. Please do so now.", vbOKOnly
      Exit Sub
   End If
   
   If colorsetclicked = False Then
      MsgBox "You have not selected a color set, please do so.", vbOKOnly
      Exit Sub
   End If
   
   If criteriaclicked = False Then
      MsgBox "You have not selected a criteria, please do so.", vbOKOnly
      Exit Sub
   End If
   
   MousePointer = vbHourglass
   progress = 0
   noCluster = 0
   noMAPPs = 0
   nolocalCluster = 0
   species = lblspecies.Caption
   TreeForm.setDatabase (databaseloc)
   TreeForm.setSpecies (species)
   TreeForm.FormLoad
   Set FSO = CreateObject("Scripting.FileSystemObject")
   
   Set dbMAPPfinder = OpenDatabase(databaseloc)
   
   'check this
   
  
  'If dbExpressionData.TableDefs.count > 11 Then 'somehow the tables didn't get deleted before
  '  Dim tbl As TableDef
   '   For Each tbl In dbExpressionData.TableDefs
    '     If tbl.name = "Results" Then
     '       dbExpressionData.Execute "DROP Table Results"
      '   ElseIf tbl.name = "NestedResults" Then
      '      dbExpressionData.Execute "DROP table NestedResults"
       '  End If
      'Next tbl
   'End If
     
   

   If geneontology Then
   'get the clustersystem from the systems table
   'get the system codes of the ED
   'find all relations
      
  
      
   If mapToClusterSystem Then
      'the mapping worked
      'you've now created primarygenes, clustergenes, and goterms
      'you've set inGO, inGOlocal, and Onchiplocal, onchipnested
      
      'now you need to select the genes meeting the criterion.
      'map those to a cluster gene using primary genes
      'use clustergenes to map to GO. keep track of which cluster genes have been
      'visited, so you only use it once.
      'once you've visited all of the CGs you're done.
      'calculate Z scores for each GO term. Do statistics. Save results and display them.

      For criterion = 0 To (lstcriteria.ListCount - 1)
         
         If lstcriteria.Selected(criterion) Then
            GOresultsExist = False
            currentcriterion = criterion
            GOresultsExist = CalculateGOResults(criterion)
            
            If Not GOresultsExist Then
               GoTo ENDSUB 'this didn't work because the species didn't select anything. end and try again.
               Exit For
            End If
         End If
         
      Next criterion
      
      
   
      
     
      Else 'maptocluster returns false
         Exit Sub
      End If
   End If 'geneontology
   
   'local mapps
   
   If LocalMAPPsLoaded Then
      Set dbLocalMAPPs = OpenDatabase(programpath & "LocalMAPPs_" & species & ".gdb ")
      MaptoLocalMAPPs
       
      If localMAPPsCollection.count = 0 Then 'the MAPtoLocalMAPPs found nothing
         MsgBox "The genes in your dataset are not found on any of the currently loaded" _
         & " Local MAPPs. This may be true, or an you may have the wrong local MAPPs" _
         & " loaded. Please check that the correct MAPPs are loaded and that you are" _
         & " using the appropriate gene database for this dataset. No results will be" _
         & " calculated.", vbOKOnly
         GoTo noerror
      End If
      
      For criterion = 0 To (lstcriteria.ListCount - 1)
         
         If lstcriteria.Selected(criterion) Then
            localResultsExist = False
            criteria = lstcriteria.List(criterion)
            localResultsExist = CalculateLocalResults(criterion)
            currentlocalcriterion = criterion
            If Not localResultsExist Then
               GoTo ENDSUB 'this didn't work because of some error
               Exit For
            End If
         End If
      Next criterion
   End If
   
   
   
   
   frmLoadFiles.setSpecies (species)
   
   If GOresultsExist Then
      frmLoadFiles.txtGO.Text = fixFileName(txtFile.Text) & "-Criterion" & currentcriterion & "-GO.txt"
   End If
   
   If localResultsExist Then
      frmLoadFiles.txtLocal.Text = fixFileName(txtFile.Text) & "-Criterion" & currentlocalcriterion & "-Local.txt"
   End If
   
   
   'rstemp.Close
   'rstemp2.Close
'   rsType.Close
   'rschip.Close
  
' dbExpressionData.Execute "DROP Table NestedResults"
'   dbExpressionData.Execute "Drop Table Results"
   
   dbExpressionData.Close
   
   dbMAPPfinder.Close
   If geneontology Then
      dbChipData.Close
      'dbExpressionData.Close
      CompactDatabase newfilename & ".gdb", newfilename & ".$tm"
      Kill newfilename & ".gdb"
      Name newfilename & ".$tm" As newfilename & ".gdb"
   End If
   
   If LocalMAPPsLoaded Then
      dbChipDataLocal.Close
      DBEngine.CompactDatabase newfilename & "-Local.gdb", newfilename & "-Local.$tm"
      Kill newfilename & "-Local.gdb"
      Name newfilename & "-Local.$tm" As newfilename & "-Local.gdb"
   End If
   frmLoadFiles.speciesselected = True
   frmLoadFiles.cmdLoadFiles_Click
ENDSUB:
   'we need to erase the collection that have been built here, so if the user comes back they
   'can start from scratch.
   Set PrimaryGenes = New Collection
   Set ClusterGenes = New Collection
   Set GOterms = New Collection
   Set LocalPrimaryGenes = New Collection
   Set LocalClusterGenes = New Collection
   Set localMAPPsCollection = New Collection
   'hopefully VB's garbage collection now frees up the memory used by those collections
   
compact:
   'On Error GoTo error
   
  
   
      'TreeForm.setFileName (txtFile.Text)
      'TreeForm.CmdExpand_Click
      'TreeForm.Show
      'frmColors.Show
      'frmNumbers.Show
   
   eraseForm
   frmCriteria.Hide
   
   GoTo noerror

error:
   Select Case Err.Number
      Case 3024 'the error for not having the database
         MsgBox "The database MAPPFinder " & species & " was not found in the folder" _
         & " containing this application. Please move it to this folder or download it from www.GenMAPP.org.", vbOKOnly
      Case 3021
         MsgBox "The local MAPPs stored in " & species & " database are different than those currently loaded." _
            & " Please reload the appropriate MAPPs. MAPP " & rstemp2![MappNameField] & " not found.", vbOKOnly
      Case 3356
         MsgBox "You have the GenMAPP Expression Dataset " & newfilename & ".gex or .gdb file open in another window. You must close" _
            & " these other instances of the files for MAPPFinder to continue. Close the other instance and click OK.", vbOKOnly
         Resume compact
      Case 3061
         MsgBox "It looks like you are trying to use a version 1.0 GenMAPP Expression dataset file. You must convert" _
            & " your file to a GenMAPP 2.0 file before you can use it in MAPPFinder 2.0.", vbOKOnly
      'Case 5
         'no primary gene in primarygenes. this means that this pg has no cluster genes
       '  Resume nextPG
      'Case Else
       '  MsgBox "An error occurred while calculating the results. Please report error " & Err.Number _
         '   & " to GenMAPP@gladstone.ucsf.edu. Error message: " & Err.Description & "."
   End Select

noerror:
   MousePointer = vbDefault
   If GOresultsExist = False And localResultsExist = False Then
      Exit_Click
   End If
End Sub

  


Private Sub chkDisplayALL_Click()
   If chkDisplayALL.Value = 1 Then
      DisplayALL = False
   Else
      DisplayALL = True
   End If
End Sub





Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
   frmInput.Show
   eraseForm
   frmCriteria.Hide
   MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
   frmStart.Show
   'dbExpressionData.Close
   eraseForm
   frmCriteria.Hide
   MousePointer = vbDefault
End Sub



Private Sub Exit_Click()
   End
End Sub

Private Sub Label9_Click()
   geneontology = False
End Sub

Private Sub loadExisting_Click()
   Unload Me
   frmLoadFiles.Show
End Sub

Private Sub localMAPPs_Click()
   Unload Me
   frmLocalMAPPs.LoadSpecies
   frmLocalMAPPs.Show
End Sub

Private Sub lstColorSet_Click()
   Dim rsCriteria As DAO.Recordset
   Dim criteria As String, record As String
   Dim pipe As Integer, endline As Integer, newend As Integer, pipe2 As Integer
   Dim i As Integer
   lstcriteria.Clear
   colorsetclicked = True
   colorset = lstColorSet.Text
   Set rsCriteria = dbExpressionData.OpenRecordset("SELECT Criteria FROM [ColorSet] WHERE" _
                  & " ColorSet = '" & colorset & "'")
   
   record = rsCriteria![criteria]
   endline = -1
   i = 0
   pipe = InStr(endline + 2, record, "|")
   While pipe > 0
      criteria = Mid(record, endline + 2, pipe - endline - 2)
      If UCase(criteria) <> "NO CRITERIA MET" And UCase(criteria) <> "NOT FOUND" Then
         lstcriteria.AddItem criteria
         pipe2 = InStr(pipe + 1, record, "|")
         sql(i) = Mid(record, pipe + 1, pipe2 - pipe - 1)
      End If
      newend = InStr(pipe, record, Chr(13))
      If newend < endline Then
         pipe = -1
      Else
         endline = newend
         pipe = InStr(endline + 2, record, "|")
         i = i + 1
      End If
   Wend
   lstcriteria.Enabled = True
   lstcriteria.Visible = True
   cmdFile.Enabled = True
   
  
End Sub

Private Sub lstcriteria_Click()
  'Debug.Print sql(lstcriteria.ListIndex)
   Dim i As Integer
   Dim criteria As String
   criteriaclicked = True
   criteria = ""
   For i = 0 To lstcriteria.ListCount - 1
      If lstcriteria.Selected(i) Then
          criteria = criteria & "-" & (i + 1)
      End If
   Next i
   
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

Private Sub newresults_Click()
   Unload Me
   frmInput.Show
End Sub



Public Sub buildOnChipLocal()
    Dim rsExpression As DAO.Recordset
    Dim rstemp As DAO.Recordset, rstemp2 As DAO.Recordset, rsGO As DAO.Recordset
    Dim genes As Integer
    Dim mapparray(MAPPCount) As Integer
    Dim MAPPName(MAPPCount) As String
    Dim progress As Integer
    
    progress = 0
      
      dbChipData.Execute ("DELETE * FROM LocalMAPPsChip")
      Set rstemp = dbExpressionData.OpenRecordset("SELECT distinct GenMAPP FROM Expression")
      While Not rstemp.EOF
         progress = progress + 1
         If progress Mod 10 = 0 Then
            lblProgress.Caption = progress & " out of the " & rstemp.RecordCount & " genes measured are linked to Local MAPPs."
            frmCriteria.Refresh
         End If
         
         Set rstemp2 = dbMAPPfinder.OpenRecordset("Select DISTINCT MAPPNameField, MAPPNumber FROM [GeneToMAPP]" _
                           & " WHERE GenMAPP = '" & rstemp![GenMAPP] & "'")
            If rstemp2.EOF = False Then
               While Not rstemp2.EOF
                  MAPPName(rstemp2![MAPPNumber]) = rstemp2![MappNameField]
                  mapparray(rstemp2![MAPPNumber]) = mapparray(rstemp2![MAPPNumber]) + 1
                  rstemp2.MoveNext
               Wend
               
            End If
         rstemp.MoveNext
      Wend
   
      'you now have a table with every one of the GenMAPPs of the expression dataset linked to a MAPP.
      'need to count how many times each MAPP is represented.
      rstemp2.Close
      rstemp.Close
      
      For i = 0 To MAPPCount
      If mapparray(i) > 0 Then
         dbChipData.Execute ("INSERT INTO LocalMAPPsChip (MAPPName, OnChip) VALUES ('" _
                     & MAPPName(i) & "', " & mapparray(i) & ")")
      End If
      Next i
   
   'dbChipData.Execute ("DELETE * FROM GO")
   'keep the tempall and goall for mapp building.
 
End Sub

Public Function fixFileName(FileName As String) As String
   Dim newname As String
   
   newname = Left(FileName, Len(FileName) - 4)
   
   fixFileName = newname


End Function
Private Sub Close_Click()
    End
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormCode Then                                            'Exiting from mnuExit
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   
 Exit_Click

End Sub


Public Sub setSpecies(name As String)
   species = name
End Sub

Public Function getSpecies() As String
   getSpecies = species
End Function



Private Sub MAPPFinderhelp_Click()
  Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, programpath & "\GenMAPP.chm::/GenMAPP.htm", HH_DISPLAY_TOPIC, 0)
   
End Sub

Public Sub eraseForm()
   specieselected = False
   geneontology = False
   LocalMAPPsLoaded = False
   'Option1(0).Value = False
   'Option1(2).Value = False
   'Option1(4).Value = False
   'cmbSpecies.Text = "Select your species"
   Check1.Value = 0
   Check2.Value = 0
   chkStats.Value = 0
   Statistics = False
   txtFile.Text = ""
   lstColorSet.Clear
   lblProgress.Caption = ""
   lblspecies.Caption = ""
End Sub

Public Function MakeGOID(i As Long) As String
   'input a number
   'output a string with 7 characters. Leading zeros will be added until 7 is reached
   Dim GOID As String
   
   GOID = Str(i)
   If InStr(1, GOID, " ") Then
      GOID = Right(GOID, Len(GOID) - 1)
   End If
   Select Case Len(GOID)
      Case 1
         GOID = "000000" & GOID
      Case 2
         GOID = "00000" & GOID
      Case 3
         GOID = "0000" & GOID
      Case 4
         GOID = "000" & GOID
      Case 5
         GOID = "00" & GOID
      Case 6
         GOID = "0" & GOID
   End Select
   
   MakeGOID = GOID
End Function

Public Sub InitializeGOArray(GOarray() As Integer)
   Dim i As Long
   
   For i = 0 To GOCount
      GOarray(i) = 0
   Next i
End Sub

Public Sub calculateRandN(criterion As Integer)
   'N = the total number of genes in this dataset that are on a MAPP in this set of local MAPPs
   'R = the total number of genes meeting the user's criterion that are on a MAPP in this set of local MAPPs
   'these numbers are necessary for calculating the test stat.
   
   Dim rsFilter As Recordset
   Dim pg As New PrimaryGene
   Dim cg As New ClusterGene
   
   
   localR = 0
   localN = LocalClusterGenes.count - nomapp
   
   Set rsFilter = dbExpressionData.OpenRecordset("SELECT DISTINCT ID FROM" _
                        & " Expression WHERE (" & sql(criterion) & ")")
  
   While rsFilter.EOF = False
      Set pg = LocalPrimaryGenes.Item(rsFilter!id)
      For Each cg In pg.getClusterGenes
         localR = localR + 1
      Next cg
      rsFilter.MoveNext
   Wend
   
End Sub

Public Sub calculateTestStat()
   Dim goterm As New goterm
   Dim r As Long, N As Long
   Dim teststat As Double
   
   
   
   Set goterm = GOterms.Item("GO")  'GO node, the root of the tree
   goterm.setValues 'set #change/measured
   bigR = goterm.getChanged
   bigN = goterm.getOnChip
   
   
   For Each go In GOterms 'step through each results and calculate TestStat
      If go.getGOID = "GO" Then
         go.setZscore (0)
      Else
         go.setValues 'set #changed/measured
         r = go.getChanged
         N = go.getOnChip
         
         'this calculate the standard test statistic under the hypergeometric distribution
         'the number changed - the number expected to changed based on background divided by the stdev of the data
         If bigR - bigN = 0 Then
            go.setZscore (0)
         Else
            numer = r - (N * bigR / bigN)
            denom = Sqr(N * (bigR / bigN) * (1 - (bigR / bigN)) * (1 - (N - 1) / (bigN - 1)))
            If numer = 0 Then
               go.setZscore (0)
            Else
               go.setZscore (numer / denom)
            End If
         End If
      End If
   Next go
      
End Sub

Public Sub calculateTestStatLocal()
   Dim goterm As New goterm
   Dim r As Long, N As Long
   Dim teststat As Double
   Dim numer As Double
   Dim denom As Double
   
   localN = LocalClusterGenes.count
   For Each MAPP In localMAPPsCollection 'step through each results and calculate TestStat
      r = MAPP.getChanged
      N = MAPP.getOnChip
         
      'this calculate the standard test statistic under the hypergeometric distribution
      'the number changed - the number expected to changed based on background divided by the stdev of the data
      If localR - localN = 0 Then
         MAPP.setZscore (0)
      Else
         
         
         numer = r - (N * localR / localN)
         denom = Sqr(N * (localR / localN) * (1 - (localR / localN)) * (1 - (N - 1) / (localN - 1)))
         If numer = 0 Then
            MAPP.setZscore (0)
         Else
            MAPP.setZscore (numer / denom)
         End If
      End If
   Next MAPP
      
End Sub


Public Function mapToClusterSystem() As Boolean
   Dim rsSystem As Recordset, rsrelations As DAO.Recordset
   Dim i As Integer, numofsystems As Integer, record As Long
   Dim tblinfo As TableDef, tblMaptoCluster As TableDef
   Dim tblGenetoGO As TableDef
   Dim tblresults As TableDef, tblnestedresults As TableDef
   Dim rsGenes As Recordset, found As Boolean
   Dim GOIDs As Collection, clusterID As String
   Dim Cluster As ClusterGene, rsSecond As Recordset, sql As String
   Dim pipe As Integer, slash As Integer, idcolumn As String
   Dim relatedItems As New Collection
   Dim rsProbes As Recordset
   Dim relationNotFound As Boolean, primary As Boolean
   
   Set rsGenes = dbExpressionData.OpenRecordset _
                     ("SELECT ID, SystemCode FROM Expression")
   rsGenes.MoveLast
   rsGenes.MoveFirst
   GEXsize = rsGenes.RecordCount
  
   Set rsSystem = dbMAPPfinder.OpenRecordset("SELECT System, SystemCode from Systems" _
                                                & " WHERE [MOD] = '" & species & "'")
   If rsSystem.EOF = False Then
      clustersystem = rsSystem!system
      clustercode = rsSystem!systemcode
      Set rsRelation = dbMAPPfinder.OpenRecordset( _
                     "SELECT Relation FROM Relations WHERE SystemCode = '" & clustercode _
                     & "' AND RelatedCode = 'T'")
      If rsRelation.EOF = False Then
         GOrelation = rsRelation!Relation
         primary = False
      Else 'check the other direction
         Set rsRelation = dbMAPPfinder.OpenRecordset( _
                     "SELECT Relation FROM Relations WHERE SystemCode = 'T'" _
                     & " AND RelatedCode = '" & clustercode & "'")
         If rsRelation.EOF = False Then
            GOrelation = rsRelation!Relation
            primary = True
         Else
            MsgBox "You do not have a table from " & clustersystem & " to GO, you need this.", vbOKOnly
            mapToClusterSystem = False
            Exit Function
         End If
      End If
   Else
      MsgBox "The database " & databaseloc & " does not have the correct tables to run" _
            & " MAPPFinder for " & cmbSpecies.Text & ". Please check the species you selected" _
            & " and the database you are using. To change your database you must return to the" _
            & " start menu.", vbOKOnly
      mapToClusterSystem = False
      Exit Function
   End If
   'buildGotermCollection
   
   If Dir(newfilename & ".gdb") = "" Then
      'build the mappings to clusterID and GO
      'else load the mappings from the gdb file
      lblProgress.Caption = "Mapping expression data to GO."
      frmCriteria.Refresh
      DoEvents
      Set rsSystem = dbExpressionData.OpenRecordset("SELECT DISTINCT SystemCode from Expression")
      i = 0
      relationNotFound = False
      While rsSystem.EOF = False
         If clustercode = rsSystem!systemcode Then
            relations(i, 0) = clustersystem
            relations(i, 1) = "S"
            relations(i, 2) = rsSystem!systemcode
         Else
            'look for clustercode-EDcode relation
            Set rsRelation = dbMAPPfinder.OpenRecordset _
                     ("SELECT Relation FROM Relations WHERE SystemCode = '" & clustercode _
                     & "' AND RelatedCode = '" & rsSystem!systemcode & "'")
            If rsRelation.EOF = False Then 'found the relation
               relations(i, 0) = rsRelation!Relation
               relations(i, 1) = "R"
               relations(i, 2) = rsSystem!systemcode
            Else 'try the other way
               Set rsRelation = dbMAPPfinder.OpenRecordset _
                        ("SELECT Relation FROM Relations WHERE RelatedCode = '" & clustercode _
                        & "' AND SystemCode = '" & rsSystem!systemcode & "'")
               If rsRelation.EOF = False Then 'found the relation
                  relations(i, 0) = rsRelation!Relation
                  relations(i, 1) = "P"
                  relations(i, 2) = rsSystem!systemcode
               Else 'no relation exists
                  relationNotFound = True
                  relations(i, 0) = "No relation"
                  relations(i, 1) = "N"
                  relations(i, 2) = rsSystem!systemcode
               End If
            End If
         End If
         If relationNotFound Then
            MsgBox "No relation exists between the system code " & rsSystem!systemcode _
               & " and " & clustersystem & ". MAPPFinder can not use this system." _
               & " Check the system code, or add a relation to your database. MAPPFinder" _
               & " will continue to calculate the results with any other systems that exist.", vbOKOnly
            GoTo endwhile
         End If
         Set rsSecond = dbMAPPfinder.OpenRecordset("SELECT Columns FROM Systems WHERE" _
                                             & " SystemCode = '" & rsSystem!systemcode & "'")
         slash = InStr(1, rsSecond!columns, "\S", vbTextCompare)
         Do While slash
            pipe = InStrRev(rsSecond!columns, "|", slash)
            relations(i, 3) = relations(i, 3) _
                                   & Mid(rsSecond!columns, pipe, slash - pipe + 2) & "|"
            slash = InStr(slash + 1, rsSecond!columns, "\S", vbTextCompare)
         Loop
         rsSystem.MoveNext
         i = i + 1
         relationNotFound = False
      Wend
      numofsystems = i
      i = 0
      'now I need to map each gene to the cluster system using the relations we just extracted
      
      record = 0
      While rsGenes.EOF = False
         record = record + 1

         If record Mod 10 = 0 Then
            lblProgress.Caption = " Mapping complete for " & record & " out of " & GEXsize
            frmCriteria.Refresh
            DoEvents
         End If
         While i < numofsystems And found = False
            If relations(i, 2) = rsGenes!systemcode Then
               found = True
            Else
               i = i + 1
            End If
         Wend
         If i < numofsystems Then 'you found a match, this system is supported, this should never fail as you already caught this error above
            Select Case relations(i, 1)
               Case "S" 'the codes are the same
                  Set rsrelations = dbMAPPfinder.OpenRecordset _
                                    ("SELECT ID FROM [" & relations(i, 0) & "] WHERE " _
                                    & " ID = '" & rsGenes!id & "'")
                  If rsrelations.EOF Then 'try any secondary IDs
                     If relations(i, 3) <> "" Then '___________________________________Check Secondary IDs
                        slash = InStr(1, relations(i, 3), "\S", vbTextCompare)
                        Do While slash
                           pipe = InStrRev(relations(i, 3), "|", slash)
                           idcolumn = Mid(relations(i, 3), pipe + 1, slash - pipe - 1)
                           Set rsSecond = dbMAPPfinder.OpenRecordset("SELECT System FROM Systems WHERE " _
                                                               & "SystemCode = '" & relations(i, 2) & "'")
                           If Mid(relations(i, 3), slash + 1, 1) = "s" Then              'Single ID, eg: P123
                              sql = "SELECT ID FROM " & rsSecond!system & _
                                    "   WHERE [" & idcolumn & "] = '" & rsGenes!id & "'"
                           Else                                             'Multiple IDs, eg: |P123|P456|P789|
                              sql = "SELECT ID FROM " & rsSecond!system & _
                                   "   WHERE [" & idcolumn & "] LIKE '*|" & rsGenes!id & "|*'"
                           End If
                           Set rsGeneID = dbMAPPfinder.OpenRecordset(sql)
                           If Not rsGeneID.EOF Then
                              Exit Do                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                           End If
                           slash = InStr(slash + 1, relations(i, 3), "\S", vbTextCompare)
                        Loop
                        If rsGeneID.EOF = False Then
                           'you found a primary ID, try linking it to a cluster ID.
                           Set rsrelations = dbMAPPfinder.OpenRecordset _
                                       ("SELECT ID FROM [" & relations(i, 0) & "] WHERE " _
                                       & " ID = '" & rsGeneID!id & "'")
                           If rsrelations.EOF = False Then
                              addtoPrimaryGenes rsGenes!id, addtoclustergenes(rsGenes!id, rsGenes!id)
                           Else
                              addNotFoundtoPrimaryGenes rsGenes!id
                              noCluster = noCluster + 1
                           End If
                        Else
                           'no primary ID for this secondary
                           addNotFoundtoPrimaryGenes rsGenes!id
                           noCluster = noCluster + 1
                        End If
                     Else ' no secondary
                        addNotFoundtoPrimaryGenes rsGenes!id
                        noCluster = noCluster + 1
                     End If
                  Else 'you found it the first time
                     addtoPrimaryGenes rsGenes!id, addtoclustergenes(rsGenes!id, rsGenes!id)
                  End If
               Case "P" 'the GEX code is the primary of the relationship
                  Set rsrelations = dbMAPPfinder.OpenRecordset _
                                    ("SELECT Related FROM [" & relations(i, 0) & "] WHERE " _
                                    & "Primary = '" & rsGenes!id & "'")
                  
                  If rsrelations.EOF Then 'try any secondary IDs
                     If relations(i, 3) <> "" Then '___________________________________Check Secondary IDs
                     slash = InStr(1, relations(i, 3), "\S", vbTextCompare)
                     Do While slash
                        pipe = InStrRev(relations(i, 3), "|", slash)
                        idcolumn = Mid(relations(i, 3), pipe + 1, slash - pipe - 1)
                        Set rsSecond = dbMAPPfinder.OpenRecordset("SELECT System FROM Systems WHERE " _
                                                            & "SystemCode = '" & relations(i, 2) & "'")
                        If Mid(relations(i, 3), slash + 1, 1) = "s" Then              'Single ID, eg: P123
                           sql = "SELECT ID FROM " & rsSecond!system & _
                                 "   WHERE [" & idcolumn & "] = '" & rsGenes!id & "'"
                        Else                                             'Multiple IDs, eg: |P123|P456|P789|
                           sql = "SELECT ID FROM " & rsSecond!system & _
                                "   WHERE [" & idcolumn & "] LIKE '*|" & rsGenes!id & "|*'"
                        End If
                        Set rsGeneID = dbMAPPfinder.OpenRecordset(sql)
                        If Not rsGeneID.EOF Then
                           Exit Do                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                        End If
                        slash = InStr(slash + 1, relations(i, 3), "\S", vbTextCompare)
                     Loop
                     If rsGeneID.EOF = False Then
                        'you found a primary ID, try linking it to a cluster ID.
                        Set rsrelations = dbMAPPfinder.OpenRecordset _
                                    ("SELECT Related FROM [" & relations(i, 0) & "] WHERE " _
                                    & "Primary = '" & rsGeneID!id & "'")
                        If rsrelations.EOF = False Then
                           While rsrelations.EOF = False
                              addtoPrimaryGenes rsGenes!id, addtoclustergenes(rsGenes!id, rsrelations!related)
                              rsrelations.MoveNext
                           Wend
                        Else
                           addNotFoundtoPrimaryGenes rsGenes!id
                           noCluster = noCluster + 1
                        End If
                     Else
                        'no primary ID for this secondary
                        addNotFoundtoPrimaryGenes rsGenes!id
                        noCluster = noCluster + 1
                     End If
                     Else ' no secondary
                     addNotFoundtoPrimaryGenes rsGenes!id
                     noCluster = noCluster + 1
                     End If
                  Else 'you found it the first time
                     While rsrelations.EOF = False
                        addtoPrimaryGenes rsGenes!id, addtoclustergenes(rsGenes!id, rsrelations!related)
                        rsrelations.MoveNext
                     Wend
                  End If
               Case "R" 'the gex code is the related of the relationship
                  Set rsrelations = dbMAPPfinder.OpenRecordset _
                                    ("SELECT Primary FROM [" & relations(i, 0) & "] WHERE " _
                                    & "Related = '" & rsGenes!id & "'")
                  
                  If rsrelations.EOF Then 'try any secondary IDs
                     If relations(i, 3) <> "" Then '___________________________________Check Secondary IDs
                     slash = InStr(1, relations(i, 3), "\S", vbTextCompare)
                     Do While slash
                        pipe = InStrRev(relations(i, 3), "|", slash)
                        idcolumn = Mid(relations(i, 3), pipe + 1, slash - pipe - 1)
                        Set rsSecond = dbMAPPfinder.OpenRecordset("SELECT System FROM Systems WHERE " _
                                                            & "SystemCode = '" & relations(i, 2) & "'")
                        If Mid(relations(i, 3), slash + 1, 1) = "s" Then              'Single ID, eg: P123
                           sql = "SELECT ID FROM " & rsSecond!system & _
                                 "   WHERE [" & idcolumn & "] = '" & rsGenes!id & "'"
                        Else                                             'Multiple IDs, eg: |P123|P456|P789|
                           sql = "SELECT ID FROM " & rsSecond!system & _
                                "   WHERE [" & idcolumn & "] LIKE '*|" & rsGenes!id & "|*'"
                        End If
                        Set rsGeneID = dbMAPPfinder.OpenRecordset(sql)
                        If Not rsGeneID.EOF Then
                           Exit Do                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                        End If
                        slash = InStr(slash + 1, relations(i, 3), "\S", vbTextCompare)
                     Loop
                     If rsGeneID.EOF = False Then
                        'you found a primary ID, try linking it to a cluster ID.
                        Set rsrelations = dbMAPPfinder.OpenRecordset _
                                    ("SELECT Primary FROM [" & relations(i, 0) & "] WHERE " _
                                    & "Related = '" & rsGeneID!id & "'")
                        If rsrelations.EOF = False Then
                           While rsrelations.EOF = False
                              addtoPrimaryGenes rsGenes!id, addtoclustergenes(rsGenes!id, rsrelations!primary)
                              rsrelations.MoveNext
                           Wend
                        Else
                           addNotFoundtoPrimaryGenes rsGenes!id
                           noCluster = noCluster + 1
                        End If
                     Else ' no secondary
                        addNotFoundtoPrimaryGenes rsGenes!id
                        noCluster = noCluster + 1
                     End If
                     Else
                        'no primary ID for this secondary
                        addNotFoundtoPrimaryGenes rsGenes!id
                        noCluster = noCluster + 1
                     End If
                  Else 'you found it the first time
                     While rsrelations.EOF = False
                        addtoPrimaryGenes rsGenes!id, addtoclustergenes(rsGenes!id, rsrelations!primary)
                        rsrelations.MoveNext
                     Wend
                  End If
               Case "N"
                  addNotFoundtoPrimaryGenes rsGenes!id
                  noCluster = noCluster + 1
            End Select
            
         Else 'this system isn't support
            addNotFoundtoPrimaryGenes rsGenes!id
            noCluster = noCluster + 1
         End If
        
         
         i = 0
endwhile:
         found = False
         
         rsGenes.MoveNext
      Wend
      
     
      
      ' at the end of this while loop we have
      ' primarygenes - a collection of primary gene objects. The key is the ID in the gex
      ' each primarygene object stores all of the related clustersystem IDs for that primaryID
      ' clusterGenes - a collection of unique clusterGene objects
      '- each element contains an object containing a collection of GOIDs if any.
      ' noCluster - a count of the number of primary IDs that did not link to the cluster system
      
       buildNestedGO
      'now for each gene we have its full(nested) path in the GO
      
      'I will write out all of the collections to a DB table later.
      'use DB because easier to see what's being linked to what and is consistent with Genmapp
      'create the database that this will be stored into.
      
      lblProgress.Caption = "Storing the GO Mappings. This will only take a minute."
      frmCriteria.Refresh
      DoEvents
      
      Set dbChipData = CreateDatabase(newfilename & ".gdb", dbLangGeneral)
      Set tblinfo = dbChipData.CreateTableDef("Info")
         With tblinfo
            .Fields.Append .CreateField("GoTable", dbText, 25)
            .Fields.Append .CreateField("Version", dbText, 15)
         End With
      dbChipData.TableDefs.Append tblinfo
      Set tblCluster = dbChipData.CreateTableDef("PrimarytoClusterSystem")
      With tblCluster
         .Fields.Append .CreateField("Primary", dbText, IDSIZE)
         .Fields.Append .CreateField("Related", dbText, IDSIZE)
      End With
      dbChipData.TableDefs.Append tblCluster
      Set tblgo = dbChipData.CreateTableDef("ClusterSystemtoGO")
      tblgo.Fields.Append tblgo.CreateField("Primary", dbText, IDSIZE)
      tblgo.Fields.Append tblgo.CreateField("Related", dbText, IDSIZE)
      dbChipData.TableDefs.Append tblgo
      
      Set tblgo = dbChipData.CreateTableDef("NestedClusterSystemtoGO")
      tblgo.Fields.Append tblgo.CreateField("Primary", dbText, IDSIZE)
      tblgo.Fields.Append tblgo.CreateField("Related", dbText, IDSIZE)
      dbChipData.TableDefs.Append tblgo
      
      Set tblGenetoGO = dbChipData.CreateTableDef("GOIDtoGene")
      tblGenetoGO.Fields.Append tblGenetoGO.CreateField("Primary", dbText, IDSIZE)
      tblGenetoGO.Fields.Append tblGenetoGO.CreateField("Related", dbText, IDSIZE)
      dbChipData.TableDefs.Append tblGenetoGO
      
      Set tblnestedresults = dbChipData.CreateTableDef("NestedResults")
      With tblnestedresults
         .Fields.Append .CreateField("GOType", dbText, 2)
         .Fields.Append .CreateField("GOName", dbMemo)
         Dim idxNestedResults As Index
         Dim idxPercentage As Index
         .Fields.Append .CreateField("GOID", dbText, 15)
         Set idxPercentage = .CreateIndex("percentageindex")
         Set idxNestedResults = .CreateIndex("nestedResults")
         idxNestedResults.primary = True
         idxNestedResults.Fields.Append .CreateField("GOID", dbText, 15)
         idxPercentage.Fields.Append .CreateField("GOType", dbText, 2)
         idxPercentage.Fields.Append .CreateField("Percentage", dbSingle)
         idxPercentage.Fields.Append .CreateField("Zscore", dbSingle)
         .Indexes.Append idxNestedResults
         .Indexes.Append idxPercentage
         .Fields.Append .CreateField("ChangedLocal", dbSingle)
         .Fields.Append .CreateField("OnChipLocal", dbSingle)
         .Fields.Append .CreateField("InGOLocal", dbSingle)
         .Fields.Append .CreateField("PercentageLocal", dbSingle)
         .Fields.Append .CreateField("PresentLocal", dbSingle)
         .Fields.Append .CreateField("Changed", dbSingle)
         .Fields.Append .CreateField("OnChip", dbSingle)
         .Fields.Append .CreateField("InGO", dbSingle)
         .Fields.Append .CreateField("Percentage", dbSingle)
         .Fields.Append .CreateField("Present", dbSingle)
         .Fields.Append .CreateField("Zscore", dbDouble)
         .Fields.Append .CreateField("permuteP", dbDouble)
         .Fields.Append .CreateField("adjustedP", dbDouble)
      End With
      Set tblresults = dbChipData.CreateTableDef("Results")
         With tblresults
            .Fields.Append .CreateField("GOType", dbText, 2)
            Dim idxResults As Index
            .Fields.Append .CreateField("GOID", dbText, 255)
            .Fields.Append .CreateField("GOName", dbMemo)
            Set idxResults = .CreateIndex("Results")
            idxResults.Fields.Append .CreateField("GOType", dbText, 2)
            idxResults.Fields.Append .CreateField("Zscore", dbSingle)
            idxResults.Fields.Append .CreateField("Changed", dbSingle)
            .Indexes.Append idxResults
            .Fields.Append .CreateField("Changed", dbSingle)
            .Fields.Append .CreateField("OnChip", dbSingle)
            .Fields.Append .CreateField("InGO", dbSingle)
            .Fields.Append .CreateField("Percentage", dbSingle)
            .Fields.Append .CreateField("Present", dbSingle)
            .Fields.Append .CreateField("Zscore", dbDouble)
            .Fields.Append .CreateField("permuteP", dbDouble)
            .Fields.Append .CreateField("adjustedP", dbDouble)
         End With
      dbChipData.TableDefs.Append tblresults
      dbChipData.TableDefs.Append tblnestedresults
      
      
      
      Set rstemp = dbMAPPfinder.OpenRecordset("Select Version from Info")
  
   
      dbChipData.Execute "INSERT INTO Info (GOTable, Version) VALUES ('" & clustersystem _
                     & "', '" & rstemp!Version & "')"
      For Each Gene In PrimaryGenes
         Set relatedItems = Gene.getClusterGenes
         If relatedItems.count = 0 Then 'no cluster gene for this guy
            dbChipData.Execute "INSERT INTO PRimarytoClusterSystem(Primary)" _
                       & " VALUES ('" & Gene.getID & "')"
         Else
            For Each related In relatedItems
               dbChipData.Execute "INSERT INTO PrimarytoClusterSystem (Primary, Related)" _
                              & " VALUES ('" & Gene.getID & "', '" & related.getID & "')"
            Next related
         End If
         
      Next Gene
      DoEvents
      For Each Cluster In ClusterGenes
         Set terms = Cluster.getGOTerms
         clusterID = Cluster.getID
         If terms.count = 0 Then 'no GOIDs for this cluster
            dbChipData.Execute "INSERT INTO ClusterSystemtoGO(Primary) VALUES" _
                              & "('" & clusterID & "')"
         Else
            For Each term In terms
               dbChipData.Execute "INSERT INTO ClustersystemtoGO(Primary, Related) VALUES ('" _
                              & clusterID & "', '" & term.getGOID & "')"
            Next term
         End If
         Set terms = Cluster.getNestedGOterms
         clusterID = Cluster.getID
         For Each term In terms
            dbChipData.Execute "INSERT INTO NestedClustersystemtoGO(Primary, Related) VALUES ('" _
                              & clusterID & "', '" & term.getGOID & "')"
         
         Next term
      Next Cluster
      
      For Each goterm In GOterms
         Set relatedItems = goterm.getGenes
         For Each Gene In relatedItems
            dbChipData.Execute "INSERT INTO GOIDtoGene(Primary, Related) VALUES" _
                     & "('" & goterm.getGOID & "', '" & Gene & "')"
         Next Gene
      Next goterm
      Dim idxGOID As Index
      Set idxGOID = tblGenetoGO.CreateIndex("Results")
      idxGOID.Fields.Append tblGenetoGO.CreateField("Primary", dbText, 15)
      tblGenetoGO.Indexes.Append idxGOID
      'now we have the GenesInME for each GOterm in a table. Indexed by GOID.
      DoEvents
      'ok, so now there are three tables. (PrimaryToClusterSytstem) and ClustertoGO and nestedclustertogo.
      'neither table is indexed, as I will only be calling select * from these tables.
      '
   '*******************************************************************************
   'END OF BUILDING THE MAPPINGS
   '*******************************************************************************
   Else 'the mapping has been done before just read in the files.
      'build the mappings to clusterID and GO
      
      frmCriteria.lblProgress.Caption = "Loading GO mappings from " & newfilename & ".gdb"
      frmCriteria.Refresh
      DoEvents
   
      Set dbChipData = OpenDatabase(newfilename & ".gdb")
      Set rstemp = dbMAPPfinder.OpenRecordset("Select Version from INFO")
      Set rstemp2 = dbChipData.OpenRecordset("SELECT GOtable, Version FROM Info")
      If clustersystem <> rstemp2!gotable Then 'they are using the wrong database
         MsgBox "The GDB file, which stores the mapping of your expression data to GO, was built" _
               & " using " & rstemp2!gotable & ". Your current database suggests that " & clustersystem _
               & " should be used. Please check this discrepancy and either select a new database" _
               & " or delete your GDB file and start over."
         mapToClusterSystem = False
         Exit Function
      ElseIf rstemp!Version <> rstemp2!Version Then
         'need to rebuild
         If MsgBox("The MAPPFinder gdb file storing the links from your Expression Dataset" _
            & " to Gene Ontology was built on a different version of the GenMAPP database." _
            & " To proceed using " & databaseloc & " MAPPFinder will need to rebuild the GDB" _
            & " file. The new file will not be reverse compatible with the other version of the database" _
            & ". Do you want to proceed?", vbYesNo) = vbNo Then
            frmStart.Show
            eraseForm
            Me.Hide
            Exit Function
            Else 'they want to rebuild the gdb file
         
            dbChipData.Close
            lblProgress.Caption = "Your gdb file is out of date. MAPPFinder is now going to rebuild it."
            frmCriteria.Refresh
            DoEvents
         
            Kill newfilename & ".gdb"
            mapToClusterSystem = mapToClusterSystem()
            
            Exit Function
         End If
         
      Else 'they match load the data
         
         Set rsGenes = dbChipData.OpenRecordset("SELECT Primary, Related FROM ClusterSystemtoGO" _
                                                & " ORDER BY Primary")
         While rsGenes.EOF = False
            If IsNull(rsGenes!related) Then
               addClusterID (rsGenes!primary)
             
            Else
               addClusterPair rsGenes!primary, rsGenes!related
            End If
            rsGenes.MoveNext
            
         Wend
         DoEvents
         Set rsGenes = dbChipData.OpenRecordset("SELECT Primary, Related FROM NestedClusterSystemtoGO" _
                                                & " ORDER BY Primary")
         While rsGenes.EOF = False
            addnestedClusterPair rsGenes!primary, rsGenes!related
            rsGenes.MoveNext
         Wend
         DoEvents
         Set rsGenes = dbChipData.OpenRecordset("SELECT Primary, Related FROM " _
                                             & "PrimaryToClusterSystem")
         While rsGenes.EOF = False
            If IsNull(rsGenes!related) Then
               addNotFoundtoPrimaryGenes (rsGenes!primary)
               ' if two probes both have the same ID, and this ID has no
               'related ClusterID then the number of probes not linked to
               'the cluster system will be off by one. Since the primarygenes
               'collection only has one copy, I'll need to query the expression table.
               'this is a hack, but it will work
               Set rsProbes = dbExpressionData.OpenRecordset("SELECT ID from Expression where ID = '" _
                                                            & rsGenes!primary & "'")
               rsProbes.MoveLast
               noCluster = noCluster + rsProbes.RecordCount
                                        
            Else
               addtoPrimaryGenes rsGenes!primary, ClusterGenes.Item(rsGenes!related)
               For Each goterm In ClusterGenes.Item(rsGenes!related).getGOTerms()
                  goterm.addGeneLocal (rsGenes!primary)
               Next goterm
               For Each goterm In ClusterGenes.Item(rsGenes!related).getNestedGOterms()
                  goterm.addGene (rsGenes!primary)
               Next goterm
            End If
            rsGenes.MoveNext
         Wend
      
      End If
      DoEvents
   End If
   mapToClusterSystem = True
End Function
Public Sub addtoPrimaryGenes(primaryID As String, clusterID As ClusterGene)
  On Error GoTo error
  Dim NewGene As New PrimaryGene
   'check to see if this gene exists in the collection
   Set NewGene = PrimaryGenes.Item(primaryID)
GeneExists:
   'Debug.Print ClusterID.getID
   NewGene.addClusterGene clusterID
   
   
error:
   Select Case Err.Number
      Case 5
         'the gene doesn't exist so add it
         NewGene.setID (primaryID)
         PrimaryGenes.Add NewGene, primaryID
         Resume GeneExists
   End Select
End Sub

Public Sub addtoLocalPrimaryGenes(primaryID As String, clusterID As ClusterGene)
  On Error GoTo error
   Dim NewGene As New PrimaryGene
   'check to see if this gene exists in the collection
   Set NewGene = LocalPrimaryGenes.Item(primaryID)
GeneExists:
   'Debug.Print ClusterID.getID
   If clusterID.getGOTerms.count > 0 Then 'don't bother adding this if it doesn't link to a MAPP
      NewGene.addClusterGene clusterID
   End If
   
error:
   Select Case Err.Number
      Case 5
         'the gene doesn't exist so add it
         NewGene.setID (primaryID)
         LocalPrimaryGenes.Add NewGene, primaryID
         Resume GeneExists
   End Select
End Sub

 Public Sub addNotFoundtoPrimaryGenes(primaryID As String)
'primarygenes should have all of the primaryIDs in it, even if they don't link to the clustersystem.
'for all of those genes we add them, but their clustergene collections are empty.
   On Error GoTo error
   Dim NewGene As New PrimaryGene
   'check to see if this gene exists in the collection
   Set NewGene = PrimaryGenes.Item(primaryID)
   
error:
   Select Case Err.Number
      Case 5
         'the gene doesn't exist so add it
         NewGene.setID (primaryID)
         PrimaryGenes.Add NewGene, primaryID
   End Select
 End Sub

Public Sub addNotFoundtoLocalPrimaryGenes(primaryID As String)
'primarygenes should have all of the primaryIDs in it, even if they don't link to the clustersystem.
'for all of those genes we add them, but their clustergene collections are empty.
   On Error GoTo error
   Dim NewGene As New PrimaryGene
   'check to see if this gene exists in the collection
   Set NewGene = LocalPrimaryGenes.Item(primaryID)
   
error:
   Select Case Err.Number
      Case 5
         'the gene doesn't exist so add it
         NewGene.setID (primaryID)
         LocalPrimaryGenes.Add NewGene, primaryID
   End Select
 End Sub

Public Sub addClusterPair(clusterID As String, GOID As String)
   On Error GoTo error
   Dim NewCG As New ClusterGene
   'check to see if this gene exists in the collection
   Set NewCG = ClusterGenes.Item(clusterID)
ClusterExists:
   NewCG.addLocalGOterm GOID, addtoLocalGOterms(GOID)
   
   
error:
   Select Case Err.Number
      Case 5
         'the gene doesn't exist so add it
        
         NewCG.setID clusterID
         ClusterGenes.Add NewCG, clusterID
         Resume ClusterExists
   End Select
End Sub
Public Sub addLocalClusterPair(clusterID As String, GOID As String)
   On Error GoTo error
   Dim NewCG As New ClusterGene
   'check to see if this gene exists in the collection
   Set NewCG = LocalClusterGenes.Item(clusterID)
ClusterExists:
   NewCG.addLocalGOterm GOID, addtoLocalMAPPs(GOID)
   
   
error:
   Select Case Err.Number
      Case 5
         'the gene doesn't exist so add it
        
         NewCG.setID clusterID
         LocalClusterGenes.Add NewCG, clusterID
         Resume ClusterExists
   End Select
End Sub
Public Sub addClusterID(clusterID As String)
   On Error GoTo error
   Dim NewCG As New ClusterGene
   'check to see if this gene exists in the collection
   Set NewCG = ClusterGenes.Item(clusterID)
ClusterExists:
   Exit Sub
   
error:
   Select Case Err.Number
      Case 5
         'the gene doesn't exist so add it
        
         NewCG.setID clusterID
         ClusterGenes.Add NewCG, clusterID
   End Select
End Sub
Public Sub addLocalClusterID(clusterID As String)
   On Error GoTo error
   Dim NewCG As New ClusterGene
   'check to see if this gene exists in the collection
   Set NewCG = ClusterGenes.Item(clusterID)
ClusterExists:
   Exit Sub
   
error:
   Select Case Err.Number
      Case 5
         'the gene doesn't exist so add it
        
         NewCG.setID clusterID
         LocalClusterGenes.Add NewCG, clusterID
   End Select
End Sub

'this seems really redundant and it probably is, but it was decided by everyone that they
'wanted both the nested and non-nested numbers. so we do everything twice.
Public Sub addnestedClusterPair(clusterID As String, GOID As String)
   On Error GoTo error
   Dim NewCG As New ClusterGene
   'check to see if this gene exists in the collection
   Set NewCG = ClusterGenes.Item(clusterID)
ClusterExists:
   NewCG.addNestedGOterm GOID, addtoNestedGOterms(GOID)
   
error:
   Select Case Err.Number
      Case 5
         'the gene doesn't exist so add it
         
         NewCG.setID clusterID
         ClusterGenes.Add NewCG, clusterID
         Resume ClusterExists
   End Select
End Sub


Public Function addtoclustergenes(primaryID As String, clusterID As String) As ClusterGene
   On Error GoTo error
   Dim found As Boolean
   Dim id As ClusterGene
   Dim rsGO As Recordset
   Dim term As goterm
   Set id = ClusterGenes.Item(clusterID)
   id.addPrimaryID (primaryID)
   Set addtoclustergenes = id
   'this will either return an ID or throw an error. If it throws and error, then this is a new
   'gene. If it doesn't, then this ID is a duplicate and we can ignore it.
   Exit Function
error:
   Select Case Err.Number
      Case 5
         Dim cg As New ClusterGene
         cg.setID clusterID
         cg.addPrimaryID (primaryID)
         If primary Then
            Set rsGO = dbMAPPfinder.OpenRecordset("SELECT Primary FROM [" _
                                       & GOrelation & "] WHERE related = '" & clusterID & "'")
         Else
            Set rsGO = dbMAPPfinder.OpenRecordset("SELECT Related FROM [" _
                                       & GOrelation & "] WHERE Primary = '" & clusterID & "'")
         End If
         If rsGO.EOF = False Then 'this ID is on a MAPP and should be stored
         
            While rsGO.EOF = False
               cg.addGOterm rsGO!related, addtoGOterms(rsGO!related, primaryID)
               rsGO.MoveNext
            Wend
           
         End If
         ClusterGenes.Add cg, clusterID
         Set addtoclustergenes = cg
   End Select
End Function

Public Function addToLocalClusterGenes(idIn As String, nohits As Boolean, primaryID As String) As ClusterGene
On Error GoTo error:
   Dim id As ClusterGene
   Dim rsMAPPs As Recordset
   Set id = LocalClusterGenes.Item(idIn)
   For Each MAPP In id.getGOTerms
      MAPP.addGene (primaryID)
   Next MAPP
   'if this returns something then this ID has been seen before, so return it
   Set addToLocalClusterGenes = id
   
error:
   Select Case Err.Number
      Case 5
         Set rsMAPPs = dbLocalMAPPs.OpenRecordset("SELECT MAPP From GeneToMAPP " _
                                                   & "WHERE ID = '" & idIn & "'")
         Dim cg As New ClusterGene
         cg.setID idIn
         If Not rsMAPPs.EOF Then
         While rsMAPPs.EOF = False
            nohits = False
            cg.addLocalGOterm rsMAPPs!MAPP, addtoLocalMAPPs(rsMAPPs!MAPP, primaryID)
            rsMAPPs.MoveNext
         Wend
         LocalClusterGenes.Add cg, idIn 'only add this CG if it's on a MAPP.
                                           'localClusterGenes.count = N for Z scores
         End If
         Set addToLocalClusterGenes = cg
   End Select

End Function

Public Function addtoGOterms(GOID As String, geneId As String) As goterm
On Error GoTo error
   Dim goterm As New goterm
   Dim rsGOcount As Recordset
   Dim GONode As Node
   
   Set goterm = GOterms.Item(GOID)
   If (goterm.addGeneLocal(geneId)) Then
      goterm.setOnChip (goterm.getOnChip + 1)
      goterm.setOnChipLocal (goterm.getOnChipLocal + 1)
   End If
   Set addtoGOterms = goterm
   'this will either return an ID or throw an error. If it throws and error, then this is a new
   'go term. If it doesn't, then this ID is a duplicate and we can ignore it.
   
error:
   Select Case Err.Number
      Case 5
         goterm.setGOID GOID
         Set rsGOcount = dbMAPPfinder.OpenRecordset("SELECT Count, total FROM [" _
                                       & clustersystem & "-GOCount] WHERE GO = '" & GOID & "'")
         If rsGOcount.EOF Then
            goterm.setingo 0
            goterm.setingolocal 0
         Else
            goterm.setingolocal rsGOcount!count
            goterm.setingo rsGOcount!total
         End If
         goterm.setOnChip (1)
         goterm.setOnChipLocal (1)
         goterm.addGene (geneId)
         GOterms.Add goterm, GOID
         
   End Select
   Set addtoGOterms = goterm
End Function

Public Function addtoLocalGOterms(GOID As String) As goterm
On Error GoTo error
   Dim goterm As New goterm
   Dim rsGOcount As Recordset
   Dim GONode As Node
   
   Set goterm = GOterms.Item(GOID)
   goterm.setOnChipLocal (goterm.getOnChipLocal + 1)
   Set addtoLocalGOterms = goterm
   'this will either return an ID or throw an error. If it throws and error, then this is a new
   'go term. If it doesn't, then this ID is a duplicate and we can ignore it.
   
error:
   Select Case Err.Number
      Case 5
         goterm.setGOID GOID
         Set rsGOcount = dbMAPPfinder.OpenRecordset("SELECT Count, total FROM [" _
                                       & clustersystem & "-GOCount] WHERE GO = '" & GOID & "'")
         If rsGOcount.EOF Then
            goterm.setingo 0
            goterm.setingolocal 0
         Else
            goterm.setingolocal rsGOcount!count
            goterm.setingo rsGOcount!total
         End If
      
         goterm.setOnChipLocal (1)
         GOterms.Add goterm, GOID
         
   End Select
   Set addtoLocalGOterms = goterm
End Function

Public Function addtoNestedGOterms(GOID As String) As goterm
On Error GoTo error
   Dim goterm As New goterm
   Dim rsGOcount As Recordset
   Dim GONode As Node
   
   Set goterm = GOterms.Item(GOID)
   goterm.setOnChip (goterm.getOnChip + 1)
   Set addtoNestedGOterms = goterm
   'this will either return an ID or throw an error. If it throws and error, then this is a new
   'go term. If it doesn't, then this ID is a duplicate and we can ignore it.
   
error:
   Select Case Err.Number
      Case 5
         goterm.setGOID GOID
         Set rsGOcount = dbMAPPfinder.OpenRecordset("SELECT Count, total FROM [" _
                                       & clustersystem & "-GOCount] WHERE GO = '" & GOID & "'")
         If rsGOcount.EOF Then
            goterm.setingo 0
            goterm.setingolocal 0
         Else
            goterm.setingolocal rsGOcount!count
            goterm.setingo rsGOcount!total
         End If
         goterm.setOnChip (1)
         GOterms.Add goterm, GOID
         
   End Select
   Set addtoNestedGOterms = goterm
End Function






Public Function addtoLocalMAPPs(MAPPName As String, Optional primaryID As String) As goterm
On Error GoTo error:
   Dim MAPP As goterm
   Dim rsMAPPCount As Recordset
   
   Set MAPP = localMAPPsCollection.Item(MAPPName)
   'if this worked, then this MAPP has been seen before. Great. Return it.
   MAPP.setOnChip (MAPP.getOnChip + 1)
   If primaryID <> "" Then
      MAPP.addGene (primaryID)
   End If
   Set addtoLocalMAPPs = MAPP
   Exit Function

error:
   Select Case Err.Number
      Case 5
         Set MAPP = New goterm
         MAPP.setGOID (MAPPName)
         Set rsMAPPCount = dbLocalMAPPs.OpenRecordset("SELECT MAPPCount from GeneTOMAPPCount WHERE" _
                                                      & " MAPPName = '" & MAPPName & "'")
         If rsMAPPCount.EOF Then
            MAPP.setingo (0)
         Else
            MAPP.setingo (rsMAPPCount!MAPPCount)
         End If
         MAPP.setOnChip (1)
         If primaryID <> "" Then
            MAPP.addGene (primaryID)
         End If
         localMAPPsCollection.Add MAPP, MAPPName
         Set addtoLocalMAPPs = MAPP 'return the new MAPP object you added
   End Select
End Function


Public Function addtoGOtermsNested(GOID As String) As goterm
On Error GoTo error
   Dim goterm As New goterm
   Dim rsGOcount As Recordset
   
   Set goterm = GOterms.Item(GOID)
   goterm.setOnChip (goterm.getOnChip + 1)
   
   'goterm.setOnChipLocal (goterm.getOnChipLocal + 1)
   Set addtoGOtermsNested = goterm
   'this will either return an ID or throw an error. If it throws and error, then this is a new
   'go term. If it doesn't, then this ID is a duplicate and we can ignore it.
   
error:
   Select Case Err.Number
      Case 5
         goterm.setGOID GOID
         Set rsGOcount = dbMAPPfinder.OpenRecordset("SELECT Count, total FROM [" _
                                       & clustersystem & "-GOCount] WHERE GO = '" & GOID & "'")
         If rsGOcount.EOF Then
            goterm.setingo 0
            goterm.setingolocal 0
         Else
            goterm.setingolocal rsGOcount!count
            goterm.setingo rsGOcount!total
         End If
         goterm.setOnChip (1)
         goterm.setOnChipLocal (0)
         GOterms.Add goterm, GOID
         
   End Select
   Set addtoGOtermsNested = goterm
End Function

Public Sub buildNestedGO()
   'Each clustergene is visited and for each goterm of that cluster genes, this cluster gene is added to
   'the go terms entire path.
   On Error GoTo error
   Dim GOs As Collection
   Dim GONode As Node
   Dim parent As Node
   Dim GOID As String, parentID As String
   Dim counter As String, IDs As String
   Dim rsGOIDs As Recordset, j As Integer
   Dim GOcollection As Collection
   Dim parentterm As goterm
   Dim parents As Collection
   Dim root As goterm
   Dim i As Integer
   Dim cgcount As Long
   Set root = addtoGOtermsNested("GO")
   lblProgress.Caption = "Building Nested GO relationships."
   frmCriteria.Refresh
   DoEvents
   cgcount = 0
   
   For Each cg In ClusterGenes
      cgcount = cgcount + 1
      If cgcount Mod 10 = 0 Then
         lblProgress.Caption = "Nested GO paths determined for " & cgcount & " of " & ClusterGenes.count & " distinct genes"
         frmCriteria.Refresh
         DoEvents
      End If
     
      Set GOcollection = cg.getGOTerms
      If GOcollection.count > 0 Then 'this CG is in a GOID find its full path
      'this is a collection of goterm objects
         Set parents = New Collection
         For Each term In GOcollection
            i = 0
            counter = ""
           
            GOID = term.getGOID
          
            Set GONode = TreeForm.TView.Nodes.Item("GO:" & GOID)
            Set parent = GONode.parent
            While parent.key <> TreeForm.rootnode.key
               parentID = Mid(parent.key, 4, 10 - 3)
               If Not visited(parents, parentID) Then  'this CG is in multiple p/c/f terms
                  parents.Add parentID, parentID 'now this GO term won't be visited Again
                  If Not cg.hasGOID(parentID) Then 'this CG could be in both the parent and the child
                                                   'so we need to make sure we don't count it twice
                     
                     Set parentterm = GOterms.Item(parentID)
                     parentterm.setOnChip (parentterm.getOnChip + 1)
                     
resumefirst:
                     For Each id In cg.getPrimaryIDs
                        parentterm.addGene (id)
                     Next id
                     cg.addNestedGOterm parentID, parentterm
                  End If
                  Set parent = parent.parent
               Else
                  Set parent = parent.parent
               End If
            Wend
         
            Set rsGOIDCount = dbMAPPfinder.OpenRecordset("Select Count FROM GeneOntologyCount WHERE " _
                                                      & " ID = '" & GOID & "'")
            
            GOIDcount = rsGOIDCount![count]
   
            For i = 2 To GOIDcount
               GOID = GOID & "I"
               Set GONode = TreeForm.TView.Nodes.Item("GO:" & GOID)
               Set parent = GONode.parent
               While parent.key <> TreeForm.rootnode.key
                  parentID = Mid(parent.key, 4, GOLENGTH - 3)
                  If Not visited(parents, parentID) Then 'this parent hasn't been seen before, add the gene to it
                      If Not cg.hasGOID(parentID) Then 'this CG could be in both the parent and the child
                                                   'so we need to make sure we don't count it twice
                        Set parentterm = GOterms.Item(parentID)
                        parentterm.setOnChip (parentterm.getOnChip + 1)
resumesecond:
                        For Each id In cg.getPrimaryIDs
                           parentterm.addGene (id)
                        Next id
                        cg.addNestedGOterm parentID, parentterm
                     End If
                     parents.Add parentID, parentID
                     Set parent = parent.parent
                  Else 'you've seen this parent before don't add, but continue up.
                     Set parent = parent.parent
                  End If
               Wend
            Next i
resumehere:
         Next term
         root.setOnChip (root.getOnChip + 1)
         cg.addNestedGOterm "GO", root 'need to add the "Gene Ontology" node (the root)
         For Each id In cg.getPrimaryIDs
            root.addGene (id)
         Next id
      End If
   Next cg
   
   root.setOnChip (root.getOnChip - 1) 'this count was off by one because the first gene is counted at the instantiation
   Exit Sub                                    'of the root and in the for each loop.
error:
   Select Case Err.Number
      Case 5 'the parent node doesn't exist we need to add it
      Set parentterm = addtoGOtermsNested(parentID)
      If i >= 2 Then
         Resume resumesecond
      Else
         Resume resumefirst
      End If
      Case 35601
         'you've come across a GOID that isn't in the tree. This happens on occasion because a GO term has
         'multiple IDs. Why would anyone do this??????
         'It turns out that the secondary IDs appear when two terms are collapsed into one, or when one term
         'is split into two. The GO annotators are supposed to only use the primary IDs, so I am just going
         'to catch this error and ignore anything that uses secondary IDs.
         'Debug.Print GOID & " is a secondary ID?"
         Resume resumehere
      Case Else
         MsgBox "error in buildNestedGO. GOID = " & GOID & " " & Err.Number & " " & Err.Description, vbOKOnly
   End Select
End Sub
Public Function visited(parentnodes As Collection, parentID As String) As Boolean
'this looks through the parentNodes collection to see if parent ID is in it already
On Error GoTo error
   Dim id As String
   
   id = parentnodes.Item(parentID)
   visited = True
   Exit Function
   
error:
   Select Case Err.Number
      Case 5
         visited = False
   End Select
End Function




Public Sub buildGotermCollection()
   Dim rsGOs As Recordset
   Dim GOID As String
   Dim goterm As New goterm
   Set rsGOs = dbMAPPfinder.OpenRecordset("SELECT DISTINCT ID from GeneOntology")
   
   While rsGOs.EOF = False
      GOID = rsGOs!id
      goterm.setGOID GOID
      Set rsGOcount = dbMAPPfinder.OpenRecordset("SELECT Count, total FROM [" _
                                    & clustersystem & "-GOCount] WHERE GO = '" & GOID & "'")
      If rsGOcount.EOF Then
         goterm.setingo 0
         goterm.setingolocal 0
      Else
         goterm.setingolocal rsGOcount!count
         goterm.setingo rsGOcount!total
      End If
      GOterms.Add goterm, GOID
      rsGOs.MoveNext
   Wend
End Sub

'This very well documented and very well written piece of code was taken directly from
'Steve Lawlor's GenMAPP v2 code. I haven't altered it at all. -SD

Sub AllRelatedGenes(ByVal idIn As String, systemIn As String, dbGene As Database, _
                    genes As Integer, geneIDs() As String, genefound As Boolean, _
                    Optional supportedSystem As Boolean = False, _
                    Optional systemsList As Variant)
   '  Entry:
   '     idIn           Gene identification received (may have to search to find primary)
   '     systemIn       Cataloging system code for passed idIn
   '     dbGene         Gene Database for this query (the Gene Database for the particular
   '                    drafter window)
   '     geneFound      If True and Specific Gene option checked, don't bother searching for the
   '                    gene. The gene is just being matched to an Expression Dataset.
   '     systemsList    List of system codes that appear in this dataset. Only passed if coloring
   '                    genes to eliminate looking for IDs in systems that don't appear in the
   '                    dataset. If not passed, defaults to "ALL".
   '                    If "EXISTS" then routine being used to find only the existence of the
   '                    specific gene, not all related genes. as soon as a gene is found,
   '                    we can exit the routine. The only return needed is genes(x, 2), where
   '                    x is the last gene found, to test to see whether the gene was found in
   '                    a [P]rimary or [S]econdary column. Use last gene because gene passed
   '                    (genes(0, x)) may not be found but leads to relational tables.
   '                    Used in converting Expression Datasets.
   '  Return:
   '     genes                   Number of related genes (counting from 1)
   '                             This is also used as the index for the geneIDs array. It is zero
   '                             based, so at exit, genes is increased by one.
   '     geneIDs(MAX_GENES, 1)   Gene ID for each related gene. Primary ones listed first
   '                             geneIDs(x, 0) ID
   '                             geneIDs(x, 1) SystemCode
   '                             geneIDs(x, 2) "P" for primary ID (ID column)
   '                                           "S" for secondary ID (eg: Accession in SwissProt)
   '                             The first gene is always the one passed to the procedure whether
   '                             it is found in relational tables or not.
   '     geneFound               True if gene found in any primary or secondary column or
   '                             relational tables.
   '     supportedSystem         systemIn is a system supported in the Gene Database. If false and
   '                             genes > 0 then search has looked in Related column of
   '                             Relational tables and found gene.
   '  The calling function can compare idIn with geneIDs(0, 0) to see if the passed gene ID exists
   '  in the database. This function will find related genes even if the passed gene does not exist
   '  in a supported system but only in a relational table.
   'For AllRelatedGenes()
   '   Dim genes as integer
   '   Dim geneIDs(MAX_GENES, 2) As String
   '   Dim geneFound as boolean
   '   'Dim supportedSystem as Boolean                'System supported in Gene Database [optional]
   '   'Dim systemsList As Variant                                    'Systems to search [optional]
   '  Call:
   '     AllRelatedGenes idIn, systemIn, dbGene, genes, geneIDs, geneFound, _
                         [supportedSystem], [systemsList]
   
   '  At this point related genes are found only if the systemIn is a supported system. We can
   '  find genes if the systemIn is represented in the Related column of a Relational table but
   '  handling that is not been defined yet. This sub should set supportedSystem = False,
   '  no matter what we do with related genes found.

   Dim Index As Integer, lastIndex As Integer
   Dim primaryIDs(MAX_GENES) As String                          'Gene IDs to search relationals for
      '  Primary IDs are those IDs in the systemIn that are used to search relational tables.
      '  A Primary ID is added to Genes() returned if it is found in the systemIn. (And if the
      '  systemIn is in the SystemsList, i.e. represented in the Expression Dataset.)
      '  The search for Primary IDs is in both the ID and Secondary ID columns of the systemIn.
      '  It is possible that more than one Primary ID may be found if the secondary ID shows
      '  up in more than one row.
   Dim primaryIndex As Integer                                                          'Zero based
   Dim firstPrimaryIndex As Integer
   Dim lastPrimaryIndex As Integer
   '  Last index for Primary IDs. May be 0 if idIn not in primary system.
   '  This might be more than zero only if secondary IDs led back to more than one primary ID.
   '  In other words, Secondary ID X1234 occurred in two rows. Not sure this is possible.
   Dim primaryIDReplaced As Integer                                                 'Either 0 or -1
   Dim system As String                            'Cataloging system code currently being examined
   Dim column As String                                                       'Any secondary column
   Dim rsSystems As Recordset                                                    'The Systems table
   Dim rsSystem As Recordset                              'A cataloging-system table, eg: SwissProt
   Dim rsrelations As Recordset                                                'The Relations table
   Dim rsRelational As Recordset                         'A relational table, eg: SwissProt-GenBank
   Dim pipe As Integer, slash As Integer
   Dim sql As String
   Dim searchSeconds As Boolean   'True if doing a GeneFinder or Backpage search or if coloring
                                  'and Expression Dataset has secondary IDs in it. If ED has
                                  'them, the Info table SystemCodes column will contain "|~|".
   Dim secondaryCols(10, 1) As String
   '           secondaryCols(x, 0)     Names of secondary columns
   '           secondaryCols(x, 1)     "M" if multiple, pipe-surrounded IDs allowed
   Dim lastSecondCol As Integer
   Dim singleGene As Boolean                   'True if systemList comes in as "EXISTS". See above.
   Dim currentMousePointer As Integer                  'Form's MousePointer on entry. Reset on exit
 
   If dbGene Is Nothing Then Exit Sub                      'No database >>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   currentMousePointer = Screen.ActiveForm.MousePointer
   Screen.ActiveForm.MousePointer = vbHourglass
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Assign Gene Passed As Gene #1
   genes = 0                                                              'Zero based at this point
   geneIDs(genes, 0) = idIn
   geneIDs(genes, 1) = systemIn
   geneIDs(genes, 2) = "P"                                                   'Default to primary ID
   
   If genefound And InStr(cfgColoring, "S") Then '++++++++++++++++++++++ Just Return Gene Passed In
      '  User has chosen the Specific Gene Option and is just matching to an Expression Dataset
      '  rather than trying to find a gene in GeneFinder or creating a Backpage. In this case,
      '  who cares whether the gene is in a supported system, found in a relational system or
      '  exists anywhere.
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   If IsMissing(systemsList) Then systemsList = "ALL"
   If VarType(systemsList) = vbNull Then systemsList = "ALL"
   If systemsList = "" Then systemsList = "ALL"
   If systemsList = "EXISTS" Then
      systemsList = "ALL"
      singleGene = True
   End If
   
   genefound = False
   supportedSystem = False
   searchSeconds = True
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ See If Supported System
'      '  Gene in (idIn and systemIn) is classified as:
'      '     Supported   Its system exists as one of the system tables in the Gene DB
'      '                 supportedSystem = True
'      '     Relational  Its system not supported but exists in a relational table
'      '                 supportedSystem = False
'      '     Neither     Exits sub with only the passed gene returned
   Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems WHERE SystemCode = '" & systemIn & "'", _
                   dbOpenForwardOnly)                                    'Get the system table name
                                                 'Eg: SELECT * FROM Systems WHERE SystemCode = 'Rm'
   If Not rsSystems.EOF Then '========================================================System Exists
      If Dat(rsSystems![Date]) <> "" Or systemIn = "O" Then        'Date or Other, supported system
         '  Other always supported. See comments at beginning of sub
         supportedSystem = True
      End If
   End If
   
   If supportedSystem Then '++++++++++++++++++++++++++++++++++++++++++++++++ Look For Specific Gene
      '  Specific gene can only be in a supported system
      Set rsSystem = dbGene.OpenRecordset( _
                     "SELECT * FROM " & rsSystems!system & _
                     "   WHERE ID = '" & idIn & "'", _
                     dbOpenForwardOnly)        'Eg: SELECT * FROM SwissProt WHERE ID = 'CALM_HUMAN'
      If Not rsSystem.EOF Then                                             'Found idIn in ID column
         genefound = True
         If singleGene Then GoTo ExitSub                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
   End If

'  Per Kam, the only time Specific Gene option has effect is in coloring. GeneFinder and
'  Backpages will always search for related genes.
'   If InStr(cfgColoring, "S") Then '++++++++++++++++++++++++++++++++++++++++ "Specific Gene" Option
'      '  If user Options specify "Specific Gene", no relations are searched for nor are
'      '  secondary columns
'      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'   End If
   
   lastSecondCol = SecondCols(rsSystems, secondaryCols)                     'Find Secondary Columns
   If supportedSystem And Not genefound Then '++++++++++++++++++ Look For Gene In Secondary Columns
      '  If the received idIn is not in the ID column of the gene system table (systemIn)
      '  go to the secondary columns and search
      '  Typical column listing in Systems!Columns
      '     ID|Accession\SMBF|Nicknames\sF|Protein|Functions\B|
      For i = 0 To lastSecondCol - 1 '========================================Each Secondary Column
         If secondaryCols(i, 1) = "M" Then                              'Multiple secondary columns
            sql = "SELECT ID" & _
                  "   FROM " & rsSystems!system & " " & _
                  "   WHERE [" & secondaryCols(i, 0) & "] LIKE '*|" & idIn & "|*'"        'Use LIKE
               'Eg: SELECT ID FROM SwissProt WHERE Accession LIKE '*|A1234|*'
         Else                                                              'Single ID without pipes
            sql = "SELECT ID" & _
                  "   FROM " & rsSystems!system & " " & _
                  "   WHERE [" & secondaryCols(i, 0) & "] = '" & idIn & "'"                  'Use =
               'Eg: SELECT ID FROM SGD WHERE Gene = 'TFC3'
         End If
         Set rsSystem = dbGene.OpenRecordset(sql, dbOpenForwardOnly)
         If Not rsSystem.EOF Then '----------------------------------Gene Found In Secondary Column
            genefound = True
            primaryIndex = 1                                                    'First primary gene
            geneIDs(0, 2) = "S"             'First gene actually in a secondary column. Change code
            If singleGene Then GoTo ExitSub                'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
               '  For singleGene, must exit here because converting Expression Datasets checks
               '  last gene found to determine if [P]rimary or [S]econdary.
            Do Until rsSystem.EOF                               'All rows where idIn in this column
               '  Add primary ID for each instance of the secondary ID in the system.
               '  This should not add another instance of idIn because that was never found
               '  in the first place
'               If systemsList = "ALL" Or InStr(systemsList, "|" & systemIn & "|") <> 0 Then
                  '  Either looking at all systems or gene system is in the
                  '  Expression Dataset (for coloring), add to returned related genes.
                  '  Add the primary ID to the list.
                  genes = genes + 1
                  geneIDs(genes, 0) = rsSystem!id
                  geneIDs(genes, 1) = systemIn
                  geneIDs(genes, 2) = "P"
'               End If
               rsSystem.MoveNext
            Loop
         End If
      Next i
   End If
   firstPrimaryIndex = primaryIndex
   lastPrimaryIndex = genes
   '  At this point we have created a list of primary IDs in the systemIn.
   '  geneIDs() up to lastPrimaryIndex has all the IDs we want to search for.
      
   If supportedSystem And Not singleGene Then '+++++++++++++++++++++ Find Secondary IDs In systemIn
      '  If singleGene, don't do this because this routine only finds secondary columns for genes
      '  that have been found above and secondary IDs are not listed in the relational tables.
      For primaryIndex = firstPrimaryIndex To lastPrimaryIndex
         '  GeneIDs(0, 0) might not be a primary ID
         '  Uses the SecondCols return from above because only dealing with systemIn at this point
         For i = 0 To lastSecondCol - 1 '=====================================Each Secondary Column
            sql = "SELECT [" & secondaryCols(i, 0) & "] AS Secondary" & _
                  "   FROM " & rsSystems!system & _
                  "   WHERE ID = '" & geneIDs(primaryIndex, 0) & "'"
               'Eg: SELECT Accession FROM SwissProt WHERE ID = 'CALM_HUMAN'
            Set rsSystem = dbGene.OpenRecordset(sql, dbOpenForwardOnly)
            If Not rsSystem.EOF Then '------------------------------------Found In Secondary Column
               If secondaryCols(i, 1) = "M" Then '_________________Multiple IDs In Secondary Column
                  '  Might return something like "|A1234|B5678|"
                  Dim nextPipe As Integer, strOut As String
                  
                  pipe = 1
                  Do While pipe < Len(rsSystem!Secondary)
                     nextPipe = InStr(pipe + 1, rsSystem!Secondary, "|")
                     If nextPipe = 0 Then nextPipe = Len(rsSystem!Secondary) + 1
                     genes = genes + 1
                     geneIDs(genes, 0) = Mid(rsSystem!Secondary, pipe + 1, nextPipe - pipe - 1)
                     geneIDs(genes, 1) = systemIn
                     geneIDs(genes, 2) = "S"
                     pipe = nextPipe
                  Loop
               Else '_________________________________________________Single ID In Secondary Column
                  genes = genes + 1
                  geneIDs(genes, 0) = rsSystem!Secondary
                  geneIDs(genes, 1) = systemIn
                  geneIDs(genes, 2) = "S"
               End If
            End If
         Next i
      Next primaryIndex
   End If
   '  We now have all the secondary IDs for the primary ones
            
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find All Applicable Relational Tables
   If systemsList <> "ALL" Then '===========================Only System Codes In Expression Dataset
      '  Look at only those relational tables where the SystemCode is represented in the
      '  Expression Dataset
      '  Matches Expression data from supported and nonsupported systems
      Set rsrelations = dbGene.OpenRecordset( _
                        "SELECT * FROM Relations" & _
                        "   WHERE SystemCode = '" & systemIn & "' " & _
                        "         AND instr('" & systemsList & "', relatedCode )" & _
                        "      OR RelatedCode = '" & systemIn & "' " & _
                        "         AND instr('" & systemsList & "', systemCode)" & _
                        "   ORDER BY Relation NOT LIKE '*GenBank*'")
                        'This query selects all relational tables, supported or not. GenBanks
                        'always listed last.
         '  Eg: SELECT * FROM Relations
         '         WHERE SystemCode = 'G'
         '               AND instr('|G|O|', 'O')                             'In Expression Dataset
         '            OR RelatedCode = 'G')
         '               AND instr('|G|O|', 'G')
   Else '==========================================================================Any System Codes
      Set rsrelations = dbGene.OpenRecordset( _
                        "SELECT * FROM Relations" & _
                        "   WHERE (SystemCode = '" & systemIn & "'" & _
                        "          OR RelatedCode = '" & systemIn & "')" & _
                        "   ORDER BY Relation NOT LIKE '*GenBank*'")
                        'This query selects all relational tables, supported or not. GenBanks
                        'always listed last.
   End If
   '  At this point, rsRelations has all the relational tables we want to look at
            
   Do Until rsrelations.EOF '+++++++++++++++++++++++++++++++++++++++++ Search All Relational Tables
'Debug.Print rsRelations!Relation
      '======================================================================Find Secondary Columns
      If rsrelations!systemcode = systemIn Then '------------------------SystemIn In Primary Column
         Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems WHERE SystemCode = '" & rsrelations!relatedCode & "'", _
                   dbOpenForwardOnly)
      Else
         Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems WHERE SystemCode = '" & rsrelations!systemcode & "'", _
                   dbOpenForwardOnly)
      End If
      lastSecondCol = SecondCols(rsSystems, secondaryCols)
      
      For primaryIndex = firstPrimaryIndex To lastPrimaryIndex '====================Each Primary ID
         '  For singleGene, there should be only one Primary ID, the gene passed in. We did
         '  not search for secondaries for the idIn, and to reach here idIn must not have
         '  been found.
         If rsrelations!systemcode = systemIn Then '---------------------SystemIn In Primary Column
            sql = "SELECT * FROM [" & rsrelations!Relation & "]" & _
                 "   WHERE Primary = '" & geneIDs(primaryIndex, 0) & "'"
                        'Eg: SELECT * FROM [SwissProt-GenBank] WHERE Primary = 'CALM_HUMAN'
            Set rsRelational = dbGene.OpenRecordset(sql)
            Do Until rsRelational.EOF
               If genes < MAX_GENES - 1 Then                          'Anything more we just forget
                  genefound = True
                  genes = genes + 1
                  geneIDs(genes, 0) = rsRelational!related
                  geneIDs(genes, 1) = rsrelations!relatedCode
                  geneIDs(genes, 2) = "P"
                  If singleGene Then GoTo ExitSub          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                     '  For singleGene, must exit here because converting Expression Datasets
                     '  checks last gene found to determine if [P]rimary or [S]econdary.
                  If searchSeconds Then
                     AddSecondIDs secondaryCols, lastSecondCol, geneIDs, genes, rsSystems, dbGene
                  End If
               End If
               rsRelational.MoveNext
            Loop
         Else '----------------------------------------------------------SystemIn In Related Column
            Set rsRelational = dbGene.OpenRecordset( _
                              "SELECT * FROM [" & rsrelations!Relation & "]" & _
                              "   WHERE Related = '" & geneIDs(primaryIndex, 0) & "'")
                              'Eg: SELECT * FROM [SwissProt-GenBank] WHERE Related = 'X1234'
            Do Until rsRelational.EOF
               If genes < MAX_GENES - 1 Then                          'Anything more we just forget
                  genefound = True
                  genes = genes + 1
                  geneIDs(genes, 0) = rsRelational!primary
                  geneIDs(genes, 1) = rsrelations!systemcode
                  geneIDs(genes, 2) = "P"
                  If singleGene Then GoTo ExitSub          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                     '  For singleGene, must exit here because converting Expression Datasets
                     '  checks last gene found to determine if [P]rimary or [S]econdary.
                  If searchSeconds Then
                     AddSecondIDs secondaryCols, lastSecondCol, geneIDs, genes, rsSystems, dbGene
                  End If
               End If
               rsRelational.MoveNext
            Loop
         End If
      Next primaryIndex
      rsrelations.MoveNext
   Loop
   
ExitSub:
   genes = genes + 1        'genes was zero based, not it should be count of actual number of genes
   Screen.ActiveForm.MousePointer = currentMousePointer

'   For i = 0 To genes - 1
'      Debug.Print geneIDs(i, 1); " "; geneIDs(i, 0)
'   Next i
End Sub

'from genmapp v2
Sub AddSecondIDs(secondaryCols() As String, lastSecondCol As Integer, geneIDs() As String, _
                 genes As Integer, rsSystems As Recordset, dbGene As Database)
   Dim rsSystem As Recordset, pipe As Integer, nextPipe As Integer, sql As String
   
   For i = 0 To lastSecondCol - 1 '===========================================Each Secondary Column
      sql = "SELECT [" & secondaryCols(i, 0) & "] AS Secondary" & _
            "   FROM " & rsSystems!system & _
            "   WHERE ID = '" & geneIDs(genes, 0) & "'"
         'Eg: SELECT Accession FROM SwissProt WHERE ID = 'CALM_HUMAN'
      Set rsSystem = dbGene.OpenRecordset(sql, dbOpenForwardOnly)
      If Not rsSystem.EOF Then '------------------------------------------Found In Secondary Column
         If secondaryCols(i, 1) = "M" Then '_______________________Multiple IDs In Secondary Column
            '  Might return something like "|A1234|B5678|"
            pipe = 1
            Do While pipe < Len(rsSystem!Secondary)
               nextPipe = InStr(pipe + 1, rsSystem!Secondary, "|")
               If nextPipe = 0 Then nextPipe = Len(rsSystem!Secondary) + 1
               genes = genes + 1
               geneIDs(genes, 0) = Mid(rsSystem!Secondary, pipe + 1, nextPipe - pipe - 1)
               geneIDs(genes, 1) = rsSystems!systemcode
               geneIDs(genes, 2) = "S"
               pipe = nextPipe
            Loop
         Else '____________________________________________________Single ID In Secondary Column
            genes = genes + 1
            geneIDs(genes, 0) = rsSystem!Secondary
            geneIDs(genes, 1) = rsSystems!systemcode
            geneIDs(genes, 2) = "S"
         End If
      End If
   Next i
End Sub

'*********************************************************** Secondary Column Names for Gene System
Function SecondCols(rsSystems As Recordset, secondaryCols() As String) As Integer
   '  Entry    rsSystem    Record from the systems table for the particular system
   '  Return   The number of secondary columns found, counting from 1
   '           secondaryCols(x, 0)     Names of secondary columns
   '           secondaryCols(x, 1)     "M" if multiple, pipe-surrounded IDs allowed
   Dim pipe As Integer, slash As Integer, column As String, columns As Integer
   
   pipe = 3                                                                           'End of "ID|"
   slash = InStr(pipe + 1, rsSystems!columns, "\")                                      'Next slash
   Do While slash '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Each Secondary Column
      pipe = InStrRev(rsSystems!columns, "|", slash)                  'Next pipe with slash in unit
      '  If the Columns column had
      '     ID|Whatever|Accession\SMBF|Nicknames\sF|Protein|Functions\B|
      '  then the pipe would be the one beginning the |Accession\SMBF|
      '  not the |Whatever| unit.
      If UCase(Mid(rsSystems!columns, slash + 1, 1)) = "S" Then '===========This is a search column
         columns = columns + 1                          'One based, used for zero-based index below
         secondaryCols(columns - 1, 0) = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)
         If Mid(rsSystems!columns, slash + 1, 1) = "S" Then       'Multiple IDs surrounded by pipes
            secondaryCols(columns - 1, 1) = "M"
'            secondaryCols(columns) = "SELECT ID " & _
'                                    "FROM " & rsSystems!system & " " & _
'                                    "WHERE [" & column & "] LIKE '*|" & idIn & "|*'"      'Use LIKE
'               'Eg: SELECT ID FROM SwissProt WHERE Accession LIKE '*|A1234|*'
         Else                                                              'Single ID without pipes
            secondaryCols(columns - 1, 1) = "S"
'            secondaryCols(columns) = "SELECT ID" & _
'                                    "FROM " & rsSystems!system & " " & _
'                                    "WHERE [" & column & "] = '" & idIn & "'"                'Use =
'               'Eg: SELECT ID FROM SGD WHERE Gene = 'TFC3'
         End If
      End If
      slash = InStr(slash + 1, rsSystems!columns, "\")                                  'Next slash
   Loop
   SecondCols = columns
End Function


Private Function FindGEXGene(idIn As String, GEXgene As ClusterGene) As Boolean
On Error GoTo error
   
   Set GEXgene = GEXgenes.Item(idIn)
   FindGEXGene = True
   Exit Function
   
error:
   Select Case Err.Number
      Case 5
         FindGEXGene = False
   End Select
   
End Function

Private Sub MaptoLocalMAPPs()
   'for each gene in the GEX find is related genes in the systemlist of the GeneTOMAPP table
   'for each related gene, find all of its MAPPs.
   'store the three collections for the criterion piece.
   Dim tblresults As TableDef, tblnestedresults As TableDef
   Dim rsGenes As Recordset, rssystemlist As Recordset
   Dim systemlist As String
   Dim genes As Integer
   Dim i As Integer
   Dim nohits As Boolean
   Dim progress As Integer
   Dim genefound As Boolean
   progress = 0
   systemlist = "|"
   Set rssystemlist = dbLocalMAPPs.OpenRecordset("SELECT DISTINCT SystemCode from GeneToMAPP")
   While Not rssystemlist.EOF
      systemlist = systemlist & rssystemlist!systemcode & "|"
      rssystemlist.MoveNext
   Wend
   'we need to add a lot of error checking.
   
   Set rsGenes = dbExpressionData.OpenRecordset _
                     ("SELECT ID, SystemCode FROM Expression")
   rsGenes.MoveLast
   rsGenes.MoveFirst
   GEXsize = rsGenes.RecordCount
   If Dir(newfilename & "-Local.gdb") = "" Then
      While rsGenes.EOF = False
         progress = progress + 1
         If progress Mod 5 = 0 Then
            lblProgress.Caption = progress & " out of " & GEXsize & " mapped to the Local MAPPs."
            frmCriteria.Refresh
            DoEvents
         End If
         nohits = True
         ReDim geneIDs(MAX_GENES, 2) As String
         genes = 0
         genefound = True 'flag for matching to expression dataset
         AllRelatedGenes rsGenes!id, rsGenes!systemcode, dbMAPPfinder, genes, geneIDs, genefound, , systemlist
         If genes = 0 Then 'no related genes found in systemlist systems
            nolocalCluster = nolocalCluster + 1
            addNotFoundtoLocalPrimaryGenes rsGenes!id
         Else
            For i = 0 To genes - 1
               addtoLocalPrimaryGenes rsGenes!id, addToLocalClusterGenes(geneIDs(i, 0), nohits, rsGenes!id)
            Next i
         End If
         If nohits Then 'the clustergene(s) for this primary gene didn't find a mapp
            nomapp = nomapp + 1
         End If
         rsGenes.MoveNext
      Wend
   
      Dim tblinfo As TableDef, tablecluster As TableDef
      Dim tblgo As TableDef, tblMAPPtoGene As TableDef
      Set dbChipDataLocal = CreateDatabase(newfilename & "-Local.gdb", dbLangGeneral)
      Set tblinfo = dbChipDataLocal.CreateTableDef("Info")
         With tblinfo
            .Fields.Append .CreateField("LocalFolder", dbText, 255)
            .Fields.Append .CreateField("Version", dbText, 255)
         End With
      dbChipDataLocal.TableDefs.Append tblinfo
      Set tblCluster = dbChipDataLocal.CreateTableDef("PrimarytoClusterSystem")
      With tblCluster
         .Fields.Append .CreateField("Primary", dbText, 15)
         .Fields.Append .CreateField("Related", dbText, 15)
      End With
      dbChipDataLocal.TableDefs.Append tblCluster
      Set tblgo = dbChipDataLocal.CreateTableDef("ClusterSystemtoGO")
      tblgo.Fields.Append tblgo.CreateField("Primary", dbText, 15)
      tblgo.Fields.Append tblgo.CreateField("Related", dbText, 255) 'need to fit the entire mapp name
      dbChipDataLocal.TableDefs.Append tblgo
      
      Set tblMAPPtoGene = dbChipDataLocal.CreateTableDef("GOIDtoGene")
      tblMAPPtoGene.Fields.Append tblMAPPtoGene.CreateField("Primary", dbText, 255)
      tblMAPPtoGene.Fields.Append tblMAPPtoGene.CreateField("Related", dbText, 15)
      dbChipDataLocal.TableDefs.Append tblMAPPtoGene
      
      Set tblnestedresults = dbChipDataLocal.CreateTableDef("NestedResults")
      With tblnestedresults
         .Fields.Append .CreateField("GOType", dbText, 2)
         .Fields.Append .CreateField("GOName", dbMemo)
         Dim idxNestedResults As Index
         Dim idxPercentage As Index
         .Fields.Append .CreateField("GOID", dbText, 15)
         Set idxPercentage = .CreateIndex("percentageindex")
         Set idxNestedResults = .CreateIndex("nestedResults")
         idxNestedResults.primary = True
         idxNestedResults.Fields.Append .CreateField("GOID", dbText, 15)
         idxPercentage.Fields.Append .CreateField("GOType", dbText, 2)
         idxPercentage.Fields.Append .CreateField("Percentage", dbSingle)
         idxPercentage.Fields.Append .CreateField("Zscore", dbSingle)
         .Indexes.Append idxNestedResults
         .Indexes.Append idxPercentage
         .Fields.Append .CreateField("ChangedLocal", dbSingle)
         .Fields.Append .CreateField("OnChipLocal", dbSingle)
         .Fields.Append .CreateField("InGOLocal", dbSingle)
         .Fields.Append .CreateField("PercentageLocal", dbSingle)
         .Fields.Append .CreateField("PresentLocal", dbSingle)
         .Fields.Append .CreateField("Changed", dbSingle)
         .Fields.Append .CreateField("OnChip", dbSingle)
         .Fields.Append .CreateField("InGO", dbSingle)
         .Fields.Append .CreateField("Percentage", dbSingle)
         .Fields.Append .CreateField("Present", dbSingle)
         .Fields.Append .CreateField("Zscore", dbDouble)
         .Fields.Append .CreateField("permuteP", dbDouble)
         .Fields.Append .CreateField("adjustedP", dbDouble)
      End With
   Set tblresults = dbChipDataLocal.CreateTableDef("Results")
      With tblresults
         .Fields.Append .CreateField("GOType", dbText, 2)
         Dim idxResults As Index
         .Fields.Append .CreateField("GOID", dbText, 255)
         .Fields.Append .CreateField("GOName", dbMemo)
         Set idxResults = .CreateIndex("Results")
         idxResults.Fields.Append .CreateField("GOType", dbText, 2)
         idxResults.Fields.Append .CreateField("Zscore", dbSingle)
         idxResults.Fields.Append .CreateField("Changed", dbSingle)
         .Indexes.Append idxResults
         .Fields.Append .CreateField("Changed", dbSingle)
         .Fields.Append .CreateField("OnChip", dbSingle)
         .Fields.Append .CreateField("InGO", dbSingle)
         .Fields.Append .CreateField("Percentage", dbSingle)
         .Fields.Append .CreateField("Present", dbSingle)
         .Fields.Append .CreateField("Zscore", dbDouble)
         .Fields.Append .CreateField("permuteP", dbDouble)
         .Fields.Append .CreateField("adjustedP", dbDouble)
      End With
   dbChipDataLocal.TableDefs.Append tblresults
   dbChipDataLocal.TableDefs.Append tblnestedresults
      
      Set rstemp = dbMAPPfinder.OpenRecordset("Select Version from Info")
  
      dbChipDataLocal.Execute "INSERT INTO Info (LocalFolder, Version) VALUES ('" _
                              & TreeForm.LocalPath & "', '" & TreeForm.localDate & "')"
      For Each Gene In LocalPrimaryGenes
         Set relatedItems = Gene.getClusterGenes
         If relatedItems.count = 0 Then 'no cluster gene for this guy
            dbChipDataLocal.Execute "INSERT INTO PRimarytoClusterSystem(Primary)" _
                       & " VALUES ('" & Gene.getID & "')"
         Else
            For Each related In relatedItems
               dbChipDataLocal.Execute "INSERT INTO PrimarytoClusterSystem (Primary, Related)" _
                              & " VALUES ('" & Gene.getID & "', '" & related.getID & "')"
            Next related
         End If
         
      Next Gene
      DoEvents
      For Each Cluster In LocalClusterGenes
         Set terms = Cluster.getGOTerms
         clusterID = Cluster.getID
         If terms.count = 0 Then 'no GOIDs for this cluster
            dbChipDataLocal.Execute "INSERT INTO ClusterSystemtoGO(Primary) VALUES" _
                              & "('" & clusterID & "')"
         Else
            For Each term In terms
               dbChipDataLocal.Execute "INSERT INTO ClustersystemtoGO(Primary, Related) VALUES ('" _
                              & clusterID & "', '" & term.getGOID & "')"
            Next term
         End If
      Next Cluster
      
      For Each goterm In localMAPPsCollection
         Set relatedItems = goterm.getGenes
         For Each Gene In relatedItems
            dbChipDataLocal.Execute "INSERT INTO GOIDtoGene(Primary, Related) VALUES" _
                     & "('" & goterm.getGOID & "', '" & Gene & "')"
         Next Gene
      Next goterm
      Dim idxGOID As Index
      Set idxGOID = tblMAPPtoGene.CreateIndex("Results")
      idxGOID.Fields.Append tblMAPPtoGene.CreateField("Primary", dbText, 15)
      tblMAPPtoGene.Indexes.Append idxGOID
      'now we have the GenesInME for each GOterm in a table. Indexed by GOID.
      
      DoEvents
      'ok, so now there are two tables. (PrimaryToClusterSytstem) and ClustertoGO and nestedclustertogo.
      'neither table is indexed, as I will only be calling select * from these tables.
      '
   '*******************************************************************************
   'END OF BUILDING THE MAPPINGS
   '*******************************************************************************
   Else 'the mapping has been done before just read in the files.
      'build the mappings to clusterID and GO
      
      frmCriteria.lblProgress.Caption = "Loading LocalMAPP mappings from " & newfilename & ".gdb"
      frmCriteria.Refresh
      DoEvents
   
      Set dbChipDataLocal = OpenDatabase(newfilename & "-Local.gdb")
      Set rstemp2 = dbChipDataLocal.OpenRecordset("SELECT LocalFolder, Version FROM Info")
      If (TreeForm.LocalPath <> rstemp2!LocalFOlder) Or _
         (TreeForm.localDate <> rstemp2!Version) Then 'they are using the wrong database
         'need to rebuild
         lblProgress.Caption = "Your local gdb file is out of date. MAPPFinder is now going to rebuild it."
         frmCriteria.Refresh
         DoEvents
         'Debug.Print newfilename
         dbChipDataLocal.Close
         Kill newfilename & "-Local.gdb"
         MaptoLocalMAPPs
         Exit Sub
      Else 'they match, load the data
         
         Set rsGenes = dbChipDataLocal.OpenRecordset("SELECT Primary, Related FROM [ClusterSystemToGO]" _
                                                & " ORDER BY Primary")
         While rsGenes.EOF = False
            If IsNull(rsGenes!related) Then
               addLocalClusterID (rsGenes!primary)
            Else
               addLocalClusterPair rsGenes!primary, rsGenes!related
            End If
            rsGenes.MoveNext
            
         Wend
         DoEvents
         Set rsGenes = dbChipDataLocal.OpenRecordset("SELECT Primary, Related FROM " _
                                             & "PrimaryToClusterSystem")
         While rsGenes.EOF = False
            If IsNull(rsGenes!related) Then
               addNotFoundtoLocalPrimaryGenes (rsGenes!primary)
               ' if two probes both have the same ID, and this ID has no
               'related ClusterID then the number of probes not linked to
               'the cluster system will be off by one. Since the primarygenes
               'collection only has one copy, I'll need to query the expression table.
               'this is a hack, but it will work
               Set rsProbes = dbExpressionData.OpenRecordset("SELECT ID from Expression where ID = '" _
                                                            & rsGenes!primary & "'")
               rsProbes.MoveLast
               nolocalCluster = nolocalCluster + rsProbes.RecordCount
                                        
            Else
               addtoLocalPrimaryGenes rsGenes!primary, LocalClusterGenes.Item(rsGenes!related)
               For Each MAPP In LocalClusterGenes.Item(rsGenes!related).getGOTerms()
                  MAPP.addGene (rsGenes!primary)
               Next MAPP
            End If
            rsGenes.MoveNext
         Wend
      
      End If
      DoEvents
   End If
End Sub

Function Dat(ByVal z As Variant) As String
   Rem************************************************************************
   Rem  CONVERTS VARIANT, PARTICULARLY DATABASE FIELD, TO STRING *************
   Rem************************************************************************
   On Error GoTo DatError
   If VarType(z) <> vbNull Then
      Dat = Trim(z)
   Else
      Dat = ""
   
   End If
DatContinue:
   Exit Function

DatError:
   Dat = ""
   Resume DatContinue
End Function


Public Function CalculateGOResults(criterion As Integer) As Boolean
   'criterion is the name of the current criterion. This is added to the file name.
   'num is the number of the criterion to extract from the SQL array
  ' On Error GoTo error
   Dim rsFilter As Recordset
   Dim metFilter As Integer 'no GeneOntology available
   Dim output As TextStream, genetoGO As TextStream
   Dim GOID As String
   Dim progress As Long
   Dim i As Long, numofsystems As Integer
   Dim noClusterC As Long, clusterC As Long
   Dim pgene As PrimaryGene, GONode As Node, cg As ClusterGene
   
   Set rsFilter = dbExpressionData.OpenRecordset("SELECT OrderNo, ID, SystemCode FROM" _
                     & " Expression WHERE (" & sql(criterion) & ")")

   If rsFilter.EOF Then
      MsgBox "There are no genes in the Expression Dataset that meet the criterion, " & sql(criterion) _
             & ". If you selected multiple criterion, the one previous to this have already been calcualted." _
              & " MAPPFinder will now exit because it can't calcualte any results.", vbOKOnly
      CalculateGOResults = False
      Exit Function
   End If

   rsFilter.MoveLast
   rsFilter.MoveFirst
   metFilter = rsFilter.RecordCount
       
   TreeForm.resetProgress
   lblProgress.Caption = "Calculating MAPPFinder results for the " & metFilter & " genes meeting criteria."
   frmCriteria.Refresh
   DoEvents
   progress = 0
   noClusterC = 0
   
   clusterC = 0
   While rsFilter.EOF = False
      progress = progress + 1
      If progress Mod 10 = 0 Then
         lblProgress.Caption = "Results calculated for " & progress & " out of " & metFilter & " genes meeting the criterion."
         frmCriteria.Refresh
         DoEvents
      End If
      Set pgene = PrimaryGenes.Item(rsFilter!id)
      Set cgs = pgene.getClusterGenes
      If cgs.count = 0 Then
         noClusterC = noClusterC + 1
      Else
         For Each cg In cgs
            If Not cg.wasVisited() Then 'two different primary genes can share the same cluster gene
               cg.visit
               clusterC = clusterC + 1
               Set GOs = cg.getGOTerms  'so you need to make sure you only use each CG once.
               For Each go In GOs
                  go.addChangedLocal (rsFilter!id)
               Next go
               Set GOs = cg.getNestedGOterms
               For Each go In GOs
                  'go.setChanged (go.getChanged + 1) see comment in GOterm class
                  go.addChangedGene (rsFilter!id)
               Next go
            End If
         Next cg
      End If
nextPG:
      rsFilter.MoveNext
   Wend
   
   lblProgress.Caption = "Calculating Z scores."
   frmCriteria.Refresh
   DoEvents
   calculateTestStat
   If Statistics Then
      CalculatePValues
   End If
   lblProgress.Caption = "Saving Results..."
   'now we create a table to store the genesChangedInME for each GOterm
   Dim tblChangedGenes As TableDef
   Dim tablename As String
   tablename = Mid(fixFileName(txtFile), InStrRev(txtFile.Text, "\") + 1, Len(txtFile.Text) _
               - InStrRev(txtFile.Text, "\") - 1) & "-Criterion" _
               & criterion & "-GO"
createtable: 'if the table exists the error is caught and the preexisting table is deleted.
   Set tblChangedGenes = dbChipData.CreateTableDef(tablename)
   tblChangedGenes.Fields.Append tblChangedGenes.CreateField("Primary", dbText, 15)
   tblChangedGenes.Fields.Append tblChangedGenes.CreateField("Related", dbText, 15)
   dbChipData.TableDefs.Append tblChangedGenes
  ' tablename = Mid(fixFileName(txtFile), InStrRev(txtFile.Text, "\") + 1, Len(txtFile.Text) _
               - InStrRev(txtFile.Text, "\") - 1)
   
   For Each goterm In GOterms
      goterm.calculateLocal
      goterm.calculateResults
      results = goterm.getResults 'this array has 11 elements. See GOterm class for more
      GOID = goterm.getGOID
      For Each Gene In goterm.getChangedGenes()
          dbChipData.Execute "INSERT INTO [" & tablename & "](Primary, Related) VALUES" _
                     & "('" & goterm.getGOID & "', '" & Gene & "')"
      Next Gene
      goterm.reset 'now we reset the GOterm for the next criterion.
      If GOID = "GO" Then
         dbChipData.Execute "INSERT INTO NestedResults (GOID, GOName, GOType, ChangedLocal" _
                              & ", OnChipLocal, InGOLocal, PercentageLocal, PresentLocal," _
                              & " Changed, OnChip, InGO, Percentage, Present, Zscore, " _
                              & "permuteP, adjustedP)" _
                              & " VALUES ('GO', 'Gene Ontology', 'r', " & results(0) _
                              & ", " & results(1) & ", " & results(2) & ", " & results(3) _
                              & ", " & results(4) & ", " & results(5) & ", " & results(6) _
                              & ", " & results(7) & ", " & results(8) & ", " & results(9) _
                              & ", " & results(10) & ", " & results(11) & ", " & results(12) & ")"
      Else
         Set rstemp = dbMAPPfinder.OpenRecordset("SELECT Name, Type FROM GeneOntology" _
                                                & " WHERE ID = '" & GOID & "'")
         If (rstemp.EOF = False) Then 'there are some cases where a GO term is annotated by a mod,
            'but not in our GO table. Why? I don't know, but I'm going to catch the error.
            dbChipData.Execute "INSERT INTO NestedResults (GOID, GOName, GOType, ChangedLocal" _
                              & ", OnChipLocal, InGOLocal, PercentageLocal, PresentLocal," _
                              & " Changed, OnChip, InGO, Percentage, Present, Zscore, " _
                              & "permuteP, adjustedP)" _
                              & " VALUES ('" & GOID & "', '" & rstemp!name & "', '" _
                              & rstemp!Type & "', " & results(0) _
                              & ", " & results(1) & ", " & results(2) & ", " & results(3) _
                              & ", " & results(4) & ", " & results(5) & ", " & results(6) _
                              & ", " & results(7) & ", " & results(8) & ", " & results(9) _
                              & ", " & results(10) & ", " & results(11) & ", " & results(12) & ")"
         End If
      End If
      DoEvents
   Next goterm
   Dim idxGOID As Index
   Set idxGOID = tblChangedGenes.CreateIndex(tablename)
   idxGOID.Fields.Append tblChangedGenes.CreateField("Primary", dbText, 15)
   tblChangedGenes.Indexes.Append idxGOID
   'now we have indexed the changedgenes table for this criterion.
   distinctgenes = ClusterGenes.count
   If Statistics Then
      Set rstemp = dbChipData.OpenRecordset("SELECT Gotype, GOID, GOName, ChangedLocal" _
               & ", OnChipLocal, InGOLocal, PercentageLocal, PresentLocal, Changed, OnChip," _
               & " InGo, Percentage, Present, Zscore, PermuteP, adjustedP FROM NestedResults " _
               & "ORDER BY permuteP, Zscore DESC")
   Else
      Set rstemp = dbChipData.OpenRecordset("SELECT Gotype, GOID, GOName, ChangedLocal" _
               & ", OnChipLocal, InGOLocal, PercentageLocal, PresentLocal, Changed, OnChip," _
               & " InGo, Percentage, Present, Zscore, permuteP, adjustedP FROM NestedResults " _
               & "ORDER BY Zscore DESC, OnChip DESC")
   End If
   Set output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Criterion" & criterion & "-GO.txt")
   output.WriteLine ("MAPPFinder 2.0 Results for the Gene Ontology")
   output.WriteLine ("File: " & filelocation)
   output.WriteLine ("Table: " & tablename)
   output.WriteLine ("Database: " & databaseloc)
   output.WriteLine ("colors:|" & colorset & "|")
   output.WriteLine (GODate)
   output.WriteLine (species)
   If Statistics Then
      output.WriteLine ("Pvalues = true")
   Else
      output.WriteLine ("Pvalues = false")
   End If
   output.WriteLine ("Calculation Summary:")
   output.WriteLine (metFilter & " probes met the " & sql(criterion) & " criteria.")
   output.WriteLine (metFilter - noClusterC & " probes meeting the filter linked to a " & clustersystem & " ID.")
   output.WriteLine (clusterC & " unique " & clustersystem & " genes met the criterion.")
   output.WriteLine (bigR & " genes meeting the criterion linked to a GO term.")
   output.WriteLine (GEXsize & " Probes in this dataset")
   output.WriteLine (GEXsize - noCluster & " Probes linked to a " & clustersystem & " ID.")
   output.WriteLine (distinctgenes & " genes in this dataset")
   output.WriteLine (bigN & " Genes linked to a GO term.")
   output.WriteLine ("The z score is based on an N of " & bigN & " and a R of " & bigR & " distinct genes in the GO.")
   output.WriteLine ("")
   output.WriteLine ("GOID" & Chr(9) & "GO Name" & Chr(9) & "GO Type" & Chr(9) & "Number Changed Local" & Chr(9) _
                     & "Number Measured Local" & Chr(9) & "Number in GO Local" & Chr(9) & "Percent Changed Local" _
                     & Chr(9) & "Percent Present Local" & Chr(9) & "Number Changed" _
                     & Chr(9) & "Number Measured" & Chr(9) & "Number in GO" _
                     & Chr(9) & "Percent Changed" & Chr(9) & "Percent Present" & Chr(9) & "Z Score" _
                     & Chr(9) & "PermuteP" & Chr(9) & "AdjustedP")
   While Not rstemp.EOF
      output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![GOName] & Chr(9) & rstemp![gotype] _
                  & Chr(9) & rstemp![ChangedLocal] & Chr(9) & rstemp![OnChipLocal] & Chr(9) _
                  & rstemp![InGOLocal] & Chr(9) & rstemp![percentagelocal] & Chr(9) _
                  & rstemp![presentLocal] & Chr(9) & rstemp![changed] & Chr(9) & rstemp![onChip] _
                  & Chr(9) & rstemp![ingo] & Chr(9) & rstemp![percentage] & Chr(9) _
                  & rstemp![present] & Chr(9) & Round(rstemp![zscore], 3) & Chr(9) & rstemp![permutep] & Chr(9) _
                  & rstemp![adjustedp])
      rstemp.MoveNext
   Wend
   output.Close
  
   CalculateGOResults = True
   For Each cg In ClusterGenes
      cg.reset 'reset for the next criterion
   Next cg
   dbChipData.Execute ("DELETE * FROM NestedResults") 'for the next criterion
   Exit Function
error:
   Select Case Err.Number
      Case 3024 'the error for not having the database
         MsgBox "The database MAPPFinder " & species & " was not found in the folder" _
         & " containing this application. Please move it to this folder or download it from www.GenMAPP.org.", vbOKOnly
      Case 3021
         MsgBox "The local MAPPs stored in " & species & " database are different than those currently loaded." _
            & " Please reload the appropriate MAPPs. MAPP " & rstemp2![MappNameField] & " not found.", vbOKOnly
      Case 5
         'no primary gene in primarygenes. this means that this pg has no cluster genes
         Resume nextPG
      Case 3010 'the tablename for the GOIDtoGene table already exists
         dbChipData.Execute "DROP table [" & tablename & "]"
         Resume createtable
      Case Else
         MsgBox "An error occurred while calculating the results. Please report error " & Err.Number _
            & " to GenMAPP@gladstone.ucsf.edu. Error message: " & Err.Description & "."
   End Select
   CalculateGOResults = False
End Function

Public Function CalculateLocalResults(criterion As Integer) As Boolean
   On Error GoTo error
   Dim noLocalMAPP As Integer
   Dim nolocalClusterC As Integer
   Dim localgene As Integer
   Dim output As TextStream, geneToMAPP As TextStream
   CalculateLocalResults = False
   Set rsFilter = dbExpressionData.OpenRecordset("SELECT OrderNo, ID, SystemCode FROM" _
                     & " Expression WHERE (" & sql(criterion) & ")")

   If rsFilter.EOF Then
      MsgBox "There are no genes in the Expression Dataset that meet the criterion you" _
            & " selected.", vbOKOnly
      Exit Function
   End If

   rsFilter.MoveLast
   rsFilter.MoveFirst
   metFilter = rsFilter.RecordCount
       
   TreeForm.resetProgress
   lblProgress.Caption = "Calculating Local MAPP results for the " & metFilter & " genes meeting criteria."
   frmCriteria.Refresh
   DoEvents
   progress = 0
   noLocalMAPP = 0
   localgene = LocalClusterGenes.count
   nolocalClusterC = 0
   localR = 0
   While rsFilter.EOF = False
      progress = progress + 1
      If progress Mod 10 = 0 Then
         lblProgress.Caption = "Local MAPP results calculated for " & progress & " out of " & metFilter & " genes meeting the criterion."
         frmCriteria.Refresh
         DoEvents
      End If
      
      
      Set pgene = LocalPrimaryGenes.Item(rsFilter!id)
      Set cgs = pgene.getClusterGenes
      If cgs.count = 0 Then
         nolocalClusterC = nolocalClusterC + 1
      Else
         For Each cg In cgs
            If Not cg.wasVisited Then
               localR = localR + 1
               cg.visit
               Set GOs = cg.getGOTerms
               If GOs.count = 0 Then
                  noLocalMAPP = noLocalMAPP + 1
               Else
                  For Each go In GOs
                     go.addChangedGene (rsFilter!id)
                     go.setChanged (go.getChanged + 1)
                  Next go
                  'Set GOs = cg.getNestedGOterms
                  'For Each go In GOs
                   '  go.setChanged (go.getChanged + 1)
                  'Next go
               End If
            End If
         Next cg
      End If
      rsFilter.MoveNext
   Wend
   lblProgress.Caption = "Calculating Z scores."
   frmCriteria.Refresh
   DoEvents
   'calculateRandN num
   localN = LocalClusterGenes.count
   calculateTestStatLocal
   If Statistics Then
      CalculateLocalPValues
   End If
   DoEvents
   
   Dim tblChangedGenes As TableDef
   Dim tablename As String
   tablename = Mid(fixFileName(txtFile), InStrRev(txtFile.Text, "\") + 1, Len(txtFile.Text) _
               - InStrRev(txtFile.Text, "\") - 1) & "-Criterion" _
               & criterion & "-Local"
createtablelocal:
   Set tblChangedGenes = dbChipDataLocal.CreateTableDef(tablename)
   tblChangedGenes.Fields.Append tblChangedGenes.CreateField("Primary", dbText, 255)
   tblChangedGenes.Fields.Append tblChangedGenes.CreateField("Related", dbText, 15)
   dbChipDataLocal.TableDefs.Append tblChangedGenes
   
   
   For Each MAPP In localMAPPsCollection
      MAPP.calculateResults
      results = MAPP.getResults 'this array has 11 elements. See GOterm class for more
      GOID = MAPP.getGOID
      For Each Gene In MAPP.getChangedGenes()
          dbChipDataLocal.Execute "INSERT INTO [" & tablename & "](Primary, Related) VALUES" _
                     & "('" & MAPP.getGOID & "', '" & Gene & "')"
      Next Gene
      MAPP.reset 'reset for the next criterion
      dbChipDataLocal.Execute "INSERT INTO Results (GOID, GOName, GOType, Changed" _
                              & ", OnChip, InGO, Percentage, Present," _
                              & " Zscore, PermuteP, AdjustedP)" _
                              & " VALUES ('" & GOID & "', '" & GOID & "', '" _
                              & "L" & "', " & results(5) _
                              & ", " & results(6) & ", " & results(7) & ", " & results(8) _
                              & ", " & results(9) & ", " & results(10) & ", " _
                              & results(11) & ", " & results(12) & ")"
      
      DoEvents
   Next MAPP
   
   Set idxGOID = tblChangedGenes.CreateIndex(tablename)
   idxGOID.Fields.Append tblChangedGenes.CreateField("Primary", dbText, 15)
   tblChangedGenes.Indexes.Append idxGOID
   'now we have indexed the changedgenes table for this criterion.
   
   
   If Statistics Then
      Set rstemp = dbChipDataLocal.OpenRecordset("SELECT GOID, Changed" _
               & ", OnChip, InGO, Percentage, Present, Zscore, permuteP, adjustedP FROM Results " _
               & "ORDER BY permuteP, Zscore DESC")
   Else
      Set rstemp = dbChipDataLocal.OpenRecordset("SELECT GOID, Changed" _
               & ", OnChip, InGO, Percentage, Present, Zscore, PermuteP, adjustedP FROM Results " _
               & "ORDER BY Zscore DESC, OnChip DESC")
   End If
   Set output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Criterion" & criterion & "-Local.txt")
   
   output.WriteLine ("MAPPFinder 2.0 Results for the Local MAPPs")
   output.WriteLine ("File: " & filelocation)
   output.WriteLine ("Table: " & tablename)
   output.WriteLine ("Database: " & databaseloc)
   output.WriteLine ("colors:|" & colorset & "|")
   output.WriteLine (GODate)
   output.WriteLine (species)
   If Statistics Then
      output.WriteLine ("Pvalues = true")
   Else
      output.WriteLine ("Pvalues = false")
   End If
   output.WriteLine ("Calculation Summary:")
   output.WriteLine (metFilter & " probes met the " & sql(criterion) & " criteria.")
   output.WriteLine (metFilter - nolocalClusterC & " probes meeting the criterion linked to a MAPP system.")
   output.WriteLine (localR & " genes linked to a MAPP.")
   output.WriteLine (GEXsize & " Probes in this dataset")
   output.WriteLine (GEXsize - nolocalCluster & " Probes linked to an ID in a MAPP system.")
   output.WriteLine (localN & " Genes are linked to the Local MAPPs.")
   output.WriteLine ("The z score is based on an N of " & localN & " and a R of " & localR & " distinct genes in the GO.")
   output.WriteLine ("")
   output.WriteLine ("MAPP Name" & Chr(9) & "Number Changed" _
                     & Chr(9) & "Number Measured" & Chr(9) & "Number On MAPP" _
                     & Chr(9) & "Percent Changed" & Chr(9) & "Percent Present" & Chr(9) & "Z Score" _
                     & Chr(9) & "PermuteP" & Chr(9) & "AdjustedP")
   While Not rstemp.EOF
      output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![changed] & Chr(9) & rstemp![onChip] _
                  & Chr(9) & rstemp![ingo] & Chr(9) & rstemp![percentage] & Chr(9) _
                  & rstemp![present] & Chr(9) & Round(rstemp![zscore], 3) & Chr(9) & rstemp![permutep] _
                  & Chr(9) & rstemp![adjustedp])
      rstemp.MoveNext
   Wend
   output.Close
  
   dbChipDataLocal.Execute "Delete * FROM Results"
   For Each cg In LocalClusterGenes
      cg.reset
   Next cg
   CalculateLocalResults = True
   Exit Function
error:
   Select Case Err.Number
      Case 3010 'the tablename for the GOIDtoGene table already exists
         dbChipData.Execute "DROP table [" & tablename & "]"
         Resume createtablelocal
      Case Else
         MsgBox "Error " & Err.Number & ", " & Err.Description & " occurred in the CalculateLocalResults function.", vbOKOnly
   End Select
End Function

Public Sub CalculatePValues()
MousePointer = vbHourglass
   Dim i As Long, j As Long, k As Long
   Dim term As goterm
   Dim GOIDs As Collection
   Dim genes() As New ClusterGene
   Dim GOarray() As New goterm
   Randomize 'initialzes random number generator
   ReDim genes(bigN - 1) As New ClusterGene
   ReDim GOarray(GOterms.count - 1) As New goterm
   totalGenes = 0
   
   For Each term In GOterms
      term.setZeroZ bigR, bigN 'each GO term initialized to its Z score if no genes are changed
   Next term
   
   'we now move all the clustergenes that link to GO into an array of bigN size.
   'this makes it easier to sort them.
   i = 0
   For Each cg In ClusterGenes
      If cg.getNestedGOterms.count > 0 Then
         Set genes(i) = cg
         i = i + 1
      End If
   Next cg
   
   
   'we also need to put the GO terms into an array so they can be sorted for the multiple hypothesis testing adjustment
   i = 0
   For Each term In GOterms
      Set GOarray(i) = term
      i = i + 1
   Next term
   
   sortGOterms GOarray, 0, GOterms.count - 1 'the GOarray is sorted by the absolute value of the real Z score\

   'now select bigR genes at random from rsGenes
   'add those genes to all of there associated GO terms. and all of those term's parents.
   'do this 1000 times, create a zsum distribution
   For i = 0 To TRIALS - 1 'bootstrap x trials
      'Debug.Print i
      
      If (i Mod 10 = 0) Then
         lblProgress.Caption = i & "out of " & TRIALS & " Bootstrap trials completed."
         frmCriteria.Refresh
         DoEvents
      End If
      
      
      For Each term In GOterms
         term.resetME
      Next term
      
      'resetGeneIndex
      For j = 0 To bigN - 1
         genes(j).setProb Rnd(Time())
      Next j
            
      genesort genes(), 0, bigN - 1 'sort the genes by the random number between 0..1
      
      
      'now just take the first bigR as the "changed" genes
      For j = 0 To bigR - 1
         Set GOIDs = genes(j).getNestedGOterms
         For Each term In GOIDs
            'Debug.Print GOterms.count
            term.addOne
         Next term
      Next j
      
      'now the bigR genes have been scattered accross the GO
      'now calculate z scores for each GO term
      For Each term In GOterms
         term.CalculatePermuteZ bigR, bigN
      Next term
      
      GOarray(GOterms.count - 1).setMonoZ (GOarray(GOterms.count - 1).getPermuteZ)
      For k = GOterms.count - 1 To 1 Step -1 'make the Zs monotonoically decreasing....
         If GOarray(k).getMonoZ > GOarray(k - 1).getPermuteZ Then
            GOarray(k - 1).setMonoZ (GOarray(k).getMonoZ)
         Else
            GOarray(k - 1).setMonoZ (GOarray(k - 1).getPermuteZ)
         End If
      Next k
      'For k = GOterms.count - 1 To 1 Step -1 'make the Zs monotonoically decreasing....
       '  Debug.Print GOarray(k).getPermuteZ & Chr(9) & GOarray(k).getMonoZ
      'Next k
      
      
      'Debug.Print zsum(i)
      'output2.WriteLine (zsum(i))
   Next i
   
   
   For Each term In GOterms
      term.CalculatePValues (TRIALS)
   Next term
   
   For k = 0 To GOterms.count - 2
      'make the mono Ps monotonically increasing
      If GOarray(k).getMonoP > GOarray(k + 1).getMonoP Then
         GOarray(k + 1).setMonoP (GOarray(k).getMonoP)
      Else
         GOarray(k + 1).setMonoP (GOarray(k + 1).getMonoP)
      End If
      'Debug.Print GOarray(i).getZscore & Chr(9) & GOarray(i).getMonoP
   Next k
   
error:
   Select Case Err.Number
      Case 5
         'if you've hit here, then there's no GOID for that gene.
    
        
   End Select
   mousepoint = vbDefault
End Sub
Public Sub CalculateLocalPValues()
'this is essentially the same code as above, but it uses the collections for the Local MAPPs
'Dim LocalClusterGenes As New Collection
   'key is an ID of a systemcode supported by the GenesToMAPP table
   'value is a clustergene object for that gene on a mapp
'Dim localMAPPsCollection As New Collection
   'key is a mappname
   'object is a GOterm object (no nested values)

MousePointer = vbHourglass
   Dim i As Long, j As Long, k As Long
   Dim term As goterm
   Dim GOIDs As Collection
   Dim genes() As New ClusterGene
   Dim GOarray() As New goterm
   Randomize 'initialzes random number generator
   ReDim genes(localN - 1) As New ClusterGene
   ReDim GOarray(localMAPPsCollection.count - 1) As New goterm
   totalGenes = 0
   
   For Each term In localMAPPsCollection
      term.setZeroZ localR, localN 'each GO term initialized to its Z score if no genes are changed
   Next term
   
   'we now move all the clustergenes that link to GO into an array of bigN size.
   'this makes it easier to sort them.
   i = 0
   For Each cg In LocalClusterGenes
      If cg.getGOTerms.count > 0 Then
         Set genes(i) = cg
         i = i + 1
      End If
   Next cg
   
   
   'we also need to put the GO terms into an array so they can be sorted for the multiple hypothesis testing adjustment
   i = 0
   For Each term In localMAPPsCollection
      Set GOarray(i) = term
      i = i + 1
   Next term
   
   sortGOterms GOarray, 0, localMAPPsCollection.count - 1 'the GOarray is sorted by the absolute value of the real Z score\

   'now select bigR genes at random from rsGenes
   'add those genes to all of there associated GO terms. and all of those term's parents.
   'do this 1000 times, create a zsum distribution
   For i = 0 To TRIALS - 1 'bootstrap x trials
      'Debug.Print i
      
      If (i Mod 10 = 0) Then
         lblProgress.Caption = i & "out of " & TRIALS & " Bootstrap trials completed for Local Results."
         frmCriteria.Refresh
         DoEvents
      End If
      
      
      For Each term In localMAPPsCollection
         term.resetME
      Next term
      
      'resetGeneIndex
      For j = 0 To localN - 1
         genes(j).setProb Rnd(Time())
      Next j
            
      genesort genes(), 0, localN - 1 'sort the genes by the random number between 0..1
      
      
      'now just take the first bigR as the "changed" genes
      For j = 0 To localR - 1
         Set GOIDs = genes(j).getGOTerms
         For Each term In GOIDs
            'Debug.Print GOterms.count
            term.addOne
         Next term
      Next j
      
      'now the bigR genes have been scattered accross the GO
      'now calculate z scores for each GO term
      For Each term In localMAPPsCollection
         term.CalculatePermuteZ localR, localN
      Next term
      
      GOarray(localMAPPsCollection.count - 1).setMonoZ (GOarray(localMAPPsCollection.count - 1).getPermuteZ)
      For k = localMAPPsCollection.count - 1 To 1 Step -1 'make the Zs monotonoically decreasing....
         If GOarray(k).getMonoZ > GOarray(k - 1).getPermuteZ Then
            GOarray(k - 1).setMonoZ (GOarray(k).getMonoZ)
         Else
            GOarray(k - 1).setMonoZ (GOarray(k - 1).getPermuteZ)
         End If
      Next k
      'For k = GOterms.count - 1 To 1 Step -1 'make the Zs monotonoically decreasing....
       '  Debug.Print GOarray(k).getPermuteZ & Chr(9) & GOarray(k).getMonoZ
      'Next k
      
      
      'Debug.Print zsum(i)
      'output2.WriteLine (zsum(i))
   Next i
   
   
   For Each term In localMAPPsCollection
      term.CalculatePValues (TRIALS)
   Next term
   
   For k = 0 To localMAPPsCollection.count - 2
      'make the mono Ps monotonically increasing
      If GOarray(k).getMonoP > GOarray(k + 1).getMonoP Then
         GOarray(k + 1).setMonoP (GOarray(k).getMonoP)
      Else
         GOarray(k + 1).setMonoP (GOarray(k + 1).getMonoP)
      End If
      'Debug.Print GOarray(i).getZscore & Chr(9) & GOarray(i).getMonoP
   Next k
   
error:
   Select Case Err.Number
      Case 5
         'if you've hit here, then there's no GOID for that gene.
       
        
   End Select
   mousepoint = vbDefault
End Sub


Public Sub sortGOterms(Garray() As goterm, start As Long, finish As Long)
   'Input - an array of GOterm
   'Output - the array sorted by the real Z score of the GO term. The sort is a mergesort (NlogN).
   
   Dim temp As goterm
   Dim split As Long
   Dim counter1 As Long, counter2 As Long
   Dim temparray() As goterm
   Dim i As Long
   
   If start = finish Then 'only one, don't sort it
      'End Sub
   ElseIf finish - start = 1 Then 'there are two left
      If Garray(start).getZscore < Garray(finish).getZscore Then 'need to swap them
         Set temp = Garray(start)
         Set Garray(start) = Garray(finish)
         Set Garray(finish) = temp
      End If
      'End Sub
   Else ' need to partition and then merge
      ReDim temparray(finish - start + 1) As goterm
      split = (finish + start) / 2
      sortGOterms Garray, start, split
      sortGOterms Garray, split + 1, finish
      counter2 = split + 1
      counter1 = start
      i = 0
      While counter1 <= split And counter2 <= finish
         If Garray(counter1).getZscore < Garray(counter2).getZscore Then 'counter2 goes into merge first swap
            Set temparray(i) = Garray(counter2)
            counter2 = counter2 + 1
         Else 'put counter 1 in first and move forward
            Set temparray(i) = Garray(counter1)
            counter1 = counter1 + 1
         End If
         i = i + 1
      Wend
      If counter1 <= split Then 'there are still first half strings to be added
         While counter1 <= split
            Set temparray(i) = Garray(counter1)
            i = i + 1
            counter1 = counter1 + 1
         Wend
      End If
      If counter2 <= finish Then 'there are still second half strings to be added
         While counter2 <= finish
            Set temparray(i) = Garray(counter2)
            i = i + 1
            counter2 = counter2 + 1
         Wend
      End If
      
      For i = 0 To finish - start
         Set Garray(i + start) = temparray(i)
      Next i
   End If
      
End Sub
Private Sub genesort(tester() As ClusterGene, start As Long, finish As Long)
    'Input - an array of strings
   'Output - the array sorted alphabetically. The sort is a mergesort (NlogN).
   'The sort is not case sensitive (ie A = a).
   Dim temp As ClusterGene
   Dim split As Long
   Dim counter1 As Long, counter2 As Long
   Dim temparray() As ClusterGene
   Dim i As Long
   
   If start = finish Then 'only one, don't sort it
      'End Sub
   ElseIf finish - start = 1 Then 'there are two left
      If tester(start).getProb > tester(finish).getProb Then 'need to swap them
         Set temp = tester(start)
         Set tester(start) = tester(finish)
         Set tester(finish) = temp
      End If
      'End Sub
   Else ' need to partition and then merge
      ReDim temparray(finish - start + 1) As ClusterGene
      split = (finish + start) / 2
      genesort tester, start, split
      genesort tester, split + 1, finish
      counter2 = split + 1
      counter1 = start
      i = 0
      While counter1 <= split And counter2 <= finish
         If tester(counter1).getProb > tester(counter2).getProb Then 'counter2 goes into merge first swap
            Set temparray(i) = tester(counter2)
            counter2 = counter2 + 1
         Else 'put counter 1 in first and move forward
            Set temparray(i) = tester(counter1)
            counter1 = counter1 + 1
         End If
         i = i + 1
      Wend
      If counter1 <= split Then 'there are still first half strings to be added
         While counter1 <= split
            Set temparray(i) = tester(counter1)
            i = i + 1
            counter1 = counter1 + 1
         Wend
      End If
      If counter2 <= finish Then 'there are still second half strings to be added
         While counter2 <= finish
            Set temparray(i) = tester(counter2)
            i = i + 1
            counter2 = counter2 + 1
         Wend
      End If
      
      For i = 0 To finish - start
         Set tester(i + start) = temparray(i)
      Next i
   End If
      
End Sub

Public Sub LoadSpecies()
   Dim dbMAPPfinder As Database
   Dim rsSpecies As Recordset
   Set dbMAPPfinder = OpenDatabase(databaseloc)
   Set rsSpecies = dbMAPPfinder.OpenRecordset("SELECT [MOD] FROM Systems WHERE [MOD] <> Null" _
                                          & " AND [Date] <> Null ORDER BY [MOD]")
      'the database should be species specific, so this will in most cases by 1, but SwissProt shows up as a MOD
      'and is also in SwissProt.
   If rsSpecies.RecordCount = 1 Then
      lblspecies.Caption = rsSpecies![Mod]
   ElseIf rsSpecies.RecordCount = 2 Then
      'SwissProt and a MOD
      While rsSpecies.EOF = False
         If rsSpecies![Mod] <> "Homo sapiens" Then
            lblspecies.Caption = rsSpecies![Mod]
         End If
         rsSpecies.MoveNext
      Wend
   Else '
      MsgBox "This database has multiple species. MAPPFinder needs you to use a species specific database.", vbOKOnly
   End If
   dbMAPPfinder.Close
End Sub

