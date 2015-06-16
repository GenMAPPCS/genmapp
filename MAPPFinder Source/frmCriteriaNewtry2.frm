VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCriteria 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Calculate New Results"
   ClientHeight    =   7815
   ClientLeft      =   6405
   ClientTop       =   1500
   ClientWidth     =   6390
   Icon            =   "frmCriteriaNewtry2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   6390
   Begin VB.ComboBox cmbSpecies 
      Height          =   315
      ItemData        =   "frmCriteriaNewtry2.frx":08CA
      Left            =   1440
      List            =   "frmCriteriaNewtry2.frx":08D1
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3480
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check2"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Dataset"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -120
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   5520
      Width           =   4575
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Browse"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ListBox lstcriteria 
      Height          =   1230
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.ListBox lstColorSet 
      Height          =   1230
      ItemData        =   "frmCriteriaNewtry2.frx":08DD
      Left            =   360
      List            =   "frmCriteriaNewtry2.frx":08DF
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdRunMAPPFinder 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Run MAPPFinder"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "(Gene Ontology Results and Local Results will be added to the file name.)"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   6000
      Width           =   5415
   End
   Begin VB.Label lblProgress 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   7080
      Width           =   5415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gray = No Local MAPPs loaded"
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Local MAPPs"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gene Ontology"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select the type of analysis you would like to run."
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   4320
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save Results as:"
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select your species:"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select Criterion to filter by:"
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select Color Set:"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Please select the Color Set and Criterion you would like to use to filter the data."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   5175
   End
   Begin VB.Menu File 
      Caption         =   "File"
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

'The code is pretty long, for the relative simplicity of the program. Unfortunately,
'each species is slightly different in how GO annotations are added, so the MAPPFinder
'algorithm could not be generalized. Each new species will require modifications
'to the existing algorithm.
 
 
 
 Const MAX_CRITERIA = 30
 Const GOCount = 46000 'the highest number assigned to a go term. This is pretty large, but oh well.
 Const MAPPCount = 5000 'the highest number of mapps that can be loaded as local
 Const MAX_RELATIONS = 30

 Dim rscolorsets As DAO.Recordset  'stores the colorsets of the .gex file
 Dim dbExpressionData As Database 'stores the expression table of the .gex file
 Dim sql(MAX_CRITERIA) As String
 Public species As String
 Dim fullname As String
 Dim newfilename As String
 Dim dbMAPPfinder As Database 'the MAPPFinder database MAPPFinder 1.0.mdb
 Dim dbChipData As Database 'the entire chips annotations.
 Dim gotable As String, FSO As Object
 Dim expressionName As String
 Dim filelocation As String
 Dim chipName As String
 Public clusterSystem As String, clusterCode As String
 Dim speciesselected As Boolean, geneontology As Boolean, localMAPPsLoaded As Boolean
 Dim colorsetclicked As Boolean, criteriaclicked As Boolean
 Dim GOrelation As String
 Dim chipbuiltOK As Boolean
 Dim localR As Long, bigR As Long
 Dim localN As Long, bigN As Long
 Dim relations(MAX_RELATIONS, 2) As String
 ' 0 = relation
 ' 1 = P or R or S(is the Expression data ID the primary or related field of the relation)
                  ' or are the Expression ID and the Cluster ID the same type
 ' 2 = systemcode of gex
 Dim genes(1, 1) As String
   '0 = the geneID in the GEX
   '1 = the geneID in the clustersystem (MGI,SwissProt, etc..)
   'this will be redimed when i know how big the GEX is.
Dim clusterGenes As New Collection
      
 
 
         
Public Sub Load(FileName As String)
   Dim colorset As String
   Dim slash As Integer
   speciesselected = False
   chipbuiltOK = False
   filelocation = FileName
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
   
   If UCase(Dir(Module1.programpath & "LocalMAPPs.txt")) <> UCase("LocalMAPPs.txt") Then
      Check2.Enabled = False
   Else
      Check2.Enabled = True
   End If
   colorsetclicked = False
   criteriaclicked = False
   geneontology = False
   localMAPPsLoaded = False
   cmbSpecies.AddItem "Select your species"
   cmbSpecies.AddItem "Arabidopsis thaliana"
   cmbSpecies.AddItem "Caenorhabditis elegans"
   cmbSpecies.AddItem "Drosophila Melanogaster"
   cmbSpecies.AddItem "Homo sapiens"
   cmbSpecies.AddItem "Mus Musculus"
   cmbSpecies.AddItem "Rattus norvegicus"
   cmbSpecies.AddItem "Saccharomyces cerevisiae"
   
   
   frmInput.Hide
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
   If localMAPPsLoaded = True Then
      localMAPPsLoaded = False
   Else
      localMAPPsLoaded = True
   End If
End Sub

Private Sub resetcheck2()
   Check2.Value = 0
End Sub






Private Sub cmdFile_Click()
   Dim FileName As String
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "Text Files|*.txt"
   CommonDialog1.ShowSave
   txtFile.Text = CommonDialog1.FileName
   
   If txtFile.Text <> "" Then
   
   FileName = Left(txtFile.Text, Len(txtFile.Text) - 4)
   If Dir(FileName & "-Gene Ontology Results.txt") <> "" Or Dir(FileName & "-Local Results.txt") <> "" Then
      If MsgBox("Overwrite the existing " & txtFile.Text & "-Gene Ontology Results and -Local Results?", vbOKCancel) = vbCancel Then
         txtFile.Text = ""
      End If
   End If
   End If
End Sub
 
Private Sub cmdRunMAPPFinder_Click()
   'On Error GoTo error

   Dim rsFilter As DAO.Recordset, rsType As DAO.Recordset, rsRelation As Recordset
   Dim tblGOAll As TableDef, rstemp2 As DAO.Recordset
   Dim tblGO As TableDef, tblResults As TableDef, rschip As DAO.Recordset, tblnestedResults As TableDef
   Dim rsFunction As DAO.Recordset, rsProcess As DAO.Recordset, rsComponent As DAO.Recordset
   Dim percentage As Single, metFilter As Integer, noGO As Integer 'no GeneOntology available
   Dim others As Integer, noSwissProt As Integer 'no swissprot counts the number of genes that can't be converted
   Dim present As Single
   Dim Output As TextStream
   Dim criteria As String, trembl As Boolean
   Dim genmappID As String, MGIID As String, GOID As String
   Dim MGIsAdded As New Collection, GenMAPPsAdded As New Collection
   Dim genecounter() As String, GOArray(GOCount) As Integer, mapparray() As Integer
   Dim progress As Integer, rsinGONested As DAO.Recordset, indata As Integer
   Dim i As Long, numofsystems As Integer, rsGenes As Recordset
   Dim arraysize As Long, relationnotfound As Boolean
   Dim r As Long
   Dim n As Long
   Dim teststat As Double
   If cmbSpecies.Text = "Select your species" Then
      MsgBox "You must select a species before proceeding.", vbOKOnly
      GoTo nospeciesselected
   End If
   
   If geneontology = False And localMAPPsLoaded = False Then
      MsgBox "You have not selected the type of MAPPFinder analysis you would like to run." _
            & " Please select Gene Ontology, Local MAPPs, or both.", vbOKOnly
      GoTo nospeciesselected
   End If
   
   If txtFile.Text = "" Then
      MsgBox "You have not selected a file to save the results to. Please do so now.", vbOKOnly
      GoTo nospeciesselected
   End If
   
   If colorsetclicked = False Then
      MsgBox "You have not selected a color set, please do so.", vbOKOnly
      GoTo nospeciesselected
   End If
   
   If criteriaclicked = False Then
      MsgBox "You have not selected a criteria, please do so.", vbOKOnly
      GoTo nospeciesselected
   End If
   
   MousePointer = vbHourglass
   progress = 0
   InitializeGOArray GOArray
   Set FSO = CreateObject("Scripting.FileSystemObject")
   
   Set dbMAPPfinder = OpenDatabase(DatabaseLoc)
   TreeForm.setDatabase (DatabaseLoc)
     
   
   Set rsFilter = dbExpressionData.OpenRecordset("SELECT OrderNo, ID, SystemCode FROM" _
                        & " Expression WHERE (" & sql(lstcriteria.ListIndex) & ")")
   
   If rsFilter.EOF Then
      MsgBox "There are no genes in the Expression Dataset that meet the criterion you" _
            & " selected.", vbOKOnly
      GoTo noerror
   End If
   
   rsFilter.MoveLast
   rsFilter.MoveFirst
   metFilter = rsFilter.RecordCount
   'check this
   If dbExpressionData.TableDefs.count > 11 Then 'somehow the tables didn't get deleted before
      dbExpressionData.TableDefs.Delete ("GO")
      dbExpressionData.TableDefs.Delete ("Results")
      dbExpressionData.TableDefs.Delete ("NestedResults")
   End If
     
   TreeForm.resetProgress
   lblProgress.Caption = "Calculating MAPPFinder results for the " & metFilter & " genes meeting criteria."
   frmCriteria.Refresh
   MousePointer = vbHourglass
   If geneontology Then
   'get the clustersystem from the systems table
   'get the system codes of the ED
   'find all relations
      
      
      
                  
         
                                                
   
   
   
   
   
   Select Case species
'=HUMAN========================================================================================
'==============================================================================================
'in the human code, MGI is a left over from the mouse code. Here MGI = SwissProt
      
      Case "human"
         While rsFilter.EOF = False
            progress = progress + 1
            If progress Mod 10 = 0 Then
               lblProgress.Caption = progress & " genes out of " & metFilter & " meeting the criterion are linked to GO."
               frmCriteria.Refresh
            End If
            If UCase(rsFilter![primaryType]) = "G" Then 'convert to MGI via MGI-GenBank table
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT primary from [SwissProt-GenBank] Where" _
                     & " related = '" & rsFilter![primary] & "'")
               If rstemp.EOF Then
                  nosp = nosp + 1
               Else
                  For Each element In MGIsAdded
                     If element = rstemp![primary] Then
                        found = True
                        GoTo foundMGIHu
                     End If
                  Next element
foundMGIHu:
               
                  If Not found Then 'this MGI hasn't been previously hit
                     MGIID = rstemp![primary] 'for some reason you need to cast the record to a string to get it into a collection
                     MGIsAdded.Add MGIID, MGIID
                     Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [GenBank-GeneOntology] where primary = '" & rsFilter![primary] & "'")
                     'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                     If rstemp2.EOF Then
                        noGO = noGO + 1
                     Else
                        While rstemp2.EOF = False
                           GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                           'each element of the array is one GOID. Add one to indicate that there was
                           'a hit for that GOID
                           dbExpressionData.Execute ("INSERT into GO (primary, GOID) VALUES ('" & rsFilter![primary] _
                                             & "', '" & rstemp2![related] & "')")
                           rstemp2.MoveNext
                        Wend
                     End If
                  End If
               End If
            
            ElseIf UCase(rsFilter![primaryType]) = "S" Or UCase(rsFilter![primaryType]) = "N" Then
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT ID from [SwissProt] Where" _
                     & " ID = '" & rsFilter![primary] & "'")
               If rstemp.EOF Then 'try trembl
                  Set rstemp = dbMAPPfinder.OpenRecordset("SELECT ID FROM SwissProt Where" _
                              & " Accession like '*|" & rsFilter![primary] & "|*'")
               End If
              
               If rstemp.EOF Then
                  nosp = nosp + 1
               Else
               For Each geneid In MGIsAdded
                  If geneid = rstemp![ID] Then
                     found = True
                     GoTo foundMGISPhu
                  End If
               Next geneid
foundMGISPhu:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rstemp![ID]
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntology] where primary = '" & rsFilter![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  
                  If rstemp2.EOF Then
                     Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntology] where primary = '" & rstemp![ID] & "'")
                     If rstemp2.EOF Then 'both swissprot and trembl failed
                        noGO = noGO + 1
                     End If
                  End If
                  While rstemp2.EOF = False
                     GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                     dbExpressionData.Execute ("INSERT into GO (primary, GOID) VALUES ('" & rsFilter![primary] _
                                             & "', '" & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               End If
               End If
            Else
               other = others + 1
            End If
            rsFilter.MoveNext
            found = False
         Wend 'now you've attached GOIDs to as many unique genes as possible. Unique is defined as
            'a individual MGI ID. Next, report the count for each GOID
         
          
         For i = 0 To GOCount
            If GOArray(i) <> 0 Then 'there was a hit for this GOID calculate the percentages
            GOID = MakeGOID(i)
            'for each GOID in GOArray - divide the GOIDCount in the array by the the chip count
             'to get the percentage changed in the dataset
            'insert into results GOType, GOID, percentage
            Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT NumberOfDups FROM [" & gotable & "Count]" _
                     & " WHERE [RelatedField] = '" & GOID & "'")
            'temp 2 now has the number of times the GO term is represented in that species' table
         
            Set rschip = dbChipData.OpenRecordset("SELECT NumOfGOID FROM [Chip] WHERE" _
                        & " [GOIDCount] = '" & GOID & "'")
            'rsChip now has the number of times the GO term is on the chip
         
            percentage = (GOArray(i) / rschip![NumofGoid]) * 100
            present = (rschip![NumofGoid] / rstemp2![numberofdups]) * 100
            'percentage = # of times changed/ # of times measured on chip * 100
            Set rsType = dbMAPPfinder.OpenRecordset("SELECT Type, Name FROM GeneOntology WHERE" _
                     & " ID = '" & GOID & "'")
            dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, GoName, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('" & rsType![Type] & "', '" _
                        & GOID & "', '" & rsType![name] & "', " & GOArray(i) & ", " _
                        & rschip![NumofGoid] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
            End If
         Next i
         GOID = "GO"
         Set rsinGONested = dbMAPPfinder.OpenRecordset("SeLECT GOIDCount FROM " & species & "HierarchyCount Where GOID = '" & GOID & "'")
         arraysize = rsinGONested![goidcount] * (4)
         ReDim genecounter(arraysize) As String 'you now have an array twice as big as the number of genes in this sub-graph. Allows for plenty of duplicates.
         indata = TreeForm.nestedResults(dbExpressionData, dbChipData, TreeForm.root, genecounter, 0)  '"GO" is the root of the tree
         calculateTestStat
      Set rstemp = dbExpressionData.OpenRecordset("SELECT Gotype, GOID, GOName, InData, OnChip, InGO," _
                  & " Percentage, Present, InDataNested, OnChipNested, InGoNested, PercentageNested," _
                  & " PresentNested, TestStat FROM NestedResults " _
                  & "ORDER BY NestedResults![TestStat] DESC, NestedResults![OnChipNested] DESC")
      Set Output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Gene Ontology Results.txt")
      Output.WriteLine ("MAPPFinder 1.0 Results for the Gene Ontology")
      Output.WriteLine ("File: " & filelocation)
      Output.WriteLine (GODate)
      Output.WriteLine ("Statistics:")
      Output.WriteLine (metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria.")
      frmCalculation.lblGOcriteria.Caption = metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria."
      Output.WriteLine (nosp & " Genes did not link to a MGI term.")
      Output.WriteLine (noGO + others & " Genes did not link to a GO term.")
      frmCalculation.lblGOnotfound.Caption = noGO + nosp + others & " Genes did not link to a GO term."
      Output.WriteLine (metFilter - noSwissProt - noGO - others & " genes were used to calculate the results shown below.")
      frmCalculation.lblGOUsed.Caption = metFilter - noSwissProt - noGO - others & " genes were used to calculate the results shown below."
      Output.WriteLine ("The z score is based on an N of " & bigN & " and a R of " & bigR & " distinct genes.")
      frmCalculation.lblGOStat = "The z score is based on an N of " & bigN & " and a R of " & bigR & " distinct genes."
      Output.WriteLine ("")
      Output.WriteLine ("GOID" & Chr(9) & "GO Name" & Chr(9) & "GO Type" & Chr(9) & "Number Changed" & Chr(9) _
                        & "Number Measured" & Chr(9) & "Number in GO" & Chr(9) & "Percent Changed" _
                        & Chr(9) & "Percent Present" & Chr(9) & "Number Changed in Hierarchy" _
                        & Chr(9) & "Number Measured in Hierarchy" & Chr(9) & "Number in GO" _
                        & " Hierarchy" & Chr(9) & "Percent Changed in Hierarchy" & Chr(9) _
                        & "Percent Present in Hierarchy" & Chr(9) & "z Score")
         While Not rstemp.EOF
            Output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![GOName] & Chr(9) & rstemp![gotype] _
                        & Chr(9) & rstemp![indata] _
                        & Chr(9) & rstemp![onChip] & Chr(9) & rstemp![ingo] & Chr(9) _
                        & rstemp![percentage] & Chr(9) & rstemp![present] & Chr(9) _
                        & rstemp![indatanested] & Chr(9) & rstemp![onchipnested] & Chr(9) _
                        & rstemp![ingonested] & Chr(9) & rstemp![percentagenested] & Chr(9) _
                        & rstemp![presentnested] & Chr(9) & rstemp![teststat])
            rstemp.MoveNext
         Wend
      Output.Close
     
      
      
'=MOUSE==========================================================================================
'===============================================================================================
'The MGI-GenBank relationship has been made a 1:1 relationship for the purposes of this
'program. This means that 996 (0.01%)genes were removed from the MGI-GenBank table.
    Case "mouse"
         While rsFilter.EOF = False
            progress = progress + 1
            If progress Mod 10 = 0 Then
               lblProgress.Caption = progress & " genes out of " & metFilter & " meeting the criterion are linked to GO."
               frmCriteria.Refresh
            End If
            If UCase(rsFilter![primaryType]) = "G" Then 'convert to MGI via MGI-GenBank table
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT primary from [MGI-GenBank] Where" _
                     & " related = '" & rsFilter![primary] & "'")
               If rstemp.EOF Then
                  nosp = nosp + 1
               Else
                  For Each element In MGIsAdded
                     If element = rstemp![primary] Then
                        found = True
                        GoTo foundMGI
                     End If
                  Next element
foundMGI:
               
                  If Not found Then 'this MGI hasn't been previously hit
                     MGIID = rstemp![primary] 'for some reason you need to cast the record to a string to get it into a collection
                     MGIsAdded.Add MGIID, MGIID
                     Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [GenBank-GeneOntology] where primary = '" & rsFilter![primary] & "'")
                     'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                     If rstemp2.EOF Then
                        noGO = noGO + 1
                     Else
                        While rstemp2.EOF = False
                           GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                           'each element of the array is one GOID. Add one to indicate that there was
                           'a hit for that GOID
                           dbExpressionData.Execute ("INSERT into GO (primary, GOID) VALUES ('" & rsFilter![primary] _
                                             & "', '" & rstemp2![related] & "')")
                           rstemp2.MoveNext
                        Wend
                     End If
                  End If
               End If
            
            ElseIf UCase(rsFilter![primaryType]) = "S" Or UCase(rsFilter![primaryType]) = "N" Then
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT related from [SwissProt-MGI] Where" _
                     & " primary = '" & rsFilter![primary] & "'")
               If rstemp.EOF Then 'try trembl
                  Set rstrembl = dbMAPPfinder.OpenRecordset("SELECT ID FROM SwissProt Where" _
                              & " Accession like '*|" & rsFilter![primary] & "|*'")
                  
                  If rstrembl.EOF = False Then
                     trembl = True
                     Set rstemp = dbMAPPfinder.OpenRecordset("SELECT related from [SwissProt-MGI] Where" _
                     & " primary = '" & rstrembl![ID] & "'")
                  End If
               End If
         
               If rstemp.EOF Then
                  nosp = nosp + 1
               Else
               For Each geneid In MGIsAdded
                  If geneid = rstemp![related] Then
                     found = True
                     GoTo foundMGISP
                  End If
               Next geneid
foundMGISP:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rstemp![related]
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rsFilter![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  
                  If rstemp2.EOF Then
                     If trembl Then
                        Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                           & " [SwissProt-GeneOntologyMAPPFInder] where primary = '" & rstrembl![ID] & "'")
                     End If
                     If rstemp2.EOF Then 'both swissprot and trembl failed
                        noGO = noGO + 1
                     End If
                  End If
                  While rstemp2.EOF = False
                     GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                     dbExpressionData.Execute ("INSERT into GO (primary, GOID) VALUES ('" & rsFilter![primary] _
                                             & "', '" & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               End If
               End If
            Else
               other = others + 1
            End If
            trembl = False
            rsFilter.MoveNext
            found = False
         Wend 'now you've attached GOIDs to as many unique genes as possible. Unique is defined as
            'a individual MGI ID. Next, report the count for each GOID
         
          
         For i = 0 To GOCount
            If GOArray(i) <> 0 Then 'there was a hit for this GOID calculate the percentages
            GOID = MakeGOID(i)
            'for each GOID in GOArray - divide the GOIDCount in the array by the the chip count
             'to get the percentage changed in the dataset
            'insert into results GOType, GOID, percentage
            Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT NumberOfDups FROM [" & gotable & "Count]" _
                     & " WHERE [RelatedField] = '" & GOID & "'")
            'temp 2 now has the number of times the GO term is represented in that species' table
         
            Set rschip = dbChipData.OpenRecordset("SELECT NumOfGOID FROM [Chip] WHERE" _
                        & " [GOIDCount] = '" & GOID & "'")
            'rsChip now has the number of times the GO term is on the chip
         
            percentage = (GOArray(i) / rschip![NumofGoid]) * 100
            present = (rschip![NumofGoid] / rstemp2![numberofdups]) * 100
            'percentage = # of times changed/ # of times measured on chip * 100
            Set rsType = dbMAPPfinder.OpenRecordset("SELECT Type, Name FROM GeneOntology WHERE" _
                     & " ID = '" & GOID & "'")
            dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, GoName, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('" & rsType![Type] & "', '" _
                        & GOID & "', '" & rsType![name] & "', " & GOArray(i) & ", " _
                        & rschip![NumofGoid] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
            End If
         Next i
         GOID = "GO"
         Set rsinGONested = dbMAPPfinder.OpenRecordset("SeLECT GOIDCount FROM " & species & "HierarchyCount Where GOID = '" & GOID & "'")
         arraysize = rsinGONested![goidcount] * (4)
         ReDim genecounter(arraysize) As String 'you now have an array twice as big as the number of genes in this sub-graph. Allows for plenty of duplicates.
         indata = TreeForm.nestedResults(dbExpressionData, dbChipData, TreeForm.root, genecounter, 0)  '"GO" is the root of the tree
         calculateTestStat
         
         Set rstemp = dbExpressionData.OpenRecordset("SELECT Gotype, GOID, GOName, InData, OnChip, InGO," _
                  & " Percentage, Present, InDataNested, OnChipNested, InGoNested, PercentageNested," _
                  & " PresentNested, TestStat FROM NestedResults " _
                  & "ORDER BY NestedResults![TestStat] DESC, NestedResults![OnChipNested] DESC")
         Set Output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Gene Ontology Results.txt")
         Output.WriteLine ("MAPPFinder 1.0 Results for the Gene Ontology")
         Output.WriteLine ("File: " & filelocation)
         Output.WriteLine (TreeForm.GODate)
         Output.WriteLine ("Statistics:")
         Output.WriteLine (metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria.")
         frmCalculation.lblGOcriteria.Caption = metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria."
         Output.WriteLine (nosp & " Genes did not link to a MGI term.")
         Output.WriteLine (noGO + others & " Genes did not link to a GO term.")
         frmCalculation.lblGOnotfound.Caption = noGO + nosp + others & " Genes did not link to a GO term."
         Output.WriteLine (metFilter - noSwissProt - noGO - others & " genes were used to calculate the results shown below.")
         frmCalculation.lblGOUsed.Caption = metFilter - noSwissProt - noGO - others & " genes were used to calculate the results shown below."
         Output.WriteLine ("The z score is based on an N of " & bigN & " and a R of " & bigR & " distinct genes.")
         frmCalculation.lblGOStat = "The z score is based on an N of " & bigN & " and a R of " & bigR & " distinct genes."
         Output.WriteLine ("")
         Output.WriteLine ("GOID" & Chr(9) & "GO Name" & Chr(9) & "GO Type" & Chr(9) & "Number Changed" & Chr(9) _
                        & "Number Measured" & Chr(9) & "Number in GO" & Chr(9) & "Percent Changed" _
                        & Chr(9) & "Percent Present" & Chr(9) & "Number Changed in Hierarchy" _
                        & Chr(9) & "Number Measured in Hierarchy" & Chr(9) & "Number in GO" _
                        & " Hierarchy" & Chr(9) & "Percent Changed in Hierarchy" & Chr(9) _
                        & "Percent Present in Hierarchy" & Chr(9) & "z Score")
         While Not rstemp.EOF
            Output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![GOName] & Chr(9) & rstemp![gotype] _
                        & Chr(9) & rstemp![indata] _
                        & Chr(9) & rstemp![onChip] & Chr(9) & rstemp![ingo] & Chr(9) _
                        & rstemp![percentage] & Chr(9) & rstemp![present] & Chr(9) _
                        & rstemp![indatanested] & Chr(9) & rstemp![onchipnested] & Chr(9) _
                        & rstemp![ingonested] & Chr(9) & rstemp![percentagenested] & Chr(9) _
                        & rstemp![presentnested] & Chr(9) & rstemp![teststat])
            rstemp.MoveNext
         Wend
         Output.Close

'=YEAST=========================================================================================
'===============================================================================================
      Case "yeast"
         While rsFilter.EOF = False
            progress = progress + 1
            If progress Mod 10 = 0 Then
               lblProgress.Caption = progress & " genes out of " & metFilter & " meeting the criterion are linked to GO."
               frmCriteria.Refresh
            End If
            If UCase(rsFilter![primaryType]) = "D" Then 'no convertion necessary. GO via SGD-GO table
               For Each element In MGIsAdded
                  If element = rsFilter![primary] Then
                     found = True
                     GoTo foundsGD
                  End If
               Next element
foundsGD:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rsFilter![primary] 'for some reason you need to cast the record to a string to get it into a collection
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                     & " [SGD-GeneOntology] where primary = '" & rsFilter![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  If rstemp2.EOF Then
                      noGO = noGO + 1
                  Else
                     While rstemp2.EOF = False
                        GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                        dbExpressionData.Execute ("INSERT into GO (primary, GOID) VALUES ('" & rsFilter![primary] _
                                          & "', '" & rstemp2![related] & "')")
                        rstemp2.MoveNext
                     Wend
                   End If
               End If
                        
            ElseIf UCase(rsFilter![primaryType]) = "S" Or UCase(rsFilter![primaryType]) = "N" Then
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT related from [SwissProt-SGD] Where" _
                     & " primary = '" & rsFilter![primary] & "'")
               If rstemp.EOF Then 'try trembl
                  Set rstrembl = dbMAPPfinder.OpenRecordset("SELECT ID FROM SwissProt Where" _
                              & " Accession like '*|" & rsFilter![primary] & "|*'")
                  If rstrembl.EOF = False Then
                     Set rstemp = dbMAPPfinder.OpenRecordset("SELECT related from [SwissProt-SGD] Where" _
                     & " primary = '" & rstrembl![ID] & "'")
                  End If
               End If
         
               If rstemp.EOF Then
                  nosp = nosp + 1
               Else
               For Each geneid In MGIsAdded
                  If geneid = rstemp![related] Then
                     found = True
                     GoTo foundSGDSP
                  End If
               Next geneid
foundSGDSP:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rstemp![related]
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rsFilter![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  
                  If rstemp2.EOF Then
                     Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rstrembl![ID] & "'")
                     If rstemp2.EOF Then 'both swissprot and trembl failed
                        noGO = noGO + 1
                     End If
                  End If
                  While rstemp2.EOF = False
                     GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                     dbExpressionData.Execute ("INSERT into GO (primary, GOID) VALUES ('" & rsFilter![primary] _
                                             & "', '" & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               End If
               End If
            Else
               other = others + 1
            End If
            rsFilter.MoveNext
            found = False
         Wend 'now you've attached GOIDs to as many unique genes as possible. Unique is defined as
            'a individual MGI ID. Next, report the count for each GOID
         
          
         For i = 0 To GOCount
            If GOArray(i) <> 0 Then 'there was a hit for this GOID calculate the percentages
            GOID = MakeGOID(i)
            'for each GOID in GOArray - divide the GOIDCount in the array by the the chip count
             'to get the percentage changed in the dataset
            'insert into results GOType, GOID, percentage
            Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT NumberOfDups FROM [" & gotable & "Count]" _
                     & " WHERE [RelatedField] = '" & GOID & "'")
            'temp 2 now has the number of times the GO term is represented in that species' table
         
            Set rschip = dbChipData.OpenRecordset("SELECT NumOfGOID FROM [Chip] WHERE" _
                        & " [GOIDCount] = '" & GOID & "'")
            'rsChip now has the number of times the GO term is on the chip
         
            percentage = (GOArray(i) / rschip![NumofGoid]) * 100
            present = (rschip![NumofGoid] / rstemp2![numberofdups]) * 100
            'percentage = # of times changed/ # of times measured on chip * 100
            Set rsType = dbMAPPfinder.OpenRecordset("SELECT Type, Name FROM GeneOntology WHERE" _
                     & " ID = '" & GOID & "'")
            dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, GoName, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('" & rsType![Type] & "', '" _
                        & GOID & "', '" & rsType![name] & "', " & GOArray(i) & ", " _
                        & rschip![NumofGoid] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
            End If
         Next i
         GOID = "GO"
         Set rsinGONested = dbMAPPfinder.OpenRecordset("SeLECT GOIDCount FROM " & species & "HierarchyCount Where GOID = '" & GOID & "'")
         arraysize = rsinGONested![goidcount] * (6)
         ReDim genecounter(arraysize) As String 'you now have an array twice as big as the number of genes in this sub-graph. Allows for plenty of duplicates.
         indata = TreeForm.nestedResults(dbExpressionData, dbChipData, TreeForm.root, genecounter, 0)  '"GO" is the root of the tree
         calculateTestStat
         Set rstemp = dbExpressionData.OpenRecordset("SELECT Gotype, GOID, GOName, InData, OnChip, InGO," _
                  & " Percentage, Present, InDataNested, OnChipNested, InGoNested, PercentageNested," _
                  & " PresentNested, TestStat FROM NestedResults " _
                  & "ORDER BY NestedResults![TestStat] DESC, NestedResults![OnChipNested] DESC")
         Set Output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Gene Ontology Results.txt")
         Output.WriteLine ("MAPPFinder 1.0 Results for the Gene Ontology")
         Output.WriteLine ("File: " & filelocation)
         Output.WriteLine (TreeForm.GODate)
         Output.WriteLine ("Statistics:")
         Output.WriteLine (metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria.")
         frmCalculation.lblGOcriteria.Caption = metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria."
         Output.WriteLine (nosp & " Genes did not link to a MGI term.")
         Output.WriteLine (noGO + others & " Genes did not link to a GO term.")
         frmCalculation.lblGOnotfound.Caption = noGO + nosp + others & " Genes did not link to a GO term."
         Output.WriteLine (metFilter - noSwissProt - noGO - others & " genes were used to calculate the results shown below.")
         frmCalculation.lblGOUsed.Caption = metFilter - noSwissProt - noGO - others & " genes were used to calculate the results shown below."
         Output.WriteLine ("The z score is based on an N of " & bigN & " and a R of " & bigR & " distinct genes.")
         frmCalculation.lblGOStat = "The z score is based on an N of " & bigN & " and a R of " & bigR & " distinct genes."
         Output.WriteLine ("")
         Output.WriteLine ("GOID" & Chr(9) & "GO Name" & Chr(9) & "GO Type" & Chr(9) & "Number Changed" & Chr(9) _
                        & "Number Measured" & Chr(9) & "Number in GO" & Chr(9) & "Percent Changed" _
                        & Chr(9) & "Percent Present" & Chr(9) & "Number Changed in Hierarchy" _
                        & Chr(9) & "Number Measured in Hierarchy" & Chr(9) & "Number in GO" _
                        & " Hierarchy" & Chr(9) & "Percent Changed in Hierarchy" & Chr(9) _
                        & "Percent Present in Hierarchy" & Chr(9) & "z Score")
         While Not rstemp.EOF
            Output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![GOName] & Chr(9) & rstemp![gotype] _
                        & Chr(9) & rstemp![indata] _
                        & Chr(9) & rstemp![onChip] & Chr(9) & rstemp![ingo] & Chr(9) _
                        & rstemp![percentage] & Chr(9) & rstemp![present] & Chr(9) _
                        & rstemp![indatanested] & Chr(9) & rstemp![onchipnested] & Chr(9) _
                        & rstemp![ingonested] & Chr(9) & rstemp![percentagenested] & Chr(9) _
                        & rstemp![presentnested] & Chr(9) & rstemp![teststat])
            rstemp.MoveNext
         Wend
         Output.Close
      End Select
      End If
'======================================================================================
'This section of the code is responsible for the MAPP Finder portion of the program. It will
'go through the rsFilter and count the number of times that each gene meeting the criteria is
'represented on a MAPP. Then the total number of genes for each MAPP is counted.
'Table 1 GenMAPP, MAPP it's on
'Table 2 MAPP, number of hits
'Table 3 MAPP, number of hits, number on Chip, number in mapp, percentage changed, percentage present

'This is much easier than the GO code because
'A) it's the same for each species because of the GenMAPP ID (this will unfortunately change come version2.0)
'B) there are no intermediate databases
'c) there is no nested data to calculate.

'======================================================================================


   If localMAPPsLoaded Then
      noGO = 0
      found = False
      ReDim mapparray(MAPPCount) As Integer
      lblProgress.Caption = "Calculating the Local Results"
      buildOnChipLocal
      calculateRandN
      progress = 0
      dbExpressionData.Execute ("DELETE * FROM GO")
      dbExpressionData.Execute ("DELETE * FROM Results")
      'rstemp will store all of the GenMAPP IDs (without duplicates) for the genes that meet the criteria.
      Set rstemp = dbExpressionData.OpenRecordset("SELECT DISTINCT GenMAPP FROM" _
                        & " Expression WHERE (" & sql(lstcriteria.ListIndex) & ")")
      rstemp.MoveLast
      rstemp.MoveFirst
      metFilter = rstemp.RecordCount
      If rstemp.EOF Then
         MsgBox "There are no genes meeting your criteria in the Expression Dataset.", vbOKOnly
         GoTo ENDSUB
      End If
   
      While rstemp.EOF = False
         progress = progress + 1
         If progress Mod 10 = 0 Then
               lblProgress.Caption = progress & " genes out of " & metFilter & " meeting the criterion are linked to local MAPPs."
               frmCriteria.Refresh
         End If
         genmappID = rstemp![GenMAPP]
         If InStr(1, genmappID, "~") <> 0 Then
            genmappID = Right(genmappID, Len(genmappID) - 1)
         End If
         Set rstemp2 = dbMAPPfinder.OpenRecordset("Select MAPPNameField, MAPPNumber from GeneToMAPP WHERE" _
                        & " GenMAPP = '" & genmappID & "'")
         
         If Not rstemp2.EOF Then
            For Each geneid In GenMAPPsAdded
               If geneid = rstemp![GenMAPP] Then
                  found = True
                  GoTo FoundGenMAPP
               End If
            Next geneid
FoundGenMAPP:
            If Not found Then
            GenMAPPsAdded.Add MGIID, MGIID
            While Not rstemp2.EOF
               MGIID = rstemp![GenMAPP]
               mapparray(rstemp2![MAPPNumber]) = mapparray(rstemp2![MAPPNumber]) + 1
               rstemp2.MoveNext
            Wend
            
            End If
         Else 'no genmapp available
            noGO = noGO + 1
         End If
         found = False
         rstemp.MoveNext
      Wend
                  
      For i = 0 To MAPPCount
         If mapparray(i) > 0 Then
         Set rstemp = dbMAPPfinder.OpenRecordset("Select MAPPNameField FROM GeneToMAPP WHERE MAPPNumber = " & i)
         Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT NumberOfDups FROM [GeneToMAPPCount]" _
                     & " WHERE [MAPPName] = '" & rstemp![MappNameField] & "'")
         'temp 2 now has the number of times the MAPP is represented in the GeneToMAPP table
         
         Set rschip = dbChipData.OpenRecordset("SELECT OnChip FROM [LocalMAPPsChip] WHERE" _
                        & " [MAPPName] = '" & rstemp![MappNameField] & "'")
         'rsChip now has the number of times the GO term is on the chip
         
         percentage = Round(((mapparray(i) / rschip![onChip]) * 100), 1)
         present = Round(((rschip![onChip] / rstemp2![numberofdups]) * 100), 1)
         'percentage = # of times change/ # of times measured on chip * 100
         
         'If rsType.EOF = False Then 'for some reason the GO db and ontology are out of sync.
      
         r = mapparray(i)
         n = rschip![onChip]
         'this calculates the standard test statistic under the hypergeometric distribution
         'the number changed - the number expected to changed based on background divided by the stdev of the data
         If localR - localN = 0 Then
            teststat = 0
         Else
            teststat = Round((r - (n * localR / localN)) / (Sqr(n * (localR / localN) * (1 - (localR / localN) * (1 - (n - 1) / (localN - 1))))), 3)
         End If
         
         dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, InData, OnChip," _
                        & " InGO, Percentage, Present, TestStat) VALUES ('L', '" _
                        & rstemp![MappNameField] & "', " & mapparray(i) & ", " _
                        & rschip![onChip] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ", " & teststat & ")")
         'End If
         
      End If
      Next i
      'TreeForm.DisplayLocalMAPPs dbExpressionData, TreeForm.GetLocalRoot()
      rstemp.Close
      Set rstemp = dbExpressionData.OpenRecordset("SELECT GOID, InData, OnChip, InGO," _
                  & " Percentage, Present, Teststat FROM Results WHERE (GOType = 'L') " _
                  & "ORDER BY Results![TestStat] DESC, Results![OnChip] DESC")
      Set Output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Local Results.txt")
      Output.WriteLine ("MAPPFinder 1.0 Results for Local MAPPs")
      Output.WriteLine ("File: " & filelocation)
      Output.WriteLine ("Statistics:")
      Output.WriteLine (metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria.")
      frmCalculation.lblLocalCriteria.Caption = metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria."
      Output.WriteLine (noGO & " Genes did not link to a MAPP.")
      frmCalculation.lblLocalNotFound.Caption = noGO & " Genes did not link to a MAPP."
      Output.WriteLine (metFilter - noGO & " genes were used to calculate the results shown below.")
      frmCalculation.LblLocalUsed.Caption = metFilter - noGO & " genes were used to calculate the results shown below."
      Output.WriteLine ("The z score is based on an N of " & localN & " and a R of " & localR & " distinct genes.")
      frmCalculation.lblLocalStat = "The z score is based on an N of " & localN & " and a R of " & localR & " distinct genes."
      Output.WriteLine ("")
      Output.WriteLine ("MAPPName" & Chr(9) & "Number Changed" & Chr(9) _
                        & "Number Measured" & Chr(9) & "Number on MAPP" & Chr(9) & "Percent Changed" _
                        & Chr(9) & "Percent Present" & Chr(9) & "z Score")
      While Not rstemp.EOF
         Output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![indata] _
                        & Chr(9) & rstemp![onChip] & Chr(9) & rstemp![ingo] & Chr(9) _
                        & rstemp![percentage] & Chr(9) & rstemp![present] & Chr(9) & rstemp![teststat])
         rstemp.MoveNext
      Wend
      Output.Close
   End If
   TreeForm.setChipDBLocation filelocation
   
   If geneontology Then
      frmLoadFiles.txtGO.Text = fixFileName(txtFile.Text) & "-Gene Ontology Results.txt"
   End If
   
   If localMAPPsLoaded Then
      frmLoadFiles.txtLocal.Text = fixFileName(txtFile.Text) & "-Local Results.txt"
   End If
   
   
   rstemp.Close
   'rstemp2.Close
'   rsType.Close
   'rschip.Close
   dbExpressionData.TableDefs.Delete ("GO")
   dbExpressionData.TableDefs.Delete ("Results")
   dbExpressionData.TableDefs.Delete ("NestedResults")
      
ENDSUB:
     
      dbExpressionData.Close
      dbChipData.Close
      dbMAPPfinder.Close
compact:
      On Error GoTo error
      DBEngine.CompactDatabase newfilename & ".gex", newfilename & ".$tm"
      Kill newfilename & ".gex"
      Name newfilename & ".$tm" As newfilename & ".gex"
      DBEngine.CompactDatabase newfilename & ".gdb", newfilename & ".$tm"
      Kill newfilename & ".gdb"
      Name newfilename & ".$tm" As newfilename & ".gdb"
      
      'TreeForm.setFileName (txtFile.Text)
      'TreeForm.CmdExpand_Click
      'TreeForm.Show
      'frmColors.Show
      'frmNumbers.Show
   frmLoadFiles.speciesselected = True
   frmLoadFiles.cmdLoadFiles_Click
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
      Case Else
         MsgBox "An error occurred while calculating the results. Please report error " & Err.Number _
            & " to GenMAPP@gladstone.ucsf.edu. Error message: " & Err.Description & "."
   End Select
nospeciesselected:
noerror:
   MousePointer = vbDefault
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
   dbExpressionData.Close
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
   frmLocalMAPPs.Show
End Sub

Private Sub lstColorSet_Click()
   Dim rsCriteria As DAO.Recordset
   Dim criteria As String, record As String
   Dim pipe As Integer, endline As Integer, newend As Integer, pipe2 As Integer
   Dim i As Integer
   lstcriteria.Clear
   colorsetclicked = True
   
   Set rsCriteria = dbExpressionData.OpenRecordset("SELECT Criteria FROM [ColorSet] WHERE" _
                  & " ColorSet = '" & lstColorSet.Text & "'")
   
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

Private Sub newresults_Click()
   Unload Me
   frmInput.Show
End Sub

Private Sub Option1_Click(Index As Integer)
   speciesselected = True
   frmLoadFiles.Option1_Click (Index)
   Select Case Index
      Case 0
         species = "human"
         gotable = "SwissProt-GeneOntology"
         TreeForm.setSpecies species, gotable
   
      Case 2
         species = "mouse"
         gotable = "MGI-GeneOntology"
         TreeForm.setSpecies species, gotable
         
      Case 4
         species = "yeast"
         gotable = "SGD-GeneOntology"
         TreeForm.setSpecies species, gotable
   End Select
End Sub

'There are two sets of tables being created. TempPrimary makes the related ID (SP,MGI,SGD) a primary ID, so that duplicate
'genes are not counted twice. TempALL and GOALL have all of the genes in the ED linked to GO so that the MAPP can be built with
'all of the genes in the dataset.


Public Sub buildOnChip()
   On Error GoTo error
    Dim rsExpression As DAO.Recordset
    Dim rstemp As DAO.Recordset, rstemp2 As DAO.Recordset, rsGO As DAO.Recordset
    Dim rstrembl As DAO.Recordset
    Dim genes As Integer, i As Long
    Dim keystring As String
    Dim MGIsAdded As New Collection
    Dim GOArray(GOCount) As Integer
    Dim SwissProtID As String
    Dim found As Boolean, trembl As Boolean
    Dim genecounter() As String
    Dim progress As Integer, rsinGONested As DAO.Recordset
   Dim onChip As Long, arraysize As Long
    found = False
    'GoTo nested
    Set rstemp = dbMAPPfinder.OpenRecordset("Select Version from Info")
    Set tblChip = dbChipData.CreateTableDef("Chip")
      With tblChip
         .Fields.Append .CreateField("GOIDCount", dbText, 15)
         Dim idxGO As Index
         .Fields.Append .CreateField("NumOfGOID", dbSingle)
         Set idxGO = .CreateIndex("idxChip")
         idxGO.Fields.Append .CreateField("GOIDCount", dbText, 15)
         idxGO.primary = False
         .Indexes.Append idxGO
      End With

   dbChipData.TableDefs.Append tblChip
   
   dbChipData.Execute ("DELETE * FROM Info")
   dbChipData.Execute ("Insert INTO Info(Version, GoTable) Values('" & rstemp![Version] _
               & "', '" & gotable & "')")
   
   
   If geneontology Then
   Set rsExpression = dbExpressionData.OpenRecordset("SELECT Primary, PrimaryType, OrderNo FROM Expression")
   rsExpression.MoveLast
   rsExpression.MoveFirst
   genes = rsExpression.RecordCount
   lblProgress.Caption = progress & "out of the " & genes & " measured linked to GO terms."
   frmCriteria.Refresh
   Select Case species
      Case "human"
         'all references in variable names to MGI actaully mean SwissProt. This is a carry over from the mosue code
          
          While rsExpression.EOF = False
         'attach GO's to as many genes as possible (via GenBank-GO or SwissProt-GO)
            progress = progress + 1
            If progress Mod 10 = 0 Then
               lblProgress.Caption = progress & " out of the " & genes & " genes measured are linked to GO terms."
               frmCriteria.Refresh
            End If
            If UCase(rsExpression![primaryType]) = "G" Then 'convert to SwissProt via SP-GenBank table
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT primary from [SwissProt-GenBank] Where" _
                     & " related = '" & rsExpression![primary] & "'")
               If rstemp.EOF Then
                  nosp = nosp + 1
               Else
               For Each element In MGIsAdded
                 If element = rstemp![primary] Then
                     found = True
                     GoTo foundMGIChipHu
                  End If
               Next element
foundMGIChipHu:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rstemp![primary] 'for some reason you need to cast the record to a string to get it into a collection
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [GenBank-GeneOntology] where primary = '" & rsExpression![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  If rstemp2.EOF Then
                     noGO = noGO + 1
                  Else
                     While rstemp2.EOF = False
                        GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                        dbChipData.Execute ("INSERT INTO GO (Primary, Related, GOID) VALUES " _
                           & "('" & rsExpression![primary] & "', '" & rstemp![primary] & "', '" _
                           & rstemp2![related] & "')")
                        dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rsExpression![primary] & "', '" & rsExpression![primaryType] _
                        & "', '" & rstemp![primary] & "', '" _
                        & rstemp2![related] & "')")
                        
                        rstemp2.MoveNext
                     Wend
                  End If
               Else 'it was found, so only add it to the GOALL table
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [GenBank-GeneOntology] where primary = '" & rsExpression![primary] & "'")
                  If rstemp2.EOF Then
                     noGO = noGO + 1
                  Else
                     While rstemp2.EOF = False
                        dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rsExpression![primary] & "', '" & rsExpression![primaryType] _
                        & "', '" & rstemp![primary] & "', '" _
                        & rstemp2![related] & "')")
                        rstemp2.MoveNext
                     Wend
                  End If
               End If
               End If
               rsExpression.MoveNext
               found = False
            ElseIf UCase(rsExpression![primaryType]) = "S" Or UCase(rsExpression![primaryType]) = "N" Then
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT ID from [SwissProt] Where" _
                     & " ID = '" & rsExpression![primary] & "'")
               If rstemp.EOF Then 'try trembl
                  Set rstemp = dbMAPPfinder.OpenRecordset("SELECT ID FROM SwissProt Where" _
                              & " Accession like '*|" & rsExpression![primary] & "|*'")
                  
               End If
              
               If rstemp.EOF Then 'both swissprot and trembl failed
                  nosp = nosp + 1
               Else
               For Each geneid In MGIsAdded
                  If geneid = rstemp![ID] Then
                     found = True
                     GoTo foundMGISPChipHu
                  End If
               Next geneid
foundMGISPChipHu:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rstemp![ID]
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntology] where primary = '" & rsExpression![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  
                  If rstemp2.EOF Then
                     Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntology] where primary = '" & rstemp![ID] & "'")
                     If rstemp2.EOF Then 'both swissprot and trembl failed
                        noGO = noGO + 1
                     End If
                  End If
                  While rstemp2.EOF = False
                     GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                     dbChipData.Execute ("INSERT INTO GO (Primary, Related, GOID) VALUES " _
                           & "('" & rsExpression![primary] & "', '" & rstemp![ID] & "', '" _
                           & rstemp2![related] & "')")
                     dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                           & "('" & rsExpression![primary] & "', 'S', '" & rstemp![ID] & "', '" _
                           & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
                  
               Else 'found
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntology] where primary = '" & rsExpression![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GOALL
                  
                  If rstemp2.EOF Then
                     Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntology] where primary = '" & rstemp![ID] & "'")
                     If rstemp2.EOF Then 'both swissprot and trembl failed
                        noGO = noGO + 1
                     End If
                  End If
                  While rstemp2.EOF = False
                        dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rsExpression![primary] & "', 'S', '" & rstemp![ID] & "', '" _
                        & rstemp2![related] & "')")
                        rstemp2.MoveNext
                  Wend
                  
               End If
               End If
               rsExpression.MoveNext
               found = False
            Else
               other = others + 1
               rsExpression.MoveNext
            End If
         Wend 'now you've attached GOIDs to as many unique genes as possible. Unique is defined as
            'a individual MGI ID. Next, report the count for each GOID
         
          
         For i = 0 To GOCount
            If GOArray(i) <> 0 Then 'there was a hit for this GOID calculate the GO counts
            GOID = MakeGOID(i)
            'you now have a table with every one of the MGIs of the expression dataset linked to GO.
            'need to count how many times each GO term is represented.
            dbChipData.Execute ("INSERT INTO Chip (GOIDCount, NumOfGOID) VALUES ('" & GOID & "', '" _
                  & GOArray(i) & "')")
            End If
         Next i
      rstemp2.Close
      rstemp.Close
   
         
      Case "mouse"
         
         While rsExpression.EOF = False
         'attach GO's to as many genes as possible (via MGI-Genbank, MGI-SwissProt, MGI-Unigene)
         'the next step is to attch GO IDs to the MGIs.
            progress = progress + 1
            If progress Mod 10 = 0 Then
               lblProgress.Caption = progress & " out of the " & genes & " genes measured are linked to GO terms."
               frmCriteria.Refresh
            End If
            If UCase(rsExpression![primaryType]) = "G" Then 'convert to MGI via MGI-GenBank table
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT primary from [MGI-GenBank] Where" _
                     & " related = '" & rsExpression![primary] & "'")
               If rstemp.EOF Then
                  nosp = nosp + 1
               Else
               For Each element In MGIsAdded
                 If element = rstemp![primary] Then
                     found = True
                     GoTo foundMGIChip
                  End If
               Next element
foundMGIChip:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rstemp![primary] 'for some reason you need to cast the record to a string to get it into a collection
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [GenBank-GeneOntology] where primary = '" & rsExpression![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  If rstemp2.EOF Then
                     noGO = noGO + 1
                  Else
                     While rstemp2.EOF = False
                        GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                        dbChipData.Execute ("INSERT INTO GO (Primary, Related, GOID) VALUES " _
                           & "('" & rsExpression![primary] & "', '" & rstemp![primary] & "', '" _
                           & rstemp2![related] & "')")
                        dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rsExpression![primary] & "', '" & rsExpression![primaryType] _
                        & "', '" & rstemp![primary] & "', '" _
                        & rstemp2![related] & "')")
                        
                        rstemp2.MoveNext
                     Wend
                  End If
               Else 'it was found, so only add it to the GOALL table
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [GenBank-GeneOntology] where primary = '" & rsExpression![primary] & "'")
                  If rstemp2.EOF Then
                     noGO = noGO + 1
                  Else
                     While rstemp2.EOF = False
                        dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rsExpression![primary] & "', '" & rsExpression![primaryType] _
                        & "', '" & rstemp![primary] & "', '" _
                        & rstemp2![related] & "')")
                        rstemp2.MoveNext
                     Wend
                  End If
               End If
               End If
               rsExpression.MoveNext
               found = False
            ElseIf UCase(rsExpression![primaryType]) = "S" Or UCase(rsExpression![primaryType]) = "N" Then
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT related from [SwissProt-MGI] Where" _
                     & " primary = '" & rsExpression![primary] & "'")
               If rstemp.EOF Then 'try trembl
                  Set rstrembl = dbMAPPfinder.OpenRecordset("SELECT ID FROM SwissProt Where" _
                              & " Accession like '*|" & rsExpression![primary] & "|*'")
                  
                  If rstrembl.EOF = False Then
                     trembl = True
                     Set rstemp = dbMAPPfinder.OpenRecordset("SELECT related from [SwissProt-MGI] Where" _
                     & " primary = '" & rstrembl![ID] & "'")
                  End If
               End If
              
               If rstemp.EOF Then 'both swissprot and trembl failed
                  nosp = nosp + 1
               Else
               For Each geneid In MGIsAdded
                  If geneid = rstemp![related] Then
                     found = True
                     GoTo foundMGISPChip
                  End If
               Next geneid
foundMGISPChip:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rstemp![related]
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rsExpression![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  
                  If rstemp2.EOF Then
                     If trembl Then 'the primary failed, maybe a name for the trembl exists?
                        Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                           & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rstrembl![ID] & "'")
                     End If
                     If rstemp2.EOF Then 'both swissprot and trembl failed
                        noGO = noGO + 1
                     End If
                  End If
                  While rstemp2.EOF = False
                     GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                     dbChipData.Execute ("INSERT INTO GO (Primary, Related, GOID) VALUES " _
                           & "('" & rsExpression![primary] & "', '" & rstemp![related] & "', '" _
                           & rstemp2![related] & "')")
                     dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                           & "('" & rsExpression![primary] & "', 'S', '" & rstemp![related] & "', '" _
                           & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
                  
               Else
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rsExpression![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GOALL
                  
                  If rstemp2.EOF Then
                     If trembl Then
                        Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rstrembl![ID] & "'")
                     End If
                     If rstemp2.EOF Then 'both swissprot and trembl failed
                        noGO = noGO + 1
                     End If
                  End If
                  While rstemp2.EOF = False
                        dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rsExpression![primary] & "', 'S', '" & rstemp![related] & "', '" _
                        & rstemp2![related] & "')")
                        rstemp2.MoveNext
                  Wend
                  
               End If
               End If
               trembl = False
               rsExpression.MoveNext
               found = False
            Else
               other = others + 1
               rsExpression.MoveNext
            End If
         Wend 'now you've attached GOIDs to as many unique genes as possible. Unique is defined as
            'a individual MGI ID. Next, report the count for each GOID
         
          
         For i = 0 To GOCount
            If GOArray(i) <> 0 Then 'there was a hit for this GOID calculate the GO counts
            GOID = MakeGOID(i)
            'you now have a table with every one of the MGIs of the expression dataset linked to GO.
            'need to count how many times each GO term is represented.
            dbChipData.Execute ("INSERT INTO Chip (GOIDCount, NumOfGOID) VALUES ('" & GOID & "', '" _
                  & GOArray(i) & "')")
            End If
         Next i
      rstemp2.Close
      rstemp.Close
   
      Case "yeast"
         'all references in variable names to MGI actaully mean SGD. This is a carry over from the mosue code
          
          While rsExpression.EOF = False
         'attach GO's to as many genes as possible (via SGD-GO or SwissProt-GO)
            progress = progress + 1
            If progress Mod 10 = 0 Then
               lblProgress.Caption = progress & " out of the " & genes & " genes measured are linked to GO terms."
               frmCriteria.Refresh
            End If
            If UCase(rsExpression![primaryType]) = "D" Then 'no convertion, use SGD-GeneOntology table
      
               For Each element In MGIsAdded
                  If element = rsExpression![primary] Then
                     found = True
                     GoTo foundSGDChip
                  End If
               Next element
foundSGDChip:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rsExpression![primary] 'for some reason you need to cast the record to a string to get it into a collection
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SGD-GeneOntology] where primary = '" & rsExpression![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  If rstemp2.EOF Then
                     noGO = noGO + 1
                  Else
                     While rstemp2.EOF = False
                        GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                        dbChipData.Execute ("INSERT INTO GO (Primary, Related, GOID) VALUES " _
                           & "('" & rsExpression![primary] & "', '" & rsExpression![primary] & "', '" _
                           & rstemp2![related] & "')")
                        dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rsExpression![primary] & "', '" & rsExpression![primaryType] _
                        & "', '" & rsExpression![primary] & "', '" _
                        & rstemp2![related] & "')")
                        
                        rstemp2.MoveNext
                     Wend
                  End If
               Else 'it was found, so only add it to the GOALL table
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SGD-GeneOntology] where primary = '" & rsExpression![primary] & "'")
                  If rstemp2.EOF Then
                     noGO = noGO + 1
                  Else
                     While rstemp2.EOF = False
                        dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rsExpression![primary] & "', '" & rsExpression![primaryType] _
                        & "', '" & rsExpression![primary] & "', '" _
                        & rstemp2![related] & "')")
                        rstemp2.MoveNext
                     Wend
                  End If
               End If
               rsExpression.MoveNext
               found = False
            ElseIf UCase(rsExpression![primaryType]) = "S" Or UCase(rsExpression![primaryType]) = "N" Then
               Set rstemp = dbMAPPfinder.OpenRecordset("SELECT ID from [SwissProt] Where" _
                     & " ID = '" & rsExpression![primary] & "'")
               If rstemp.EOF Then 'try trembl
                  Set rstemp = dbMAPPfinder.OpenRecordset("SELECT ID FROM SwissProt Where" _
                              & " Accession like '*|" & rsExpression![primary] & "|*'")
                  
               End If
              
               If rstemp.EOF Then 'both swissprot and trembl failed
                  nosp = nosp + 1
               Else
               For Each geneid In MGIsAdded
                  If geneid = rstemp![ID] Then
                     found = True
                     GoTo foundSGDSPChip
                  End If
               Next geneid
foundSGDSPChip:
               
               If Not found Then 'this MGI hasn't been previously hit
                  MGIID = rstemp![ID]
                  MGIsAdded.Add MGIID, MGIID
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rsExpression![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GO node's bucket
                  
                  If rstemp2.EOF Then
                     Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rstemp![ID] & "'")
                     If rstemp2.EOF Then 'both swissprot and trembl failed
                        noGO = noGO + 1
                     End If
                  End If
                  While rstemp2.EOF = False
                     GOArray(Val(rstemp2![related])) = GOArray(Val(rstemp2![related])) + 1
                        'each element of the array is one GOID. Add one to indicate that there was
                        'a hit for that GOID
                     dbChipData.Execute ("INSERT INTO GO (Primary, Related, GOID) VALUES " _
                           & "('" & rsExpression![primary] & "', '" & rstemp![ID] & "', '" _
                           & rstemp2![related] & "')")
                     dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                           & "('" & rsExpression![primary] & "', 'S', '" & rstemp![ID] & "', '" _
                           & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
                  
               Else 'found
                  Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFinder] where primary = '" & rsExpression![primary] & "'")
                  'now you've got all the GO's this GB is in, so add one to the specific that GOALL
                  
                  If rstemp2.EOF Then
                     Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT DISTINCT related from" _
                        & " [SwissProt-GeneOntologyMAPPFINDER] where primary = '" & rstemp![ID] & "'")
                     If rstemp2.EOF Then 'both swissprot and trembl failed
                        noGO = noGO + 1
                     End If
                  End If
                  While rstemp2.EOF = False
                        dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rsExpression![primary] & "', 'S', '" & rstemp![ID] & "', '" _
                        & rstemp2![related] & "')")
                        rstemp2.MoveNext
                  Wend
                  
               End If
               End If
               rsExpression.MoveNext
               found = False
            Else
               other = others + 1
               rsExpression.MoveNext
            End If
         Wend 'now you've attached GOIDs to as many unique genes as possible. Unique is defined as
            'a individual MGI ID. Next, report the count for each GOID
         
          
         For i = 0 To GOCount
            If GOArray(i) <> 0 Then 'there was a hit for this GOID calculate the GO counts
            GOID = MakeGOID(i)
            'you now have a table with every one of the MGIs of the expression dataset linked to GO.
            'need to count how many times each GO term is represented.
            dbChipData.Execute ("INSERT INTO Chip (GOIDCount, NumOfGOID) VALUES ('" & GOID & "', '" _
                  & GOArray(i) & "')")
            End If
         Next i
      rstemp2.Close
      rstemp.Close
   
         
   End Select
nested:
   onChip = 0
   
  GOID = "GO"
  Set rsinGONested = dbMAPPfinder.OpenRecordset("SeLECT GOIDCount FROM " & species & "HierarchyCount Where GOID = '" & GOID & "'")
   arraysize = rsinGONested![goidcount] * (4)
   ReDim genecounter(arraysize) As String 'you now have an array twice as big as the number of genes in this sub-graph. Allows for plenty of duplicates.
   onChip = TreeForm.NestedChipData(dbChipData, TreeForm.root, genecounter, 0)  '"GO" is the root of the tree
   dbChipData.Execute ("INSERT INTO NestedChip (GOID, OnchipNested) VALUES ('GO', " & onChip & ")")
   
   End If
   chipbuiltOK = True
   'keep goall for mapp building.
error:
   Select Case Err.Number
      Case 91 'rstemp2 not set
         MsgBox "No genes in your dataset can be linked to gene ontology. You have" _
            & " probably selected the wrong species. The current species is " & species & ".", vbOKOnly
         
   End Select
End Sub
Public Sub buildOnChipLocal()
    Dim rsExpression As DAO.Recordset
    Dim rstemp As DAO.Recordset, rstemp2 As DAO.Recordset, rsGO As DAO.Recordset
    Dim genes As Integer
    Dim mapparray(MAPPCount) As Integer
    Dim MappName(MAPPCount) As String
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
                  MappName(rstemp2![MAPPNumber]) = rstemp2![MappNameField]
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
                     & MappName(i) & "', " & mapparray(i) & ")")
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
   frmHelp.Show
End Sub

Public Sub eraseForm()
   specieselected = False
   geneontology = False
   localMAPPsLoaded = False
   Option1(0).Value = False
   Option1(2).Value = False
   Option1(4).Value = False
   Check1.Value = 0
   Check2.Value = 0
   txtFile.Text = ""
   lstColorSet.Clear
   lblProgress.Caption = ""
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

Public Sub InitializeGOArray(GOArray() As Integer)
   Dim i As Long
   
   For i = 0 To GOCount
      GOArray(i) = 0
   Next i
End Sub

Public Sub calculateRandN()
   'N = the total number of genes in this dataset that are on a MAPP in this set of local MAPPs
   'R = the total number of genes meeting the user's criterion that are on a MAPP in this set of local MAPPs
   'these numbers are necessary for calculating the test stat.
   
   Dim rsFilter As DAO.Recordset
   Dim rsAll As DAO.Recordset
   Dim rsgene As DAO.Recordset
   
   localR = 0
   localN = 0
   
   Set rsAll = dbExpressionData.OpenRecordset("Select DISTINCT GenMAPP from Expression")
   
   While rsAll.EOF = False
      Set rsgene = dbMAPPfinder.OpenRecordset("SELECT GenMAPP from GenesOnMAPP WHERE" _
                     & " GenMAPP = '" & rsAll![GenMAPP] & "'")
      If rsgene.EOF = False Then 'this gene is in the dataset
         localN = localN + 1
      End If
      rsAll.MoveNext
   Wend
   
   Set rsFilter = dbExpressionData.OpenRecordset("SELECT DISTINCT GenMAPP FROM" _
                        & " Expression WHERE (" & sql(lstcriteria.ListIndex) & ")")
   rsFilter.MoveLast
   rsFilter.MoveFirst
   While rsFilter.EOF = False
      Set rsgene = dbMAPPfinder.OpenRecordset("Select GenMAPP from GenesOnMAPP WHERE" _
                     & " GenMAPP = '" & rsFilter![GenMAPP] & "'")
      If rsgene.EOF = False Then 'this gene is in the dataset
         localR = localR + 1
      End If
      rsFilter.MoveNext
   Wend
   
End Sub

Public Sub calculateTestStat()
   Dim rsResults As DAO.Recordset
   Dim rsGOterm As DAO.Recordset
   Dim r As Long, n As Long
   Dim teststat As Double
   
   Set rsGOterm = dbExpressionData.OpenRecordset("SELECT indatanested, onchipnested FROM NestedResults Where GOID = 'GO'")
   bigR = rsGOterm![indatanested]
   bigN = rsGOterm![onchipnested]
   
   Set rsResults = dbExpressionData.OpenRecordset("Select GOID, indatanested, onchipnested FROM NestedResults")
   While rsResults.EOF = False 'step through each results and calculate TestStat
      If rsResults![GOID] = "GO" Then
         teststat = 0
         dbExpressionData.Execute ("Update NestedResults Set TestStat = 0 WHere GOID = 'GO'")
      Else
         r = rsResults![indatanested]
         n = rsResults![onchipnested]
         
         'this calculate the standard test statistic under the hypergeometric distribution
         'the number changed - the number expected to changed based on background divided by the stdev of the data
         If bigR - bigN = 0 Then
            teststat = 0
         Else
            teststat = Round((r - (n * bigR / bigN)) / (Sqr(n * (bigR / bigN) * (1 - (bigR / bigN)) * (1 - (n - 1) / (bigN - 1)))), 3)
         End If
         dbExpressionData.Execute ("Update NestedResults Set TestStat = " & teststat & " Where GOID = '" _
            & rsResults![GOID] & "'")
      End If
      rsResults.MoveNext
   Wend
      
End Sub

Public Sub junk()
 
   Set tblGO = dbExpressionData.CreateTableDef("GO")
      With tblGO
         .Fields.Append .CreateField("Primary", dbText, 15)
         .Fields.Append .CreateField("Related", dbText, 15)
         Dim idxGO2 As Index
         Dim idxRelated As Index
         Set idxGO2 = .CreateIndex("idxGO2")
         idxGO2.Fields.Append .CreateField("Primary", dbText, 15)
         idxGO2.Fields.Append .CreateField("GOID", dbText, 255)
         .Indexes.Append idxGO2
         Set idxRelated = .CreateIndex("idxRelated")
         idxRelated.Fields.Append .CreateField("Related", dbText, 15)
         .Indexes.Append idxRelated
         .Fields.Append .CreateField("GOID", dbText, 255)
      End With
   Set tblResults = dbExpressionData.CreateTableDef("Results")
      With tblResults
         .Fields.Append .CreateField("GOType", dbText, 2)
         Dim idxResults As Index
         .Fields.Append .CreateField("GOID", dbText, 255)
         .Fields.Append .CreateField("GOName", dbMemo)
         Set idxResults = .CreateIndex("Results")
         idxResults.Fields.Append .CreateField("GOType", dbText, 2)
         idxResults.Fields.Append .CreateField("TestStat", dbSingle)
         idxResults.Fields.Append .CreateField("InData", dbSingle)
         .Indexes.Append idxResults
         .Fields.Append .CreateField("InData", dbSingle)
         .Fields.Append .CreateField("OnChip", dbSingle)
         .Fields.Append .CreateField("InGO", dbSingle)
         .Fields.Append .CreateField("Percentage", dbSingle)
         .Fields.Append .CreateField("Present", dbSingle)
         .Fields.Append .CreateField("TestStat", dbDouble)
      End With
   Set tblnestedResults = dbExpressionData.CreateTableDef("NestedResults")
      With tblnestedResults
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
         idxPercentage.Fields.Append .CreateField("PercentageNested", dbSingle)
         idxPercentage.Fields.Append .CreateField("TestStat", dbSingle)
         .Indexes.Append idxNestedResults
         .Indexes.Append idxPercentage
         .Fields.Append .CreateField("InData", dbSingle)
         .Fields.Append .CreateField("OnChip", dbSingle)
         .Fields.Append .CreateField("InGO", dbSingle)
         .Fields.Append .CreateField("Percentage", dbSingle)
         .Fields.Append .CreateField("Present", dbSingle)
         .Fields.Append .CreateField("InDataNested", dbSingle)
         .Fields.Append .CreateField("OnChipNested", dbSingle)
         .Fields.Append .CreateField("InGONested", dbSingle)
         .Fields.Append .CreateField("PercentageNested", dbSingle)
         .Fields.Append .CreateField("PresentNested", dbSingle)
         .Fields.Append .CreateField("TestStat", dbDouble)
      End With
   dbExpressionData.TableDefs.Append tblGO
   dbExpressionData.TableDefs.Append tblResults
   dbExpressionData.TableDefs.Append tblnestedResults

 'now need to check if the chip data has already been calculated. Only do this once
   If Dir(newfilename & ".gdb") = "" Then
      Set dbChipData = CreateDatabase(newfilename & ".gdb", dbLangGeneral)
      Set tblInfo = dbChipData.CreateTableDef("Info")
         With tblInfo
            .Fields.Append .CreateField("Version", dbText, 15)
            .Fields.Append .CreateField("GoTable", dbText, 25)
         End With
      dbChipData.TableDefs.Append tblInfo
   Set tblGO = dbChipData.CreateTableDef("GO")
      With tblGO
         .Fields.Append .CreateField("Primary", dbText, 15)
         Dim idxGOChip As Index
         .Fields.Append .CreateField("Related", dbText, 15)
         .Fields.Append .CreateField("PrimaryType", dbText, 2)
         Set idxGOChip = .CreateIndex("idxGO")
         idxGOChip.Fields.Append .CreateField("Related", dbText, 15)
         idxGOChip.Fields.Append .CreateField("GOID", dbText, 255)
         .Indexes.Append idxGOChip
         .Fields.Append .CreateField("GOID", dbText, 255)
      End With
   Set tblGOAll = dbChipData.CreateTableDef("GOAll")
      With tblGOAll
         .Fields.Append .CreateField("Primary", dbText, 15)
         Dim idxGOChipALL As Index
         .Fields.Append .CreateField("Related", dbText, 15)
         .Fields.Append .CreateField("PrimaryType", dbText, 2)
         Set idxGOChipALL = .CreateIndex("idxGO")
         idxGOChipALL.Fields.Append .CreateField("Related", dbText, 15)
         idxGOChipALL.Fields.Append .CreateField("GOID", dbText, 15)
         .Indexes.Append idxGOChipALL
         .Fields.Append .CreateField("GOID", dbText, 15)
      End With
      dbChipData.TableDefs.Append tblGOAll
      Set tblNestedChip = dbChipData.CreateTableDef("NestedChip")
      With tblNestedChip
         .Fields.Append .CreateField("GOID", dbText, 15)
         Dim idxNestedChip As Index
         Dim idxGOID As Index
         .Fields.Append .CreateField("OnChipNested", dbLong)
         
         Set idxGOID = .CreateIndex("idxGOID")
         idxGOID.Fields.Append .CreateField("GOID", dbText, 15)
         idxGOID.primary = True
         .Indexes.Append idxGOID
      End With
      Set tblLocalMAPPsChip = dbChipData.CreateTableDef("LocalMAPPsChip")
      With tblLocalMAPPsChip
         .Fields.Append .CreateField("MAPPName", dbText, 255)
         Dim idxGOID2 As Index
         .Fields.Append .CreateField("OnChip", dbLong)
         
         Set idxGOID2 = .CreateIndex("idxMAPPName")
         idxGOID2.Fields.Append .CreateField("MAPPName", dbText, 255)
         idxGOID2.primary = True
         .Indexes.Append idxGOID2
      End With
      dbChipData.TableDefs.Append tblNestedChip
      dbChipData.TableDefs.Append tblGO
      dbChipData.TableDefs.Append tblLocalMAPPsChip
      If geneontology Then
         buildOnChip  'this sub rountine builds the Chip table. This is information for the denominator of the ChipPercentage
         If chipbuiltOK = False Then 'the chip file failed.
            dbChipData.Close
            FSO.DeleteFile newfilename & ".gdb"
            
            GoTo nospeciesselected
         End If
      End If
    Else
      Set dbChipData = OpenDatabase(newfilename & ".gdb")
      If geneontology Then
         Set rstemp = dbChipData.OpenRecordset("SELECT Version, GOTable FROM Info")
         Set rstemp2 = dbMAPPfinder.OpenRecordset("SELECT Version FROM Info")
         If rstemp.EOF Then 'nothing has been written to info, so the chipDB wasn't built
            buildOnChip
            If chipbuiltOK = False Then 'the chip file failed.
               dbChipData.Close
               FSO.DeleteFile newfilename & ".gdb"
               
               GoTo nospeciesselected
            End If
         Else
         If (rstemp![gotable] <> gotable) Then 'the Chip table was built from a different version of the MAPPFinder database. need to rebuild
         Select Case rstemp![gotable]
            Case "MGI-GeneOntology"
               MsgBox "The Chip file for this Expression Dataset was calculated with mouse data" _
                  & " and you have selected " & species & " as your species. Please make sure you have" _
                  & " the right species selected. If the Chip file was calculated using the wrong" _
                  & " species, you should delete " & newfilename & ".gdb and re-run MAPPFinder.", vbOKOnly
            Case "SwissProt-GeneOntology"
               MsgBox "The Chip file for this Expression Dataset was calculated with human data" _
                  & " and you have selected " & species & " as your species. Please make sure you have" _
                  & " the right species selected. If the Chip file was calculated using the wrong" _
                  & " species, you should delete " & newfilename & ".gdb and re-run MAPPFinder.", vbOKOnly
            Case "SGD-GeneOntology"
               MsgBox "The Chip file for this Expression Dataset was calculated with yeast data" _
                  & " and you have selected " & species & " as your species. Please make sure you have" _
                  & " the right species selected. If the Chip file was calculated using the wrong" _
                  & " species, you should delete " & newfilename & ".gdb and re-run MAPPFinder.", vbOKOnly
         End Select
         GoTo nospeciesselected
         End If
         If (rstemp![Version] <> rstemp2![Version]) Then 'the Chip table was built from a different version of the Pathfinder database. need to rebuild
            dbChipData.Execute "DROP TABLE Chip"
            dbChipData.Execute ("DELETE * FROM LocalMAPPsChip")
            dbChipData.Execute ("Delete * from NestedChip")
            buildOnChip
            If buildonchipok = False Then 'the chip file failed.
               GoTo nospeciesselected
            End If
         End If

      
         Set rstemp = dbChipData.OpenRecordset("SElect * FROM nestedchip")
         If rstemp.EOF = True Then 'the go chip info didn't get built before
            buildOnChip
         End If
         rstemp.Close
         rstemp2.Close
      End If
      End If
   End If
   'buildOnChip"

End Sub

Public Function mapToClusterSystem() As Boolean
   Dim rssystem As Recordset, rsrelations As DAO.Recordset
   Dim i As Integer, numofsystems As Integer, record As Integer
   Dim tblInfo As TableDef, tblMaptoCluster As TableDef
   Dim rsGenes As Recordset, found As Boolean
   
   mapToClusterSystem = True
   If Dir(newfilename & ".gdb") = "" Then
      Set dbChipData = CreateDatabase(newfilename & ".gdb", dbLangGeneral)
      Set tblInfo = dbChipData.CreateTableDef("Info")
         With tblInfo
            .Fields.Append .CreateField("GoTable", dbText, 25)
         End With
      dbChipData.TableDefs.Append tblInfo
      Set tblCluster = dbChipData.CreateTableDef("MaptoCluster")
      With tblCluster
         .Fields.Append .CreateField("Primary", dbText, 15)
         .Fields.Append .CreateField("Related", dbText, 15)
      End With
      dbChipData.TableDefs.Append tblCluster
      Set tblGO = dbChipData.CreateTableDef("ClustertoGO")
      tblGO.fields.
   
      Set rssystem = dbMAPPfinder.OpenRecordset("SELECT System, SystemCode from Systems" _
                                                & " WHERE MAPPFinder = '" & cmbSpecies.Text _
                                                & "'")
      If rssystem.EOF = False Then
         clusterSystem = rssystem!System
         clusterCode = rssystem!SystemCode
         dbChipData.Execute "INSERT INTO Info (GOTable) VALUES ('" & clusterSystem & "')"
         Set rsRelation = dbMAPPfinder.OpenRecordset( _
                        "SELECT Relation FROM Relations WHERE SystemCode = '" & clusterCode _
                        & "' AND RelatedCode = 'T'")
         'this can't be empty but check it anyway
         If rsRelation.EOF = False Then
            GOrelation = rsRelation!Relation
         Else
            MsgBox "You do not have a table from " & clusterSystem & " to GO, you need this.", vbOKOnly
            mapToClusterSystem = False
            Exit Function
         End If
      Else
         MsgBox "The database " & DatabaseLoc & " does not have the correct tables to run" _
               & " MAPPFinder for " & cmbSpecies.Text & ". Please check the species you selected" _
               & " and the database you are using. To change your database you must return to the" _
               & " start menu.", vbOKOnly
         mapToClusterSystem = False
         Exit Function
      End If
      
               
      Set rssystem = dbExpressionData.OpenRecordset("SELECT DISTINCT SystemCode from Expression")
      i = 0
      While rssystem.EOF = False
         If clusterCode = rssystem!SystemCode Then
            relations(i, 0) = clusterSystem
            relations(i, 1) = "S"
            relations(i, 2) = rssystem!SystemCode
         Else
            'look for clustercode-EDcode relation
            Set rsRelation = dbMAPPfinder.OpenRecordset _
                     ("SELECT Relation FROM Relations WHERE SystemCode = '" & clusterCode _
                     & "' AND RelatedCode = '" & rssystem!SystemCode & "'")
            If rsRelation.EOF = False Then 'found the relation
               relations(i, 0) = rsRelation!Relation
               relations(i, 1) = "R"
               relations(i, 2) = rssystem!SystemCode
            Else 'try the other way
               Set rsRelation = dbMAPPfinder.OpenRecordset _
                        ("SELECT Relation FROM Relations WHERE RelatedCode = '" & clusterCode _
                        & "' AND SystemCode = '" & rssystem!SystemCode & "'")
               If rsRelation.EOF = False Then 'found the relation
                  relations(i, 0) = rsRelation!Relation
                  relations(i, 1) = "P"
                  relations(i, 2) = rssystem!SystemCode
               Else 'no relation exists
                  relationnotfound = True
               End If
            End If
         End If
         If relationnotfound Then
            MsgBox "No relation exists between the system code " & rssystem!SystemCode _
               & " and " & clusterSystem & ". MAPPFinder can not use this system." _
               & " Check the system code, or add a relation to your database.", vbOKOnly
         End If
         rssystem.MoveNext
         i = i + 1
      Wend
      numofsystems = i
      i = 0
      'now I need to map each gene to the cluster system using the relations we just extracted
      
      Set rsGenes = dbExpressionData.OpenRecordset _
                     ("SELECT OrderNo, ID, SystemCode FROM Expression")
      rsGenes.MoveLast
      rsGenes.MoveFirst
      ReDim genes(rsGenes.RecordCount, 1)
      record = 0
      While rsGenes.EOF = False
         While i < numofsystems And found = False
            If relations(i, 2) = rsGenes!SystemCode Then
               found = True
            Else
               i = i + 1
            End If
         Wend
         If i < numofsystems Then 'you found a match, this system is supported
            genes(record, 0) = rsGenes!ID
            Select Case relations(i, 1)
               Case "S" 'the codes are the same
                  genes(record, 1) = rsGenes!ID
                  addToClusterGenes rsGenes!ID
               Case "P" 'the GEX code is the primary of the relationship
                  Set rsrelations = dbMAPPfinder.OpenRecordset _
                                    ("SELECT Related FROM " & relations(i, 0) & " WHERE " _
                                    & "Primary = '" & rsGenes!ID & "'")
                  If rsrelations.EOF = False Then
                     genes(record, 1) = rsrelations!related
                     addToClusterGenes rsGenes!ID
                  Else
                     NoCluster = NoCluster + 1
                  End If
               Case "R" 'the gex code is the related of the relationship
                  Set rsrelations = dbMAPPfinder.OpenRecordset _
                                    ("SELECT Primary FROM " & relations(i, 0) & " WHERE " _
                                    & "Related = '" & rsGenes!ID & "'")
                  If rsrelations.EOF = falst Then
                     genes(record, 1) = rsrelations!primary
                     addToClusterGenes rsGenes!ID
                  Else
                     NoCluster = NoCluster + 1
            End Select
            record = record + 1
         Else 'this system isn't support
            NoCluster = NoCluster + 1
         End If
         i = 0
         
         rsGenes.MoveNext
      Wend
      ' at the end of this while loop we have
      ' genes () which has two columns the GEX primary and it's related clusterID if applicable
      ' clusterGenes - a collection of unique clusterGene objects
         '- each element contains an object containing a collection of GOIDs if any.
      ' noCluster - a count of the number of primary IDs that did not link to the cluster system
      
      'I should save this now so I don't have to go through all of this the next time.
      'I wonder if text is faster than a database? I don't need to access this information.
      'I just want to spit it out and read it back an a later date.
      
   
   
   
   Set idxCluster = tblMaptoCluster.CreateIndex("idxGO")
   idxCluster.Fields.Append .CreateField("Related", dbText, 15)
   tblMaptoCluster.Indexes.Append idxCluster
      
      
End Function


Public Sub addToClusterGenes(ClusterID As String)
   On Error GoTo error:
   Dim found As Boolean
   Dim ID As New ClusterGene
   Dim rsGO As Recordset
    
   ID = clusterGenes.Item(ClusterID)
   'this will either return an ID or throw an error. If it throws and error, then this is a new
   'gene. If it doesn't, then this ID is a duplicate and we can ignore it.
   
error:
   Select Case Err.Number
      Case 5
         Dim cg As New ClusterGene
         cg.setID ClusterID
         Set rsGO = dbMAPPfinder.OpenRecordset("SELECT related from " & GOrelation _
                                             & " WHERE Primary = '" & ClusterID & "'")
         While rsGO.EOF = False
            cg.addGOID (rsGO!related)
            rsGO.MoveNext
         Wend
         
         clusterGenes.Add cg, ClusterID
   End Select
End Sub
