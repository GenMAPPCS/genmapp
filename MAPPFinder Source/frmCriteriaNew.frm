VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCriteriaOld 
   BackColor       =   &H80000000&
   Caption         =   "PathFinder 1.0"
   ClientHeight    =   9630
   ClientLeft      =   6405
   ClientTop       =   1500
   ClientWidth     =   6300
   Icon            =   "frmCriteriaNew.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   6300
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   5880
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   5880
      Width           =   255
   End
   Begin VB.CheckBox chkDisplayALL 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   6480
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose New File"
      Height          =   735
      Left            =   1440
      TabIndex        =   15
      Top             =   8520
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFile 
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   7800
      Width           =   3975
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Browse"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   7800
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000000&
      Caption         =   "Yeast (GO Link made from SGD table)"
      Height          =   615
      Index           =   4
      Left            =   1200
      TabIndex        =   8
      Top             =   4320
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000000&
      Caption         =   "Mouse (GO Link made from MGI table)"
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   7
      Top             =   3720
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000000&
      Caption         =   "Human (GO Link made from EBI table)"
      Height          =   615
      Index           =   0
      Left            =   1200
      TabIndex        =   6
      Top             =   3240
      Width           =   3255
   End
   Begin VB.ListBox lstcriteria 
      Height          =   1230
      Left            =   3120
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.ListBox lstColorSet 
      Height          =   1230
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdRunPathFinder 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Run GenMAPP-GO PathFinder"
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Local MAPPs"
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Gene Ontology"
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Select the MAPP sets you would like to use for the calculations:"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   5400
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000000&
      Caption         =   "20020314"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000000&
      Caption         =   "Save Results as: (Component, Function, Process will be added to the file name.)"
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   7320
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000000&
      Caption         =   "Click here to limit the output to only those GO terms that were measured at least twice in your Expression Dataset."
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   6480
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000000&
      Caption         =   "Select your species and GO Table:"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000000&
      Caption         =   "Select Criteria to Filter by:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      Caption         =   "Select color set:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000000&
      Caption         =   "Please select the color set and criteria you would like to use to filter the data for significant genes."
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
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmCriteriaOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PathFinder 1.0
'Written by Scott Doniger
'Completed 12/11/2001
'This file contains the Pathfinder algorithm. It links Expression Data to GO terms and then
'calculates the percentage of genes in each GO term that changed based on the criteria selected
'by the user.

'Input: a .Gex file (GenMAPP expression Data file)
'Output: 3 text files, Component, Function, Process with the results.
'The first time pathfinder is run on a .gex file, a .gdb file is created storing
'the GO annotations for that entire chip. This total chip data is used for subsequent
'criteria.

'The code is pretty long, for the relative simplicity of the program. Unfortunately,
'each species is slightly different in how GO annotations are added, so the Pathfinder
'algorithm could not be generalized. Each new species will require modifications
'to the existing algorithm.
 
 
 
 Const MAX_CRITERIA = 30
 
 Dim rscolorsets As DAO.Recordset  'stores the colorsets of the .gex file
 Dim dbExpressionData As Database 'stores the expression table of the .gex file
 Dim sql(MAX_CRITERIA) As String
 Public species As String
 Dim fullname As String
 Dim newfilename As String
 Dim dbPathFinder As Database 'the pathfinder database pathfinder 1.0.mdb
 Dim dbChipData As Database 'the entire chips annotations.
 Dim DisplayALL As Boolean
 Dim GoTable As String
 Dim expressionName As String
 Dim filelocation As String
 Dim chipName As String
 Dim speciesselected As Boolean, geneontology As Boolean, LocalMAPPs As Boolean
   
Public Sub Load(FileName As String)
   Dim colorset As String
   Dim slash As Integer
   speciesselected = False
   filelocation = FileName
   slash = InStrRev(FileName, "\")
   newfilename = Mid(FileName, slash + 1, Len(FileName) - slash - 4) 'everything but .gex
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
   frmInput.Hide
End Sub

Private Sub Frame2_DragDrop(Source As Control, x As Single, y As Single)

End Sub



Private Sub Check1_Click()
   geneontology = True
End Sub

Private Sub Check2_Click()
   LocalMAPPs = True
End Sub

Private Sub cmdFile_Click()
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "Text Files|*.txt"
   CommonDialog1.ShowSave
   txtFile.Text = CommonDialog1.FileName
End Sub
 
Private Sub cmdRunPathFinder_Click()
  
   Dim tblNestedGO As TableDef
   Dim rsfilter As DAO.Recordset, rsType As DAO.Recordset
   Dim tblTempPrimary As TableDef, tbltempAll As TableDef, tblGOAll As TableDef, rstemp2 As DAO.Recordset
   Dim tblGO As TableDef, tblResults As TableDef, rschip As DAO.Recordset, tblnestedResults As TableDef
   Dim rsFunction As DAO.Recordset, rsProcess As DAO.Recordset, rsComponent As DAO.Recordset
   Dim percentage As Single, metFilter As Integer, noGO As Integer 'no GeneOntology available
   Dim others As Integer, noSwissProt As Integer 'no swissprot counts the number of genes that can't be converted
   Dim FSO As Object, present As Single
   Dim Output As TextStream
   Dim criteria As String
   Dim genmappID As String
   
   If speciesselected = False Then
      MsgBox "You must select a species before proceeding.", vbOKOnly
      GoTo nospeciesselected
   End If
   
   If geneontology = False And LocalMAPPs = False Then
      MsgBox "You have not select a MAPP Set(s) to calculate PathFinder results for. Please do so now.", vbOKOnly
      GoTo nospeciesselected
   End If
   
   MousePointer = vbHourglass
   Set FSO = CreateObject("Scripting.FileSystemObject")
   Select Case species
      Case "human"
         Set dbPathFinder = OpenDatabase(programpath & "PathFinder Human.mdb")
         
      Case "mouse"
         Set dbPathFinder = OpenDatabase(programpath & "PathFinder Mouse.mdb")
      Case "yeast"
         Set dbPathFinder = OpenDatabase(programpath & "PathFinder Mouse.mdb")
   End Select
   Set rsfilter = dbExpressionData.OpenRecordset("SELECT OrderNo, Primary, PrimaryType FROM" _
                        & " Expression WHERE (" & sql(lstcriteria.ListIndex) & ")")
   
   If rsfilter.EOF Then
      MsgBox "There are no genes in the Expression Dataset that meet the criteria you" _
            & " selected.", vbOKOnly
      GoTo ENDSUB
   End If
   
   rsfilter.MoveLast
   rsfilter.MoveFirst
   metFilter = rsfilter.RecordCount
   'check this
   If dbExpressionData.TableDefs.count > 10 Then 'somehow the tables didn't get deleted before
      dbExpressionData.TableDefs.Delete ("TempPrimary")
      'dbExpressionData.TableDefs.Delete ("TempAll")
      dbExpressionData.TableDefs.Delete ("GO")
      dbExpressionData.TableDefs.Delete ("Results")
      dbExpressionData.TableDefs.Delete ("NestedResults")
      dbExpressionData.TableDefs.Delete ("NestedGO")
   End If
   
   
   
   'Create the temporary tables that will be used to store the conversions and GO terms
   'There are two temp tables, TempPrimary and TempAll.
   'TempPrimary stores the ID that is related to GO as a primary key. This will be used to
   'to calculate the results of Pathfinder. This way, each MGI/SGD/SwissProt term is only counted
   'once, even if several ESTs representing the same full length gene are on the chip.
   
   'TempALL is used to show the backpage information(future plan). In this case est that links to a MGI/SGD/SwissProt
   'is stored, so that for each GO term, a list of all of the genes on the chip that are included
   'in that go term can be shown.
   
   'For example: If 4 ests all link to the same MGI in a mouse dataset, and that MGI links to
   'the go term Fatty Acid Degredation, then this will be counted as one occurence of Fatty Acid
   'Degredation on the chip. Not 4 occurences of the GO term. However, on the "backpage" the
   'expression data of all 4 genes will be shown.
   
   Set tblTempPrimary = dbExpressionData.CreateTableDef("TempPrimary")
      With tblTempPrimary
         .Fields.Append .CreateField("Primary", dbText, 15)
         Dim idxOrderNo As Index
         .Fields.Append .CreateField("OrderNo", dbLong)
         Set idxOrderNo = .CreateIndex("ixOrder")
         idxOrderNo.Fields.Append .CreateField("Primary")
         idxOrderNo.primary = True
         .Indexes.Append idxOrderNo
      End With
   ''Set tbltempAll = dbExpressionData.CreateTableDef("TempAll")
     ' With tbltempAll
      '   .Fields.Append .CreateField("Primary", dbText, 15)
       '  Dim idxAll As Index
        ' .Fields.Append .CreateField("OrderNo", dbLong)
        ' Set idxAll = .CreateIndex("ixOrder")
        ' idxAll.Fields.Append .CreateField("Primary")
        ' .Indexes.Append idxAll
        ' .Fields.Append .CreateField("Related", dbText, 15)
      'End With
   Set tblGO = dbExpressionData.CreateTableDef("GO")
      With tblGO
         .Fields.Append .CreateField("Primary", dbText, 15)
         Dim idxGO2 As Index
         .Fields.Append .CreateField("OrderNo", dbLong)
         Set idxGO2 = .CreateIndex("idxGO2")
         idxGO2.Fields.Append .CreateField("Primary", dbText, 15)
         idxGO2.Fields.Append .CreateField("GOID", dbText, 255)
         .Indexes.Append idxGO2
         .Fields.Append .CreateField("GOID", dbText, 255)
      End With
   Set tblNestedGO = dbExpressionData.CreateTableDef("NestedGO")
      With tblNestedGO
         .Fields.Append .CreateField("Primary", dbText, 15)
         .Fields.Append .CreateField("OrderNo", dbLong)
         Dim idxNestedGO As Index
         Dim idxNestedPrimary As Index
         Set idxNestedPrimary = .CreateIndex("idxPrimary")
         idxNestedPrimary.Fields.Append .CreateField("Primary", dbText, 15)
         idxNestedPrimary.primary = True
         Set idxNestedGO = .CreateIndex("idxNestedGO")
         idxNestedGO.Fields.Append .CreateField("GOID", dbText, 15)
         .Indexes.Append idxNestedPrimary
         .Indexes.Append idxNestedGO
         .Fields.Append .CreateField("GOID", dbText, 15)
      End With
   Set tblResults = dbExpressionData.CreateTableDef("Results")
      With tblResults
         .Fields.Append .CreateField("GOType", dbText, 2)
         Dim idxResults As Index
         .Fields.Append .CreateField("GOID", dbText, 255)
         .Fields.Append .CreateField("GOName", dbMemo)
         Set idxResults = .CreateIndex("Results")
         idxResults.Fields.Append .CreateField("GOType", dbText, 2)
         idxResults.Fields.Append .CreateField("Percentage", dbSingle)
         .Indexes.Append idxResults
         .Fields.Append .CreateField("InData", dbSingle)
         .Fields.Append .CreateField("OnChip", dbSingle)
         .Fields.Append .CreateField("InGO", dbSingle)
         .Fields.Append .CreateField("Percentage", dbSingle)
         .Fields.Append .CreateField("Present", dbSingle)
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
      End With
   'dbExpressionData.TableDefs.Append tbltempAll
   dbExpressionData.TableDefs.Append tblTempPrimary
   dbExpressionData.TableDefs.Append tblGO
   dbExpressionData.TableDefs.Append tblNestedGO
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
      Set tblTempPrimary = dbChipData.CreateTableDef("TempPrimary")
      With tblTempPrimary
         .Fields.Append .CreateField("Primary", dbText, 15)
         Dim idxOrderNoChip As Index
         .Fields.Append .CreateField("Related", dbText, 15)
         .Fields.Append .CreateField("PrimaryType", dbText, 2)
         Set idxOrderNoChip = .CreateIndex("ixOrder")
         idxOrderNoChip.Fields.Append .CreateField("Primary")
         idxOrderNoChip.primary = True
         .Indexes.Append idxOrderNoChip
      End With
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
      Set tbltempAll = dbChipData.CreateTableDef("TempAll")
      With tbltempAll
         .Fields.Append .CreateField("Primary", dbText, 15)
         Dim idxOrderNoChipALL As Index
         .Fields.Append .CreateField("Related", dbText, 15)
         .Fields.Append .CreateField("PrimaryType", dbText, 2)
         Set idxOrderNoChipALL = .CreateIndex("ixOrder")
         idxOrderNoChipALL.Fields.Append .CreateField("Primary")
         .Indexes.Append idxOrderNoChipALL
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
      dbChipData.TableDefs.Append tblTempPrimary
      dbChipData.TableDefs.Append tbltempAll
      dbChipData.TableDefs.Append tblGOAll
      Set tblNestedGOChip = dbChipData.CreateTableDef("NestedGO")
      With tblNestedGOChip
         .Fields.Append .CreateField("Primary", dbText, 15)
         Dim idxNestedGOChip As Index
         .Fields.Append .CreateField("OrderNo", dbLong)
         Set idxprimary = .CreateIndex("idxPrimary")
         idxprimary.Fields.Append .CreateField("Primary", dbText, 15)
         idxprimary.primary = True
         Set idxNestedGOChip = .CreateIndex("idxGO")
         idxNestedGOChip.Fields.Append .CreateField("GOID", dbText, 15)
         .Indexes.Append idxprimary
         .Indexes.Append idxNestedGOChip
         .Fields.Append .CreateField("GOID", dbText, 15)
      End With
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
      dbChipData.TableDefs.Append tblNestedGOChip
      dbChipData.TableDefs.Append tblLocalMAPPsChip
     buildOnChip  'this sub rountine builds the Chip table. This is information for the denominator of the ChipPercentage
   'idxOrderNo.Primary = True 'want to make Temp's index primary, so you don't get duplicate MGI's. This prevent's you from have a result of 200%
   Else
      Set dbChipData = OpenDatabase(newfilename & ".gdb")
      Set rstemp = dbChipData.OpenRecordset("SELECT Version, GOTable FROM Info")
      Set rstemp2 = dbPathFinder.OpenRecordset("SELECT Version FROM Info")
      If (rstemp![Version] <> rstemp2![Version]) Or (rstemp![GoTable] <> GoTable) Then 'the Chip table was built from a different version of the Pathfinder database. need to rebuild
         dbChipData.TableDefs.Delete "Chip"
         dbChipData.Execute ("DELETE * FROM LocalMAPPsChip")
         dbChipData.Execute ("Delete * from NestedChip")
         'dbChipData.TableDefs.Delete "TempPrimary"
         'dbChipData.TableDefs.Delete "GO"
         buildOnChip
      End If
      'you don't want to have to build the go chip info if someone is only using the local mapps.
      If geneontology Then
         Set rstemp = dbChipData.OpenRecordset("SElect * FROM nestedchip")
         If rstemp.EOF = True Then 'the go chip info didn't get built before
            buildOnChip
         End If
      End If
      
      If LocalMAPPs Then
         Set rstemp = dbChipData.OpenRecordset("SElect * FROM localmappschip")
         If rstemp.EOF = True Then 'the go chip info didn't get built before
            buildOnChip
         End If
      End If
      rstemp.Close
      rstemp2.Close
   End If
   
   If geneontology Then
   Select Case species
'=HUMAN========================================================================================
'==============================================================================================
      Case "human"
         While rsfilter.EOF = False
            If UCase(rsfilter![primaryType]) = "G" Then 'convert to SP via SP-GenBank table
               Set rstemp = dbPathFinder.OpenRecordset("Select Primary FROM [SwissProt-GenBank]" _
                           & " WHERE (related = '" & rsfilter![primary] & "')")
               If rstemp.EOF Then 'the genbank isn't in the SP-GB table.
                  noSwissProt = noSwissProt + 1
               Else 'the SP comes from the SP-GB table
               dbExpressionData.Execute ("INSERT INTO TempPrimary (Primary, OrderNo) VALUES ('" _
                           & rstemp![primary] & "', '" & rsfilter![OrderNo] & "')")
               End If
               rsfilter.MoveNext
            ElseIf UCase(rsfilter![primaryType]) = "S" Then
               dbExpressionData.Execute ("INSERT INTO TempPrimary (Primary, OrderNo) VALUES ('" _
                           & rsfilter![primary] & "', '" & rsfilter![OrderNo] & "')")
               rsfilter.MoveNext
            Else
               other = others + 1
               rsfilter.MoveNext
            End If
         Wend 'now you've attached SP's to as many genes as possible
         'the next step is to attch GO IDs to the SPs.
         
         Set rstemp = dbExpressionData.OpenRecordset("Select * from TempPrimary")
         If rstemp.EOF Then
            MsgBox "No genes have been linked to SwissProt. Check the data you are inputing " _
               & "into PathFinder.", vbOKOnly
            GoTo ENDSUB
         Else 'there are some SPs
            While Not rstemp.EOF
               Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT Related FROM [" & GoTable & " ]" _
                           & " WHERE Primary = '" & rstemp![primary] & "'")
               If Not rstemp2.EOF Then
                  While Not rstemp2.EOF
                     dbExpressionData.Execute ("INSERT INTO GO (Primary, OrderNo, GOID) VALUES " _
                        & "('" & rstemp![primary] & "', '" & rstemp![OrderNo] & "', '" _
                        & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               Else 'no geneontology available for that SP
                  noGO = noGO + 1
               End If
               rstemp.MoveNext
            Wend
          End If
      'you now have a table with every one of the SPs of the expression dataset linked to GO.
      'need to count how many times each GO term is represented.
      rstemp2.Close
      rstemp.Close
      Set rstemp = dbExpressionData.OpenRecordset("SELECT First([GOID]) as GOIDCount, Count([GOID]) AS" _
                  & " NumOfGOID FROM [GO] GROUP BY [GOID]")
      While Not rstemp.EOF
         'for each GOID in rstemp - select that GOID and it's count from GO-SwissProt Count in dbPathfinder
         'divide the GOIDCount in rstemp by GO-SwissProt Count to get the percentage changed in the dataset
         'insert into results GOType, GOID, percentage
         Set rstemp2 = dbPathFinder.OpenRecordset("SELECT NumberOfDups FROM [" & GoTable & "Count]" _
                     & " WHERE [RelatedField] = '" & rstemp![goidcount] & "'")
         'temp 2 now has the number of times the GO term is represented in that species' table
         
         Set rschip = dbChipData.OpenRecordset("SELECT NumOfGOID FROM [Chip] WHERE" _
                        & " [GOIDCount] = '" & rstemp![goidcount] & "'")
         'rsChip now has the number of times the GO term is on the chip
         
         percentage = (rstemp![numofGOID] / rschip![numofGOID]) * 100
         present = (rschip![numofGOID] / rstemp2![numberofdups]) * 100
         'percentage = # of times change/ # of times measured on chip * 100
         Set rsType = dbPathFinder.OpenRecordset("SELECT Type, Name FROM GeneOntology WHERE" _
                     & " GO = '" & rstemp![goidcount] & "'")
         If rsType.EOF = False Then 'for some reason the GO db and ontology are out of sync.
            If DisplayALL Then
               dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, GoName, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('" & rsType![Type] & "', '" _
                        & rstemp![goidcount] & "', '" & rsType![name] & "', " & rstemp![numofGOID] & ", " _
                        & rschip![numofGOID] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
            Else
               If rschip![numofGOID] > 1 Then
                  dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, GoName, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('" & rsType![Type] & "', '" _
                        & rstemp![goidcount] & "', '" & rsType![name] & "', " & rstemp![numofGOID] & ", " _
                        & rschip![numofGOID] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
               End If
            End If
         End If
         
   
         rstemp.MoveNext
      Wend
      TreeForm.createNestedGOTable dbExpressionData, TreeForm.root
      TreeForm.nestedResults dbExpressionData, dbChipData, TreeForm.root
            
      dbExpressionData.Execute ("DROP TABLE ROOT")
      Set rstemp = dbExpressionData.OpenRecordset("SELECT Gotype, GOID, GOName, InData, OnChip, InGO," _
                  & " Percentage, Present, InDataNested, OnChipNested, InGoNested, PercentageNested," _
                  & " PresentNested FROM NestedResults " _
                  & "ORDER BY NestedResults![PercentageNested] DESC, NestedResults![OnChipNested] DESC")
      Set Output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Gene Ontology Results.txt")
      Output.WriteLine ("PathFinder 1.0 Results for the Gene Ontology")
      Output.WriteLine ("File: " & filelocation)
      Output.WriteLine ("Statistics:")
      Output.WriteLine (metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria.")
      Output.WriteLine (others & " Genes had neither SwissProt or GenBank IDs.")
      Output.WriteLine (noSwissProt & " Genes did not link to an SwissProt ID.")
      Output.WriteLine (noGO & " Genes with SwissProt IDs did not link to a GO Term.")
      Output.WriteLine (metFilter - noSwissProt - noGO - others & " genes were used to calculate the results shown below.")
      Output.WriteLine ("")
      Output.WriteLine ("GOID" & Chr(9) & "GO Name" & Chr(9) & "GO Type" & Chr(9) & "Number Changed" & Chr(9) _
                        & "Number Measured" & Chr(9) & "Number in GO" & Chr(9) & "Percent Changed" _
                        & Chr(9) & "Percent Present" & Chr(9) & "Nuber Changed in Hierarchy" _
                        & Chr(9) & "Number Measured in Hierarchy" & Chr(9) & "Number in GO" _
                        & " Hierarchy" & Chr(9) & "Percent Changed in Hierarchy" & Chr(9) _
                        & "Percent Present in Hierarchy")
      While Not rstemp.EOF
         Output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![GOName] & Chr(9) & rstemp![gotype] _
                        & Chr(9) & rstemp![indata] _
                        & Chr(9) & rstemp![onChip] & Chr(9) & rstemp![ingo] & Chr(9) _
                        & rstemp![percentage] & Chr(9) & rstemp![present] & Chr(9) _
                        & rstemp![indatanested] & Chr(9) & rstemp![onchipnested] & Chr(9) _
                        & rstemp![ingonested] & Chr(9) & rstemp![percentagenested] & Chr(9) _
                        & rstemp![presentnested])
         rstemp.MoveNext
      Wend
      Output.Close
     
      
      
'=MOUSE==========================================================================================
'===============================================================================================
      Case "mouse"
         While rsfilter.EOF = False
            If UCase(rsfilter![primaryType]) = "G" Then 'convert to MGI via MGI-GenBank table
               Set rstemp = dbPathFinder.OpenRecordset("Select Primary FROM [MGI-GenBank]" _
                           & " WHERE (related = '" & rsfilter![primary] & "')")
               If rstemp.EOF Then 'the genbank isn't in the MGI-GB table. Try going through Unigene
                  Set rstemp = dbPathFinder.OpenRecordset("Select Primary FROM [Unigene-GenBank]" _
                           & " WHERE (related = '" & rsfilter![primary] & "')")
                  If rstemp.EOF = False Then
                     Set rstemp2 = dbPathFinder.OpenRecordset("Select Primary FROM [MGI-Unigene]" _
                           & " WHERE (related = '" & rstemp![primary] & "')")
                     If rstemp2.EOF Then
                        noSwissProt = noSwissProt + 1
                     Else 'you found an MGI via Unigene!
                        dbExpressionData.Execute ("INSERT INTO TempPrimary (Primary, OrderNo) VALUES ('" _
                           & rstemp2![primary] & "', '" & rsfilter![OrderNo] & "')")
                     End If
                  Else 'not in unigene either
                     noSwissProt = noSwissProt + 1
                  End If
               Else 'the MGI comes from the MGI-GB table
               dbExpressionData.Execute ("INSERT INTO TempPrimary (Primary, OrderNo) VALUES ('" _
                           & rstemp![primary] & "', '" & rsfilter![OrderNo] & "')")
               End If
               rsfilter.MoveNext
            ElseIf UCase(rsfilter![primaryType]) = "S" Then
               Set rstemp = dbPathFinder.OpenRecordset("Select Related FROM [SwissProt-MGI]" _
                           & " WHERE (primary = '" & rsfilter![primary] & "')")
               If rstemp.EOF Then
                  noSwissProt = noSwissProt + 1
               Else 'MGI from SP-MGI
                  dbExpressionData.Execute ("INSERT INTO TempPrimary (Primary, OrderNo) VALUES ('" _
                           & rstemp![related] & "', '" & rsfilter![OrderNo] & "')")
               End If
               rsfilter.MoveNext
            Else
               other = others + 1
               rsfilter.MoveNext
            End If
         Wend 'now you've attached MGI's to as many genes as possible (via MGI-Genbank, MGI-SwissProt, MGI-Unigene)
         'the next step is to attch GO IDs to the MGIs.
         
         Set rstemp = dbExpressionData.OpenRecordset("Select * from TempPrimary")
         If rstemp.EOF Then
            MsgBox "No genes have been linked to MGI. Check the data you are inputing " _
               & "into PathFinder.", vbOKOnly
            GoTo ENDSUB
         Else 'there are some MGIs
            While Not rstemp.EOF
               Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT Related FROM [" & GoTable & " ]" _
                           & " WHERE Primary = '" & rstemp![primary] & "'")
               If Not rstemp2.EOF Then
                  While Not rstemp2.EOF
                     dbExpressionData.Execute ("INSERT INTO GO (Primary, OrderNo, GOID) VALUES " _
                        & "('" & rstemp![primary] & "', '" & rstemp![OrderNo] & "', '" _
                        & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               Else 'no geneontology available for that MGI
                  noGO = noGO + 1
               End If
               rstemp.MoveNext
            Wend
          End If
      'you now have a table with every one of the MGIs of the expression dataset linked to GO.
      'need to count how many times each GO term is represented.
      rstemp2.Close
      rstemp.Close
      Set rstemp = dbExpressionData.OpenRecordset("SELECT First([GOID]) as GOIDCount, Count([GOID]) AS" _
                  & " NumOfGOID FROM [GO] GROUP BY [GOID]")
      While Not rstemp.EOF
         'for each GOID in rstemp - select that GOID and it's count from GO-MGI Count in dbPathfinder
         'divide the GOIDCount in rstemp by GO-MGI Count to get the percentage changed in the dataset
         'insert into results GOType, GOID, percentage
         Set rstemp2 = dbPathFinder.OpenRecordset("SELECT NumberOfDups FROM [" & GoTable & "Count]" _
                     & " WHERE [RelatedField] = '" & rstemp![goidcount] & "'")
         'temp 2 now has the number of times the GO term is represented in that species' table
         
         Set rschip = dbChipData.OpenRecordset("SELECT NumOfGOID FROM [Chip] WHERE" _
                        & " [GOIDCount] = '" & rstemp![goidcount] & "'")
         'rsChip now has the number of times the GO term is on the chip
         
         percentage = (rstemp![numofGOID] / rschip![numofGOID]) * 100
         present = (rschip![numofGOID] / rstemp2![numberofdups]) * 100
         'percentage = # of times change/ # of times measured on chip * 100
         Set rsType = dbPathFinder.OpenRecordset("SELECT Type, Name FROM GeneOntology WHERE" _
                     & " ID = '" & rstemp![goidcount] & "'")
         If DisplayALL Then
            dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, GoName, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('" & rsType![Type] & "', '" _
                        & rstemp![goidcount] & "', '" & rsType![name] & "', " & rstemp![numofGOID] & ", " _
                        & rschip![numofGOID] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
         Else
            If rschip![numofGOID] > 1 Then
               dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, GoName, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('" & rsType![Type] & "', '" _
                        & rstemp![goidcount] & "', '" & rsType![name] & "', " & rstemp![numofGOID] & ", " _
                        & rschip![numofGOID] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
            End If
         End If
         
         
   
         rstemp.MoveNext
      Wend
      TreeForm.createNestedGOTable dbExpressionData, TreeForm.root
      TreeForm.nestedResults dbExpressionData, dbChipData, TreeForm.root
      dbExpressionData.Execute ("DROP TABLE ROOT")
      Set rstemp = dbExpressionData.OpenRecordset("SELECT Gotype, GOID, GOName, InData, OnChip, InGO," _
                  & " Percentage, Present, InDataNested, OnChipNested, InGoNested, PercentageNested," _
                  & " PresentNested FROM NestedResults " _
                  & "ORDER BY NestedResults![PercentageNested] DESC, NestedResults![OnChipNested] DESC")
      Set Output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Gene Ontology Results.txt")
      Output.WriteLine ("PathFinder 1.0 Results for Results for the Gene Ontology")
      Output.WriteLine ("File: " & filelocation)
      Output.WriteLine ("Statistics:")
      Output.WriteLine (metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria.")
      Output.WriteLine (others & " Genes had neither SwissProt or GenBank IDs.")
      Output.WriteLine (noSwissProt & " Genes did not link to an SwissProt ID.")
      Output.WriteLine (noGO & " Genes with SwissProt IDs did not link to a GO Term.")
      Output.WriteLine (metFilter - noSwissProt - noGO - others & " genes were used to calculate the results shown below.")
      Output.WriteLine ("")
      Output.WriteLine ("GOID" & Chr(9) & "GO Name" & Chr(9) & "GO Type" & Chr(9) & "Number Changed" & Chr(9) _
                        & "Number Measured" & Chr(9) & "Number in GO" & Chr(9) & "Percent Changed" _
                        & Chr(9) & "Percent Present" & Chr(9) & "Nuber Changed in Hierarchy" _
                        & Chr(9) & "Number Measured in Hierarchy" & Chr(9) & "Number in GO" _
                        & " Hierarchy" & Chr(9) & "Percent Changed in Hierarchy" & Chr(9) _
                        & "Percent Present in Hierarchy")
      While Not rstemp.EOF
         Output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![GOName] & Chr(9) & rstemp![gotype] _
                        & Chr(9) & rstemp![indata] _
                        & Chr(9) & rstemp![onChip] & Chr(9) & rstemp![ingo] & Chr(9) _
                        & rstemp![percentage] & Chr(9) & rstemp![present] & Chr(9) _
                        & rstemp![indatanested] & Chr(9) & rstemp![onchipnested] & Chr(9) _
                        & rstemp![ingonested] & Chr(9) & rstemp![percentagenested] & Chr(9) _
                        & rstemp![presentnested])
         rstemp.MoveNext
      Wend
      Output.Close

'=YEAST=========================================================================================
'===============================================================================================
      Case "yeast"
      While rsfilter.EOF = False
            If UCase(rsfilter![primaryType]) = "S" Then 'convert to SGD via SP-SGD table
               Set rstemp = dbPathFinder.OpenRecordset("Select Related FROM [SwissProt-SGD]" _
                           & " WHERE (primary = '" & rsfilter![primary] & "')")
               If rstemp.EOF Then 'the SwissProt isn't in the SP-SGD table.
                  noSwissProt = noSwissProt + 1
               Else 'the SGD comes from the SP-SGD table
               dbExpressionData.Execute ("INSERT INTO TempPrimary (Primary, OrderNo) VALUES ('" _
                           & rstemp![primary] & "', '" & rsfilter![OrderNo] & "')")
               End If
               rsfilter.MoveNext
            ElseIf UCase(rsfilter![primaryType]) = "D" Then
               dbExpressionData.Execute ("INSERT INTO TempPrimary (Primary, OrderNo) VALUES ('" _
                           & rsfilter![primary] & "', '" & rsfilter![OrderNo] & "')")
               rsfilter.MoveNext
            Else
               other = others + 1
               rsfilter.MoveNext
            End If
         Wend 'now you've attached SGD's to as many genes as possible
         'the next step is to attch GO IDs to the SGDs.
         
         Set rstemp = dbExpressionData.OpenRecordset("Select * from TempPrimary")
         If rstemp.EOF Then
            MsgBox "No genes have been linked to SGD. Check the data you are inputing " _
               & "into PathFinder.", vbOKOnly
            GoTo ENDSUB
         Else 'there are some SPs
            While Not rstemp.EOF
               Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT Related FROM [SGD-GeneOntology]" _
                           & " WHERE Primary = '" & rstemp![primary] & "'")
               If Not rstemp2.EOF Then
                  While Not rstemp2.EOF
                     dbExpressionData.Execute ("INSERT INTO GO (Primary, OrderNo, GOID) VALUES " _
                        & "('" & rstemp![primary] & "', '" & rstemp![OrderNo] & "', '" _
                        & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               Else 'no geneontology available for that SP
                  noGO = noGO + 1
               End If
               rstemp.MoveNext
            Wend
          End If
      'you now have a table with every one of the SPs of the expression dataset linked to GO.
      'need to count how many times each GO term is represented.
      rstemp2.Close
      rstemp.Close
      Set rstemp = dbExpressionData.OpenRecordset("SELECT First([GOID]) as GOIDCount, Count([GOID]) AS" _
                  & " NumOfGOID FROM [GO] GROUP BY [GOID]")
      While Not rstemp.EOF
         'for each GOID in rstemp - select that GOID and it's count from GO-SwissProt Count in dbPathfinder
         'divide the GOIDCount in rstemp by GO-SwissProt Count to get the percentage changed in the dataset
         'insert into results GOType, GOID, percentage
         Set rstemp2 = dbPathFinder.OpenRecordset("SELECT NumberOfDups FROM [SGD-GOCount]" _
                     & " WHERE [RelatedField] = '" & rstemp![goidcount] & "'")
         'temp 2 now has the number of times the GO term is represented in that species' table
         
         Set rschip = dbChipData.OpenRecordset("SELECT NumOfGOID FROM [Chip] WHERE" _
                        & " [GOIDCount] = '" & rstemp![goidcount] & "'")
         'rsChip now has the number of times the GO term is on the chip
         
         percentage = (rstemp![numofGOID] / rschip![numofGOID]) * 100
         present = (rschip![numofGOID] / rstemp2![numberofdups]) * 100
         'percentage = # of times change/ # of times measured on chip * 100
         Set rsType = dbPathFinder.OpenRecordset("SELECT Type, Name FROM GeneOntology WHERE" _
                     & " GO = '" & rstemp![goidcount] & "'")
         If DisplayALL Then
            dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, GoName, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('" & rsType![Type] & "', '" _
                        & rstemp![goidcount] & "', '" & rsType![name] & "', " & rstemp![numofGOID] & ", " _
                        & rschip![numofGOID] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
         Else
            If rschip![numofGOID] > 1 Then
               dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, GoName, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('" & rsType![Type] & "', '" _
                        & rstemp![goidcount] & "', '" & rsType![name] & "', " & rstemp![numofGOID] & ", " _
                        & rschip![numofGOID] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
            End If
         End If
         
         
   
         rstemp.MoveNext
      Wend
      TreeForm.createNestedGOTable dbExpressionData, TreeForm.root
      TreeForm.nestedResults dbExpressionData, dbChipData, TreeForm.root
      dbExpressionData.Execute ("DROP TABLE ROOT")
      Set rstemp = dbExpressionData.OpenRecordset("SELECT Gotype, GOID, GOName, InData, OnChip, InGO," _
                  & " Percentage, Present, InDataNested, OnChipNested, InGoNested, PercentageNested," _
                  & " PresentNested FROM NestedResults WHERE " _
                  & "ORDER BY NestedResults![PercentageNested] DESC, NestedResults![OnChipNested] DESC")
      Set Output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Gene Ontology Results.txt")
      Output.WriteLine ("PathFinder 1.0 Results for Results for the Gene Ontology")
      Output.WriteLine ("File: " & filelocation)
      Output.WriteLine ("Statistics:")
      Output.WriteLine (metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria.")
      Output.WriteLine (others & " Genes had neither SwissProt or GenBank IDs.")
      Output.WriteLine (noSwissProt & " Genes did not link to an SwissProt ID.")
      Output.WriteLine (noGO & " Genes with SwissProt IDs did not link to a GO Term.")
      Output.WriteLine (metFilter - noSwissProt - noGO - others & " genes were used to calculate the results shown below.")
      Output.WriteLine ("")
      Output.WriteLine ("GOID" & Chr(9) & "GO Name" & Chr(9) & "GO Type" & Chr(9) & "Number Changed" & Chr(9) _
                        & "Number Measured" & Chr(9) & "Number in GO" & Chr(9) & "Percent Changed" _
                        & Chr(9) & "Percent Present" & Chr(9) & "Nuber Changed in Hierarchy" _
                        & Chr(9) & "Number Measured in Hierarchy" & Chr(9) & "Number in GO" _
                        & " Hierarchy" & Chr(9) & "Percent Changed in Hierarchy" & Chr(9) _
                        & "Percent Present in Hierarchy")
      While Not rstemp.EOF
         Output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![GOName] & Chr(9) & rstemp![gotype] _
                        & Chr(9) & rstemp![indata] _
                        & Chr(9) & rstemp![onChip] & Chr(9) & rstemp![ingo] & Chr(9) _
                        & rstemp![percentage] & Chr(9) & rstemp![present] & Chr(9) _
                        & rstemp![indatanested] & Chr(9) & rstemp![onchipnested] & Chr(9) _
                        & rstemp![ingonested] & Chr(9) & rstemp![percentagenested] & Chr(9) _
                        & rstemp![presentnested])
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


   If LocalMAPPs Then
      noGO = 0
      dbExpressionData.Execute ("DELETE * FROM TempPrimary")
      dbExpressionData.Execute ("DELETE * FROM GO")
      dbExpressionData.Execute ("DELETE * FROM Results")
      
      'rstemp will store all of the GenMAPP IDs (without duplicates) for the genes that meet the criteria.
      Set rstemp = dbExpressionData.OpenRecordset("SELECT DISTINCT GenMAPP FROM" _
                        & " Expression WHERE (" & sql(lstcriteria.ListIndex) & ")")
      metFilter = rstemp.RecordCount
      If rstemp.EOF Then
         MsgBox "There are no genes meeting your criteria in the Expression Dataset.", vbOKOnly
         GoTo ENDSUB
      End If
   
      While rstemp.EOF = False
         genmappID = rstemp![GenMAPP]
         If InStr(1, genmappID, "~") <> 0 Then
            genmappID = Right(genmappID, Len(genmappID) - 1)
         End If
         Set rstemp2 = dbPathFinder.OpenRecordset("Select MAPPNameField from GeneToMAPP WHERE" _
                        & " GenMAPP = '" & genmappID & "'")
         
         If Not rstemp2.EOF Then
            While Not rstemp2.EOF
               dbExpressionData.Execute ("INSERT INTO GO (Primary, GOID) VALUES " _
                        & "('" & genmappID & "', '" & rstemp2![MappNameField] & "')")
               rstemp2.MoveNext
            Wend
         Else 'no geneontology available for that SP
            noGO = noGO + 1
         End If
         rstemp.MoveNext
      Wend
         
      
      'you now have a table with every one of the GenMAPP IDs of the expression dataset linked to a MAPP.
      'need to count how many times each MAPP term is represented.
      rstemp2.Close
      rstemp.Close
      'here GOID = MAPPName (a left over from the GO portion of PathFinder)
      Set rstemp = dbExpressionData.OpenRecordset("SELECT First([GOID]) as GOIDCount, Count([GOID]) AS" _
                  & " NumOfGOID FROM [GO] GROUP BY [GOID]")
      Debug.Print
      While Not rstemp.EOF
         'for each GOID in rstemp -> select that GOID and it's count from GeneToMAPPCount in dbPathfinder
         'divide the GOIDCount in rstemp by GeneToMAPPCount to get the percentage changed in the dataset
         'insert into results GOType, GOID, percentage
        
         
         Set rstemp2 = dbPathFinder.OpenRecordset("SELECT NumberOfDups FROM [GeneToMAPPCount]" _
                     & " WHERE [MAPPName] = '" & rstemp![goidcount] & "'")
         'temp 2 now has the number of times the MAPP is represented in the GeneToMAPP table
         
         Set rschip = dbChipData.OpenRecordset("SELECT OnChip FROM [LocalMAPPsChip] WHERE" _
                        & " [MAPPName] = '" & rstemp![goidcount] & "'")
         'rsChip now has the number of times the GO term is on the chip
         
         percentage = (rstemp![numofGOID] / rschip![onChip]) * 100
         present = (rschip![onChip] / rstemp2![numberofdups]) * 100
         'percentage = # of times change/ # of times measured on chip * 100
         
         'If rsType.EOF = False Then 'for some reason the GO db and ontology are out of sync.
            If DisplayALL Then
               dbExpressionData.Execute ("INSERT INTO Results(GOType, GOID, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('L', '" _
                        & rstemp![goidcount] & "', " & rstemp![numofGOID] & ", " _
                        & rschip![onChip] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
            Else
               If rschip![numofGOID] > 1 Then
                  dbExpressionData.Execute ("INSERT INTO Results(GOType, GoID, InData, OnChip," _
                        & " InGO, Percentage, Present) VALUES ('L', '" _
                        & rstemp![goidcount] & "', " & rstemp![numofGOID] & ", " _
                        & rschip![numofGOID] & ", " & rstemp2![numberofdups] & ", " & percentage _
                        & ", " & present & ")")
               End If
            End If
         'End If
         
   
         rstemp.MoveNext
      Wend
      TreeForm.DisplayLocalMAPPs dbExpressionData, TreeForm.GetLocalRoot()
      rstemp.Close
      Set rstemp = dbExpressionData.OpenRecordset("SELECT GOID, InData, OnChip, InGO," _
                  & " Percentage, Present FROM Results WHERE (GOType = 'L') " _
                  & "ORDER BY Results![Percentage] DESC, Results![OnChip] DESC")
      Set Output = FSO.CreateTextFile(fixFileName(txtFile.Text) & "-Local Results.txt")
      Output.WriteLine ("PathFinder 1.0 Results for Local MAPPs")
      Output.WriteLine ("Statistics:")
      Output.WriteLine (metFilter & " Genes met the " & sql(lstcriteria.ListIndex) & " criteria.")
      Output.WriteLine (noGO & " Genes did not link to a MAPP.")
      Output.WriteLine (metFilter - noGO & " genes were used to calculate the results shown below.")
      Output.WriteLine ("")
      Output.WriteLine ("MAPPName" & Chr(9) & "Number Changed" & Chr(9) _
                        & "Number Measured" & Chr(9) & "Number on MAPP" & Chr(9) & "Percent Changed" _
                        & Chr(9) & "Percent Present")
      While Not rstemp.EOF
         Output.WriteLine (rstemp![GOID] & Chr(9) & rstemp![indata] _
                        & Chr(9) & rstemp![onChip] & Chr(9) & rstemp![ingo] & Chr(9) _
                        & rstemp![percentage] & Chr(9) & rstemp![present])
         rstemp.MoveNext
      Wend
      Output.Close
   End If
   TreeForm.setChipDBLocation filelocation
   rstemp.Close
   rstemp2.Close
'   rsType.Close
   rschip.Close
   'dbExpressionData.TableDefs.Delete ("TempAll")
   dbExpressionData.TableDefs.Delete ("TempPrimary")
   dbExpressionData.TableDefs.Delete ("GO")
   dbExpressionData.TableDefs.Delete ("Results")
   dbExpressionData.TableDefs.Delete ("NestedGO")
   dbExpressionData.TableDefs.Delete ("NestedResults")
      
      
     
ENDSUB:
     
      dbExpressionData.Close
      dbChipData.Close
      dbPathFinder.Close
      
      DBEngine.CompactDatabase newfilename & ".gex", newfilename & ".$tm"
      Kill newfilename & ".gex"
      Name newfilename & ".$tm" As newfilename & ".gex"
      DBEngine.CompactDatabase newfilename & ".gdb", newfilename & ".$tm"
      Kill newfilename & ".gdb"
      Name newfilename & ".$tm" As newfilename & ".gdb"
      MousePointer = vbDefault
      TreeForm.Show
      TreeForm.Refresh
nospeciesselected:

error:
   Select Case Err.Number
      Case 3024 'the error for not having the database
         MsgBox "The database PathFinder " & species & " was not found in the folder" _
         & " containing this application. Please move it to this folder, or downloaded from GenMAPP.org.", vbOKOnly
   End Select
End Sub

  


Private Sub chkDisplayALL_Click()
   If chkDisplayALL.Value = 1 Then
      DisplayALL = False
   Else
      DisplayALL = True
   End If
End Sub





Private Sub Command1_Click()
   frmInput.Show
   frmCriteria.Hide
   MousePointer = vbDefault
End Sub

Private Sub Exit_Click()
   End
End Sub

Private Sub Label9_Click()
   geneontology = False
End Sub

Private Sub lstColorSet_Click()
   Dim rsCriteria As DAO.Recordset
   Dim criteria As String, record As String
   Dim pipe As Integer, endline As Integer, newend As Integer, pipe2 As Integer
   Dim i As Integer
   lstcriteria.Clear
   
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
   criteria = ""
   For i = 0 To lstcriteria.ListCount - 1
      If lstcriteria.Selected(i) Then
          criteria = criteria & "-" & (i + 1)
      End If
   Next i
   txtFile.Text = programpath & newfilename & "-ColorSet" & lstColorSet.ListIndex _
                  & "-Criteria" & criteria & ".txt"
End Sub

Private Sub Option1_Click(Index As Integer)
   speciesselected = True
      
   Select Case Index
      Case 0
         species = "human"
         GoTable = "SwissProt-GeneOntology"
         TreeForm.setSpecies (species)
      Case 1
         species = "human"
         GoTable = "HumanCompugen"
      Case 2
         species = "mouse"
         GoTable = "MGI-GeneOntology"
         TreeForm.setSpecies (species)
      Case 3
         species = "mouse"
         GoTable = "MouseCompugen"
      Case 4
         species = "yeast"
         TreeForm.setSpecies (species)
   End Select
End Sub

'There are two sets of tables being created. TempPrimary makes the related ID (SP,MGI,SGD) a primary ID, so that duplicate
'genes are not counted twice. TempALL and GOALL have all of the genes in the ED linked to GO so that the MAPP can be built with
'all of the genes in the dataset.


Public Sub buildOnChip()
    Dim rsExpression As DAO.Recordset
    Dim rstemp As DAO.Recordset, rstemp2 As DAO.Recordset, rsGO As DAO.Recordset
   
    Set rstemp = dbPathFinder.OpenRecordset("Select Version from Info")
    
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
               & "', '" & GoTable & "')")
   
   If geneontology Then
   Set rsExpression = dbExpressionData.OpenRecordset("SELECT Primary, PrimaryType, OrderNo FROM Expression")
   Select Case species
      Case "human"
        While rsExpression.EOF = False
            If UCase(rsExpression![primaryType]) = "G" Then 'convert to SP via SP-GenBank table
               Set rstemp = dbPathFinder.OpenRecordset("Select Primary FROM [SwissProt-GenBank]" _
                           & " WHERE (related = '" & rsExpression![primary] & "')")
               If rstemp.EOF = False Then 'you found a sp
                  dbChipData.Execute ("INSERT INTO TempPrimary (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![primary] & "', '" & rsExpression![primary] & "', 'G')")
                  dbChipData.Execute ("INSERT INTO TempALL (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![primary] & "', '" & rsExpression![primary] & "', 'G')")
               End If
               rsExpression.MoveNext
            ElseIf UCase(rsExpression![primaryType]) = "S" Then
               dbChipData.Execute ("INSERT INTO TempPrimary (Related, Primary, PrimaryType) VALUES ('" _
                           & rsExpression![primary] & "', '" & rsExpression![primary] & "', 'S')")
                dbChipData.Execute ("INSERT INTO TempALL (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![primary] & "', '" & rsExpression![primary] & "', 'S')")
               rsExpression.MoveNext
            Else
               rsExpression.MoveNext
            End If
         Wend 'now you've attached SP's to as many genes as possible
         'the next step is to attch GO IDs to the MGIs.
         
         Set rstemp = dbChipData.OpenRecordset("Select * from TempPrimary")
         While Not rstemp.EOF
            Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT Related FROM [" & GoTable & " ]" _
                           & " WHERE Primary = '" & rstemp![related] & "'")
               If Not rstemp2.EOF Then
                  While Not rstemp2.EOF
                     dbChipData.Execute ("INSERT INTO GO (Primary, Related, GOID) VALUES " _
                        & "('" & rstemp![primary] & "', '" & rstemp![related] & "', '" _
                        & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               End If
               rstemp.MoveNext
         Wend
   
         Set rstemp = dbChipData.OpenRecordset("Select * from TempAll")
         While Not rstemp.EOF
            Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT Related FROM [" & GoTable & " ]" _
                           & " WHERE Primary = '" & rstemp![related] & "'")
               If Not rstemp2.EOF Then
                  While Not rstemp2.EOF
                     dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rstemp![primary] & "', '" & "', '" & rstemp![primaryType] _
                        & "', '" & rstemp![related] & "', '" _
                        & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               End If
               rstemp.MoveNext
         Wend
      'you now have a table with every one of the genes in the expression dataset linked to GO.
      'need to count how many times each GO term is represented from the primary table.
      rstemp2.Close
      rstemp.Close
      
      dbChipData.Execute ("INSERT INTO Chip SELECT First([GOID]) as GOIDCount, Count([GOID]) AS" _
                  & " NumOfGOID FROM [GO] GROUP BY [GOID]")
         
      Case "mouse"
        While rsExpression.EOF = False
            If UCase(rsExpression![primaryType]) = "G" Then 'convert to MGI via MGI-GenBank table
               Set rstemp = dbPathFinder.OpenRecordset("Select Primary FROM [MGI-GenBank]" _
                           & " WHERE (related = '" & rsExpression![primary] & "')")
               If rstemp.EOF Then 'the genbank isn't in the MGI-GB table. Try going through Unigene
                  Set rstemp = dbPathFinder.OpenRecordset("Select Primary FROM [Unigene-GenBank]" _
                           & " WHERE (related = '" & rsExpression![primary] & "')")
                  If rstemp.EOF = False Then
                     Set rstemp2 = dbPathFinder.OpenRecordset("Select Primary FROM [MGI-Unigene]" _
                           & " WHERE (related = '" & rstemp![primary] & "')")
                     If Not rstemp2.EOF Then
                        'you found an MGI via Unigene!
                        dbChipData.Execute ("INSERT INTO TempPrimary (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp2![primary] & "', '" & rsExpression![primary] & "', 'G')")
                        dbChipData.Execute ("INSERT INTO TempAll (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp2![primary] & "', '" & rsExpression![primary] & "', 'G')")
                     End If
                  End If
               Else 'the MGI comes from the MGI-GB table
               dbChipData.Execute ("INSERT INTO TempPrimary (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![primary] & "', '" & rsExpression![primary] & "', 'G')")
               dbChipData.Execute ("INSERT INTO TempAll (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![primary] & "', '" & rsExpression![primary] & "', 'G')")
               End If
               rsExpression.MoveNext
            ElseIf UCase(rsExpression![primaryType]) = "S" Then
               Set rstemp = dbPathFinder.OpenRecordset("Select Related FROM [SwissProt-MGI]" _
                           & " WHERE (primary = '" & rsExpression![primary] & "')")
               'MGI from SP-MGI
               dbChipData.Execute ("INSERT INTO TempPrimary (related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![related] & "', '" & rsExpression![primary] & "', 'S')")
               dbChipData.Execute ("INSERT INTO TempALL (related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![related] & "', '" & rsExpression![primary] & "', 'S')")
               rsExpression.MoveNext
            Else
               rsExpression.MoveNext
            End If
         Wend 'now you've attached MGI's to as many genes as possible (via MGI-Genbank, MGI-SwissProt, MGI-Unigene)
         'the next step is to attch GO IDs to the MGIs.
         
         Set rstemp = dbChipData.OpenRecordset("SeLECT * FROM TempPrimary")
         While Not rstemp.EOF
            Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT Related FROM [" & GoTable _
                           & "] WHERE Primary = '" & rstemp![related] & "'")
               If Not rstemp2.EOF Then
                  While Not rstemp2.EOF
                     dbChipData.Execute ("INSERT INTO GO (Primary, related, GOID) VALUES " _
                        & "('" & rstemp![primary] & "', '" & rstemp![related] & "', '" _
                        & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               End If
               rstemp.MoveNext
         Wend
   
         Set rstemp = dbChipData.OpenRecordset("SeLECT * FROM TempAll")
         While Not rstemp.EOF
            Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT Related FROM [" & GoTable _
                           & "] WHERE Primary = '" & rstemp![related] & "'")
               If Not rstemp2.EOF Then
                  While Not rstemp2.EOF
                     dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rstemp![primary] & "', '" & rstemp![primaryType] _
                        & "', '" & rstemp![related] & "', '" _
                        & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               End If
               rstemp.MoveNext
         Wend
      'you now have a table with every one of the MGIs of the expression dataset linked to GO.
      'need to count how many times each GO term is represented.
      rstemp2.Close
      rstemp.Close
      
      dbChipData.Execute ("INSERT INTO Chip SELECT First([GOID]) as GOIDCount, Count([GOID]) AS" _
                  & " NumOfGOID FROM [GO] GROUP BY [GOID]")
         
      Case "yeast"
           While rsExpression.EOF = False
            If UCase(rsExpression![primaryType]) = "S" Then 'convert to SGD via SP-SGD table
               Set rstemp = dbPathFinder.OpenRecordset("Select Related FROM [SwissProt-SGD]" _
                           & " WHERE (primary = '" & rsExpression![primary] & "')")
               If rstemp.EOF = False Then 'you found a sgd
                  dbChipData.Execute ("INSERT INTO TempPrimary (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![primary] & "', '" & rsExpression![primary] & "', 'S')")
                  dbChipData.Execute ("INSERT INTO TempAll (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![primary] & "', '" & rsExpression![primary] & "', 'S')")
               End If
               rsExpression.MoveNext
            ElseIf UCase(rsExpression![primaryType]) = "D" Then
              dbChipData.Execute ("INSERT INTO TempPrimary (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![primary] & "', '" & rsExpression![primary] & "', 'D')")
                  dbChipData.Execute ("INSERT INTO TempAll (Related, Primary, PrimaryType) VALUES ('" _
                           & rstemp![primary] & "', '" & rsExpression![primary] & "', 'D')")
               rsExpression.MoveNext
            Else
               rsExpression.MoveNext
            End If
         Wend 'now you've attached SGD's to as many genes as possible
         'the next step is to attch GO IDs to the SGDs.
         
         Set rstemp = dbChipData.OpenRecordset("Select * from TempPrimary")
         While Not rstemp.EOF
            Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT Related FROM [SGD-GeneOntology]" _
                           & " WHERE Primary = '" & rstemp![primary] & "'")
               If Not rstemp2.EOF Then
                  While Not rstemp2.EOF
                     dbChipData.Execute ("INSERT INTO GO (Primary, Related, GOID) VALUES " _
                        & "('" & rstemp![primary] & "', '" & rstemp![related] & "', '" _
                        & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               End If
               rstemp.MoveNext
         Wend
         
          Set rstemp = dbChipData.OpenRecordset("Select * from TempAll")
         While Not rstemp.EOF
            Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT Related FROM [SGD-GeneOntology]" _
                           & " WHERE Primary = '" & rstemp![primary] & "'")
               If Not rstemp2.EOF Then
                  While Not rstemp2.EOF
                    dbChipData.Execute ("INSERT INTO GOAll (Primary, PrimaryType, Related, GOID) VALUES " _
                        & "('" & rstemp![primary] & "', '" & "', '" & rstemp![primaryType] _
                        & "', '" & rstemp![related] & "', '" _
                        & rstemp2![related] & "')")
                     rstemp2.MoveNext
                  Wend
               End If
               rstemp.MoveNext
         Wend
      rstemp2.Close
      rstemp.Close
      'you now have a table with every one of the SPs of the expression dataset linked to GO.
      'need to count how many times each GO term is represented.
      
      
      dbChipData.Execute ("INSERT INTO Chip SELECT First([GOID]) as GOIDCount, Count([GOID]) AS" _
                  & " NumOfGOID FROM [GO] GROUP BY [GOID]")
         
   End Select
   TreeForm.createNestedGOTable dbChipData, TreeForm.root
   TreeForm.NestedChipData dbChipData, TreeForm.root '"GO" is the root of the tree
   End If
   '=========================================================================================
   'Now you need to calculate for the
       
   If LocalMAPPs Then
      dbChipData.Execute ("DELETE * FROM GO")
      Set rstemp = dbExpressionData.OpenRecordset("SELECT GenMAPP FROM Expression")
      While Not rstemp.EOF
         Set rstemp2 = dbPathFinder.OpenRecordset("Select DISTINCT MAPPNameField FROM [GeneToMAPP]" _
                           & " WHERE GenMAPP = '" & rstemp![GenMAPP] & "'")
            If Not rstemp2.EOF Then
               While Not rstemp2.EOF
                  'GOID = MAPPName
                  dbChipData.Execute ("INSERT INTO GO (Primary, GOID) VALUES " _
                        & "('" & rstemp![GenMAPP] & "', '" & rstemp2![MappNameField] & "')")
                  rstemp2.MoveNext
               Wend
            End If
         rstemp.MoveNext
      Wend
   
      'you now have a table with every one of the GenMAPPs of the expression dataset linked to a MAPP.
      'need to count how many times each MAPP is represented.
      rstemp2.Close
      rstemp.Close
      
      dbChipData.Execute ("INSERT INTO LocalMAPPsChip SELECT First([GOID]) as MAPPName, Count([GOID]) AS" _
                  & " OnChip FROM [GO] GROUP BY [GOID]")
         
   End If
   
   'dbChipData.Execute ("DELETE * FROM TempPrimary")
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


