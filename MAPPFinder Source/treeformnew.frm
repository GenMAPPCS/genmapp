VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form TreeForm 
   BackColor       =   &H00C0FFFF&
   Caption         =   "MAPPFinder Browser"
   ClientHeight    =   10395
   ClientLeft      =   375
   ClientTop       =   735
   ClientWidth     =   13950
   Icon            =   "treeformnew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   13950
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treeformnew.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treeformnew.frx":151C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtStat 
      Height          =   285
      Left            =   10560
      TabIndex        =   19
      Text            =   "0.05"
      Top             =   1440
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exact Match"
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Keyword"
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "treeformnew.frx":216E
      Left            =   6360
      List            =   "treeformnew.frx":2170
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "Select Gene ID type"
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdcollapse 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Collapse Tree"
      Height          =   375
      Left            =   12600
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdGeneSearch 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gene ID Search"
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtgeneID 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton CmdExpand 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Expand Tree"
      Height          =   375
      Left            =   11280
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtNumberChanged 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6840
      TabIndex        =   8
      Text            =   "3"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtPercent 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Text            =   "0"
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton CmdSearchGO 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Word Search"
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtGoTerm 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin MSComctlLib.TreeView TView 
      Height          =   8055
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   14208
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblGOdate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "GO date: "
      Height          =   375
      Left            =   10320
      TabIndex        =   20
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblFile 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   14775
   End
   Begin VB.Label lblColors 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"treeformnew.frx":2172
      Height          =   375
      Left            =   1028
      TabIndex        =   18
      Top             =   10080
      Width           =   12015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Search for a specific Gene ID"
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblstat 
      BackColor       =   &H00C0FFFF&
      Caption         =   "genes changed and an absolute Z score >="
      Height          =   855
      Left            =   7320
      TabIndex        =   15
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "percent of its genes changed and >="
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Expand the tree to show all GO terms with >="
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Sea 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Search for GO term or MAPP"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   480
      Width           =   2175
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu LoadLocalMAPPs 
         Caption         =   "Load Local MAPPs"
      End
      Begin VB.Menu calculateNew 
         Caption         =   "Calculate New Results"
      End
      Begin VB.Menu loadExisting 
         Caption         =   "Load Existing Results"
      End
      Begin VB.Menu exportMAPPs 
         Caption         =   "Export Terms Matching Numerical Filter"
      End
      Begin VB.Menu exportgreenterms 
         Caption         =   "Export Terms Matching Word/Gene ID Search"
      End
      Begin VB.Menu mnuExportTree1 
         Caption         =   "Export Highlighted Portion of Tree as Text"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu calculation 
      Caption         =   "Calculation Summary"
   End
   Begin VB.Menu rankedList 
      Caption         =   "Show Ranked List"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnusetnodeclick 
         Caption         =   "Set Click Response"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu MAPPfinderhelp 
         Caption         =   "MAPPFinder Help"
      End
      Begin VB.Menu whatdocolorsmean 
         Caption         =   "What do the colors mean?"
      End
      Begin VB.Menu whatdonumbersmean 
         Caption         =   "What do the numbers mean?"
      End
      Begin VB.Menu aboutMAPPfinder 
         Caption         =   "About MAPPFinder"
      End
   End
End
Attribute VB_Name = "TreeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const GOLENGTH = 10
Const LABEL_LENGTH = 35
Const YELLOW = 8388607 'rgb(255,255,127)
Const BLUE = 14998707 '179,220,228
Const GREEN = 12380098 '194,231,188
Const WHITE = 16777215 '255,255,255

Public species As String
Public OpenMAPPWhenClicked As Boolean 'if true a node click opens a mapp.
                                      'if false, export the gene list
Public tablename As String 'the name of the table that stores the genesCHangedinME for each GO
Public LocalTablename As String
Dim Fsys As Object
Dim ontology As TextStream
Dim dbMAPPfinder As Database, dbChip As Database, dbchiplocal As Database, dbLocalMAPPs As Database
Public rootnode As Node, LocalRoot As Node 'rootnode = "Gene Ontology" localroot = "base folder for local mapps"
Dim gotable As String, clustersystem As String, labelfield As String, clustercode As String
Dim oldGOID As String, databaselocation As String
Public LocalPath As String, localDate As String
Dim dbChipLocation As String
Dim localMAPPsOK As Boolean, localloaded As Boolean
Dim progress As Long, keyword As Boolean
Dim Statistics As Boolean, ED As String, CS As String, localED As String, localCS As String
Public GODate As String, GOprimary As Boolean, StatisticsLocal As Boolean
Dim godataloaded As Boolean, localdataloaded As Boolean, treeindex As Long
Public gopaths As New Collection
Public localpaths As New Collection

Private Sub cmdGetExpressionData_Click()
   frmInput.Show
End Sub

Public Sub FormLoad()
   On Error GoTo error
   Dim line As String, name As String
   Dim i As Integer, start As Integer, oldstart As Integer, semicolon As Integer
   Dim currentnode As Node, GONode As Node
   Dim rsGOTree As Recordset
   Dim parent As Node
   Dim level As Integer, oldlevel As Integer, difference As Integer
   Dim newnode As Node
   oldGOID = "empty"
   Set Fsys = New FileSystemObject
   continue = True
   Combo1.Text = "Select Gene ID type"
   Option1(0).Value = True
   keyword = True
   OpenMAPPWhenClicked = True
   ImageList1.ListImages.Add , , LoadPicture(App.Path & "\arrow.bmp")
   TView.ImageList = ImageList1
   TView.Nodes.Clear
   
   'add local nodes
   If UCase(Dir(Module1.programpath & "LocalMAPPs_" & species & ".txt")) = UCase("LocalMAPPs_" & species & ".txt") Then
      Set dbLocalMAPPs = OpenDatabase(programpath & "LocalMAPPs_" & species & ".gmf")
      localloaded = True
      Set ontology = Fsys.OpenTextFile(Module1.programpath & "LocalMAPPs_" & species & ".txt")
      line = ontology.ReadLine '
      While InStr(1, line, "!date") = 0
         line = ontology.ReadLine
      Wend
      localDate = line
      While InStr(1, line, "LocalPath") = 0
         line = ontology.ReadLine
      Wend
      LocalPath = Mid(line, 12, Len(line) - 11)
   
      i = 1
      oldstart = 0
      line = ontology.ReadLine 'read the root line
      semicolon = InStr(1, line, ";")
      name = Mid(line, 2, semicolon - start - 3) 'start from 2 to remove the <
   
      Set LocalRoot = AddNode(name, name)
      Set currentnode = LocalRoot
      ReadLineGenMAPP currentnode
   End If
   DoEvents
   
   'add gene ontology root node
   Set currentnode = AddNode("GO", "Gene Ontology")
   Set GONode = currentnode
   Set rootnode = currentnode
   Set rsGOTree = dbMAPPfinder.OpenRecordset("SELECT [Level], [ID], [Name] FROM GeneOntologyTree ORDER BY OrderNO")
   'add the first child by hand
   
   Set newnode = AddChild(currentnode.key, "GO:" & rsGOTree!id, rsGOTree!name)
   Set parent = currentnode
   oldlevel = rsGOTree!level 'this had better be 1
   Set currentnode = newnode
   rsGOTree.MoveNext
   'now parent = the root and currentnode = the first branch (biological process)
   While rsGOTree.EOF = False
   
      If rsGOTree!level > oldlevel Then
         Set newnode = AddChild(currentnode.key, "GO:" & rsGOTree!id, rsGOTree!name)
      ElseIf rsGOTree!level = oldlevel Then
         Set parent = currentnode.parent
         Set newnode = AddChild(parent.key, "GO:" & rsGOTree!id, rsGOTree!name)
      Else
         difference = oldlevel - rsGOTree!level  'this is the number of steps to take up the tree
         For i = 1 To difference + 1
            Set currentnode = currentnode.parent
         Next i
         Set newnode = AddChild(currentnode.key, "GO:" & rsGOTree!id, rsGOTree!name)
      End If
      
      oldlevel = rsGOTree!level
      Set currentnode = newnode
     'Debug.Print currentnode.parent
      rsGOTree.MoveNext
     ' Debug.Print TView.Nodes.count
   Wend
 
   Set rsdate = dbMAPPfinder.OpenRecordset("SELECT Date from GeneOntology")
   frmCriteria.GODate = rsdate!Date 'all three come as unit, so use process as the go date
   lblGOdate.Caption = "GO Date: " & rsdate!Date
   
   
   
   
error:
   Select Case Err.Number
      Case 3024 'the error for not having the database
         MsgBox "The file LocalMAPPs.gmf was not found in the application folder." _
         & " This file was created when you loaded local MAPPs. You will need to reload your" _
         & " local MAPPs.", vbOKOnly
      Case 53
         MsgBox "MAPPFinder can not find the file process_ontology_" & dbDate & ".txt. MAPPFinder thinks" _
         & " the files should be in " & Module1.programpath & "process_ontology_" & dbDate & ".txt. Where do you think it should be? This and the correspond" _
         & "ing component and function ontology files are needed to run MAPPFinder. If you have " _
         & "recently updated your GenMAPP database, you will also need to update the Ontology files" _
         & " that represent the Gene Ontology. These files are available at GenMAPP.org and should be" _
         & " placed in the program files folder (C:\Program Files\GenMAPP 2\ is the default) containing" _
         & " MAPPFinder.exe. You need to go get these files and restart MAPPFinder or select the previous" _
         & " GenMAPP database.", vbOKOnly
         
   End Select
   
   
End Sub

Private Sub loadGOTree(currentnode As Node, currentlevel As Integer, parentnode As Node, _
                        parentlevel As Integer, rsGOTree As Recordset)
'this function will read through the GeneOntologyTree table of the GenMAPP Gene Database
'and build the tree representation of the GO DAG structure.
   Dim newlevel As Integer
   Dim newID As String
   Dim newnode As Node
   
   
                        
                        
End Sub

'this is the code that reads in the geneontology text files. this has been replaced by a database table
'that represents the GO tree structure.
'Private Sub readtextfiles()
' Dim line As String
 'Dim currentnode  As Node
 
 'Set ontology = Fsys.OpenTextFile(Module1.programpath & "process_ontology_" & dbDate & ".txt")
 '  i = 1
 '  For i = 1 To 3
 '     line = ontology.ReadLine '
 '  Next i
 '  frmCriteria.GODate = line 'all three come as unit, so use process as the go date
 '  lblGOdate.Caption = "GO Date: " & Right(line, InStrRev(line, "   ") + 3)
 '  line = ontology.ReadLine
 '  line = ontology.ReadLine
 '  line = ontology.ReadLine
 '  i = 1
  ' oldstart = 0
  ' line = ontology.ReadLine 'read the geneontology line
  '
  ' ReadLine currentnode
  ' DoEvents
  ' Set ontology = Fsys.OpenTextFile(Module1.programpath & "function_ontology_" & dbDate & ".txt")
  ' i = 1
  ' For i = 1 To 6
  '    line = ontology.ReadLine '
  ' Next i
  ' i = 1
  ' oldstart = 0
  ' line = ontology.ReadLine 'read the geneontology line
   'Set currentnode = GONode
   'ReadLine currentnode
   'DoEvents
   'Set ontology = Fsys.OpenTextFile(Module1.programpath & "component_ontology_" & dbDate & ".txt")
   'i = 1
   'For i = 1 To 6
   '   line = ontology.ReadLine '
   'Next i
   'i = 1
   'oldstart = 0
   'line = ontology.ReadLine 'read the geneontology line
   'Set currentnode = GONode
   'ReadLine currentnode
'End  Sub
Public Function AddNode(key As String, name As String) As Node
    Dim newnode As Node
    
    Set newnode = TView.Nodes.Add(, , key, name, 2, 1)
    newnode.Sorted = True
    Set AddNode = newnode
End Function

Public Function AddChild(parent As String, key As String, name As String) As Node
   On Error GoTo errorhandler
   Dim newnode As Node
   Set newnode = TView.Nodes.Add(parent, tvwChild, key, name, 2, 1)
   newnode.Sorted = True
   Set AddChild = newnode
   GoTo aftererror
errorhandler:
   Select Case Err.Number
      Case 35602
         key = key & "I"
         Set AddChild = AddChild(parent, key, name)
   End Select
aftererror:
End Function

Public Sub ReadLine(currentnode As Node) 'accepts the key of the current node
   Dim line As String
   Dim start As Integer, semicolon As Integer, difference As Integer, i As Integer
   Dim key As String, name As String
   Dim newnode As Node, parent As Node
   
   While ontology.AtEndOfStream = False
      line = ontology.ReadLine
      start = InStr(1, line, "<")
      If start = 0 Or (start > InStr(1, line, "%") And InStr(1, line, "%") > 0) Then
         start = InStr(1, line, "%")
      End If
      semicolon = InStr(1, line, ";")
      name = fixName(Mid(line, start + 1, semicolon - start - 2))
      key = Mid(line, semicolon + 2, GOLENGTH)

      If start > oldstart Then
         Set newnode = AddChild(currentnode.key, key, name)
      ElseIf start = oldstart Then
         Set parent = currentnode.parent
         Set newnode = AddChild(parent.key, key, name)
      Else
         difference = oldstart - start 'this is the number of steps to take up the tree
         For i = 1 To difference + 1
            Set currentnode = currentnode.parent
         Next i
         Set newnode = AddChild(currentnode.key, key, name)
      End If
      oldstart = start
      Set currentnode = newnode
   Wend
End Sub
Public Sub ReadLineGenMAPP(currentnode As Node) 'accepts the key of the current node
   Dim line As String
   Dim start As Integer, semicolon As Integer, difference As Integer, i As Integer
   Dim key As String, name As String
   Dim newnode As Node, parent As Node
   
   While ontology.AtEndOfStream = False
      line = ontology.ReadLine
      start = InStr(1, line, "<")
      If start = 0 Or (start > InStr(1, line, "%") And InStr(1, line, "%") > 0) Then
         start = InStr(1, line, "%")
      End If
      semicolon = InStr(1, line, ";")
      name = fixName(Mid(line, start + 1, semicolon - start - 2))
      key = Mid(line, semicolon + 2, Len(line) - semicolon - 1)
      
      If start > oldstart Then
         Set newnode = AddChild(currentnode.key, key, name)
      ElseIf start = oldstart Then
         Set parent = currentnode.parent
         Set newnode = AddChild(parent.key, key, name)
      Else
         difference = oldstart - start 'this is the number of steps to take up the tree
         For i = 1 To difference + 1
            Set currentnode = currentnode.parent
         Next i
        
         Set newnode = AddChild(currentnode.key, key, name)
      End If
      oldstart = start
      Set currentnode = newnode
   Wend
End Sub


Public Function fixName(name As String) As String
   fixName = Replace(name, "\", "")
End Function



Private Sub calculation_Click()
   frmCalculation.Show
End Sub

Private Sub exportgreenterms_Click()
    On Error GoTo error
   MousePointer = vbHourglass
   Dim MAPPFolder As String
   Dim NameEnd As Integer
   Dim MAPPName As String, slash As Integer
   Dim currentnode As Node
   Dim mappbuilderfile As TextStream
   Dim gomapp As Boolean
    gomapp = False
    frmFolderCreator.Cancel = 0
   frmFolderCreator.Show vbModal
   
   If (frmFolderCreator.Cancel = 2) Then
      MAPPFolder = mapploc & species & "\" & frmFolderCreator.txtFolderName.Text & "\"
      
greendirectoryok:
      If Dir(MAPPFolder) = "" Then 'it's a new directory
         MkDir (MAPPFolder)
      Else 'the user wants to overwrite an existing directory
         ClearDirectory (MAPPFolder)
      End If
         
      Set mappbuilderfile = Fsys.CreateTextFile(MAPPFolder & "MAPPFinderTempMAPPBuilder.txt")
      mappbuilderfile.WriteLine ("geneId" & Chr(9) & "systemcode" & Chr(9) & "Label" _
                                 & Chr(9) & "Head" & Chr(9) & "Remarks" & Chr(9) & "MappName")
      
      For Each Node In TView.Nodes
         Set currentnode = Node
         If (currentnode.BackColor = BLUE Or currentnode.BackColor = GREEN) Then 'it meets the filter export it
            If InStr(1, Node.FullPath, "Gene Ontology") <> 0 Then 'it's a GO MAPP do the GO routine
                gomapp = True
                If InStr(1, currentnode.key, "I") = 0 Then 'this is the second or more occurence of this term, ignore it
                  
                  NameEnd = InStr(1, Node.Text, "     ")
                  If NameEnd = 0 Then 'nothing has been added, ie. this node has no MAPPFinder data
                     MAPPName = fixName(currentnode.Text)
                  Else
                     MAPPName = fixName(Mid(Node.Text, 1, NameEnd - 1))
                  End If
                  BuildMAPPs currentnode, MAPPName, mappbuilderfile
               End If
            Else 'local mapp
               NameEnd = InStr(1, Node.FullPath, "     ")
               If NameEnd = 0 Then 'nothing has been added, ie. this node has no MAPPFinder data
                  slash = InStr(1, Node.FullPath, "\")
                  MAPPName = Mid(Node.FullPath, slash, Len(Node.FullPath) - slash + 1)
               Else
                 slash = InStr(1, Node.FullPath, "\")
                  MAPPName = Mid(Node.FullPath, slash, NameEnd - slash)
               End If
              Fsys.CopyFile LocalPath & MAPPName & ".mapp", MAPPFolder, True
            End If
            DoEvents
         End If
         
     Next Node
  
      'now we have copied the local MAPPs and have a csv file ready to build the GO MAPPs
    If gomapp Then
        MappBuilderForm_Normal.setBaseMapp MAPPFolder, databaselocation
        MappBuilderForm_Normal.setFileName MAPPFolder & "MAPPFinderTempMAPPBuilder.txt"
        MappBuilderForm_Normal.MakeMapps_Click
      
        mappbuilderfile.Close
    End If
      frmFolderCreator.txtFolderName = ""
      Kill MAPPFolder & "MAPPFinderTempMAPPBuilder.txt"
   End If
   MousePointer = vbDefault
   Exit Sub
error:
   Select Case Err.Number
      Case 76
         MkDir (mapploc & species & "\")
         Resume greendirectoryok
   End Select
   MousePointer = vbDefault
End Sub



Private Sub exportMAPPs_Click()
   On Error GoTo error
   MousePointer = vbHourglass
   Dim MAPPFolder As String
   Dim NameEnd As Integer
   Dim MAPPName As String, slash As Integer
   Dim currentnode As Node
   Dim mappbuilderfile As TextStream
   Dim gomapp As Boolean
    gomapp = False
    frmFolderCreator.Cancel = 0
   frmFolderCreator.Show vbModal
   
   If (frmFolderCreator.Cancel = 2) Then
      MAPPFolder = mapploc & species & "\" & frmFolderCreator.txtFolderName.Text & "\"
      
directoryok:
      If Dir(MAPPFolder) = "" Then 'it's a new directory
         MkDir (MAPPFolder)
      Else 'the user wants to overwrite an existing directory
         ClearDirectory (MAPPFolder)
      End If
         
      Set mappbuilderfile = Fsys.CreateTextFile(MAPPFolder & "MAPPFinderTempMAPPBuilder.txt")
      mappbuilderfile.WriteLine ("geneId" & Chr(9) & "systemcode" & Chr(9) & "Label" _
                                 & Chr(9) & "Head" & Chr(9) & "Remarks" & Chr(9) & "MappName")
      
      For Each Node In TView.Nodes
         Set currentnode = Node
         If (currentnode.BackColor = YELLOW Or currentnode.BackColor = GREEN) Then 'it meets the filter export it
            If InStr(1, Node.FullPath, "Gene Ontology") <> 0 Then 'it's a GO MAPP do the GO routine
                gomapp = True
                If InStr(1, currentnode.key, "I") = 0 Then 'this is the second or more occurence of this term, ignore it
                  
                  NameEnd = InStr(1, Node.Text, "     ")
                  If NameEnd = 0 Then 'nothing has been added, ie. this node has no MAPPFinder data
                     MAPPName = fixName(currentnode.Text)
                  Else
                     MAPPName = fixName(Mid(Node.Text, 1, NameEnd - 1))
                  End If
                  BuildMAPPs currentnode, MAPPName, mappbuilderfile
               End If
            Else 'local mapp
               NameEnd = InStr(1, Node.FullPath, "     ")
               If NameEnd = 0 Then 'nothing has been added, ie. this node has no MAPPFinder data
                  slash = InStr(1, Node.FullPath, "\")
                  MAPPName = Mid(Node.FullPath, slash, Len(Node.FullPath) - slash + 1)
               Else
                 slash = InStr(1, Node.FullPath, "\")
                  MAPPName = Mid(Node.FullPath, slash, NameEnd - slash)
               End If
              Fsys.CopyFile LocalPath & MAPPName & ".mapp", MAPPFolder, True
            End If
            DoEvents
         End If
         
     Next Node
  
      'now we have copied the local MAPPs and have a csv file ready to build the GO MAPPs
    If gomapp Then
        MappBuilderForm_Normal.setBaseMapp MAPPFolder, databaselocation
        MappBuilderForm_Normal.setFileName MAPPFolder & "MAPPFinderTempMAPPBuilder.txt"
        MappBuilderForm_Normal.MakeMapps_Click
      
        mappbuilderfile.Close
    End If
      frmFolderCreator.txtFolderName = ""
      Kill MAPPFolder & "MAPPFinderTempMAPPBuilder.txt"
   End If
   MousePointer = vbDefault
   Exit Sub
error:
   Select Case Err.Number
      Case 76
         MkDir (mapploc & species & "\")
         Resume directoryok
   End Select
   MousePointer = vbDefault
End Sub


Private Sub Form_Resize()
   If Me.Width > 735 Then
      TView.Width = Me.Width - 735 'offset to give it a border
      lblColors.Left = (Me.Width - lblColors.Width) / 2
   End If
   If Me.Height > 3016 Then
      TView.Height = Me.Height - 3015
      lblColors.Top = Me.Height - 990
   End If
   TView.Refresh
   DoEvents
End Sub


Private Sub mnuExportTree1_Click()
   On Error GoTo error
   MousePointer = vbHourglass
   Dim FileName As String
   Dim Fsys As New FileSystemObject
   Dim output As TextStream, outputlocal As TextStream
   frmCriteria.CommonDialog1.FileName = ""
   frmCriteria.CommonDialog1.Filter = "Text Files|*.txt"
   frmCriteria.CommonDialog1.ShowSave
   FileName = frmCriteria.CommonDialog1.FileName
   treeindex = 1
   If FileName <> "" Then
      If invalidFileName(FileName) Then
         MsgBox "A filename cannot contain any of the following characters: /\:*?" & Chr(34) & "<>| are not", vbOKOnly
         Exit Sub
      End If
      If Dir(FileName) <> "" Then
         If MsgBox("Overwrite the existing " & FileName & "?", vbOKCancel) = vbCancel Then
            Exit Sub
         End If
      End If
      'export list
      If godataloaded Then
         Set output = Fsys.CreateTextFile(FileName)
      
         output.WriteLine ("Index" & Chr(9) & "Path" & Chr(9) & "GOID" & Chr(9) & "GO Name" & Chr(9) & "GO Type" & Chr(9) & "Number Changed Local" & Chr(9) _
                     & "Number Measured Local" & Chr(9) & "Number in GO Local" & Chr(9) & "Percent Changed Local" _
                     & Chr(9) & "Percent Present Local" & Chr(9) & "Number Changed" _
                     & Chr(9) & "Number Measured" & Chr(9) & "Number in GO" _
                     & Chr(9) & "Percent Changed" & Chr(9) & "Percent Present" & Chr(9) & "Z Score" _
                     & Chr(9) & "PermuteP" & Chr(9) & "AdjustedP")
      
      
         ExportGONode rootnode, 0, output
         output.Close
      End If
      DoEvents
      treeindex = 1
      If localdataloaded Then
         FileName = Left(FileName, Len(FileName) - 4)
         FileName = FileName & "_Local.txt"
         Set outputlocal = Fsys.CreateTextFile(FileName)
         
         outputlocal.WriteLine ("Index" & Chr(9) & "Path" & Chr(9) & "MAPP Name" & Chr(9) & "Number Changed" _
                     & Chr(9) & "Number Measured" & Chr(9) & "Number On MAPP" _
                     & Chr(9) & "Percent Changed" & Chr(9) & "Percent Present" & Chr(9) & "Z Score" _
                     & Chr(9) & "PermuteP" & Chr(9) & "AdjustedP")
                     
         ExportLocalNode LocalRoot, 0, outputlocal
         outputlocal.Close
      End If
      
      
      
   End If
   
error:
   Select Case Err.Number
      Case 70
         MsgBox "Permission Denied. Perhaps you have the file " & FileName & " open in another program (e.g. Excel)?", vbOKOnly
   End Select
   MousePointer = vbDefault
   
End Sub

Private Sub mnusetnodeclick_Click()
  ' frmNodeClickResponse.Load
   frmNodeClickResponse.Show vbModal
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         keyword = True
      Case 1
         keyword = False
   End Select
End Sub


'called by nestedchipdata, call onchip.
'will build the nested chip table that contains the number of genes on the chip in each go term (hierarchical)
Public Sub CountChipGenes(currentnode As Node, ByRef dbChipData As Database)
   Dim chip As Integer, ingo As Long
   Dim GOID As String
   Dim rstemp As DAO.Recordset
   GOID = currentnode.key
   If GOID <> "GO" Then
      GOID = Mid(GOID, 4, GOLENGTH - 3) 'remove GO:
   End If
   dbChipData.Execute ("DELETE * from nestedgo")
   'onChip currentnode, dbChipData
   Set rstemp = dbChipData.OpenRecordset("SELECT * FROM nestedGO")
   rstemp.MoveLast
   chip = rstemp.RecordCount
   dbChipData.Execute ("INSERT INTO NestedChip (GOID, OnchipNested) VALUES ('" & GOID & "', " & chip & ")")
   
End Sub

'builds the Nested Chip table. Traverses the tree, counting genes on the chip for each node.

Public Function NestedChipData(dbChipData As Database, currentnode As Node, parentcounter() As String _
               , parentgenes As Long) As Long  'calculates the chip data for parent and child.
   Dim i As Long, count As Long
   Dim GOID As String
   Dim childnd As Node, gotype As String, rsGOType As DAO.Recordset
   Dim rsCount As DAO.Recordset, metcriteria As Integer, rsinGONested As DAO.Recordset
   Dim genecounter() As String
   Dim genes As Long, arraysize As Long
   
   progress = progress + 1
   If progress Mod 10 = 0 Then
      frmCriteria.lblProgress.Caption = "Genes associated with the GO nodes calculated for " & progress & " nodes out of " & TView.Nodes.count & "."
      frmCriteria.Refresh
   End If
   
   genes = 0
   GOID = currentnode.key
   If GOID <> "GO" Then
      GOID = Mid(GOID, 4, GOLENGTH - 3) 'remove GO:
      'If GOID = "0019936" Then
         'Debug.Print "GOID = 0019936"
      'End If
   End If
   Set rsinGONested = dbMAPPfinder.OpenRecordset("SeLECT GOIDCount FROM " & species & "HierarchyCount Where GOID = '" & GOID & "'")
   'there are a lot of duplicate gene associations in yeast (see GOID 0006913 and its children)
   'so yeast gets a bigger array per node
   If species = "yeast" Then
      arraysize = rsinGONested![GOIDcount] * (currentnode.children + 5)
   Else
      arraysize = rsinGONested![GOIDcount] * (currentnode.children + 3)
   End If
   'add one for those with no children
   ReDim genecounter(arraysize) As String 'you now have an array twice as big as the number of genes in this sub-graph. Allows for plenty of duplicates.
   Set childnd = currentnode.Child
   For i = 0 To currentnode.children - 1
      'count genes in children nodes
      genes = genes + NestedChipData(dbChipData, childnd, genecounter, genes)
      Set childnd = childnd.Next
   
   Next i
   'no more children -> so genecounter now has all of the children's genes in it.
   'add this node's genes
   'then add all of the genes in genecounter to parentcounter
   Set rsCount = dbChipData.OpenRecordset("Select Primary From GO WHERE GOID = '" & GOID & "'")
   If rsCount.EOF = False Then
      rsCount.MoveLast
      rsCount.MoveFirst
      count = rsCount.RecordCount
   Else
      count = 0 'no genes are directly associated with this node
   End If
   For i = genes To genes + count - 1
      genecounter(i) = rsCount![primary]
      'Debug.Print genecounter(i)
      rsCount.MoveNext
   Next i
   'genecounter now has all of its children's genes and it's own genes
   genes = genes + count 'the total number of genes now in genecounter
   'now add all of these genes to the parent's array
   For i = 0 To genes - 1
          parentcounter(i + parentgenes) = genecounter(i)
   Next i
   'now you've added all of the genes in this node and it's children to the parent node
   'now we must sort and remove duplicates
   If genes = 0 Then
      metcriteria = 0
   ElseIf genes = 1 Then
      metcriteria = 1
   Else 'more than 1 make sure they're unique
      metcriteria = unique(genecounter, genes)
   End If
   If GOID = "root" Then
      GOID = "GO"
   End If
   dbChipData.Execute ("INSERT INTO NestedChip (GOID, OnchipNested) VALUES ('" & GOID _
                              & "', " & metcriteria & ")")
   
   NestedChipData = genes
   'Debug.Print currentnode.Text & "calculated. Genes ar:"; e
   'For i = 0 To genes
      'Debug.Print genecounter(i)
   'Next i
  
End Function

Public Sub nestedResults(GOterms As Collection, currentnode As Node)
   'input : both databases, the currentnode, and the current growing array
   'output : the number of genes added
   Dim i As Long
   Dim GOID As String, parent As Node
   Dim percentage As Single, present As Single
   Dim childnd As Node, gotype As String, rsGOType As DAO.Recordset
   Dim rsCount As DAO.Recordset, metcriteria As Integer, rsinGONested As DAO.Recordset
   Dim nestedpercentage As Single, nestedpresent As Single
   Dim genecounter() As String, count As Long
   Dim genes As New Collection, currentTerm As goterm
   Dim rsGOcount As DAO.Recordset, rsChipCount As DAO.Recordset
   progress = progress + 1
   If progress Mod 10 = 0 Then
      frmCriteria.lblProgress.Caption = "Results calculated for " & progress & " nodes out of " & TView.Nodes.count & "."
      frmCriteria.Refresh
   End If
   
   GOID = currentnode.key
   If GOID <> "GO" Then
      GOID = Mid(GOID, 4, GOLENGTH - 3) 'remove GO:
  
      Set currentTerm = GOterms.Item(GOID)
      If currentTerm.visited = False Then 'the dag makes it possible that this term was already
                                          'hit once. Don't bother traversing this branch of the
                                          'tree again
         Set childnd = currentnode.Child
         For i = 0 To currentnode.children - 1
         'the the child's genes to this node's collection of genes
            nestedResults GOterms, childnd
            Set childnd = childnd.Next

         Next i
         'no more children -> so now we add all of this nodes genes and genes changed to its parent
      
         Set parentnode = currentnode.parent
         Set parentterm = GOterms.Item(parentnode.key)
         'at some point the nodes in the tree should store all of this information, but that's
         'too much for now. I don't know how to extend a VB class. Can it be done?
         Set genes = currentTerm.getGenes()
         currentTerm.setOnChip (genes.count)
         For Each gene In genes
            parentterm.addGene (gene)
         Next gene
         
         Set genes = currentTerm.getChangedGenes()
         currentTerm.setChanged (genes.count)
         For Each gene In genes
            parentterm.addChangedGene (gene)
         Next gene
         currentTerm.setvisited
      End If
      
  Else 'this is the root node. Only visit the children.
      Set currentTerm = GOterms.Item(GOID)
      Set childnd = currentnode.Child
      For i = 0 To currentnode.children - 1
      'the the child's genes to this node's collection of genes
         nestedResults GOterms, childnd
         Set childnd = childnd.Next

      Next i
   End If
      
   'calculate all results
   'save them in the node
   currentTerm.calculateResults

End Sub

Public Sub CountCriteriaGenes(currentnode As Node, dbChipData As Database, dbExpressionData As Database)
   Dim metcriteria As Integer
   Dim GOID As String
   Dim rsResults As DAO.Recordset, rsNestedGO As DAO.Recordset
   Dim rsNestedChip As DAO.Recordset, rsinGONested As DAO.Recordset
   Dim nestedpercentage As Single, nestedpresent As Single
   If GOID <> "GO" Then
      GOID = Mid(GOID, 4, GOLENGTH - 3) 'remove GO:
   End If
   dbExpressionData.Execute ("DELETE * FROM NestedGO")
   countChangedGenes currentnode, dbExpressionData
   Set rsResults = dbExpressionData.OpenRecordset("Select * FROM Results WHERE GOID = '" & GOID & "'")
   Set rsNestedGO = dbExpressionData.OpenRecordset("SELECT * from nestedgo")
   Set rsNestedChip = dbChipData.OpenRecordset("SELECT GOID, OnChipNested FROM NestedChip" _
                                             & " WHERE GOID = '" & GOID & "'")
   Set rsinGONested = dbMAPPfinder.OpenRecordset("SELECT GOCount FROM " & species & "HierarchyCount" _
                                                & " WHERE GOID = '" & GOID & "'")
   rsNestedGO.MoveLast
   metcriteria = rsNestedGO.RecordCount
   nestedpercentage = metcriteria / rsNestedChip![onchipnested]
   nestedpresent = rsNestedChip![onchipnested] / rsinGONested![GOCount]
   
   
   dbExpressionData.Execute ("INSERT INTO NestedResults (GOID, Name, indata, Onchip, inGO, " _
                           & "Percentage, Present, InDataNested, OnChipNested, InGONested, " _
                           & "PercentageNested, PresentNested) VALUES ('" & GOID & "', '" _
                           & TextToSql(currentnode.Text) & "', " & rsResults![indata] & ", " & rsResults![onChip] _
                           & ", " & rsResults![ingo] & ", " & rsResults![percentage] & ", " _
                           & rsResults![present] & ", " & metcriteria & ", " & rsNestedChip![onchipnested] _
                           & ", " & rsinGONested![GOCount] & ", " & nestedpercentage & ", " & nestedpresent _
                           & ")")
               
End Sub

Public Sub countChangedGenes(currentnode As Node, dbExpressionData As Database)
   Dim rsGeneNumber As DAO.Recordset
   Dim GOID As String
   Dim i As Integer
   Dim childnd As Node
   
   GOID = currentnode.key
   If Len(GOID) <= GOLENGTH Then 'this is the first occurence of the GOID
      GOID = Mid(GOID, 4, GOLENGTH - 3)
      Set rsGeneNumber = dbExpressionData.OpenRecordset("Select Primary FROM GO WHERE " _
                         & "GOID = '" & GOID & "'")
      While rsGeneNumber.EOF = False
         dbExpressionData.Execute ("INSERT INTO NestedGO (Primary, GOID) VALUES ('" & _
                           rsGeneNumber![primary] & "', '" & GOID & "')")
         rsGeneNumber.MoveNext
      Wend
         
      If currentnode.children > 0 Then
         Set childnd = currentnode.Child
         For i = 0 To currentnode.children - 1
            countChangedGenes childnd, dbExpressionData
            Set childnd = childnd.Next
         Next i
      End If
   End If
End Sub


Public Function root() As Node
   Set root = rootnode
End Function

Public Function GetLocalRoot() As Node
   Set GetLocalRoot = LocalRoot
End Function

Public Sub createNestedGOTable(dbGO As Database, currentnode As Node)
   Dim tblNestedGO As TableDef
   Dim GOID As String
   
   If currentnode.key = "GO" Then
      GOID = "root"
   Else
      GOID = Mid(currentnode.key, 4, GOLENGTH - 3)
   End If
   
   Set tblNestedGO = dbGO.CreateTableDef(GOID)
      With tblNestedGO
         .Fields.Append .CreateField("Primary", dbText, 30)
         Dim idxNestedPrimary As Index
         Set idxNestedPrimary = .CreateIndex("idx" & GOID)
         idxNestedPrimary.Fields.Append .CreateField("Primary", dbText, 30)
         idxNestedPrimary.primary = True
         .Indexes.Append idxNestedPrimary
         .Fields.Append .CreateField("GOID", dbText, 30)
      End With
   dbGO.TableDefs.Append tblNestedGO
   
End Sub

Public Sub setcolor(ByRef currentnode As Node, percentage As Single)
   
   If (percentage >= 5 And percentage < 15) Then
         currentnode.ForeColor = RGB(129, 23, 136)
   ElseIf (percentage >= 15 And percentage < 25) Then
         currentnode.ForeColor = RGB(10, 80, 161)
   ElseIf (percentage >= 25 And percentage < 35) Then
         currentnode.ForeColor = RGB(103, 198, 221)
   ElseIf (percentage >= 35 And percentage < 45) Then
         currentnode.ForeColor = RGB(102, 187, 80)
   ElseIf (percentage >= 45 And percentage < 55) Then
         currentnode.ForeColor = RGB(255, 127, 0)
   ElseIf (percentage >= 55) Then
         currentnode.ForeColor = RGB(255, 0, 0)
   End If
   
End Sub

Public Sub LoadFiles(GOFileName As String, LocalName As String)
   On Error GoTo error
   Dim i As Integer, count As Integer
   Dim resultsfile As TextStream
   Dim line As String, GONode As Node
   Dim GOID As String, GOName As String, gotype As String
   Dim indata As Integer, onChip As Integer, ingo As Integer
   Dim percentage As Single, present As Single
   Dim indatanested As Integer, onchipnested As Integer, ingonested As Integer
   Dim percentagenested As Single, presentnested As Single
   Dim permutep As Double, adjustedp As Double
   Dim tab1 As Long, tab2 As Long, tab3 As Long
   Dim abridgedGOname As String, abridgedLocalName As String, dbset As Boolean, GOfile As Boolean
   Dim goterm As goterm, genetoGOFile As String, openingLocal As Boolean
   Dim rsMOD As Recordset
   openingLocal = False
   localMAPPsOK = False
   GOfile = False
   localdataloaded = False
   godataloaded = False
   If GOFileName <> "" Then
   If Dir(GOFileName) = "" Then
      MsgBox "No GO results file of that name exists. Please select a different file.", vbOKOnly
      GoTo endfile
   Else
      GOfile = True
      'MakeGoPaths rootnode, "0"
      Set resultsfile = Fsys.OpenTextFile(GOFileName)
      line = resultsfile.ReadLine
      While InStr(1, line, "File:") = 0 And resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
      Wend
      ED = Mid(line, 7, Len(line) - 6)
      If Dir(ED) = "" Then 'this expression dataset does not exist on this computer
         MsgBox "The GenMAPP Expression Dataset used to calculate these results, " _
            & ED & ", can not be found on this computer. Please select the GEX file" _
            & " used to calculate these results.", vbOKOnly
         CommonDialog1.FileName = ""
         CommonDialog1.Filter = "GenMAPP Expression Dataset|*.gex"
         CommonDialog1.ShowOpen
         CommonDialog1.CancelError = False
         If CommonDialog1.FileName = "" Then
            MsgBox "You did not select a database. Please do so."
            Resume tryagain2
         End If
         ED = CommonDialog1.FileName
      End If
      'Debug.Print ED
      abridgedGOname = Right(GOFileName, Len(GOFileName) - InStrRev(GOFileName, "\"))
      'abridgedGOname = filename-criterion-GO.txt
      
      dbChipLocation = Left(ED, Len(ED) - 3) & "gmf" 'everything but the gex
      line = resultsfile.ReadLine 'line = table: tablename
      tablename = Mid(line, 8, Len(line) - 7) 'subtract .txt
opendbchip:
      
      Set dbChip = OpenDatabase(dbChipLocation)
      
       godataloaded = True
      line = resultsfile.ReadLine
      databaselocation = Mid(line, 11, Len(line) - 10 + 1)
      If databaselocation <> frmStart.lblDB.Caption Then
         If MsgBox("The database used to calculate the results, " & databaselocation & ", is not the gene database you " _
            & "currently have loaded. Should MAPPFinder continue with the currently loaded database, " _
            & frmStart.lblDB.Caption & "?", vbYesNo) = vbNo Then
            MsgBox "Please select the correct database.", vbOKOnly
tryagain2:
            CommonDialog1.FileName = ""
            CommonDialog1.Filter = "GenMAPP Gene Databasee|*.gdb"
            CommonDialog1.ShowOpen
            CommonDialog1.CancelError = False
            If CommonDialog1.FileName = "" Then
               MsgBox "You did not select a database. Please do so."
               Resume tryagain2
            End If
            databaselocation = CommonDialog1.FileName
            Dim rsdate As Recordset
            Set dbMAPPfinder = OpenDatabase(databaselocation)
            setSpecies frmLoadFiles.lblspecies.Caption
            Set rsdate = dbMAPPfinder.OpenRecordset("SELECT version FROM info")
            If dbDate <> rsdate!Version Then
               dbDate = rsdate!Version
               TreeForm.FormLoad 'need to reload the treeform with the correct ontology files
            End If
         Else 'set databaselocation as currently loaded database
            databaselocation = frmStart.lblDB.Caption
         End If
      End If
    
      DoEvents
      Set rsMOD = dbMAPPfinder.OpenRecordset("SELECT modsystem FROM INFO")
      'frmCalculation.lblMODgenes.Caption = rsMOD!modsystem & " genes linked to probes"
      'frmCalculation.lblgeneingo.Caption = rsMOD!modsystem & " genes in GO"
      
      
      frmRank.setDB dbMAPPfinder
      dbset = True
      line = resultsfile.ReadLine
      CS = line
      
      line = resultsfile.ReadLine
      If StrComp(line, frmCriteria.GODate) <> 0 Then 'the string are not equal
         MsgBox "These Gene Ontology results are based on a different version of the GO than is currently loaded." _
               & " If you have updated the Ontology" _
               & " files and the gene database, you should recalculate the results to maintain" _
               & " data consistency. MAPPFinder will display the out of date data, but it is strongly" _
               & " recommended that you recalculate these results.", vbOKOnly
      End If
      line = resultsfile.ReadLine
      species = line
      line = resultsfile.ReadLine
      If InStr(1, line, "true") > 0 Then 'pvalue = true
         Statistics = True
         lblstat.Caption = "genes changed and a p value <"
         txtStat.Text = "0.05"
         txtStat.Left = 9910
      Else
         Statistics = False
         lblstat.Caption = "genes changed and a absolute Z score >= "
         txtStat.Text = "2"
         txtStat.Left = 10560
      End If
      'read through all of the header lines and the stats.
      While InStr(1, line, "Calculation Summary:") = 0 And resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
      Wend
      line = resultsfile.ReadLine
      tab1 = InStr(1, line, " ")
      frmCalculation.lblprobeC.Caption = Left(line, tab1 - 1)
      line = resultsfile.ReadLine
      tab1 = InStr(1, line, " ")
      frmCalculation.lblnoClusterC.Caption = Left(line, tab1 - 1)
      'old files have # of mgi genes met the criterion
      'we took that out, so if you read that line, you should skip it and read the next one.
      line = resultsfile.ReadLine
      If (InStr(1, line, "genes met the criterion")) Then
        line = resultsfile.ReadLine
      End If
        'tab1 = InStr(1, line, " ")
        'frmCalculation.lblGenesC.Caption = Left(line, tab1 - 1)
        'line = resultsfile.ReadLine
      tab1 = InStr(1, line, " ")
      frmCalculation.lblGenesinGOC.Caption = Left(line, tab1 - 1)
      line = resultsfile.ReadLine
       tab1 = InStr(1, line, " ")
      frmCalculation.lblprobeE.Caption = Left(line, tab1 - 1)
      line = resultsfile.ReadLine
       tab1 = InStr(1, line, " ")
      frmCalculation.lblnoClusterE.Caption = Left(line, tab1 - 1)
      line = resultsfile.ReadLine
       If (InStr(1, line, "in this dataset")) Then
            line = resultsfile.ReadLine
        End If
        'tab1 = InStr(1, line, " ")
      'frmCalculation.lblGenesE.Caption = Left(line, tab1 - 1)
      'line = resultsfile.ReadLine
       tab1 = InStr(1, line, " ")
      frmCalculation.lblGenesinGOE.Caption = Left(line, tab1 - 1)
      
      
      While (InStr(1, line, "GOID" & Chr(9) & "GO Name" & Chr(9) & "GO Type" _
                        & Chr(9) & "Number Changed Local" & Chr(9) _
                        & "Number Measured Local" & Chr(9) & "Number in GO Local" & Chr(9) & "Percent Changed Local" _
                        & Chr(9) & "Percent Present Local" & Chr(9) & "Number Changed" _
                        & Chr(9) & "Number Measured" & Chr(9) & "Number in GO" _
                        & Chr(9) & "Percent Changed" & Chr(9) & "Percent Present" & Chr(9) & "Z Score" _
                        & Chr(9) & "PermuteP" & Chr(9) & "AdjustedP") = 0) _
                        And resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
      Wend
   ' line now equals the header line of the results file
      
      While resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
         tab1 = InStr(1, line, Chr(9))
         GOID = Mid(line, 1, tab1 - 1)
         tab2 = InStr(tab1 + 1, line, Chr(9))
         GOName = Mid(line, tab1 + 1, tab2 - tab1 - 1)
         tab3 = InStr(tab2 + 1, line, Chr(9))
         gotype = Mid(line, tab2 + 1, tab3 - tab2 - 1)
         tab1 = InStr(tab3 + 1, line, Chr(9))
         indata = Val(Mid(line, tab3 + 1, tab1 - tab3 - 1))
         tab2 = InStr(tab1 + 1, line, Chr(9))
         onChip = Val(Mid(line, tab1 + 1, tab2 - tab1 - 1))
         tab1 = InStr(tab2 + 1, line, Chr(9))
         ingo = Val(Mid(line, tab2 + 1, tab1 - tab2 - 1))
         tab2 = InStr(tab1 + 1, line, Chr(9))
         percentage = Val(Mid(line, tab1 + 1, tab2 - tab1 - 1))
         percentage = Round(percentage, 1)
         tab1 = InStr(tab2 + 1, line, Chr(9))
         present = Val(Mid(line, tab2 + 1, tab1 - tab2 - 1))
         present = Round(present, 1)
         tab2 = InStr(tab1 + 1, line, Chr(9))
         indatanested = Val(Mid(line, tab1 + 1, tab2 - tab1 - 1))
         tab1 = InStr(tab2 + 1, line, Chr(9))
         onchipnested = Val(Mid(line, tab2 + 1, tab1 - tab2 - 1))
         tab2 = InStr(tab1 + 1, line, Chr(9))
         ingonested = Val(Mid(line, tab1 + 1, tab2 - tab1 - 1))
         tab1 = InStr(tab2 + 1, line, Chr(9))
         percentagenested = Val(Mid(line, tab2 + 1, tab1 - tab2 - 1))
         percentagenested = Round(percentagenested, 1)
         tab2 = InStr(tab1 + 1, line, Chr(9))
         presentnested = Val(Mid(line, tab1 + 1, tab2 - tab1))
         presentnested = Round(presentnested, 1)
         tab1 = InStr(tab2 + 1, line, Chr(9))
         teststat = Val(Mid(line, tab2 + 1, tab1 - tab2 - 1))
         tab2 = InStr(tab1 + 1, line, Chr(9))
         permutep = Val(Mid(line, tab1 + 1, tab2 - tab1 - 1))
         tab1 = Len(line)
         adjustedp = Val(Right(line, tab1 - tab2))
         If Statistics Then 'display p values
            If GOID <> "GO" Then
               Set rsCount = dbMAPPfinder.OpenRecordset("SELECT Count FROM GeneOntologyCount WHERE ID = '" & GOID & "'")
               count = rsCount![count]
               GOID = "GO:" & GOID
               Set GONode = TView.Nodes.Item(GOID)
               GONode.Text = GOName & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                      & " " & present & "%   NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat _
                      & " permute p = " & permutep & " adjusted p = " & adjustedp
               setcolor GONode, percentagenested
               frmRank.lstGO.AddItem (GOName & "     " & "NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat _
                      & " permute p = " & permutep & " adjusted p = " & adjustedp)
             
               For i = 2 To count
                  GOID = GOID & "I"
                  Set GONode = TView.Nodes.Item(GOID)
                  GONode.Text = GOName & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                      & " " & present & "%   NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat _
                      & " permute p = " & permutep & " adjusted p = " & adjustedp
                  setcolor GONode, percentagenested
               Next i
            Else 'GO node
                Set GONode = TView.Nodes.Item(GOID)
                GONode.Text = GOName & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                      & " " & present & "%   NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat _
                      & " permute p = " & permutep & " adjusted p = " & adjustedp
                setcolor GONode, percentagenested
                frmRank.lstGO.AddItem (GOName & "     " & "NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat _
                      & " permute p = " & permutep & " adjusted p = " & adjustedp)
      
            End If
         Else ' no p values
            If GOID <> "GO" Then
               Set rsCount = dbMAPPfinder.OpenRecordset("SELECT COUNT FROM GeneOntologyCount WHERE ID = '" & GOID & "'")
               count = rsCount![count]
               GOID = "GO:" & GOID
               Set GONode = TView.Nodes.Item(GOID)
               GONode.Text = GOName & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                      & " " & present & "%   NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat
               setcolor GONode, percentagenested
               frmRank.lstGO.AddItem (GOName & "     " & "NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat)
             
               For i = 2 To count
                  GOID = GOID & "I"
                  Set GONode = TView.Nodes.Item(GOID)
                  GONode.Text = GOName & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                      & " " & present & "%   NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat
                  setcolor GONode, percentagenested
               Next i
            Else 'GO node
                Set GONode = TView.Nodes.Item(GOID)
                GONode.Text = GOName & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                      & " " & present & "%   NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat
                setcolor GONode, percentagenested
                frmRank.lstGO.AddItem (GOName & "     " & "NESTED " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
                      & onchipnested & "/" & ingonested & " " & presentnested & "% z score = " & teststat)
      
            End If
         End If
NextTerm:
      Wend
     
      TView.Refresh
      GOfile = False
   End If
   End If
   If LocalName <> "" Then
   If Dir(LocalName) = "" Then
      MsgBox "The local results file you select does not exist. Please select another file.", vbOKOnly
      GoTo endfile
   Else
      'MakeLocalPaths LocalRoot, "0"
      openingLocal = True
      abridgedLocalName = Right(LocalName, Len(LocalName) - InStrRev(LocalName, "\"))
      localMAPPsOK = True
      Set resultsfile = Fsys.OpenTextFile(LocalName)
      While InStr(1, line, "File:") = 0 And resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
      Wend
      localED = Mid(line, 7, Len(line) - 6)

      If Dir(localED) = "" Then 'this expression dataset does not exist on this computer
         MsgBox "The GenMAPP Expression Dataset used to calculate these results, " _
            & localED & ", can not be found on this computer. Please select the GEX file" _
            & " used to calculate these results.", vbOKOnly
         CommonDialog1.FileName = ""
         CommonDialog1.Filter = "GenMAPP Expression Dataset|*.gex"
         CommonDialog1.ShowOpen
         CommonDialog1.CancelError = False
         If CommonDialog1.FileName = "" Then
            MsgBox "You did not select a database. Please do so."
            Resume tryagain2
         End If
         localED = CommonDialog1.FileName
      End If
      localdbChipLocation = Left(localED, Len(localED) - 4) & "-Local.gmf" 'everything but the file: and gex
      line = resultsfile.ReadLine 'line = table: tablename
      LocalTablename = Mid(line, 8, Len(line) - 7)
openLocaldbChip:
      Set dbchiplocal = OpenDatabase(localdbChipLocation)
       localdataloaded = True
      
      
      
      If dbset = False Then 'need to open database
         line = resultsfile.ReadLine
         While InStr(1, line, "Database:") = 0 And resultsfile.AtEndOfStream = False
            line = resultsfile.ReadLine
         Wend
         databaselocation = Mid(line, 11, Len(line) - 10 + 1)
         If StrComp(databaselocation, frmStart.lblDB.Caption) <> 0 Then
            If MsgBox("The database used to calculate the results, " & databaselocation & ", is not the gene database you " _
               & "currently have loaded. Is " & frmStart.lblDB.Caption & " the correct database?.", vbYesNo) = vbNo Then
               MsgBox "Please select the correct database.", vbOKOnly
tryagain3:
               CommonDialog1.FileName = ""
               CommonDialog1.Filter = "GenMAPP Gene Databasee|*.gdb"
               CommonDialog1.ShowOpen
               CommonDialog1.CancelError = False
               If CommonDialog1.FileName = "" Then
                  MsgBox "You did not select a database. Please do so."
                  Resume tryagain3
               End If
               databaselocation = CommonDialog1.FileName
               Set dbMAPPfinder = OpenDatabase(databaselocation)
            Else
               Set dbMAPPfinder = OpenDatabase(frmStart.lblDB.Caption)
            End If
         Else
            Set dbMAPPfinder = OpenDatabase(frmStart.lblDB.Caption)
         End If
         dbset = True
      End If
      line = resultsfile.ReadLine
      While InStr(1, line, "colors:") = 0 And resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
      Wend
      localCS = line
      While InStr(1, line, "Pvalues") = 0 And resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
      Wend
      If InStr(1, line, "true") > 0 Then 'pvalue = true
         StatisticsLocal = True
         lblstat.Caption = "genes changed and a p value <"
         txtStat.Text = "0.05"
         txtStat.Left = 9910
      Else
         StatisticsLocal = False
         lblstat.Caption = "genes changed and a absolute Z score >= "
         txtStat.Text = "2"
         txtStat.Left = 10560
      End If
      While InStr(1, line, "Calculation Summary:") = 0 And resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
      Wend
      
      line = resultsfile.ReadLine
      tab1 = InStr(1, line, " ")
      frmCalculation.lblLocalProbeC.Caption = Left(line, tab1 - 1)
      line = resultsfile.ReadLine
      tab1 = InStr(1, line, " ")
      frmCalculation.lblinClusterLocalC.Caption = Left(line, tab1 - 1)
      line = resultsfile.ReadLine
      tab1 = InStr(1, line, " ")
      frmCalculation.lblgenesonMAPPC.Caption = Left(line, tab1 - 1)
      line = resultsfile.ReadLine
       tab1 = InStr(1, line, " ")
      frmCalculation.lblLocalProbeE.Caption = Left(line, tab1 - 1)
      line = resultsfile.ReadLine
       tab1 = InStr(1, line, " ")
      frmCalculation.lbinClusterLocalE.Caption = Left(line, tab1 - 1)
      line = resultsfile.ReadLine
       tab1 = InStr(1, line, " ")
      frmCalculation.lblGenesOnMAPPE.Caption = Left(line, tab1 - 1)
      
      
      
      line = resultsfile.ReadLine
      frmCalculation.lblLocalCriteria = line
     
      
      'read through all of the header lines.
      While InStr(1, line, "MAPP Name" & Chr(9) & "Number Changed" & Chr(9) _
                        & "Number Measured" & Chr(9) & "Number On MAPP" & Chr(9) & "Percent Changed" _
                        & Chr(9) & "Percent Present" & Chr(9) & "Z Score") = 0 And resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
      Wend
   ' line now equals the header line of the results file
      
      While resultsfile.AtEndOfStream = False
         line = resultsfile.ReadLine
         tab3 = InStr(1, line, Chr(9))
         GOID = Mid(line, 1, tab3 - 1)
         tab1 = InStr(tab3 + 1, line, Chr(9))
         indata = Val(Mid(line, tab3 + 1, tab1 - tab3 - 1))
         tab2 = InStr(tab1 + 1, line, Chr(9))
         onChip = Val(Mid(line, tab1 + 1, tab2 - tab1 - 1))
         tab1 = InStr(tab2 + 1, line, Chr(9))
         ingo = Val(Mid(line, tab2 + 1, tab1 - tab2 - 1))
         tab2 = InStr(tab1 + 1, line, Chr(9))
         percentage = Val(Mid(line, tab1 + 1, tab2 - tab1 - 1))
         percentage = Round(percentage, 1)
         tab1 = InStr(tab2 + 1, line, Chr(9))
         present = Val(Mid(line, tab2 + 1, tab1 - tab2 - 1))
         present = Round(present, 1)
         tab2 = InStr(tab1 + 1, line, Chr(9))
         teststat = Val(Mid(line, tab1 + 1, tab2 - tab1 - 1))
         tab1 = InStr(tab2 + 1, line, Chr(9))
         permutep = Val(Mid(line, tab2 + 1, tab1 - tab2 - 1))
         tab2 = Len(line)
         adjustedp = Val(Right(line, tab2 - tab1))
         
         'GOID = MAPPNAME
         Set GONode = TView.Nodes.Item(GOID)
         If StatisticsLocal Then
            GONode.Text = GOID & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                  & " " & present & "% z score = " & teststat & " Permute p = " & permutep & " Adjusted P = " _
                  & adjustedp
            setcolor GONode, percentage
            frmRank.lstLocal.AddItem (GOID & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                  & " " & present & "% z score = " & teststat & " Permute p = " & permutep & " Adjusted P = " _
                  & adjustedp)
         Else
            GONode.Text = GOID & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                  & " " & present & "% z score = " & teststat
            setcolor GONode, percentage
            frmRank.lstLocal.AddItem (GOID & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
                  & " " & present & "% z score = " & teststat)
         End If
         'For i = 2 To count
          '  GOID = GOID & "I"
           ' Set GONode = TView.Nodes.Item(GOID)
            'GONode.Text = GOName & "     " & indata & "/" & onChip & " " & percentage & "%, " & onChip & "/" & ingo _
             '     & ", " & present & "% Nested " & indatanested & "/" & onchipnested & " " & percentagenested & "%, " _
              '    & onchipnested & "/" & ingonested & " " & presentnested & "%"
            'setcolor GONode, percentagenested
         'Next i
         
      Wend
   End If
   End If
   CmdExpand_Click
   lblFile.Caption = "MAPPFinder Results for " & abridgedGOname & ", " & abridgedLocalName
   TView.Refresh
   TView.Visible = True
error:
   Select Case Err.Number
      Case 35601
         If GOfile Then
            'a secondary ID is being loaded. Can't load Secondary IDs.
            Resume NextTerm
         Else
            'a bad local MAPP
            MsgBox "The results you are loading containing Local MAPPs not currently loaded in MAPPFinder. Please" _
               & " re-run MAPPFinder and load the appropriate local MAPPs. MAPP " & GOID & " not found.", vbOKOnly
         End If
      Case 62
         MsgBox "MAPPFinder can not load the file you selected. The results file has been corrupted." _
            & " This probably occured while viewing the data directly in a spreadsheet format. You should" _
            & " re-run the analysis for this criteria. In the future, you should make a copy of the results" _
            & " file if you plan to work with the results outside of MAPPFinder.", vbOKOnly
      Case 3021
         'a secondary ID is in the results file. Can't load that data into the tree, because that ID doesn't exist
         Resume NextTerm
      Case 3044
         MsgBox "MAPPFinder is looking for " & dbChipLocation & ", but can not find the file." _
               & " You have either removed this file, or these results were calculated on a " _
               & "different computer. Please locate this file.", vbOKOnly
tryagain:
         CommonDialog1.FileName = ""
         CommonDialog1.Filter = "MAPPFinder Chip File|*.gmf"
         CommonDialog1.ShowOpen
         CommonDialog1.CancelError = False
         If CommonDialog1.FileName = "" Then
            MsgBox "You did not select the chip file. Please do so."
            Resume tryagain
         End If
         If openingLocal Then
            localdbChipLocation = CommonDialog1.FileName
            Resume openLocaldbChip
         Else
            dbChipLocation = CommonDialog1.FileName
            Resume opendbchip
         End If
         Err.Clear
         
      Case 3024
         MsgBox "MAPPFinder is looking for " & dbChipLocation & ", but can not find the file." _
               & " You have either removed this file, or these results were calculated on a " _
               & "different computer. Please locate this file.", vbOKOnly
twotryagain:
         CommonDialog1.FileName = ""
         CommonDialog1.Filter = "MAPPFinder Chip File|*.gmf"
         CommonDialog1.ShowOpen
         CommonDialog1.CancelError = False
         If CommonDialog1.FileName = "" Then
            MsgBox "You did not select the chip file. Please do so."
            Resume twotryagain
         End If
         If openingLocal Then
            localdbChipLocation = CommonDialog1.FileName
            Resume openLocaldbChip
         Else
            dbChipLocation = CommonDialog1.FileName
            Resume opendbchip
         End If
         Err.Clear
   End Select

endfile:
End Sub
Public Sub MakeGoPaths(current As Node, Path As String)
    Dim childnode As Node
    gopaths.Add Path, current.key
    Set childnode = current.Child
    For i = 0 To current.children - 1
         MakeGoPaths childnode, (Path & "." & i)
         Set childnode = childnode.Next
    Next i
End Sub

Private Sub MakeLocalPaths(current As Node, Path As String)
    Dim childnode As Node
    localpaths.Item(current.Index) = Path
    Set childnode = childnode.Child
    For i = 0 To current.children - 1
         MakeLocalPaths childnode, (Path & "." & i)
         Set childnode = childnode.Next
    Next i
End Sub
Public Sub setSpecies(newspecies As String)
   On Error GoTo error
   'this procedure does everything necessary to make this species specific
   Dim columns As String, label As String
   Dim pipe1 As Integer, pipe2 As Integer
   Dim found As Boolean
   Dim rsSystem As Recordset, rsMOD As Recordset
   Dim colum
   species = newspecies
   Set rsMOD = dbMAPPfinder.OpenRecordset("Select MODsystem From INFO")
   Set rsSystem = dbMAPPfinder.OpenRecordset("SELECT SystemCode, Columns from Systems where" _
                  & " System = '" & rsMOD!modsystem & "'")
   If rsSystem.EOF = False Then
      clustersystem = rsMOD!modsystem
      clustercode = rsSystem!systemcode
      'a typical column field looks like ID|Accession\SBF|Name\BF|Protein\BF|Functions\BF| we need to extract
      'the label field which contains the gene symbol. It is always the second column, unless the second is accession
      
      Set rsRelation = dbMAPPfinder.OpenRecordset( _
                        "SELECT Relation FROM Relations WHERE SystemCode = '" & clustercode _
                        & "' AND RelatedCode = 'T'")
         'this can't be empty but check it anyway
      If rsRelation.EOF = False Then
         gotable = rsRelation!Relation
         GOprimary = False
      Else
         Set rsRelation = dbMAPPfinder.OpenRecordset( _
                        "SELECT Relation FROM Relations WHERE SystemCode = 'T'" _
                        & " AND RelatedCode = '" & clustercode & "'")
         If rsRelation.EOF = False Then
            gotable = rsRelation!Relation
            GOprimary = True
         End If
      End If
      
      
      columns = rsSystem!columns
      pipe1 = InStr(1, columns, "|")
      pipe2 = InStr(pipe1 + 1, columns, "\")
      label = Mid(columns, pipe1 + 1, pipe2 - pipe1 - 1)
      If UCase(label) = "ACCESSION" Then
         pipe1 = InStr(pipe2 + 1, columns, "|")
         pipe2 = InStr(pipe1 + 1, columns, "\")
         label = Mid(columns, pipe1 + 1, pipe2 - pipe1 - 1)
      End If
      labelfield = label
      
   'Else
   '   MsgBox "The database " & databaseloc & " does not have the correct tables to run" _
   '         & " MAPPFinder for " & cmbSpecies.Text & ". Please check the species you selected" _
   '         & " and the database you are using. To change your database you must return to the" _
    '        & " start menu.", vbOKOnly
   End If
   Exit Sub
error:
   Select Case Err.Number
      Case 91
         Exit Sub 'trying to open a database that hasn't been set. No file was loaded in LoadFiles
      Case 3078
         MsgBox "The database you loaded does not appear to be a gene database. Please reload the" _
            & " results files and this time make sure you select gene database.", vbOKOnly
      Case Else
         MsgBox "Error in set species. " & Err.Number, vbOKOnly
   End Select
End Sub

Private Sub cmdColors_Click()
   frmColors.Show
End Sub

Private Sub aboutPathfinder_Click()
   frmAbout.Show
End Sub

Private Sub aboutMAPPfinder_Click()
   frmAbout.Show
End Sub

Private Sub calculateNew_Click()
   ResetTree
   TreeForm.Hide
   frmInput.Show

   
   
   
End Sub

Private Sub cmdcollapse_Click()
   TView.Visible = False
   For Each Node In TView.Nodes
      Node.Expanded = False
      Node.BackColor = RGB(255, 255, 255)
      Node.Bold = False
   Next Node
   TView.Visible = True
   TreeForm.Refresh
End Sub

Public Sub CmdExpand_Click()
   Dim i As Integer, count As Integer
   Dim currentnode As Node
   MousePointer = vbHourglass
   'TreeForm.Visible = False
   TView.Visible = False
   For Each Node In TView.Nodes
      Set currentnode = Node
      Node.Expanded = False
      
      If currentnode.BackColor = YELLOW Then 'its yellow make it white
         currentnode.BackColor = RGB(255, 255, 255)
      ElseIf currentnode.BackColor = GREEN Then  'it's green make it blue
         currentnode.BackColor = BLUE
      End If
   Next Node
      
   For Each Node In TView.Nodes 'a second time to make sure the non-white are visible
      Set currentnode = Node
      If currentnode.BackColor = BLUE Then 'it's not white so it needs to be opened
          'need to cast it to a tview.node
         openNode currentnode
      End If
   Next Node
   
   ExpandGONode rootnode
   If localloaded Then
      ExpandLocalNode LocalRoot
   End If
   TView.Visible = True
   TreeForm.Refresh
   If localloaded Then
      LocalRoot.EnsureVisible
   Else
      rootnode.EnsureVisible
   End If
   TreeForm.Visible = True
   MousePointer = vbDefault
End Sub

Private Sub cmdNumbers_Click()
   frmNumbers.Show
End Sub

Private Sub cmdGeneSearch_Click()
   On Error GoTo error
   MousePointer = vbHourglass
   TView.Visible = False
   Dim rstemp As DAO.Recordset, rstemp2 As DAO.Recordset
   Dim notfoundGO As Boolean, notfoundLocal As Boolean
   Dim currentnode As Node
   Dim dbGenMAPP As Database
   Dim key As String, i As Integer
   Dim rsSystem As Recordset, systemlist As String, rssystemlist As Recordset
   Dim geneIDs(MAX_GENES, 2) As String, genes As Integer, genefound As Boolean
   notfoundGO = True
   notfoundLocal = True
   
   For Each Node In TView.Nodes
      Set currentnode = Node
      currentnode.Expanded = False
      currentnode.Bold = False
      If currentnode.BackColor = GREEN Then
         currentnode.BackColor = YELLOW
      ElseIf Not currentnode.BackColor = YELLOW Then 'ie it's blue or white
         currentnode.BackColor = RGB(255, 255, 255)
      End If
   Next
   
   For Each Node In TView.Nodes
      Set currentnode = Node
      If currentnode.BackColor = YELLOW Then
         openNode currentnode
      End If
   Next Node
   
   'find all related genes for the input for those systems in the GeneToMAPP table
   'for each related genes
   '  does it have a mapp?
   '  if yes, then highlight that node
   Set rsSystem = dbMAPPfinder.OpenRecordset("SELECT SystemCode From [Systems] Where System = '" _
                                                & Combo1.Text & "'")
   If localMAPPsOK Then
      systemlist = "|"
      Set rssystemlist = dbLocalMAPPs.OpenRecordset("SELECT DISTINCT SystemCode from GeneToMAPP")
      While Not rssystemlist.EOF
         systemlist = systemlist & rssystemlist!systemcode & "|"
         rssystemlist.MoveNext
      Wend
      
      genes = 0
      genefound = True
      frmCriteria.AllRelatedGenes txtgeneID.Text, rsSystem!systemcode, dbMAPPfinder, _
                                    genes, geneIDs, genefound, , systemlist
      For i = 0 To genes - 1
         Set rstemp = dbLocalMAPPs.OpenRecordset("SELECT DISTINCT MAPP FROM GeneToMAPP Where " _
                                                & "ID = '" & geneIDs(i, 0) & "' AnD " _
                                                & "SystemCode = '" & geneIDs(i, 1) & "'")
         While rstemp.EOF = False
            notfoundLocal = False
            key = rstemp!MAPP
            Set currentnode = TView.Nodes.Item(key)
            If Not ((currentnode.BackColor = GREEN) Or (currentnode.BackColor = BLUE)) Then 'this mapp was already hit
               openNode currentnode
               currentnode.Bold = True
               If currentnode.BackColor = YELLOW Then
                  currentnode.BackColor = GREEN
               Else
                  currentnode.BackColor = BLUE
               End If
               currentnode.EnsureVisible
            End If
            rstemp.MoveNext
         Wend
      Next i
   End If
   
   genes = 0
   genefound = True
   frmCriteria.AllRelatedGenes txtgeneID.Text, rsSystem!systemcode, dbMAPPfinder, _
                                    genes, geneIDs, genefound, , "|T|"
   
  

For i = 0 To genes
      If geneIDs(i, 1) = "T" Then
         notfoundGO = False
         Set currentnode = TView.Nodes.Item("GO:" & geneIDs(i, 0))
         If Not ((currentnode.BackColor = GREEN) Or (currentnode.BackColor = BLUE)) Then 'this mapp was already hit
            currentnode.Bold = True
            openNode currentnode
            If currentnode.BackColor = YELLOW Then 'this term is currently yellow
               currentnode.BackColor = GREEN 'yellow and blue make green (it's sealed?)
            Else ' it's white, make it blue
               currentnode.BackColor = BLUE
            End If
            currentnode.EnsureVisible
         End If
      End If
   Next i
       
      
      
   If notfoundLocal = True And notfoundGO = True Then
      MsgBox "No MAPPs or GO terms could be found for that " & Combo1.Text & " ID.", vbOKOnly
   End If
error:
   Select Case Err.Number
      Case 35601
         MsgBox "The MAPPs loaded in MAPPFinder do not agree with the MAPPs in the " & species & " database. " _
            & "Please reload the local MAPPs.", vbOKOnly
   End Select
   TView.Visible = True
   TreeForm.Refresh
   MousePointer = vbDefault
   
End Sub

Private Sub CmdSearchGO_Click()
   Dim rsGOID As DAO.Recordset, rsCount As DAO.Recordset
   Dim GONode As Node
   Dim GOString As String
   Dim count As Integer, i As Integer, currentnode As Node
   MousePointer = vbHourglass
   TView.Visible = False
   
   For Each Node In TView.Nodes
      Set currentnode = Node
      currentnode.Expanded = False
      currentnode.Bold = False
      If currentnode.BackColor = GREEN Then
         currentnode.BackColor = YELLOW
      ElseIf Not currentnode.BackColor = YELLOW Then 'ie it's blue or white
         currentnode.BackColor = RGB(255, 255, 255)
      End If
   Next
   
   For Each Node In TView.Nodes
      Set currentnode = Node
      If currentnode.BackColor = YELLOW Then
         openNode currentnode
      End If
   Next Node
   
   If keyword Then
   For Each Node In TView.Nodes
      Set GONode = Node
      If InStr(1, UCase(GONode.Text), UCase(txtGoTerm.Text)) <> 0 Then
         GONode.Expanded = True
         GONode.Bold = True
         found = True
         openNode GONode
         GONode.EnsureVisible
         If GONode.BackColor = YELLOW Then
            GONode.BackColor = GREEN
         Else
            GONode.BackColor = BLUE
         End If
      End If
   Next
   Else 'exact search
   For Each Node In TView.Nodes
      Set GONode = Node
      GOString = getName(GONode)
      If StrComp(UCase(GOString), UCase(txtGoTerm.Text)) = 0 Then
         GONode.Expanded = True
         GONode.Bold = True
         found = True
         openNode GONode
         GONode.EnsureVisible
         If GONode.BackColor = YELLOW Then
            GONode.BackColor = GREEN
         Else
            GONode.BackColor = BLUE
         End If
      End If
   Next
   End If
   If found = False Then
      MsgBox "The word you searched for could not be found. Please check that you spelled it correctly.", vbOKOnly
      GoTo getout
   End If
getout:
  
   TView.Visible = True
   TreeForm.Refresh
    MousePointer = vbDefault
End Sub

Public Sub closenode(GONode As Node)
   Dim parentnode As Node
   
   If localloaded Then
      If GONode.key = rootnode.key Or GONode.key = LocalRoot.key Then
         GONode.Expanded = False
      Else
      
         Set parentnode = GONode.parent
         If parentnode.Expanded = True Then 'else it's already been closed, so don't bother doing it again
            parentnode.Expanded = False
            While (InStr(1, parentnode.Text, rootnode.Text) = 0) And (InStr(1, parentnode.Text, LocalRoot.Text) = 0)
               Set parentnode = parentnode.parent
               parentnode.Expanded = False
            Wend
         End If
      End If
   Else
      If GONode.key = rootnode.key Then
         GONode.Expanded = False
      Else
         Set parentnode = GONode.parent
         If parentnode.Expanded = True Then 'else it's already been closed, so don't bother doing it again
            parentnode.Expanded = False
            While (InStr(1, parentnode.Text, rootnode.Text) = 0)
               Set parentnode = parentnode.parent
               parentnode.Expanded = False
            Wend
         End If
      End If
   End If
End Sub
   
Public Sub openNode(GONode As Node)
   Dim parentnode As Node
   
   If localloaded Then
      If GONode.key = rootnode.key Or GONode.key = LocalRoot.key Then
         GONode.Expanded = True
      Else
      
         Set parentnode = GONode.parent
         If parentnode.Expanded = False Then 'else it's already been closed, so don't bother doing it again
            parentnode.Expanded = True
            While (InStr(1, parentnode.Text, rootnode.Text) = 0) And (InStr(1, parentnode.Text, LocalRoot.Text) = 0)
               Set parentnode = parentnode.parent
               parentnode.Expanded = True
            Wend
         End If
      End If
   Else
      If GONode.key = rootnode.key Then
         GONode.Expanded = True
      Else
         Set parentnode = GONode.parent
         If parentnode.Expanded = False Then 'else it's already been closed, so don't bother doing it again
            parentnode.Expanded = True
            While (InStr(1, parentnode.Text, rootnode.Text) = 0)
               Set parentnode = parentnode.parent
               parentnode.Expanded = True
            Wend
         End If
      End If
   End If
End Sub


Private Sub Exit_Click()
   dbMAPPfinder.Close
   If godataloaded Then
      dbChip.Close
   End If
   If localdataloaded Then
      dbchiplocal.Close
   End If
   End
End Sub

Private Sub Command1_Click()
   
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormCode Then                                            'Exiting from mnuExit
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
  
 Unload frmColors
 Unload frmNumbers
 
   
 Exit_Click

End Sub
Public Sub ExportGONode(currentnode As Node, Path As String, File As TextStream)
   Dim children As Integer, i As Integer
   Dim childnode As Node
   Dim start As Integer, finish As Integer
   Dim s As String, percent As String
   Dim line As String, rstemp As Recordset
   If (currentnode.BackColor = YELLOW) Then
      'parse the text and write it out
      start = InStr(1, currentnode.Text, "    ")
      s = Left(currentnode.Text, start - 1)
      Set rstemp = dbMAPPfinder.OpenRecordset("SELECT ID, Type from GeneOntology where Name = '" & s & "'")
      line = Path & Chr(9) & rstemp!id & Chr(9) & s & Chr(9) & rstemp!Type
      rstemp.Close
      start = start + 5
      finish = InStr(start, currentnode.Text, "/")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s
      start = finish + 1
      finish = InStr(start, currentnode.Text, " ")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s
      start = finish + 1
      finish = InStr(start, currentnode.Text, "%")
      percent = Mid(currentnode.Text, start, finish - start)
      'line = line & Chr(9) & s
      start = finish + 3
      finish = InStr(start, currentnode.Text, "/")
      's = Mid(currentnode.Text, start, finish - start)
      'line = line & Chr(9) & s
      start = finish + 1
      finish = InStr(start, currentnode.Text, " ")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s & Chr(9) & percent
      start = finish + 1
      finish = InStr(start, currentnode.Text, "%")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s
      start = InStr(finish + 1, currentnode.Text, "NESTED") + 7
      finish = InStr(start, currentnode.Text, "/")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s
      start = finish + 1
      finish = InStr(start, currentnode.Text, " ")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s
      start = finish + 1
       finish = InStr(start, currentnode.Text, "%")
      percent = Mid(currentnode.Text, start, finish - start)
      'line = line & Chr(9) & s
      start = finish + 3
      finish = InStr(start, currentnode.Text, "/")
      's = Mid(currentnode.Text, start, finish - start)
      'line = line & Chr(9) & s
      start = finish + 1
      finish = InStr(start, currentnode.Text, " ")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s & Chr(9) & percent
      start = finish + 1
      finish = InStr(start, currentnode.Text, "%")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s
      start = InStr(finish, currentnode.Text, "=") + 2
      finish = InStr(start, currentnode.Text, " ")
      If (finish = 0) Then
         s = Mid(currentnode.Text, start, Len(currentnode.Text) - start + 1)
      Else
         s = Mid(currentnode.Text, start, finish - start)
      End If
      line = line & Chr(9) & s
      If Statistics Then
         start = InStr(finish, currentnode.Text, "=") + 2
         finish = InStr(start, currentnode.Text, " ")
         s = Mid(currentnode.Text, start, finish - start)
         line = line & Chr(9) & s
         start = InStr(finish, currentnode.Text, "=") + 2
         finish = Len(currentnode.Text)
         s = Mid(currentnode.Text, start, finish - start + 1)
         line = line & Chr(9) & s
      End If
      File.WriteLine treeindex & Chr(9) & line
      treeindex = treeindex + 1
   End If
   children = currentnode.children
   Set childnode = currentnode.Child
   For i = 0 To children - 1
      ExportGONode childnode, Path & "." & i, File
      Set childnode = childnode.Next
   Next i
      
      
End Sub
Public Sub ExportLocalNode(currentnode As Node, Path As String, File As TextStream)
   Dim children As Integer, i As Integer
   Dim childnode As Node
   Dim start As Integer, finish As Integer
   Dim s As String, percent As String
   Dim line As String
   If (currentnode.BackColor = YELLOW) Then
      
      'parse the text and write it out
      start = InStr(1, currentnode.Text, "    ")
      s = Left(currentnode.Text, start - 1)
      line = Path & Chr(9) & s
      start = start + 5
      finish = InStr(start, currentnode.Text, "/")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s
      start = finish + 1
      finish = InStr(start, currentnode.Text, " ")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s
      start = finish + 1
      finish = InStr(start, currentnode.Text, "%")
      percent = Mid(currentnode.Text, start, finish - start)
      'line = line & Chr(9) & s
      start = finish + 3
      finish = InStr(start, currentnode.Text, "/")
      s = Mid(currentnode.Text, start, finish - start)
      'line = line & Chr(9) & s
      start = finish + 1
      finish = InStr(start, currentnode.Text, " ")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s & Chr(9) & percent
      start = finish + 1
      finish = InStr(start, currentnode.Text, "%")
      s = Mid(currentnode.Text, start, finish - start)
      line = line & Chr(9) & s
      start = InStr(finish, currentnode.Text, "=") + 2
      finish = InStr(start, currentnode.Text, " ")
      If (finish = 0) Then
         s = Mid(currentnode.Text, start, Len(currentnode.Text) - start + 1)
      Else
         s = Mid(currentnode.Text, start, finish - start)
      End If
      line = line & Chr(9) & s
      If StatisticsLocal Then
         start = InStr(finish, currentnode.Text, "=") + 2
         finish = InStr(start, currentnode.Text, " ")
         s = Mid(currentnode.Text, start, finish - start)
         line = line & Chr(9) & s
         start = InStr(finish, currentnode.Text, "=") + 2
         finish = Len(currentnode.Text)
         s = Mid(currentnode.Text, start, finish - start + 1)
         line = line & Chr(9) & s
      End If
      File.WriteLine treeindex & Chr(9) & line
      treeindex = treeindex + 1
   End If
   children = currentnode.children
   Set childnode = currentnode.Child
   For i = 0 To children - 1
      ExportLocalNode childnode, Path & "." & i, File
      Set childnode = childnode.Next
   Next i
      
      
End Sub
Public Sub ExpandGONode(currentnode As Node)
   Dim children As Integer, i As Integer
   Dim childnode As Node
   Dim nested As Integer
   Dim start As String, finish As String
   Dim indata As Integer, percentage As Single
   Dim teststat As Double, permutep As Double
   
   
   
   children = currentnode.children
   Set childnode = currentnode.Child
   For i = 0 To children - 1
      ExpandGONode childnode
      Set childnode = childnode.Next
   Next i
   
   nested = InStr(1, currentnode.Text, "NESTED")
   If nested <> 0 Then 'else this node has no associated genes and no data, ignore it
      If Not Statistics Then
         start = nested + 7
         finish = InStr(nested, currentnode.Text, "/")
         indata = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, " ") + 1
         finish = InStr(start, currentnode.Text, "%")
         percentage = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, "=")
         teststat = Val(Mid(currentnode.Text, start + 2, Len(currentnode.Text) - start + 1))
         If indata >= Val(txtNumberChanged) And percentage >= Val(txtPercent) And Abs(teststat) >= Abs(Val(txtStat)) Then
         
         
            If currentnode.Expanded = False Then 'if it's already expanded then someone below met the criteria, so don't bother with recursive expanding
               openNode currentnode
            End If
            If currentnode.BackColor = BLUE Then 'it's blue, need to make it green
               currentnode.BackColor = GREEN
            Else 'it's not blue, so it must be white, make it yellow
               currentnode.BackColor = YELLOW
            End If
         End If
      Else
         start = nested + 7
         finish = InStr(nested, currentnode.Text, "/")
         indata = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, " ") + 1
         finish = InStr(start, currentnode.Text, "%")
         percentage = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, "=") + 2
         finish = InStr(start + 2, currentnode.Text, " ")
         teststat = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, "=") + 2
         finish = InStr(start + 2, currentnode.Text, " ")
         permutep = Val(Mid(currentnode.Text, start, finish - start))
         
         If indata >= Val(txtNumberChanged) And percentage >= Val(txtPercent) And permutep < Val(txtStat) Then
         
         
            If currentnode.Expanded = False Then 'if it's already expanded then someone below met the criteria, so don't bother with recursive expanding
               openNode currentnode
            End If
            If currentnode.BackColor = BLUE Then 'it's blue, need to make it green
               currentnode.BackColor = GREEN
            Else 'it's not blue, so it must be white, make it yellow
               currentnode.BackColor = YELLOW
            End If
         End If
      End If
   End If
   
End Sub
Public Sub ExpandLocalNode(currentnode As Node)
   Dim children As Integer, i As Integer
   Dim childnode As Node
   Dim nested As Integer
   Dim start As String, finish As String
   Dim indata As Integer, percentage As Single
   
   children = currentnode.children
   Set childnode = currentnode.Child
   For i = 0 To children - 1
      ExpandLocalNode childnode
      Set childnode = childnode.Next
   Next i
   nested = InStr(1, currentnode.Text, "     ")
   If nested <> 0 Then 'else this node has no associated genes and no data, ignore it
      If Not StatisticsLocal Then
         start = nested + 1
         finish = InStr(nested, currentnode.Text, "/")
         indata = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, " ") + 1
         finish = InStr(start, currentnode.Text, "%")
         percentage = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, "=")
         teststat = Val(Mid(currentnode.Text, start + 2, Len(currentnode.Text) - start + 1))
         If indata >= Val(txtNumberChanged) And percentage >= Val(txtPercent) And Abs(teststat) >= Abs(Val(txtStat)) Then
            If currentnode.Expanded = False Then 'if it's already expanded then someone below met the criteria, so don't bother
               openNode currentnode
            End If
            'if currentnode
            
            currentnode.BackColor = RGB(255, 255, 127)
         End If
      Else
         start = nested + 1
         finish = InStr(nested, currentnode.Text, "/")
         indata = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, " ") + 1
         finish = InStr(start, currentnode.Text, "%")
         percentage = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, "=") + 2
         finish = InStr(start + 2, currentnode.Text, " ")
         teststat = Val(Mid(currentnode.Text, start, finish - start))
         start = InStr(finish + 1, currentnode.Text, "=") + 2
         finish = InStr(start + 2, currentnode.Text, " ")
         permutep = Val(Mid(currentnode.Text, start, finish - start))
         
         If indata >= Val(txtNumberChanged) And percentage >= Val(txtPercent) And permutep < Val(txtStat) Then
         
         
            If currentnode.Expanded = False Then 'if it's already expanded then someone below met the criteria, so don't bother with recursive expanding
               openNode currentnode
            End If
            If currentnode.BackColor = BLUE Then 'it's blue, need to make it green
               currentnode.BackColor = GREEN
            Else 'it's not blue, so it must be white, make it yellow
               currentnode.BackColor = YELLOW
            End If
         End If
      End If
   End If
  
End Sub


Private Sub loadExisting_Click()
   ResetTree
   frmLoadFiles.Show
   TreeForm.Hide
End Sub

Private Sub LoadLocalMAPPs_Click()
   ResetTree
   frmLocalMAPPs.Show
   frmLocalMAPPs.LoadSpecies
   TreeForm.Hide
End Sub

Private Sub MAPPFinderhelp_Click()
   Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, programpath & "\GenMAPP.chm::/MAPPFinder.htm", HH_DISPLAY_TOPIC, 0)

End Sub

Private Sub rankedList_Click()
   frmRank.Show
End Sub

Private Sub TView_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo error
   MousePointer = vbHourglass
   Dim MAPPName As String
   Dim NameEnd As Integer, slash As Integer
   Dim mappMade As Boolean
   Dim rsGenes As Recordset
   Dim GOID As String
   
   Dim abrev As String, space As Integer
   
   abrev = Mid(species, 1, 1)
   space = InStr(1, species, " ")
   abrev = UCase(abrev & Mid(species, space + 1, 1))
   If OpenMAPPWhenClicked Then 'go through the mapp opening process
      If Node = rootnode Then
         MsgBox "A MAPP can not be created for the entire Gene Ontology. Please select a more specific MAPP.", vbOKOnly
      Else
         If InStr(1, Node.FullPath, "Gene Ontology") <> 0 Then 'it's a GO MAPP do the GO routine
            NameEnd = InStr(1, Node.Text, "     ")
            If NameEnd = 0 Then 'nothing has been added, ie. this node has no MAPPFinder data
               MAPPName = MappBuilderForm_Normal.fixName(Node.Text)
            Else
               MAPPName = MappBuilderForm_Normal.fixName(Mid(Node.Text, 1, NameEnd - 1))
            End If
            mappMade = MakeMapp(MAPPName, Node)
            If mappMade Then
               'the location is c:\GenMAPPv2\MAPPs\MAPPFinder\Homo Sapiens\Fatty Acid Metabolism.mapp
               'EDlocation = """C:\Genmappv2\Expression Datasets\Development Data MAS 5-fixed.gex"""
               'CS = "12.5 Day Embryo"""
               'Debug.Print ggenmapploc & " """ & mapploc & species & "\" & fixPath(MAPPName) & ".mapp"" """ _
                             & ED & """ """ & CS & """ """ & databaselocation & """"
               GenMAPP = Shell(genmapploc & " """ & mapploc & abrev & " GO\" & fixPath(MAPPName) & ".mapp"" """ _
                             & ED & """ """ & CS & """ """ & databaselocation & """", vbNormalFocus)
            End If
         Else 'its a local mapp
            'the mapp already exists, so open it
            NameEnd = InStr(1, Node.FullPath, "     ")
            If NameEnd = 0 Then 'nothing has been added, ie. this node has no MAPPFinder data
               slash = InStr(1, Node.FullPath, "\")
               MAPPName = Mid(Node.FullPath, slash, Len(Node.FullPath) - slash + 1)
            Else
               slash = InStr(1, Node.FullPath, "\")
               MAPPName = Mid(Node.FullPath, slash, NameEnd - slash)
            End If
            
           ' Debug.Print LocalPath & MAPPName & ".mapp"
            If Dir(LocalPath & MAPPName & ".mapp") <> "" Then
               GenMAPP = Shell(genmapploc & " """ & LocalPath & MAPPName & ".mapp"" """ _
                              & localED & """ """ & databaselocation & """ """ & CS & """", vbNormalFocus)
            End If
         End If
      End If
   Else 'open the list of genes form. No MAPP
      frmGeneList.LstChanged.Clear
      frmGeneList.lstAllgenes.Clear
      If InStr(1, Node.FullPath, "Gene Ontology") <> 0 Then 'it's a GO MAPP do the GO routine
         
         If Node.key = "GO" Then
            GOID = "GO"
         Else
            GOID = Mid(Node.key, 4, GOLENGTH - 3) 'Each GO term is GO:1234567
         End If
         Set rsgene = dbChip.OpenRecordset("SELECT Related from [" & tablename _
                     & "] WHERE Primary = '" & GOID & "' ORDER BY Related")
         While rsgene.EOF = False
            frmGeneList.LstChanged.AddItem rsgene!related
            rsgene.MoveNext
         Wend
         
         Set rsgene = dbChip.OpenRecordset("SELECT Related from [GOIDtoGene]" _
                     & " WHERE Primary = '" & GOID & "' ORDER BY Related")
         While rsgene.EOF = False
            frmGeneList.lstAllgenes.AddItem rsgene!related
            rsgene.MoveNext
         Wend
         
      Else ' a local node
         GOID = Node.key
         Set rsgene = dbchiplocal.OpenRecordset("SELECT Related from [" & LocalTablename _
                     & "] WHERE Primary = '" & GOID & "' ORDER BY Related")
         While rsgene.EOF = False
            frmGeneList.LstChanged.AddItem rsgene!related
            rsgene.MoveNext
         Wend
         
         Set rsgene = dbchiplocal.OpenRecordset("SELECT Related from [GOIDtoGene]" _
                     & " WHERE Primary = '" & GOID & "' ORDER BY Related")
         While rsgene.EOF = False
            frmGeneList.lstAllgenes.AddItem rsgene!related
            rsgene.MoveNext
         Wend
      End If
      frmGeneList.lblNodeText.Caption = Node.Text
      frmGeneList.ExpressionDataset = ED
      frmGeneList.Refresh
      frmGeneList.Show
   End If
error:
   Select Case Err.Number
      Case 5
         MsgBox "No MAPP exists for that node.", vbOKOnly
      Case 3078 'that database table doesn't exist?
         MsgBox "The database table " & tablename & " was not found in the MAPPFinder data" _
           & " file for this dataset, " & dbChipLocation & ". It looks like you've either" _
          & " changed the file name of a MAPPFinder results file, or deleted a table from" _
           & " the MAPPFinder data file. Neither of which is good. You'll have to rerun the" _
          & " analysis for this criterion.", vbOKOnly
            
   End Select
         
   MousePointer = vbDefault
End Sub

Public Function MakeMapp(MAPPName As String, currentnode As Node) As Boolean
  On Error GoTo error
   Dim mappbuilderfile As TextStream
   Dim build As Boolean
   Dim chip As Boolean
restart:
   chip = False
   'create mapp builder template
   Set mappbuilderfile = Fsys.CreateTextFile(mapploc & "MAPPFinderTempMAPPBuilder.txt")
   mappbuilderfile.WriteLine ("geneId" & Chr(9) & "systemcode" & Chr(9) & "Label" _
                              & Chr(9) & "Head" & Chr(9) & "Remarks" & Chr(9) & "MappName")
   build = BuildMAPPs(currentnode, MAPPName, mappbuilderfile)
   If build Then
      build = OpenMAPPBuilder(mapploc & "MAPPFinderTempMAPPBuilder.txt")
   End If
   mappbuilderfile.Close
   Fsys.DeleteFile (mapploc & "MAPPFinderTempMAPPBuilder.txt")
   MakeMapp = build

error:
   Select Case Err.Number
      Case 3044
         MsgBox "MAPPFinder is looking for " & dbChipLocation & ", but can not find the file." _
               & " You have either removed this file, or these results were calculated on a " _
               & "different computer. Please locate this file.", vbOKOnly
tryagain:
         CommonDialog1.FileName = ""
         CommonDialog1.Filter = "MAPPFinder Chip File|*.gmf"
         CommonDialog1.ShowOpen
         CommonDialog1.CancelError = False
         If CommonDialog1.FileName = "" Then
            MsgBox "You did not select the chip file. Please do so."
            Resume tryagain
         End If
         dbChipLocation = CommonDialog1.FileName
         
         Err.Clear
         Resume restart
         
      Case 3024
         MsgBox "MAPPFinder is looking for " & dbChipLocation & ", but can not find the file." _
               & " You have either removed this file, or these results were calculated on a " _
               & "different computer. Please locate this file.", vbOKOnly
tryagain2:
         CommonDialog1.FileName = ""
         CommonDialog1.Filter = "MAPPFinder Chip File|*.gmf"
         CommonDialog1.ShowOpen
         CommonDialog1.CancelError = False
         If CommonDialog1.FileName = "" Then
            MsgBox "You did not select the chip file. Please do so."
            Resume tryagain2
         End If
         dbChipLocation = CommonDialog1.FileName
         Err.Clear
         Resume restart
      Case 76
         MsgBox "MAPPFinder can not find the path for the GO MAPPs. Be sure that there is a MAPPFinder" _
            & " folder in your base MAPPs folder (default would be c:\GenMAPP 2 Data\MAPPs\MAPPFinder\).", vbOKOnly
      
   End Select
   If chip Then
      dbChip.Close
   End If
End Function
'this function will return the number of genes associated with a GO term and all of its children
Public Function countGenes(currentnode As Node) As Long
   Dim rsGeneNumber As DAO.Recordset
   Dim GOID As String
   
   GOID = currentnode.key
   If Len(GOID) > GOLENGTH Then 'a "I" has been added to this ID.
      GOID = Mid(GOID, 1, GOLENGTH)
   End If
   GOID = Mid(GOID, 4, GOLENGTH - 3)
   
   Set rsGeneNumber = dbMAPPfinder.OpenRecordset("Select Total FROM [" & clustersystem _
                           & "-GOCount] WHERE GO = '" & GOID & "'")
   If rsGeneNumber.EOF Then
      countGenes = 0
   Else
      countGenes = rsGeneNumber![total]
   End If
End Function


'this function will build the mapps for this node and the mapps for all of its children.

Public Function BuildMAPPs(currentnode As Node, MAPPName As String, mappbuilderfile As TextStream) As Boolean
   Dim childnd As Node
   Dim i As Integer, key As String
   Dim MAPP As TextStream
   Dim genecount As Long
   Dim rsGOcount As DAO.Recordset
   BuildMAPPs = True
   'If Len(currentnode.key) = GOLENGTH Then 'if currentnode is longer than GOLENGTH an I has been added and the MAPP is a duplicate, so don't make it
   'build the mapp with the genes of this node and all of its children
   key = Mid(currentnode.key, 4, GOLENGTH - 3) 'Each GO term is GO:1234567
   genecount = countGenes(currentnode)
   MappNameList = "NEW"
   'this will build the mapp for the node that was clicked. If that node has more than 300 children,
   'a recursive mapp will not be made. Instead a message saying "too many children to be displayed,
   'look at a more specific mapp." will be shown.

   If genecount > 0 Then 'this is a artifact from the GO MAPP Builder
         If genecount < 300 Then
            AddGenesToMAPP currentnode, TextToSql(MAPPName), mappbuilderfile, True
         Else
         'building the recursive mapps would have to many genes, so we have to build it the
         'old way with only the genes from that GO term, not its children too. Currently we're
         'going to build all GO MAPPs this way regardless of size
            AddGenesToMAPP currentnode, TextToSql(MAPPName), mappbuilderfile, False
         End If
   Else
      MsgBox "There are no genes associated with this GO term. No MAPP can be created.", vbOKOnly
      BuildMAPPs = False
   End If
     
End Function

Public Sub AddGenesToMAPP(currentnode As Node, MAPPName As String, mappbuilderfile As TextStream, recurse As Boolean)
   Dim rsGenes As DAO.Recordset, rsName As DAO.Recordset, rsMGI As DAO.Recordset
   Dim rschip As DAO.Recordset
   Dim GOID As String
   Dim i As Integer
   Dim childnd As Node, name As String
   Dim rsMAPP As DAO.Recordset
   Dim rsGOName As DAO.Recordset
   Dim dbLocalMAPPs As Database
   Dim childnames As String, childID As String
   
   Set dbLocalMAPPs = OpenDatabase(programpath & "LocalMAPPTmpl.gtp")
   GOID = currentnode.key
   If GOID = "GO" Then
      MsgBox "A MAPP for all of the Gene Ontology cannot be produced. It is too large.", vbOKOnly
      GoTo root
   Else
      GOID = Mid(GOID, 4, GOLENGTH - 3) 'remove GO:
   End If
      
   Set rsGenes = dbMAPPfinder.OpenRecordset("Select Distinct Primary FROM [" & gotable & "]" _
                              & " WHERE Related = '" & GOID & "'")
   Set rsGOName = dbMAPPfinder.OpenRecordset("Select Name From GeneOntology" _
                              & " WHERE ID = '" & GOID & "'")
   
   If rsGenes.EOF Then 'there are no genes for that GO Term
      'do something if you want to?
   Else
      
      addLabel mappbuilderfile, UCase(TextToSql(rsGOName![name])), MAPPName
      While rsGenes.EOF = False
         Set rsName = dbMAPPfinder.OpenRecordset("Select " & labelfield & " From [" _
                     & clustersystem & "] WHERE ID = '" & rsGenes![primary] & "'")
         If rsName.EOF = True Then
               name = rsGenes![primary] 'SP doesn't have a symbol for this gene
               Head = rsGenes![primary]
         Else
            name = TextToSql(rsName.Fields(0))
            If name = "" Then
               name = rsGenes![primary]
               Head = rsGenes![primary]
            ElseIf InStr(2, name, "|") > 0 Then 'looks like |abcd|EFGH| (TAIR does this)
               pipe1 = InStr(2, name, "|")
               name = Mid(name, 2, pipe1 - 2)
            Else
               name = TextToSql(rsName.Fields(0))
               name = Replace(name, "|", "")
            End If
            If InStr(1, name, ",") <> 0 Then
               name = Mid(name, 1, InStr(1, name, ","))
            End If
            Head = TextToSql(name)
         End If
         
         dbLocalMAPPs.Execute ("INSERT INTO MAPPTemplate (ID, SystemCode, Label, Head, Remarks, MAPPName)" _
                           & " VALUES('" & rsGenes![primary] & "', '" & clustercode & "', '" & name & "', '" _
                           & Head & "', ' ', '" & MAPPName & "')")
               
         rsGenes.MoveNext
      Wend
   End If
   
   
   Set rsMAPP = dbLocalMAPPs.OpenRecordset("SELECT * FROM MAPPTemplate Order By Label")
   While rsMAPP.EOF = False
      mappbuilderfile.WriteLine (rsMAPP![id] & Chr(9) & rsMAPP![systemcode] & Chr(9) & rsMAPP![label] _
                        & Chr(9) & rsMAPP![Head] & Chr(9) & Chr(9) & rsMAPP![MAPPName])
      rsMAPP.MoveNext
   Wend
   'End If
   dbLocalMAPPs.Execute "DELETE * FROM MAPPTemplate"
   If currentnode.children > 0 And recurse Then 'recurse
      Set childnd = currentnode.Child
      For i = 0 To currentnode.children - 1
         AddGenesToMAPP childnd, MAPPName, mappbuilderfile, True
         Set childnd = childnd.Next
      Next i
   ElseIf currentnode.children > 0 And recurse = False Then
      
      addLabel mappbuilderfile, "The number of children associated with this term is greater than the number that" _
                              & " can be usefully displayed in a MAPP file. Please look at one of the children of this GO term:", MAPPName
      
      Set childnd = currentnode.Child
      For i = 0 To currentnode.children - 1
        childID = childnd.key
        childID = Mid(childID, 4, GOLENGTH - 3) 'remove GO:
         Set rsGOName = dbMAPPfinder.OpenRecordset("Select Name From GeneOntology" _
                              & " WHERE ID = '" & childID & "'")
         
         addLabel mappbuilderfile, UCase(rsGOName![name]), MAPPName
         Set childnd = childnd.Next
      Next i
      
   End If
   dbLocalMAPPs.Close
root:
End Sub

Function TextToSql(txt As String) As String '**************************** Makes Text SQL Compatible
   Dim Index As Integer                     'copied from GenMAPP 1.0 Source code
   Dim sql As String
   
   sql = txt
   For Index = 1 To Len(txt)
     Select Case Mid(txt, Index, 1)
     Case "'"                            'Convert single quote to typographer's close single quote
        Mid(sql, Index, 1) = Chr(146)
     Case "!"                            'Convert exclamation quote to blank
        Mid(sql, Index, 1) = Chr(32)
     End Select
   Next Index
   TextToSql = sql
End Function

Public Function fixPath(Path As String) As String
    Dim Index As Integer
    For Index = 1 To Len(Path)
      Select Case Mid(Path, Index, 1)
      Case "/"
        Mid(Path, Index, 1) = Chr(32)
      Case ":"                            'Convert single quote to typographer's close single quote
         Mid(Path, Index, 1) = Chr(32)
      End Select
   Next Index
    Path = TextToSql(Path)
    fixPath = Path

End Function


Public Sub addLabel(mappbuilderfile As TextStream, label As String, MAPPName As String)
   Dim temp As String
   'labels can't be wider than LABEL_LENGTH characters or they run into the next column.
   'this procedure truncates them and creates multiple lines.
   If Len(label) > LABEL_LENGTH Then
      temp = Mid(label, 1, LABEL_LENGTH)
      mappbuilderfile.WriteLine (Chr(9) & "Label" & Chr(9) & temp & "-" _
                                 & Chr(9) & Chr(9) & Chr(9) & MAPPName)
      addLabel mappbuilderfile, Mid(label, LABEL_LENGTH + 1, Len(label) - LABEL_LENGTH), MAPPName
   Else
      mappbuilderfile.WriteLine (Chr(9) & "Label" & Chr(9) & label _
                                 & Chr(9) & Chr(9) & Chr(9) & MAPPName)
   End If
End Sub







Public Function OpenMAPPBuilder(File As String) As Boolean
   Dim mappCFG As TextStream
   Dim baseMAPP As String, line As String
   Dim datalocation As String
    Dim abrev As String, space As Integer
  abrev = Mid(species, 1, 1)
   space = InStr(1, species, " ")
   abrev = UCase(abrev & Mid(species, space + 1, 1))
   MappBuilderForm_Normal.setBaseMapp mapploc & abrev & " GO\", databaselocation
   MappBuilderForm_Normal.setFileName File
   MappBuilderForm_Normal.MakeMapps_Click
   OpenMAPPBuilder = True
   
End Function

Public Function makeRecursiveMAPPs(currentnode As Node) As Boolean
'this function will return true if the genecount of currentnode is < 300
   Dim count As Long
   Dim recurse As Boolean
   count = countGenes(currentnode)
   If count < 300 Then
      recurse = True
   Else 'count >= 300
     recurse = False
   End If
   makeRecursiveMAPPs = recurse
   
End Function



Public Sub DisplayLocalMAPPs(dbExpressionData As Database, currentnode As Node)
'take the data from the results table and append it to the text of each of the nodes representing local MAPPs.
   Dim rsResults As DAO.Recordset
   Dim childnd As Node
   Dim i As Integer
   
   Set rsResults = dbExpressionData.OpenRecordset("SELECT * FROM Results WHERE GOID = '" _
                  & currentnode.Text & "'")
   If rsResults.EOF = False Then
      currentnode.Text = currentnode.Text & "     " & rsResults![indata] & "/" & rsResults![onChip] _
                     & " " & rsResults![percentage] & "%, " & rsResults![onChip] & "/" & rsResults![ingo] _
                     & ", " & rsResults![present] & "%"
      setcolor currentnode, rsResults![percentage]
   End If
   Set childnd = currentnode.Child
   For i = 0 To currentnode.children - 1
      DisplayLocalMAPPs dbExpressionData, childnd
      Set childnd = childnd.Next
   Next i

End Sub



Private Sub whatdocolorsmean_Click()
    frmColors.Show
End Sub

Private Sub whatdonumbersmean_Click()
   frmNumbers.Show
End Sub

Public Sub setLocalPath(Path As String)
   LocalPath = Path
End Sub

Public Sub setChipDBLocation(loc As String)
   
   dbChipLocation = Mid(loc, 1, Len(loc) - 3) & "gdb" 'replace gex with gdb

End Sub

Public Sub setDatabase(dbname As String)
   On Error GoTo error
   Dim rsGO As Recordset
   Dim rsSystems As Recordset
   Set dbMAPPfinder = OpenDatabase(dbname)
   databaselocation = dbname
   frmRank.setDB dbMAPPfinder
      
   'we need to find out all of the systems who related to Gene Ontology for the search button.
   Set rsGO = dbMAPPfinder.OpenRecordset("SELECT SystemCode FROM Relations WHERE relatedCode = 'T'")
   While rsGO.EOF = False
      Set rsSystem = dbMAPPfinder.OpenRecordset("SELECT System from Systems WHERE SystemCOde = '" _
                                                & rsGO!systemcode & "'")
      Combo1.AddItem rsSystem!system
      rsGO.MoveNext
   Wend
   Set rsGO = dbMAPPfinder.OpenRecordset("SELECT RelatedCode FROM Relations WHERE SystemCode = 'T'")
   While rsGO.EOF = False
      Set rsSystem = dbMAPPfinder.OpenRecordset("SELECT System from Systems WHERE SystemCOde = '" _
                                                & rsGO!relatedCode & "'")
      Combo1.AddItem rsSystem!system
      rsGO.MoveNext
   Wend
error:
   Select Case Err.Number
      Case 3078 'the error for not having the database
         MsgBox "The database does not appear to be a GenMAPP gene database. Please make sure you have the correct" _
         & " database loaded.", vbOKOnly
   End Select
End Sub

Public Sub setFileName(label As String)
   lblFile.Caption = label
End Sub

Public Function unique(incoming() As String, ByVal genes As Long) As Integer
   'this function counts the number of unique elements in the array
   'incoming is the array, genes is the number of elements in the array
   'since all we need is a count of the number of unique elements, not the actual list of unique there is no need to fix the array
   Dim i As Long, j As Integer
   Dim split As Integer
   mergesort incoming, 0, genes - 1 'array starts at 0, so offset by 1
   
   For i = 0 To genes - 2
      
      If StrComp(incoming(i), incoming(i + 1)) = 0 Then 'the two are equal there are duplicates
            genes = genes - 1
      End If
   Next i
   unique = genes
End Function

Private Sub mergesort(tester() As String, start As Long, finish As Long)
   'Input - an array of strings
   'Output - the array sorted alphabetically. The sort is a mergesort (NlogN).
   'The sort is not case sensitive (ie A = a).
   Dim temp As String
   Dim split As Long
   Dim counter1 As Long, counter2 As Long
   Dim temparray() As String
   Dim i As Long
   
   If start = finish Then 'only one, don't sort it
      'End Sub
   ElseIf finish - start = 1 Then 'there are two left
      If StrComp(UCase(tester(start)), UCase(tester(finish))) > 0 Then 'need to swap them
         temp = tester(start)
         tester(start) = tester(finish)
         tester(finish) = temp
      End If
      'End Sub
   Else ' need to partition and then merge
      ReDim temparray(finish - start + 1) As String
      split = (finish + start) / 2
      mergesort tester, start, split
      mergesort tester, split + 1, finish
      counter2 = split + 1
      counter1 = start
      i = 0
      While counter1 <= split And counter2 <= finish
         If StrComp(UCase(tester(counter1)), UCase(tester(counter2))) > 0 Then 'counter2 goes into merge first swap
            temparray(i) = tester(counter2)
            counter2 = counter2 + 1
         Else 'put counter 1 in first and move forward
            temparray(i) = tester(counter1)
            counter1 = counter1 + 1
         End If
         i = i + 1
      Wend
      If counter1 <= split Then 'there are still first half strings to be added
         While counter1 <= split
            temparray(i) = tester(counter1)
            i = i + 1
            counter1 = counter1 + 1
         Wend
      End If
      If counter2 <= finish Then 'there are still second half strings to be added
         While counter2 <= finish
            temparray(i) = tester(counter2)
            i = i + 1
            counter2 = counter2 + 1
         Wend
      End If
      
      For i = 0 To finish - start
         tester(i + start) = temparray(i)
      Next i
   End If
      
End Sub


Public Sub resetProgress()
   progress = 0
End Sub

Public Sub ResetTree()
   'all data from the nodes are removed and all nodes are black and not highlighted
   MousePointer = vbHourglass
   
   Dim space As Integer
   TView.Visible = False
   TView.Nodes.Clear
   'FormLoad this will be done after the user enters the species they're using so that the correct local mapps are loaded.
   
   'dbMAPPfinder.Close
   frmRank.lstGO.Clear
   frmRank.lstLocal.Clear
   Combo1.Clear
   TView.Visible = True
   frmNumbers.Hide
   frmColors.Hide
   frmRank.Hide
   frmCalculation.Clear_Form
   frmCalculation.Hide
   MousePointer = vbDefault
   dbMAPPfinder.Close
   If godataloaded Then
      dbChip.Close
   End If
   If localdataloaded Then
      dbchiplocal.Close
    End If
    If localloaded Then
      dbLocalMAPPs.Close
   End If
   'dbExpressionData.Close
      
End Sub

Public Function getName(GONode As Node) As String
   Dim name As String
   Dim space As Integer
   
   
   space = InStr(1, GONode.Text, "     ") 'a node with results
   If space <> 0 Then
      name = Left(GONode.Text, space - 1)
   Else
      name = GONode.Text
   End If
   
   getName = name
   
End Function

Public Sub setSearchCombo()
   Dim rsGO As Recordset
   Dim rsSystem As Recordset
   
   
   'we need to find out all of the systems who related to Gene Ontology for the search button.
   Set rsGO = dbMAPPfinder.OpenRecordset("SELECT SystemCode FROM Relations WHERE relatedCode = 'T'")
   While rsGO.EOF = False
      Set rsSystem = dbMAPPfinder.OpenRecordset("SELECT System from Systems WHERE SystemCOde = '" _
                                                & rsGO!systemcode & "'")
      Combo1.AddItem rsSystem!system
      rsGO.MoveNext
   Wend
   Set rsGO = dbMAPPfinder.OpenRecordset("SELECT RelatedCode FROM Relations WHERE SystemCode = 'T'")
   While rsGO.EOF = False
      Set rsSystem = dbMAPPfinder.OpenRecordset("SELECT System from Systems WHERE SystemCOde = '" _
                                                & rsGO!relatedCode & "'")
      Combo1.AddItem rsSystem!system
      rsGO.MoveNext
   Wend
End Sub

Public Function TviewCount() As Long
   TviewCount = TView.Nodes.count
End Function
