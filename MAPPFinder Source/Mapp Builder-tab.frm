VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MappBuilderForm_Normal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   Caption         =   "MAPP Builder 2.0 "
   ClientHeight    =   6555
   ClientLeft      =   2115
   ClientTop       =   825
   ClientWidth     =   8400
   FillStyle       =   0  'Solid
   Icon            =   "Mapp Builder-tab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8400
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "MAPP Information"
      Height          =   3495
      Left            =   600
      TabIndex        =   13
      Top             =   2280
      Width           =   7335
      Begin VB.TextBox author 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Adapted from Gene Ontology"
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox maintain 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "GenMAPP.org"
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox email 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "genmapp@gladstone.ucsf.edu"
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtRemarks 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "Right click here for notes."
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox txtCopyright 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1800
         Width           =   5055
      End
      Begin VB.TextBox txtNotes 
         Height          =   1095
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "Mapp Builder-tab.frx":08CA
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Author"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Maintained by"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E-mail"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Remarks"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Copyright"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Notes"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   975
      End
   End
   Begin VB.CheckBox chkOther 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox destination 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Text            =   "C:\GenMAPP\Mapps"
      Top             =   720
      Width           =   5055
   End
   Begin VB.CommandButton SelectDestination 
      Caption         =   "Select Destination Folder"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton MakeMapps 
      Caption         =   "Make MAPPs"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   5880
      Width           =   2775
   End
   Begin VB.TextBox FileName 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton SelectFile 
      Caption         =   "Select File"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "20021113"
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"Mapp Builder-tab.frx":0A3C
      Height          =   975
      Left            =   2760
      TabIndex        =   12
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Helpfile 
         Caption         =   "MAPP Builder Help"
      End
   End
End
Attribute VB_Name = "MappBuilderForm_Normal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MAPP Builder 2.0
'Created November 2002
'Beta Version released March 15, 2003
'Author: Scott Doniger with help from Steve Lawlor



Option Explicit
Const SECONDX = 0
Const SECONDY = 0
Const GENEWIDTH = 900
Const GENEHEIGHT = 300
Const ROTATION = 0
Const COLOR = -1
Const gene = "Gene"
Const ROWSEPERATOR = 2300
Const MAXGENE = 29  'makes columns of MAXGENE + 1
Const STARTX = 1500
Const STARTY = 1850
Const MAXWIDTH = 28000
Const BOARDHEIGHTMAX = 11000
Const BOARDWIDTHMAX = 15000
Const MAX_SYSTEMS = 30

Dim databasechange As Boolean
Dim blankDB As Database
Dim dbGene As Database
Dim rsGeneID As DAO.Recordset, rstempGeneID As DAO.Recordset
Dim rsBlank As DAO.Recordset
Dim rsInfo As DAO.Recordset
Dim geneId As String, currentMAPP As String, MappPath As String, csvPath As String
Dim Fsys As Object
Dim centerX, centerY As Integer
Dim Head As String, Remarks As String, labeltext As String
Dim strline As String, oldmappname As String
Dim slash As Integer, dot As Integer
Dim modify As String
Dim systemcode As String, label As String, MAPPName As String, MAPPFileName As String
Dim comma1 As Integer, comma2 As Integer, comma3 As Integer, Index As Integer
Dim overflow As Boolean
Dim CSVfile As TextStream, databaseloc As String, othersEX As TextStream
Dim addtoother As Boolean, keepex As Boolean
Dim systems(MAX_SYSTEMS, 3) As String, lastsystem As Integer, rsSystems As Recordset
      '  systems(x, 0)     Name of cataloging system. Eg: GenBank
      '  systems(x, 1)     System code. Eg: G
      '  systems(x, 2)     Additional search columns. Eg: |Gene\sBF|Orf\SBF|
      ' systems(x, 3)   Species
Dim rsRelational As Recordset, rsrelations As Recordset
Dim GenBankRelations(20, 1) As String, lastGenBankRelation As Integer
      '  Names of all GenBank relational files
      '  GenBankRelations(x, 0)    Name of cataloging system. Eg: GenBank
      '  GenBankRelations(x, 1)    System code. Eg: G


Private Sub chkOther_Click()
   If chkOther.Value = 1 Then
      addtoother = True
   Else
      addtoother = False
   End If
End Sub

Private Sub Close_Click()
    End
End Sub

Private Sub SelectDestination_Click()
    frmFindFolder.Load
    frmFindFolder.Show
End Sub

Private Sub SelectFile_Click() 'Select File
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Text Files(TAB Delimited)|*.txt"
    CommonDialog1.ShowOpen
    FileName.Text = CommonDialog1.FileName
End Sub



 Public Sub MakeMapps_Click() 'Make MAPPs
   On Error GoTo MappError
    MousePointer = vbHourglass
   Dim titlewidth As Integer, Index As Integer
   Dim newmapp As Boolean, continuelabel As String
   Dim boardwidth As Long, boardheight As Integer, windowwidth As Integer, windowheight As Integer
   Dim filetype As String
   Dim filepath As String, continue As Boolean, othergene As Boolean
   Dim tempcsv As TextStream, labelbool As Boolean, idcolumn As String
   Dim blankdbexists As Boolean, species As String, objkey As Integer
   Dim pipe As Integer, slash As Integer, i As Integer, sql As String
   Dim addtomapp As Boolean
   
   Set Fsys = CreateObject("Scripting.FileSystemObject")

   '************************************************************************'
   'Open the file containing the MAPP data
   'the file names all have CSV in them because of a previous version of MAPP builder
   'that used CSV files, rather than TAB delimited files. TABs make more sense because
   'mapp names, labels, and head/remarks have commas in them.
   labelbool = False
   overflow = False
   
   lastGenBankRelation = -1
   
   
   If InStrRev(destination.Text, "\") = Len(destination.Text) Then
pathbuilt:
      Set tempcsv = Fsys.CreateTextFile(destination.Text & "tempcsv.txt")
   Else
      Set tempcsv = Fsys.CreateTextFile(destination.Text & "\" & "tempcsv.txt")
   End If
 
   If FileName.Text = "" Then
      MsgBox "You have not selected a file to build the MAPPs from. Please do so.", vbOKOnly
      GoTo noinput
   End If
   Set CSVfile = Fsys.OpenTextFile(FileName.Text)
   While CSVfile.AtEndOfStream = False
      tempcsv.WriteLine (CSVfile.ReadLine)
   Wend
   tempcsv.WriteLine ("end") 'this is necessary becuase I'm losing the last line. I need to read in the first
   tempcsv.Close             'line before the loop so I'm off by one.
   
   If InStrRev(destination.Text, "\") = Len(destination.Text) Then
      Set CSVfile = Fsys.OpenTextFile(destination.Text & "tempcsv.txt")
   Else
      Set CSVfile = Fsys.OpenTextFile(destination.Text & "\" & "tempcsv.txt")
   End If

   strline = CSVfile.ReadLine
   If UCase(strline) <> UCase("geneId" & Chr(9) & "systemcode" & Chr(9) & "Label" _
                              & Chr(9) & "Head" & Chr(9) & "Remarks" & Chr(9) & "MappName") Then
       MsgBox "The column headings are incorrect. Please check your text file. They should be GeneID" _
       & ", SystemCode, Label, Head, Remarks, MappName. No other columns are allowed.", vbOKOnly
       CSVfile.Close
       GoTo csvFailed
   End If
   
   If addtoother = False Then
      Set othersEX = Fsys.CreateTextFile(Left(FileName.Text, Len(FileName.Text) - 4) & ".EX.txt")
      othersEX.WriteLine ("geneId" & Chr(9) & "systemcode" & Chr(9) & "Label" _
                              & Chr(9) & "Head" & Chr(9) & "Remarks" & Chr(9) & "MappName")
      keepex = False
   End If
   
   'check to see that the most recently used db is the one they want
   
   Set dbGene = OpenDatabase(databaseloc)
   
   Set rsInfo = dbGene.OpenRecordset("SELECT * FROM Info")
   
   'Start making the MAPP files
   modify = Format(Now, "Short Date")
   centerX = STARTX
   centerY = STARTY - GENEHEIGHT
   'Read the first line and parse it into the 6 data fields
   
   strline = CSVfile.ReadLine
   If CSVfile.AtEndOfStream Then
      'MsgBox "The file is blank. MAPP Builder can not build anything.", vbOKOnly
      GoTo NoNameFirst
   End If
   comma1 = InStr(1, strline, Chr(9))
   comma2 = InStr(comma1 + 1, strline, Chr(9))
   geneId = Left(strline, comma1 - 1)
   systemcode = Mid(strline, comma1 + 1, comma2 - comma1 - 1)
   comma1 = InStr(comma2 + 1, strline, Chr(9))
   label = CheckLength(Mid(strline, comma2 + 1, comma1 - comma2 - 1))
   comma2 = InStr(comma1 + 1, strline, Chr(9))
   Head = CheckLength(Mid(strline, comma1 + 1, comma2 - comma1 - 1))
   comma1 = InStr(comma2 + 1, strline, Chr(9))
   Remarks = Mid(strline, comma2 + 1, comma1 - comma2 - 1)
   If Remarks = "" Then ' cant insert null values into database. need to make them a space.
       Remarks = " "
   End If
   If Head = "" Then
       Head = " "
   End If
   If geneId = "" Then
      If UCase(systemcode) = "LABEL" Then 'labels have null geneId field
         geneId = " "
      Else
         MsgBox "You have not entered a geneId ID for a gene object. Please do so.", vbOKOnly
         GoTo NoNameFirst
      End If
   End If
   MAPPFileName = Mid(strline, comma1 + 1, Len(strline) - comma1)
   MAPPName = CheckLength(MAPPFileName)
   If MAPPName = "" Then
      MsgBox "You have not entered anything in the MappName Field. Please do so.", vbOKOnly
      GoTo NoNameFirst
   End If
      
   currentMAPP = MAPPFileName
   If InStrRev(destination.Text, "\") = Len(destination.Text) Then
      MappPath = destination.Text & currentMAPP & ".mapp"
   Else
      MappPath = destination.Text & "\" & currentMAPP & ".mapp"
   End If
  
   MappPath = fixPath(MappPath)
   
   
   Fsys.CopyFile App.Path & "\" & "MAPPTmpl.gtp", MappPath, True
   
   'taken from Steve's version 2 code
   Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems ORDER BY System", dbOpenForwardOnly)
   lastsystem = -1
   Do Until rsSystems.EOF
      If VarType(rsSystems!Date) <> vbNull Or rsSystems!species <> vbNull Then     'Supported system
         '  Other system is always supported, date or not
         lastsystem = lastsystem + 1
         systems(lastsystem, 0) = rsSystems!system
         systems(lastsystem, 1) = rsSystems!systemcode
         systems(lastsystem, 3) = rsSystems!species
         slash = InStr(1, rsSystems!columns, "\S", vbTextCompare)
         Do While slash
            pipe = InStrRev(rsSystems!columns, "|", slash)
            systems(lastsystem, 2) = systems(lastsystem, 2) _
                                   & Mid(rsSystems!columns, pipe, slash - pipe + 2) & "|"
            slash = InStr(slash + 1, rsSystems!columns, "\S", vbTextCompare)
         Loop
      End If
      rsSystems.MoveNext
   Loop
   
   Set blankDB = OpenDatabase(MappPath)
   blankdbexists = True
   objkey = 1
   While strline <> "end" 'step through every row of the data file and add it to the
   '                                correct MAPP file
      i = 0
      addtomapp = True
      othergene = False
      If currentMAPP = MAPPFileName Then
         centerY = centerY + GENEHEIGHT 'add the height to make a vertical column.
         If centerY > (STARTY + (GENEHEIGHT * MAXGENE)) Then 'columns of MAXGENE
            centerX = centerX + ROWSEPERATOR
            If centerX > MAXWIDTH Then
               continuelabel = "This MAPP is continued on " & currentMAPP & " 2."
               blankDB.Execute "INSERT INTO Objects (ID, ObjKey, SystemCode, Type, CenterX, CenterY, SecondX," _
                          & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks)" _
                          & "VALUES ('Arial', '" & objkey & "', '" & Chr(1) & "', 'Label', " _
                          & centerX & ", " & centerY & ", 8, 0, " & (Len(label) * 100) & ", 285, 0, 0, '" _
                          & TextToSql(continuelabel) & "', '" & TextToSql(Head) & "', '" & TextToSql(Remarks) & "')"
               
               centerX = centerX - ROWSEPERATOR 'put centerX back so that adding the legend doesn't overflow
               continue = HandleOverflow(2)
                  If continue Then
                     GoTo CloseMapp
                  Else
                     GoTo NoName
                  End If
            End If
            centerY = STARTY
         End If
         'select the appropriate table from Systems
         'select the primaryID from that system.
         'if it's not there add it to other or add it to the exception file.
            
         If UCase(systemcode) = "LABEL" Then ' this is a label
            If labelbool Then 'previous line was a label, only move down half a row
               centerY = centerY - (GENEHEIGHT / 2) 'labels don't need to be that far apart
            End If
            blankDB.Execute "INSERT INTO Objects (ID, ObjKey, SystemCode, Type, CenterX, CenterY, SecondX," _
                          & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks)" _
                          & "VALUES ('Arial', '" & objkey & "', '" & Chr(1) & "', 'Label', " _
                          & centerX & ", " & centerY & ", 8, 0, " & (Len(label) * 100) & ", 285, 0, 0, '" _
                          & TextToSql(label) & "', '" & TextToSql(Head) & "', '" & TextToSql(Remarks) & "')"
            labelbool = True
         Else 'it's a gene do the gene stuff
            Do Until systemcode = systems(i, 1) Or i > lastsystem
               i = i + 1
            Loop
            If i <= lastsystem Then '-------------------------------------------------Try System For Gene
               Set rsGeneID = dbGene.OpenRecordset( _
                             "SELECT ID FROM " & systems(i, 0) & " WHERE ID = '" & geneId & "'")
               species = systems(i, 3)
               If rsGeneID.EOF Then                                                         'ID not found
                  'I think that this will never get called since we added all related Genbanks to the GB
                  'table, but it's still here in case.
                  If systemcode = "G" Then '_____________________________________Special Case For GenBank
                 '  For GenBank, try all relational files, too
                    If lastGenBankRelation = -1 Then '..............Get List Of GenBank Relational Files
                        Set rsrelations = dbGene.OpenRecordset( _
                                          "SELECT * FROM Relations WHERE RelatedCode = 'G'")
                       '  Assumes that GenBank is always the dependent system. This is true of
                       '  all but the GenBank-GenBank table but all those genes are in the
                        '  GenBank table itself.
                       Do Until rsrelations.EOF
                           '  Search through Relations table to find all System tables related to GenBank
                           '  Do not include any unsupported tables
                          If rsrelations!Relation <> "GenBank-GenBank" Then
                             For i = 0 To lastsystem                   'Check to see if supported system
                                If rsrelations!systemcode = systems(i, 1) Then   'Only supported systems
                                   lastGenBankRelation = lastGenBankRelation + 1
                                   GenBankRelations(lastGenBankRelation, 0) = rsrelations!Relation
                                   GenBankRelations(lastGenBankRelation, 1) = rsrelations!systemcode
                                   Exit For
                                End If
                             Next i
                          End If
                          rsrelations.MoveNext
                        Loop
                       If lastGenBankRelation = -1 Then                              'No relations found
                          lastGenBankRelation = -2                       'Don't enter this routine again
                       End If
                     End If
                     For i = 0 To lastGenBankRelation                       'Search each Relational table
                        Set rsRelational = dbGene.OpenRecordset( _
                              "SELECT * FROM [" & GenBankRelations(i, 0) & "]" & _
                              "    WHERE Related = '" & geneId & "'")
                        If Not rsRelational.EOF Then                'Found GenBank ID in relational table
                           Exit For                              'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                        End If
                     Next i
                     If i > lastGenBankRelation Then '.Not Found In Any Existing Table Related To GenBank
                        If addtoother Then
                           othergene = True
                        Else
                           othersEX.WriteLine (strline & Chr(9) & "ID is not in genbank or related tables.")
                           keepex = True
                           centerY = centerY - GENEHEIGHT
                           addtomapp = False
                        End If
                     End If
                  ElseIf systems(i, 2) <> "" Then '___________________________________Check Secondary IDs
                     slash = InStr(1, systems(i, 2), "\S", vbTextCompare)
                     Do While slash
                        pipe = InStrRev(systems(i, 2), "|", slash)
                        idcolumn = Mid(systems(i, 2), pipe + 1, slash - pipe - 1)
                        If Mid(systems(i, 2), slash + 1, 1) = "s" Then               'Single ID, eg: P123
                           sql = "SELECT ID FROM " & systems(i, 0) & _
                                 "   WHERE [" & idcolumn & "] = '" & geneId & "'"
                        Else                                             'Multiple IDs, eg: |P123|P456|P789|
                           sql = "SELECT ID FROM " & systems(i, 0) & _
                                "   WHERE [" & idcolumn & "] LIKE '*|" & geneId & "|*'"
                        End If
                        Set rsGeneID = dbGene.OpenRecordset(sql)
                        If Not rsGeneID.EOF Then
                           Exit Do                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                        End If
                        slash = InStr(slash + 1, systems(i, 2), "\S", vbTextCompare)
                     Loop
                     If rsGeneID.EOF Then
                        '  Because there will always be a slash, there will always be at least one
                        '  query. if any query results in a hit, the loop is exited and EOF will
                        '  be false.
                        If addtoother Then
                           othergene = True
                        Else
                           othersEX.WriteLine (strline & Chr(9) & "ID is not found in geneId" _
                                       & " or secondary accession numbers for this system.")
                           keepex = True
                           centerY = centerY - GENEHEIGHT
                           addtomapp = False
                        End If
                     End If
                  
                  Else 'not found and no secondary ids to check
                     If addtoother Then
                        othergene = True
                     Else
                        othersEX.WriteLine (strline & Chr(9) & "ID not found in this system.")
                        keepex = True
                        centerY = centerY - GENEHEIGHT
                        addtomapp = False
                     End If
                  End If
               End If
                  
            Else  '-----------------------------------------------------------System Not In Gene Database
               If addtoother Then
                  othergene = True
                  species = "unknown"
               Else
                  othersEX.WriteLine (strline & Chr(9) & "This sytem code is not supported.")
                  keepex = True
                  centerY = centerY - GENEHEIGHT
                  addtomapp = False
               End If
            End If
      
      'rsGeneID now holds the ID or it's EOF and it needs to be added to other.
            
            If othergene Then 'you need to add the gene to the other table and then the MAPP
               dbGene.Execute "INSERT INTO Other (ID, SystemCode, Species, Local, [Date]) " _
                              & "VALUES ('" & geneId & "', '" & systemcode & "', '" _
                              & species & "', 'MAPPBuilder', '" & Format(Now, "Short Date") & "')"
                      
               blankDB.Execute "INSERT INTO Objects (ID, ObjKey, SystemCode, Type, CenterX, CenterY, SecondX," _
                             & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks)" _
                                 & "VALUES ('" & geneId & "', '" & objkey & "', 'O', 'Gene', " _
                                 & centerX & ", " & centerY & ", " & SECONDX & ", " & SECONDY & ", " _
                                 & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                 & COLOR & ", '" & TextToSql(label) & "', '" & TextToSql(Head) & "', '" _
                                 & TextToSql(Remarks) & "')"
            ElseIf addtomapp Then 'it was found. the gene exists add it to the mapp!
               blankDB.Execute "INSERT INTO Objects (ID, ObjKey, SystemCode, Type, CenterX, CenterY, SecondX," _
                          & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks)" _
                          & "VALUES ('" & rsGeneID!id & "', '" & objkey & "', '" & systems(i, 1) & "', 'Gene', " _
                          & centerX & ", " & centerY & ", " & SECONDX & ", " & SECONDY & ", " _
                          & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                          & COLOR & ", '" & TextToSql(label) & "', '" & TextToSql(Head) & "', '" _
                          & TextToSql(Remarks) & "')"
            End If
            labelbool = False
         End If
         'parse the next line of the csv file
         strline = CSVfile.ReadLine
         If strline <> "end" Then 'a line containing end has been added to the csv file.
            comma1 = InStr(1, strline, Chr(9))
            comma2 = InStr(comma1 + 1, strline, Chr(9))
            geneId = Left(strline, comma1 - 1)
            systemcode = Mid(strline, comma1 + 1, comma2 - comma1 - 1)
            comma1 = InStr(comma2 + 1, strline, Chr(9))
            label = CheckLength(Mid(strline, comma2 + 1, comma1 - comma2 - 1))
            comma2 = InStr(comma1 + 1, strline, Chr(9))
            Head = CheckLength(Mid(strline, comma1 + 1, comma2 - comma1 - 1))
            comma1 = InStr(comma2 + 1, strline, Chr(9))
            Remarks = Mid(strline, comma2 + 1, comma1 - comma2 - 1)
            If Remarks = "" Then ' cant insert null values into database. need to make them a space.
               Remarks = " "
            End If
            If Head = "" Then
               Head = " "
            End If
            If geneId = "" Then
               If UCase(systemcode) = "LABEL" Then 'labels have null geneId field
                  geneId = " "
               Else
                  MsgBox "You have not entered a geneId ID for a gene object. Please do so.", vbOKOnly
                  GoTo NoName
               End If
            End If
            MAPPFileName = Mid(strline, comma1 + 1, Len(strline) - comma1)
            oldmappname = MAPPName
            MAPPName = TextToSql(CheckLength(MAPPFileName))
            If MAPPName = "" Then
               MsgBox "You have not entered anything in the MappName Field. Please do so.", vbOKOnly
               GoTo NoName
            End If
            objkey = objkey + 1
         End If 'don't parse the last line
CloseMapp:
      Else
         blankDB.Execute "INSERT INTO Objects (ObjKey, Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                & "Rotation, Color, Remarks) VALUES (" & objkey + 1 & ", 'InfoBox', 76.5, 325, 0, 0, 45, 675, 0, -1, " _
                & "'<font name=Arial, size=9><p><i><b>Author:</b></i>" & TextToSql(author.Text) _
               & "</p><p><i><b>Maintained by:</b></i>" & TextToSql(maintain.Text) _
                  & "</p><p><i><b>Last modified:</b></i> " & TextToSql(modify) & "</p>')"
         blankDB.Execute "INSERT INTO Objects (objkey, Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                     & "Rotation, Color) VALUES (" & objkey + 2 & ", 'Legend', " & centerX + 1520 & ", 1778, 0, 0, 0, 0, 0, -1)"
         titlewidth = Len(currentMAPP) * 150
         If centerX = STARTX Then 'for a mapp with one column and a long title, you need to accomodate the title
            If titlewidth > 5800 Then
               boardwidth = titlewidth
            Else
               boardwidth = 5800 'need to accomodate the info box and legend. minimum width possible
            End If
         Else
            boardwidth = centerX + 4475
         End If
         
         If boardwidth < BOARDWIDTHMAX Then
            windowwidth = boardwidth + 300
         Else
            windowwidth = BOARDWIDTHMAX
         End If
      
         If centerX > STARTX Then 'more than one column
            boardheight = STARTY + (MAXGENE * GENEHEIGHT) + 200
         Else 'one incomplete column
            boardheight = centerY + 200
            If boardheight < 3700 Then
               boardheight = 3700
            End If
         End If
         If boardheight < BOARDHEIGHTMAX Then
            windowheight = boardheight + 1100
         Else
            windowheight = BOARDHEIGHTMAX
         End If
         
         blankDB.Execute "DELETE * FROM Info"
         blankDB.Execute "INSERT INTO Info (Title, MAPP, Version, Author, Maint, Email, Copyright, Modify, Remarks, BoardWidth, " _
                     & "BoardHeight, WindowWidth, WindowHeight, Notes) " _
                     & "VALUES ('" & TextToSql(oldmappname) & "', '', '" & rsInfo![Version] & "', '" _
                     & TextToSql(author.Text) & "', '" & TextToSql(maintain.Text) & "', '" & TextToSql(email.Text) & "', '" _
                     & TextToSql(txtCopyright.Text) & "', '" & modify & "', '" & TextToSql(txtRemarks.Text) & "', " _
                     & boardwidth & ", " & boardheight & ", " & windowwidth & ", " & windowheight _
                     & ", '" & TextToSql(txtNotes.Text) & "')"
         blankDB.Close
         centerX = STARTX
         centerY = STARTY - GENEHEIGHT
         currentMAPP = MAPPFileName
         If InStrRev(destination.Text, "\") = Len(destination.Text) Then
            MappPath = destination.Text & currentMAPP & ".mapp"
         Else
            MappPath = destination.Text & "\" & currentMAPP & ".mapp"
         End If
         If Dir(MappPath) = currentMAPP & ".mapp" Then
            frmOverWrite.setMAPPName currentMAPP
            frmOverWrite.Show vbModal
            If frmOverWrite.overwrite = False Then
               GoTo NoName
            End If
         End If
         MappPath = fixPath(MappPath)
         
         Fsys.CopyFile App.Path & "\" & "MAPPTmpl.gtp", MappPath, True
         Set blankDB = OpenDatabase(MappPath)
         objkey = 1
      End If
   Wend
    'need to to the else case one more time for the last mapp
    blankDB.Execute "INSERT INTO Objects (ObjKey, Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                & "Rotation, Color, Remarks) VALUES (" & objkey + 1 & ", 'InfoBox', 76.5, 325, 0, 0, 45, 675, 0, -1, " _
                & "'<font name=Arial, size=9><p><i><b>Author:</b></i>" & TextToSql(author.Text) _
               & "</p><p><i><b>Maintained by:</b></i>" & TextToSql(maintain.Text) _
                  & "</p><p><i><b>Last modified:</b></i> " & TextToSql(modify) & "</p>')"
   blankDB.Execute "INSERT INTO Objects (objkey, Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                     & "Rotation, Color) VALUES (" & objkey + 2 & ", 'Legend', " & centerX + 1520 & ", 1778, 0, 0, 0, 0, 0, -1)"
   
    titlewidth = Len(currentMAPP) * 140
    If centerX = STARTX Then 'for a mapp with one column and a long title, you need to accomodate the title
        If titlewidth > 5800 Then
            boardwidth = titlewidth
        Else
            boardwidth = 5800 'need to accomodate the info box and legend. minimum width possible
        End If
    Else
        boardwidth = centerX + 4475
    End If
            
    If boardwidth < BOARDWIDTHMAX Then
        windowwidth = boardwidth + 300
    Else
        windowwidth = BOARDWIDTHMAX
    End If
            
    If centerX > STARTX Then 'more than one column
        boardheight = STARTY + (MAXGENE * GENEHEIGHT) + 200
    Else 'one incomplete column
        boardheight = centerY + 200
        If boardheight < 3700 Then
            boardheight = 3700
        End If
    End If
           
   If boardheight < BOARDHEIGHTMAX Then
       windowheight = boardheight + 1100
   Else
       windowheight = BOARDHEIGHTMAX
   End If
            
   blankDB.Execute "DELETE * FROM Info"
   blankDB.Execute "INSERT INTO Info (Title, MAPP, Version, Author, Maint, Email, Copyright, Modify, Remarks, BoardWidth, " _
                          & "BoardHeight, WindowWidth, WindowHeight, Notes) " _
                          & "VALUES ('" & TextToSql(MAPPName) & "', '', '" & rsInfo![Version] & "', '" _
                          & TextToSql(author.Text) & "', '" & TextToSql(maintain.Text) & "', '" & TextToSql(email.Text) & "', '" _
                          & TextToSql(txtCopyright.Text) & "', '" & modify & "', '" & TextToSql(txtRemarks.Text) & "', " _
                          & boardwidth & ", " & boardheight & ", " & windowwidth & ", " & windowheight _
                          & ", '" & TextToSql(txtNotes.Text) & "')"
   blankDB.Close
    
   dbGene.Close
   CSVfile.Close
   If addtoother = False Then
      othersEX.Close
      Fsys.DeleteFile (Left(FileName.Text, Len(FileName.Text) - 4) & ".EX.txt")
   End If
   If InStrRev(destination.Text, "\") = Len(destination.Text) Then
      Kill destination.Text & "tempcsv.txt"
   Else
      Kill destination.Text & "\" & "tempcsv.txt"
   End If
   
    
   If overflow Then
      MsgBox "Some MAPPs have overflowed the maximum number of genes per MAPP. Overflow MAPPs labelled," _
      & " MappName2 (or 3, etc.) have been created and are in the destination directory.", vbOKOnly
   End If
csvFailed:
   MousePointer = vbDefault
   Exit Sub
    
MappError:
   Select Case Err.Number
   Case 5
        MsgBox "You must enter a filename to save the new MAPP to.", vbOKOnly
   Case 3134 'no geneId value
        MsgBox "MAPP Builder has encountered a null geneId ID. You must enter a geneId ID " _
               & " for every gene (Labels are excluded)", vbOKOnly
   Case 94
        MsgBox "The input file contains a blank MAPP Name. Each entry must have a mapp name. Please fix this and rerun the program.", vbOKOnly
   Case 70
      MsgBox "Permission denied. Check to see if you have MAPP Builder file or Exception file open. They must be closed before you can run MAPP Builder.", vbOKOnly
      GoTo NoNameFirst
   Case 76 'path not found.
      'need to make a directory for the MAPP
    
      frmConfigure.AddFolder (destination.Text)
      Resume pathbuilt
   'Case Else
      'FatalError "MakeMapps", Err.Description
   End Select
NoName:
   If blankdbexists Then
   blankDB.Close
   End If
NoNameFirst: 'blankDB hasn't been created yet if the first line of data has an error, so you jump here instead.
   If addtoother = False Then
      othersEX.Close
      Fsys.DeleteFile (Left(FileName.Text, Len(FileName.Text) - 4) & ".EX.txt")
   End If
   CSVfile.Close
   If InStrRev(destination.Text, "\") = Len(destination.Text) Then
      CSVfile.Close
      Kill destination.Text & "\" & "tempcsv.txt"
   Else
      Fsys.deleteTextFile (destination.Text & "\" & "tempcsv.txt")
   End If
   
   dbGene.Close
   
   
noinput:
   MousePointer = vbDefault
   'MappBuilderForm_Normal.Hide
End Sub





Function TextToSql(txt As String) As String '**************************** Makes Text SQL Compatible
    Dim Index As Integer                     'copied from GenMAPP 1.0 Source code
    Dim sql As String
   
    sql = txt
    For Index = 1 To Len(txt)
      Select Case Mid(txt, Index, 1)
      Case "'"                            'Convert single quote to typographer's close single quote
         Mid(sql, Index, 1) = Chr(146)
      Case "!"                            'Convert single quote to typographer's close single quote
         Mid(sql, Index, 1) = Chr(32)
      Case Chr(34)
          Mid(sql, Index, 1) = Chr(146)
      Case "$"
          Mid(sql, Index, 1) = Chr(32)
      Case Else
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
      Case "*"
         Mid(Path, Index, 1) = Chr(32)
      Case "?"
         Mid(Path, Index, 1) = Chr(32)
      Case Chr(34)
          Mid(Path, Index, 1) = Chr(146)
      Case "<"
          Mid(Path, Index, 1) = Chr(32)
      Case ">"
          Mid(Path, Index, 1) = Chr(32)
      Case "|"
          Mid(Path, Index, 1) = Chr(32)
      Case Else
      End Select
   Next Index
    For Index = 3 To Len(Path) 'start at three to ignore c:
      Select Case Mid(Path, Index, 1)
      Case ":"                            'Convert single quote to typographer's close single quote
         Mid(Path, Index, 1) = Chr(32)
      End Select
   Next Index
   Path = TextToSql(Path)
   fixPath = Path

End Function

Public Function HandleOverflow(mappnum As Integer) As Boolean
   On Error GoTo MappError
   Dim blankDB2 As Database
   Dim currentMAPP2 As String, oldMAPP As String
   Dim titlewidth As Integer, othergene As Boolean
   Dim newmapp As Boolean, continue As Boolean, continuelabel As String
   Dim boardwidth As Long, boardheight As Integer, windowwidth As Integer, windowheight As Integer
   Dim centerY As Integer, centerX As Long, labelbool As Boolean
   Dim pipe As Integer, slash As Integer, i As Integer, sql As String
   Dim addtomapp As Boolean, species As String, objkey As Integer, idcolumn As String
   
   labelbool = False
   overflow = True
   objkey = 1
   currentMAPP2 = MAPPFileName & Str(mappnum)
   If InStrRev(destination.Text, "\") = Len(destination.Text) Then
      MappPath = destination.Text & currentMAPP2 & ".mapp"
   Else
      MappPath = destination.Text & "\" & currentMAPP2 & ".mapp"
   End If
   MappPath = fixPath(MappPath)
   
   Fsys.CopyFile App.Path & "\" & "MAPPTmpl.gtp", MappPath, True
   Set blankDB2 = OpenDatabase(MappPath)
   centerY = STARTY - GENEHEIGHT
   centerX = STARTX
   While CSVfile.AtEndOfStream = False And newmapp = False
      i = 0
      addtomapp = True
      If currentMAPP = MAPPFileName Then
         centerY = centerY + GENEHEIGHT 'add the height to make a vertical column.
         If centerY > (STARTY + (GENEHEIGHT * MAXGENE)) Then 'columns of MAXGENE
            centerX = centerX + ROWSEPERATOR
            If centerX > MAXWIDTH Then
               continuelabel = "This MAPP is continued on " & currentMAPP & " " & mappnum + 1 & "."
               blankDB2.Execute "INSERT INTO Objects (ID, ObjKey, SystemCode, Type, CenterX, CenterY, SecondX," _
                          & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks)" _
                          & "VALUES ('Arial', '" & objkey & "', '" & Chr(1) & "', 'Label', " _
                          & centerX & ", " & centerY & ", 8, 0, " & (Len(label) * 100) & ", 285, 0, 0, '" _
                          & TextToSql(continuelabel) & "', '" & TextToSql(Head) & "', '" & TextToSql(Remarks) & "')"
               centerX = centerX - ROWSEPERATOR 'put centerX back so that adding the legend doesn't overflow
               continue = HandleOverflow(mappnum + 1)  'recursively handle the overflow. Can make mapps forever.
                  If continue Then
                     GoTo closeMAPP2
                  Else
                     GoTo NoName
                  End If
               End If
               centerY = STARTY
           End If
           If UCase(systemcode) = "LABEL" Then ' this is a label
            If labelbool Then 'previous line was a label, only move down half a row
               centerY = centerY - (GENEHEIGHT / 2) 'labels don't need to be that far apart
            End If
            blankDB2.Execute "INSERT INTO Objects (ID, ObjKey, SystemCode, Type, CenterX, CenterY, SecondX," _
                          & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks)" _
                          & "VALUES ('Arial', '" & objkey & "', '" & Chr(1) & "', 'Label', " _
                          & centerX & ", " & centerY & ", 8, 0, " & (Len(label) * 100) & ", 285, 0, 0, '" _
                          & TextToSql(label) & "', '" & TextToSql(Head) & "', '" & TextToSql(Remarks) & "')"
            labelbool = True
         Else 'it's a gene do the gene stuff
            Do Until systemcode = systems(i, 1) Or i > lastsystem
               i = i + 1
            Loop
            If i <= lastsystem Then '-------------------------------------------------Try System For Gene
               Set rsGeneID = dbGene.OpenRecordset( _
                             "SELECT ID FROM " & systems(i, 0) & " WHERE ID = '" & geneId & "'")
               species = systems(i, 3)
               If rsGeneID.EOF Then                                                         'ID not found
                  'I think that this will never get called since we added all related Genbanks to the GB
                  'table, but it's still here in case.
                  If systemcode = "G" Then '_____________________________________Special Case For GenBank
                 '  For GenBank, try all relational files, too
                    If lastGenBankRelation = -1 Then '..............Get List Of GenBank Relational Files
                        Set rsrelations = dbGene.OpenRecordset( _
                                          "SELECT * FROM Relations WHERE RelatedCode = 'G'")
                       '  Assumes that GenBank is always the dependent system. This is true of
                       '  all but the GenBank-GenBank table but all those genes are in the
                        '  GenBank table itself.
                       Do Until rsrelations.EOF
                           '  Search through Relations table to find all System tables related to GenBank
                           '  Do not include any unsupported tables
                          If rsrelations!Relation <> "GenBank-GenBank" Then
                             For i = 0 To lastsystem                   'Check to see if supported system
                                If rsrelations!systemcode = systems(i, 1) Then   'Only supported systems
                                   lastGenBankRelation = lastGenBankRelation + 1
                                   GenBankRelations(lastGenBankRelation, 0) = rsrelations!Relation
                                   GenBankRelations(lastGenBankRelation, 1) = rsrelations!systemcode
                                   Exit For
                                End If
                             Next i
                          End If
                          rsrelations.MoveNext
                        Loop
                       If lastGenBankRelation = -1 Then                              'No relations found
                          lastGenBankRelation = -2                       'Don't enter this routine again
                       End If
                     End If
                     For i = 0 To lastGenBankRelation                       'Search each Relational table
                        Set rsRelational = dbGene.OpenRecordset( _
                              "SELECT * FROM [" & GenBankRelations(i, 0) & "]" & _
                              "    WHERE Related = '" & geneId & "'")
                        If Not rsRelational.EOF Then                'Found GenBank ID in relational table
                           Exit For                              'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                        End If
                     Next i
                     If i > lastGenBankRelation Then '.Not Found In Any Existing Table Related To GenBank
                        If addtoother Then
                           othergene = True
                        Else
                           othersEX.WriteLine (strline & Chr(9) & "ID is not in genbank or related tables.")
                           keepex = True
                           centerY = centerY - GENEHEIGHT
                           addtomapp = False
                        End If
                     End If
                  ElseIf systems(i, 2) <> "" Then '___________________________________Check Secondary IDs
                     slash = InStr(1, systems(i, 2), "\S", vbTextCompare)
                     Do While slash
                        pipe = InStrRev(systems(i, 2), "|", slash)
                        idcolumn = Mid(systems(i, 2), pipe + 1, slash - pipe - 1)
                        If Mid(systems(i, 2), slash + 1, 1) = "s" Then               'Single ID, eg: P123
                           sql = "SELECT ID FROM " & systems(i, 0) & _
                                 "   WHERE [" & idcolumn & "] = '" & geneId & "'"
                        Else                                             'Multiple IDs, eg: |P123|P456|P789|
                           sql = "SELECT ID FROM " & systems(i, 0) & _
                                "   WHERE [" & idcolumn & "] LIKE '*|" & geneId & "|*'"
                        End If
                        Set rsGeneID = dbGene.OpenRecordset(sql)
                        If Not rsGeneID.EOF Then
                           Exit Do                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                        End If
                        slash = InStr(slash + 1, systems(i, 2), "\S", vbTextCompare)
                     Loop
                     If rsGeneID.EOF Then
                        '  Because there will always be a slash, there will always be at least one
                        '  query. if any query results in a hit, the loop is exited and EOF will
                        '  be false.
                        If addtoother Then
                           othergene = True
                        Else
                           othersEX.WriteLine (strline & Chr(9) & "ID is not found in geneId" _
                                       & " or secondary accession numbers for this system.")
                           keepex = True
                           centerY = centerY - GENEHEIGHT
                           addtomapp = False
                        End If
                     End If
                  
                  Else 'not found and no secondary ids to check
                     If addtoother Then
                        othergene = True
                     Else
                        othersEX.WriteLine (strline & Chr(9) & "ID not found in this system.")
                        keepex = True
                        centerY = centerY - GENEHEIGHT
                        addtomapp = False
                     End If
                  End If
               End If
                  
            Else  '-----------------------------------------------------------System Not In Gene Database
               If addtoother Then
                  othergene = True
                  species = "unknown"
               Else
                  othersEX.WriteLine (strline & Chr(9) & "This sytem code is not supported.")
                  keepex = True
                  centerY = centerY - GENEHEIGHT
                  addtomapp = False
               End If
            End If
      
      'rsGeneID now holds the ID or it's EOF and it needs to be added to other.
            
            If othergene Then 'you need to add the gene to the other table and then the MAPP
               dbGene.Execute "INSERT INTO Other (ID, SystemCode, Species, Local, [Date]) " _
                              & "VALUES ('" & geneId & "', '" & systemcode & "', '" _
                              & species & "', 'MAPPBuilder', '" & Format(Now, "Short Date") & "')"
                      
               blankDB2.Execute "INSERT INTO Objects (ID, ObjKey, SystemCode, Type, CenterX, CenterY, SecondX," _
                             & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks)" _
                                 & "VALUES ('" & geneId & "', '" & objkey & "', 'O', 'Gene', " _
                                 & centerX & ", " & centerY & ", " & SECONDX & ", " & SECONDY & ", " _
                                 & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                 & COLOR & ", '" & TextToSql(label) & "', '" & TextToSql(Head) & "', '" _
                                 & TextToSql(Remarks) & "')"
            ElseIf addtomapp Then 'it was found. the gene exists add it to the mapp!
               blankDB2.Execute "INSERT INTO Objects (ID, ObjKey, SystemCode, Type, CenterX, CenterY, SecondX," _
                          & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks)" _
                          & "VALUES ('" & rsGeneID!id & "', '" & objkey & "', '" & systems(i, 1) & " ', 'Gene', " _
                          & centerX & ", " & centerY & ", " & SECONDX & ", " & SECONDY & ", " _
                          & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                          & COLOR & ", '" & TextToSql(label) & "', '" & TextToSql(Head) & "', '" _
                          & TextToSql(Remarks) & "')"
            End If
            labelbool = False
         End If
           
           
           
         strline = CSVfile.ReadLine
         objkey = objkey + 1
         If strline <> "end" Then 'a line containing end has been added to the csv file.
            comma1 = InStr(1, strline, Chr(9))
            comma2 = InStr(comma1 + 1, strline, Chr(9))
            geneId = Left(strline, comma1 - 1)
            systemcode = Mid(strline, comma1 + 1, comma2 - comma1 - 1)
            comma1 = InStr(comma2 + 1, strline, Chr(9))
            label = CheckLength(Mid(strline, comma2 + 1, comma1 - comma2 - 1))
            comma2 = InStr(comma1 + 1, strline, Chr(9))
            Head = CheckLength(Mid(strline, comma1 + 1, comma2 - comma1 - 1))
            comma1 = InStr(comma2 + 1, strline, Chr(9))
            Remarks = Mid(strline, comma2 + 1, comma1 - comma2 - 1)
            If Remarks = "" Then ' cant insert null values into database. need to make them a space.
               Remarks = " "
            End If
            If Head = "" Then
               Head = " "
            End If
            If geneId = "" Then
               If UCase(systemcode) = "LABEL" Then 'labels have null geneId field
                  geneId = " "
               Else
                  MsgBox "You have not entered a geneId ID for a gene object. Please do so.", vbOKOnly
                  GoTo NoName
               End If
            End If
            MAPPFileName = Mid(strline, comma1 + 1, Len(strline) - comma1)
            oldmappname = MAPPName
            MAPPName = CheckLength(MAPPFileName)
            If MAPPName = "" Then
               MsgBox "You have not entered anything in the MappName Field. Please do so.", vbOKOnly
               GoTo NoName
            End If
         End If
      Else
           newmapp = True
      End If

   Wend
closeMAPP2:
  blankDB2.Execute "INSERT INTO Objects (Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                    & "Rotation, Color, Remarks) VALUES ('InfoBox', 76.5, 325, 0, 0, 45, 675, 0, -1, " _
                    & "'<font name=Arial, size=9><p><i><b>Author:</b></i>" & TextToSql(author.Text) _
                    & "</p><p><i><b>Maintained by:</b></i>" & TextToSql(maintain.Text) _
                    & "</p><p><i><b>Last modified:</b></i> " & TextToSql(modify) & "</p>')"
   blankDB2.Execute "INSERT INTO Objects (Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                       & "Rotation, Color) VALUES ('Legend', " & centerX + 1520 & ", 1778, 0, 0, 0, 0, 0, -1)"
   titlewidth = Len(currentMAPP) * 150
   If centerX = STARTX Then 'for a mapp with one column and a long title, you need to accomodate the title
       If titlewidth > 5800 Then
           boardwidth = titlewidth
       Else
           boardwidth = 5800 'need to accomodate the info box and legend. minimum width possible
       End If
   Else
       boardwidth = centerX + 4475
   End If
           
   If boardwidth < BOARDWIDTHMAX Then
       windowwidth = boardwidth + 300
   Else
       windowwidth = BOARDWIDTHMAX
   End If
           
   If centerX > STARTX Then 'more than one column
       boardheight = STARTY + (MAXGENE * GENEHEIGHT) + 200
   Else 'one incomplete column
       boardheight = centerY + 200
       If boardheight < 3700 Then
          boardheight = 3700
       End If
   End If
           
   If boardheight < BOARDHEIGHTMAX Then
       windowheight = boardheight + 1100
   Else
       windowheight = BOARDHEIGHTMAX
   End If
           
   blankDB2.Execute "DELETE * FROM Info"
   blankDB2.Execute "INSERT INTO Info (Title, MAPP, Version, Author, Maint, Email, Copyright, Modify, Remarks, BoardWidth, " _
                          & "BoardHeight, WindowWidth, WindowHeight, Notes) " _
                          & "VALUES ('" & TextToSql(oldmappname) & "', '', '" & rsInfo![Version] & "', '" _
                          & TextToSql(author.Text) & "', '" & TextToSql(maintain.Text) & "', '" & TextToSql(email.Text) & "', '" _
                          & TextToSql(txtCopyright.Text) & "', '" & modify & "', '" & TextToSql(txtRemarks.Text) & "', " _
                          & boardwidth & ", " & boardheight & ", " & windowwidth & ", " & windowheight _
                          & ", '" & TextToSql(txtNotes.Text) & "')"
   blankDB2.Close
   HandleOverflow = True
 Exit Function
MappError:
   Select Case Err.Number
   Case 5
       MsgBox "You must enter a filename to save the new MAPP to.", vbOKOnly
   Case 3078
       MsgBox "You have entered an incorrect table name. Check your Access file and enter the correct name.", vbOKOnly
   Case 3134 'no geneId value
        MsgBox "MAPP Builder has encountered a null geneId ID. You must enter a geneId ID " _
               & " for every gene (Labels are excluded)", vbOKOnly
   Case 94
       MsgBox "The input file contains a blank MAPP Name. Each entry must have a mapp name. Please fix this and rerun the program.", vbOKOnly
   Case Else
   MsgBox "An Error has occured. Please check the format of the CSV file and then try again.", vbOKOnly
   End Select
NoName:
   blankDB2.Close
   'dbGene.Close
   'CSVfile.Close
   HandleOverflow = False
End Function


Public Sub setMappPath(Path As String)
    MappPath = Path
    destination.Text = MappPath
End Sub

Public Sub setBaseMapp(baseMAPP As String, datalocation As String)
      destination.Text = baseMAPP
      databaseloc = datalocation
End Sub

Public Function CheckLength(line As String) As String
   'MAPP files limit filename and label and head to 50 characters.
   If line <> "" Then
      If Len(line) > 50 Then
         CheckLength = Left(line, 50)
         Exit Function
      End If
   End If
   CheckLength = line

End Function

Public Sub setFileName(File As String)
   FileName = File
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



Public Function fixName(name As String) As String
   fixName = Replace(name, "\", "")
End Function
