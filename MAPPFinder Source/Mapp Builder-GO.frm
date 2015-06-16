VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MappBuilderForm_Normal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   Caption         =   "MAPP Builder 1.0 Beta"
   ClientHeight    =   6255
   ClientLeft      =   2115
   ClientTop       =   825
   ClientWidth     =   8295
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8295
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "MAPP Information"
      Height          =   3495
      Left            =   480
      TabIndex        =   13
      Top             =   1920
      Width           =   7335
      Begin VB.TextBox author 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Text            =   "Adapted from Gene Ontology"
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox maintain 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "GenMAPP.org"
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox email 
         Height          =   285
         Left            =   1440
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
         Text            =   "Right click here for Notes."
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox txtCopyright 
         Height          =   285
         Left            =   1440
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
         Text            =   "Mapp Builder-GO.frx":0000
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
      Left            =   2400
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
      Caption         =   "Make MAPP"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox FileName 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton SelectFile 
      Caption         =   "Select CSV File"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Add gene identifications that are not found in the GenMAPP Database to ""Other"" category?"
      Height          =   615
      Left            =   2760
      TabIndex        =   12
      Top             =   1320
      Width           =   5055
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
'MAPP Maker 1.1


Option Explicit
Const SECONDX = 0
Const SECONDY = 0
Const GENEWIDTH = 900
Const GENEHEIGHT = 300
Const ROTATION = 0
Const COLOR = -1
Const GENE = "Gene"
Const ROWSEPERATOR = 3800
Const MAXGENE = 29  'makes columns of MAXGENE + 1
Const STARTX = 1500
Const STARTY = 2450
Const MAXWIDTH = 28000
Const BOARDHEIGHTMAX = 11000
Const BOARDWIDTHMAX = 15000
Dim databasechange As Boolean
Dim blankDB As Database
Dim MAPPDB As Database
Dim rsGeneID As DAO.Recordset, rstempGeneID As DAO.Recordset
Dim rsBlank As DAO.Recordset
Dim rsinfo As DAO.Recordset
Dim primary As String, currentMAPP As String, MappPath As String, csvPath As String
Dim fsys As Object
Dim centerX, centerY As Integer
Dim Head As String, Remarks As String, labeltext As String
Dim strline As String
Dim slash As Integer, dot As Integer
Dim modify As String
Dim primaryType As String, label As String, MappName As String, MAPPFileName As String, oldMAPPName As String
Dim comma1 As Integer, comma2 As Integer, comma3 As Integer, Index As Integer
Dim overflow As Boolean
Dim CSVfile As TextStream, DatabaseLoc As String, othersEX As TextStream
Dim addothers As Boolean, keepex As Boolean


Public Sub chkOther_Click()
   If chkOther.Value = 1 Then
      addothers = True
   Else
      addothers = False
   End If
End Sub

Private Sub Close_Click()
    End
End Sub

Private Sub SelectDestination_Click()
   ' Form2.Show
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
    Dim newmapp As Boolean
    Dim boardwidth As Long, boardheight As Integer, windowwidth As Integer, windowheight As Integer
    Dim filetype As String
    Dim filepath As String, continue As Boolean
    Dim tempcsv As TextStream, labelbool As Boolean
    Set fsys = CreateObject("Scripting.FileSystemObject")
     
   '************************************************************************'
   'Open the file containing the MAPP data
   'the file names all have CSV in them because of a previous version of MAPP builder
   'that used CSV files, rather than TAB delimited files. TABs make more sense because
   'mapp names, labels, and head/remarks have commas in them.
   labelbool = False
   Set tempcsv = fsys.CreateTextFile("c:\tempcsv.txt")
   Set CSVfile = fsys.OpenTextFile(FileName.Text)
   While CSVfile.AtEndOfStream = False
      tempcsv.WriteLine (CSVfile.ReadLine)
   Wend
   tempcsv.WriteLine ("end") 'this was necessary because for some reason the file was not ending properly
   tempcsv.Close
   
   Set CSVfile = fsys.OpenTextFile("c:\tempCSV.txt")
   strline = CSVfile.ReadLine
   If UCase(strline) <> UCase("Primary" & Chr(9) & "PrimaryType" & Chr(9) & "Label" _
                              & Chr(9) & "Head" & Chr(9) & "Remarks" & Chr(9) & "MappName") Then
       MsgBox "The column headings are incorrect. Please check your CSV file. They should be Primary, PrimaryType, Label, Head, Remarks, MappName", vbOKOnly
       CSVfile.Close
       GoTo csvFailed
   End If
   
   If addothers = False Then
      Set othersEX = fsys.CreateTextFile(Left(FileName.Text, Len(FileName.Text) - 4) & ".EX.txt")
      othersEX.WriteLine ("Primary" & Chr(9) & "PrimaryType" & Chr(9) & "Label" _
                              & Chr(9) & "Head" & Chr(9) & "Remarks" & Chr(9) & "MappName")
      keepex = False
   End If
   'Open the genMAPP Database
   Set MAPPDB = OpenDatabase(DatabaseLoc)
   
   Set rsinfo = MAPPDB.OpenRecordset("SELECT * FROM Info")
   
   'Start making the MAPP files
   modify = Format(Now, "Short Date")
   centerX = STARTX
   centerY = STARTY - GENEHEIGHT
   'Read the first line and parse it into the 6 data fields
   strline = CSVfile.ReadLine
   comma1 = InStr(1, strline, Chr(9))
   comma2 = InStr(comma1 + 1, strline, Chr(9))
   primary = Left(strline, comma1 - 1)
   primaryType = Mid(strline, comma1 + 1, comma2 - comma1 - 1)
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
   If primary = "" Then
      If UCase(primaryType) = "L" Then 'labels have null primary field
         primary = " "
      Else
         MsgBox "You have not entered a Primary ID for a gene object. Please do so.", vbOKOnly
         GoTo NoNameFirst
      End If
   End If
   MAPPFileName = Mid(strline, comma1 + 1, Len(strline) - comma1)
   MappName = CheckLength(MAPPFileName)
   If MappName = "" Then
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
   fsys.CopyFile mapptmpl, MappPath, True
   Set blankDB = OpenDatabase(MappPath)
   
   While Not CSVfile.AtEndOfStream 'step through every row of the data file and add it to the
   '                                correct MAPP file
        If currentMAPP = MAPPFileName Then
            centerY = centerY + GENEHEIGHT 'add the height to make a vertical column.
            If centerY > (STARTY + (GENEHEIGHT * MAXGENE)) Then 'columns of MAXGENE
                centerX = centerX + ROWSEPERATOR
                If centerX > MAXWIDTH Then
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
            If primaryType = "G" Or primaryType = "g" Then
               labelbool = False
                Set rsGeneID = MAPPDB.OpenRecordset("SELECT GenMAPP FROM GenBank WHERE" _
                                                   & " GenBank = '" & primary & "'")
                If rsGeneID.EOF Then 'the primary ID doesn't exist in GenMAPP add it to other table
                  If addothers Then
                     Set rstempGeneID = MAPPDB.OpenRecordset("Select Other from Other Where " _
                                    & "Other = '" & primary & "'")
                     If rstempGeneID.EOF Then 'this gene hasn't already been added to the other table
                        MAPPDB.Execute "INSERT INTO Other (Other, GenMAPP, Local) VALUES ('" & _
                            primary & "', '" & primary & "', '')"
                        MAPPDB.Execute "INSERT INTO GenMAPP (GenMAPP) VALUES ('" & primary & "')"
                     End If
                     blankDB.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type," _
                           & " centerX , centerY, SecondX, SecondY, Width, Height," _
                           & " Rotation, Color, Label, Head, Remarks) VALUES " _
                           & "('" & primary & "', '" & primary & "', 'O', 'Gene', " & centerX _
                           & ", " & centerY & ", " _
                           & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " _
                           & ROTATION & " , " & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                  Else 'don't want to add others
                     othersEX.WriteLine (strline)
                     keepex = True
                     centerY = centerY - GENEHEIGHT
                  End If
                Else    'the primary ID does exist add a genbank to the objects table
                    blankDB.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                    & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & rsGeneID![GenMAPP] & "', '" _
                                    & primary & "', '" & primaryType & "', 'Gene', " & centerX & ", " & centerY & ", " _
                                    & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                    & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                End If
            ElseIf primaryType = "S" Or primaryType = "s" Then
               labelbool = False
               Set rsGeneID = MAPPDB.OpenRecordset("Select GenMAPP FROM SwissProt Where SwissName = '" _
                              & primary & "'")
               If rsGeneID.EOF Then 'a swissname wasn't found. Check SwissNo
                  Set rstempGeneID = MAPPDB.OpenRecordset("SELECT SwissName from SwissNO where SwissNo" _
                                    & "= '" & primary & "'")
                  If rstempGeneID.EOF = False Then 'you found a swissno, now get the genMAPP ID
                     Set rsGeneID = MAPPDB.OpenRecordset("SELECT GenMAPP from SwissProt where SwissName" _
                                 & "= '" & rstempGeneID![SwissName] & "'")
                     blankDB.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                    & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & rsGeneID![GenMAPP] & "', '" _
                                    & primary & "', '" & primaryType & "', 'Gene', " & centerX & ", " & centerY & ", " _
                                    & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                    & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                  Else 'no swissno or swissname, add other
                     If addothers Then
                        Set rstempGeneID = MAPPDB.OpenRecordset("Select * from Other Where " _
                                    & "Other = '" & primary & "'")
                        If rstempGeneID.EOF Then 'this gene hasn't already been added to the other table
                           MAPPDB.Execute "INSERT INTO Other (Other, GenMAPP, Local) VALUES ('" & _
                              primary & "', '" & primary & "', '')"
                           MAPPDB.Execute "INSERT INTO GenMAPP (GenMAPP) VALUES ('" & primary & "')"
                        End If
                        blankDB.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                   & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & primary & "', '" _
                                   & primary & "', 'O', 'Gene', " & centerX & ", " & centerY & ", " _
                                   & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                   & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                     Else
                        othersEX.WriteLine (strline)
                        keepex = True
                        centerY = centerY - GENEHEIGHT
                     End If
                  End If
               Else    'the primary ID exists add it to the objects table
                    blankDB.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                    & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & rsGeneID![GenMAPP] & "', '" _
                                    & primary & "', '" & primaryType & "', 'Gene', " & centerX & ", " & centerY & ", " _
                                    & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                    & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
               End If
            ElseIf primaryType = "O" Or primaryType = "o" Then    'it's not g and not s, so it should be other
               labelbool = False
               If addothers Then
                  Set rsGeneID = MAPPDB.OpenRecordset("SELECT GenMAPP FROM Other WHERE Other = '" & primary & "'")
                  If rsGeneID.EOF Then 'the primary ID doesn't exist in GenMAPP add it to other table
                    MAPPDB.Execute "INSERT INTO Other (Other, GenMAPP, Local) VALUES ('" & _
                            primary & "', '" & primary & "', '')"
                    MAPPDB.Execute "INSERT INTO GenMAPP (GenMAPP) VALUES ('" & primary & "')"
                    blankDB.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                    & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & primary & "', '" _
                                    & primary & "', 'O', 'Gene', " & centerX & ", " & centerY & ", " _
                                    & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                    & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                  Else 'it's already in the other table, so just add it to objects
                    blankDB.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                    & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & rsGeneID![GenMAPP] & "', '" _
                                    & primary & "', '" & primaryType & "', 'Gene', " & centerX & ", " & centerY & ", " _
                                    & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                 & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                  End If
               Else ' don't add the others, instead, write it to the exception file
                  othersEX.WriteLine (strline)
                  keepex = True
                  centerY = centerY - GENEHEIGHT
               End If
            ElseIf primaryType = "L" Or primaryType = "l" Then
                     If labelbool Then 'previous line was a label, only move down half a row
                        centerY = centerY - (GENEHEIGHT / 2) 'labels don't need to be that far apart
                     End If
                     blankDB.Execute "INSERT INTO Objects (Primary, PrimaryType, Type, CenterX, CenterY, SecondX, SecondY, " _
                                & "Width, Height, Rotation, Color, Label) VALUES ('Arial', '" & Chr(1) & "', 'Label', " _
                                & centerX & ", " & centerY & ", 8, 0, " & (Len(label) * 100) & ", 285, 0, 0, '" _
                                & label & "')"
                     labelbool = True
            Else 'it's not a valid primary type
                MsgBox "MAPP Builder has encountered an unknown (possibly blank) Primary Type", vbOKOnly
                GoTo NoName
            End If 'all options have been tried
            'parse the next line of the csv file
            strline = CSVfile.ReadLine
            If strline <> "end" Then 'a line containing end has been added to the csv file.
               comma1 = InStr(1, strline, Chr(9))
               comma2 = InStr(comma1 + 1, strline, Chr(9))
               primary = Left(strline, comma1 - 1)
               primaryType = Mid(strline, comma1 + 1, comma2 - comma1 - 1)
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
               If primary = "" Then
                  If UCase(primaryType) = "L" Then 'labels have null primary field
                     primary = " "
                  Else
                     MsgBox "You have not entered a Primary ID for a gene object. Please do so.", vbOKOnly
                     GoTo NoName
                  End If
               End If
               MAPPFileName = Mid(strline, comma1 + 1, Len(strline) - comma1)
               oldMAPPName = MappName
               MappName = CheckLength(MAPPFileName)
               If MappName = "" Then
                  MsgBox "You have not entered anything in the MappName Field. Please do so.", vbOKOnly
                  GoTo NoName
               End If
            End If
CloseMapp:
        Else
           blankDB.Execute "INSERT INTO Objects (Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                    & "Rotation, Color, Remarks) VALUES ('InfoBox', 76.5, 640, 0, 0, 45, 675, 0, -1, " _
                    & "'<font name=Arial, size=9><p><i><b>Author:</b></i>" & author.Text & "</p><p><i><b>Maintained by:</b></i>" & maintain.Text & "</p><p><i><b>Last modified:</b></i> " & modify & "</p>')"
            blankDB.Execute "INSERT INTO Objects (Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                        & "Rotation, Color) VALUES ('Legend', " & centerX + 1320 & ", 1778, 0, 0, 0, 0, 0, -1)"
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
                          & "BoardHeight, WindowWidth, WindowHeight, Expression, ColorSet, Notes) " _
                          & "VALUES ('" & TextToSql(oldMAPPName) & "', '', '" & rsinfo![Version] & "', '" _
                          & author.Text & "', '" & maintain.Text & "', '" & email.Text & "', '" _
                          & txtCopyright.Text & "', '" & modify & "', '" & txtRemarks.Text & "', " _
                          & boardwidth & ", " & boardheight & ", " & windowwidth & ", " & windowheight _
                          & ", '', '', '" & txtNotes.Text & "')"
            blankDB.Close
            centerX = STARTX
            centerY = STARTY - GENEHEIGHT
            currentMAPP = MAPPFileName
            If InStrRev(destination.Text, "\") = Len(destination.Text) Then
               MappPath = destination.Text & currentMAPP & ".mapp"
            Else
               MappPath = destination.Text & "\" & currentMAPP & ".mapp"
            End If
            MappPath = fixPath(MappPath)
            fsys.CopyFile mapptmpl, MappPath, True
            Set blankDB = OpenDatabase(MappPath)
        End If
    Wend
    'need to to the else case one more time for the last mapp
    blankDB.Execute "INSERT INTO Objects (Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                    & "Rotation, Color, Remarks) VALUES ('InfoBox', 76.5, 640, 0, 0, 45, 675, 0, -1, " _
                    & "'<font name=Arial, size=9><p><i><b>Author:</b></i>" & author.Text & "</p><p><i><b>Maintained by:</b></i>" & maintain.Text & "</p><p><i><b>Last modified:</b></i> " & modify & "</p>')"
    blankDB.Execute "INSERT INTO Objects (Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                        & "Rotation, Color) VALUES ('Legend', " & centerX + 1320 & ", 1778, 0, 0, 0, 0, 0, -1)"
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
                          & "BoardHeight, WindowWidth, WindowHeight, Expression, ColorSet, Notes) " _
                          & "VALUES ('" & TextToSql(MappName) & "', '', '" & rsinfo![Version] & "', '" _
                          & author.Text & "', '" & maintain.Text & "', '" & email.Text & "', '" _
                          & txtCopyright.Text & "', '" & modify & "', '" & txtRemarks.Text & "', " _
                          & boardwidth & ", " & boardheight & ", " & windowwidth & ", " & windowheight _
                          & ", '', '', '" & txtNotes.Text & "')"
    blankDB.Close
    
    MAPPDB.Close
    CSVfile.Close
    If addothers = False Then
      othersEX.Close
      If keepex = False Then
         fsys.DeleteFile (Left(FileName.Text, Len(FileName.Text) - 4) & ".EX.txt")
      End If
   End If
   
      
    
    If overflow Then
      'MsgBox "Some MAPPs have overflowed the maximum number of genes per MAPP. Overflow MAPPs labelled," _
      '& " MappName2 (or 3, etc.) have been created and are in the destination directory.", vbOKOnly
    End If
csvFailed:
    MousePointer = vbDefault
    Exit Sub
    
MappError:
    Select Case Err.Number
    Case 5
        MsgBox "You must enter a filename to save the new MAPP to.", vbOKOnly
    Case 3078
        MsgBox "You have entered an incorrect table name. Check your Access file and enter the correct name.", vbOKOnly
    Case 3134 'no primary value
        MsgBox "MAPP Builder has encountered a null Primary ID. You must enter a Primary ID " _
               & " for every gene (Labels are excluded)", vbOKOnly
    Case 94
        MsgBox "The input file contains a blank MAPP Name. Each entry must have a mapp name. Please fix this and rerun the program.", vbOKOnly
    Case Else
    MsgBox "An Error has occured. Please check the format of the Excel file and the sheet name an then try again.", vbOKOnly
    End Select
NoName:
   blankDB.Close
NoNameFirst: 'blankDB hasn't been created yet if the first line of data has an error, so you jump here instead.
   MAPPDB.Close
   CSVfile.Close
   MousePointer = vbDefault
   MappBuilderForm_Normal.Hide
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


Public Function fixPath(path As String) As String
    Dim Index As Integer
    For Index = 1 To Len(path)
      Select Case Mid(path, Index, 1)
      Case "/"
        Mid(path, Index, 1) = Chr(32)
      Case Else
      End Select
   Next Index
    For Index = 3 To Len(path) 'start at three to ignore c:
      Select Case Mid(path, Index, 1)
      Case ":"                            'Convert single quote to typographer's close single quote
         Mid(path, Index, 1) = Chr(32)
      Case Else
      End Select
   Next Index
    path = TextToSql(path)
    fixPath = path

End Function

Public Function HandleOverflow(mappnum As Integer) As Boolean
   On Error GoTo MappError
   Dim blankDB2 As Database
   Dim currentMAPP2 As String, oldMAPPName As String
   Dim titlewidth As Integer
   Dim newmapp As Boolean, continue As Boolean
   Dim boardwidth As Long, boardheight As Integer, windowwidth As Integer, windowheight As Integer
   Dim centerY As Integer, centerX As Long, labelbool2 As Boolean
   
   labelbool2 = False
   overflow = True
   
   currentMAPP2 = MAPPFileName & Str(mappnum)
   If InStrRev(destination.Text, "\") = Len(destination.Text) Then
      MappPath = destination.Text & currentMAPP2 & ".mapp"
   Else
      MappPath = destination.Text & "\" & currentMAPP2 & ".mapp"
   End If
   MappPath = fixPath(MappPath)
   fsys.CopyFile mapptmpl, MappPath, True
   Set blankDB2 = OpenDatabase(MappPath)
   centerY = STARTY - GENEHEIGHT
   centerX = STARTX
   While CSVfile.AtEndOfStream = False And newmapp = False
       If currentMAPP = MAPPFileName Then
           centerY = centerY + GENEHEIGHT 'add the height to make a vertical column.
           If centerY > (STARTY + (GENEHEIGHT * MAXGENE)) Then 'columns of MAXGENE
               centerX = centerX + ROWSEPERATOR
               If centerX > MAXWIDTH Then
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
           If primaryType = "G" Or primaryType = "g" Then
               labelbool2 = False
               Set rsGeneID = MAPPDB.OpenRecordset("SELECT GenMAPP FROM GenBank WHERE GenBank = '" & primary & "'")
               If rsGeneID.EOF Then 'the primary ID doesn't exist in GenMAPP add it to other table
                  If addothers Then
                     Set rstempGeneID = MAPPDB.OpenRecordset("Select Other from Other Where " _
                                    & "Other = '" & primary & "'")
                     If rstempGeneID.EOF Then 'this gene hasn't already been added to the other table
                        MAPPDB.Execute "INSERT INTO Other (Other, GenMAPP, Local) VALUES ('" & _
                            primary & "', '" & primary & "', '')"
                        MAPPDB.Execute "INSERT INTO GenMAPP (GenMAPP) VALUES ('" & primary & "')"
                     End If
                     blankDB2.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                   & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & primary & "', '" _
                                   & primary & "', 'O', 'Gene', " & centerX & ", " & centerY & ", " _
                                   & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                   & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                  Else
                     othersEX.WriteLine (strline)
                     keepex = True
                     centerY = centerY - GENEHEIGHT
                  End If
               Else    'the primary ID does exist add a genbank to the objects table
                   blankDB2.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                   & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & rsGeneID![GenMAPP] & "', '" _
                                   & primary & "', '" & primaryType & "', 'Gene', " & centerX & ", " & centerY & ", " _
                                   & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                   & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
               End If
           ElseIf primaryType = "S" Or primaryType = "s" Then
               labelbool2 = False
               Set rsGeneID = MAPPDB.OpenRecordset("Select GenMAPP FROM SwissProt Where SwissName = '" _
                              & primary & "'")
               If rsGeneID.EOF Then 'a swissname wasn't found. Check SwissNo
                  Set rstempGeneID = MAPPDB.OpenRecordset("SELECT SwissName from SwissNO where SwissNo" _
                                    & "= '" & primary & "'")
                  If rstempGeneID.EOF = False Then 'you found a swissno, now get the genMAPP ID
                     Set rsGeneID = MAPPDB.OpenRecordset("SELECT GenMAPP from SwissProt where SwissName" _
                                 & "= '" & rstempGeneID![SwissName] & "'")
                     blankDB2.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                    & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & rsGeneID![GenMAPP] & "', '" _
                                    & primary & "', '" & primaryType & "', 'Gene', " & centerX & ", " & centerY & ", " _
                                    & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                    & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                  Else 'no swissno or swissname, add other
                     If addothers Then
                        Set rstempGeneID = MAPPDB.OpenRecordset("Select Other from Other Where " _
                                    & "Other = '" & primary & "'")
                        If rstempGeneID.EOF Then 'this gene hasn't already been added to the other table
                           MAPPDB.Execute "INSERT INTO Other (Other, GenMAPP, Local) VALUES ('" & _
                              primary & "', '" & primary & "', '')"
                           MAPPDB.Execute "INSERT INTO GenMAPP (GenMAPP) VALUES ('" & primary & "')"
                        End If
                        blankDB2.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                   & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & primary & "', '" _
                                   & primary & "', 'O', 'Gene', " & centerX & ", " & centerY & ", " _
                                   & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                   & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                     Else
                        othersEX.WriteLine (strline)
                        keepex = True
                        centerY = centerY - GENEHEIGHT
                     End If
                  End If
               Else    'the primary ID exists add it to the objects table
                    blankDB2.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                    & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & rsGeneID![GenMAPP] & "', '" _
                                    & primary & "', '" & primaryType & "', 'Gene', " & centerX & ", " & centerY & ", " _
                                    & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                    & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
               End If
           ElseIf primaryType = "O" Or primaryType = "o" Then    'it's not g and not s, so it should be other
               labelbool2 = False
               If addothers Then
                  Set rsGeneID = MAPPDB.OpenRecordset("SELECT GenMAPP FROM Other WHERE Other = '" & primary & "'")
                  If rsGeneID.EOF Then 'the primary ID doesn't exist in GenMAPP add it to other table
                     
                        MAPPDB.Execute "INSERT INTO Other (Other, GenMAPP, Local) VALUES ('" & _
                           primary & "', '" & primary & "', '')"
                        MAPPDB.Execute "INSERT INTO GenMAPP (GenMAPP) VALUES ('~" & primary & "')"
                        blankDB2.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                   & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & primary & "', '" _
                                   & primary & "', 'O', 'Gene', " & centerX & ", " & centerY & ", " _
                                   & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                   & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                     
                  Else 'it's already in the other table, so just add it to objects
                     blankDB2.Execute "INSERT INTO Objects (GenMAPP, Primary, PrimaryType, Type, CenterX, CenterY, SecondX," _
                                   & " SecondY, Width, Height, Rotation, Color, Label, Head, Remarks) VALUES ('" & rsGeneID![GenMAPP] & "', '" _
                                   & primary & "', '" & primaryType & "', 'Gene', " & centerX & ", " & centerY & ", " _
                                   & SECONDX & ", " & SECONDY & ", " & GENEWIDTH & ", " & GENEHEIGHT & ", " & ROTATION & " , " _
                                & COLOR & ", '" & TextToSql(label) & "', '" & Head & "', '" & Remarks & "')"
                  End If
               Else
                  othersEX.WriteLine (strline)
                  keepex = True
                  centerY = centerY - GENEHEIGHT
               End If
           ElseIf primaryType = "L" Or primaryType = "l" Then
               If labelbool2 Then
                  centerY = centerY - (GENEHEIGHT / 2)
               End If
               blankDB2.Execute "INSERT INTO Objects (Primary, PrimaryType, Type, CenterX, CenterY, SecondX, SecondY, " _
                               & "Width, Height, Rotation, Color, Label) VALUES ('Arial', '" & Chr(1) & "', 'Label', " _
                               & centerX & ", " & centerY & ", 8, 0, " & (Len(label) * 132) & ", 285, 0, 0, '" _
                               & label & "')"
               labelbool2 = True
           End If 'all options have been tried
       
         strline = CSVfile.ReadLine
         If strline <> "end" Then 'a line containing end has been added to the csv file.
            comma1 = InStr(1, strline, Chr(9))
            comma2 = InStr(comma1 + 1, strline, Chr(9))
            primary = Left(strline, comma1 - 1)
            primaryType = Mid(strline, comma1 + 1, comma2 - comma1 - 1)
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
            If primary = "" Then
               If UCase(primaryType) = "L" Then 'labels have null primary field
                  primary = " "
               Else
                  MsgBox "You have not entered a Primary ID for a gene object. Please do so.", vbOKOnly
                  GoTo NoName
               End If
            End If
            MAPPFileName = Mid(strline, comma1 + 1, Len(strline) - comma1)
            oldMAPPName = MappName
            MappName = CheckLength(MAPPFileName)
            If MappName = "" Then
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
                   & "Rotation, Color, Remarks) VALUES ('InfoBox', 76.5, 640, 0, 0, 45, 675, 0, -1, " _
                   & "'<font name=Arial, size=9><p><i><b>Author:</b></i>" & author.Text & "</p><p><i><b>Maintained by:</b></i>" & maintain.Text & "</p><p><i><b>Last modified:</b></i> " & modify & "</p>')"
   blankDB2.Execute "INSERT INTO Objects (Type, CenterX, CenterY, SecondX, SecondY, Width, Height, " _
                       & "Rotation, Color) VALUES ('Legend', " & centerX + 1320 & ", 1778, 0, 0, 0, 0, 0, -1)"
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
                          & "BoardHeight, WindowWidth, WindowHeight, Expression, ColorSet, Notes) " _
                          & "VALUES ('" & TextToSql(oldMAPPName) & "', '', '" & rsinfo![Version] & "', '" _
                          & author.Text & "', '" & maintain.Text & "', '" & email.Text & "', '" _
                          & txtCopyright.Text & "', '" & modify & "', '" & txtRemarks.Text & "', " _
                          & boardwidth & ", " & boardheight & ", " & windowwidth & ", " & windowheight _
                          & ", '', '', '" & txtNotes.Text & "')"
   blankDB2.Close
   HandleOverflow = True
 Exit Function
MappError:
   Select Case Err.Number
   Case 5
       MsgBox "You must enter a filename to save the new MAPP to.", vbOKOnly
   Case 3078
       MsgBox "You have entered an incorrect table name. Check your Access file and enter the correct name.", vbOKOnly
   Case 3134 'no primary value
        MsgBox "MAPP Builder has encountered a null Primary ID. You must enter a Primary ID " _
               & " for every gene (Labels are excluded)", vbOKOnly
   Case 94
       MsgBox "The input file contains a blank MAPP Name. Each entry must have a mapp name. Please fix this and rerun the program.", vbOKOnly
   Case Else
   MsgBox "An Error has occured. Please check the format of the CSV file and then try again.", vbOKOnly
   End Select
NoName:
   blankDB2.Close
   'mappDB.Close
   'CSVfile.Close
   HandleOverflow = False
End Function


Public Sub setMappPath(path As String)
    MappPath = path
    destination.Text = MappPath
End Sub

Private Sub Helpfile_Click()
    'Form4.Show
End Sub


Public Sub setBaseMapp(baseMAPP As String, datalocation As String)
      destination.Text = baseMAPP
      DatabaseLoc = datalocation
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

