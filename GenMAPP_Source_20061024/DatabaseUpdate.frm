VERSION 5.00
Begin VB.Form frmDatabaseUpdate_Old 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Update"
   ClientHeight    =   6612
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   7656
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6612
   ScaleWidth      =   7656
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLater 
      Caption         =   "How to Update Later"
      Height          =   252
      Left            =   5520
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5760
      Width           =   2052
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1332
   End
   Begin VB.CommandButton cmdDont 
      Caption         =   "More Information"
      Height          =   252
      Left            =   5880
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1692
   End
   Begin VB.TextBox txtDont 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   540
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "DatabaseUpdate.frx":0000
      Top             =   4620
      Width           =   6972
   End
   Begin VB.CheckBox chkDont 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Don't Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4380
      Width           =   2232
   End
   Begin VB.CommandButton cmdAdjust 
      Caption         =   "More Information"
      Height          =   252
      Left            =   5880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2460
      Width           =   1692
   End
   Begin VB.TextBox txtAdjust 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1752
      Left            =   540
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "DatabaseUpdate.frx":000A
      Top             =   2700
      Width           =   6972
   End
   Begin VB.CheckBox chkAdjust 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Adjust MAPPs and Expression Datasets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2460
      Value           =   1  'Checked
      Width           =   5352
   End
   Begin VB.CommandButton cmdDb 
      Caption         =   "More Information"
      Height          =   252
      Left            =   5880
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1692
   End
   Begin VB.TextBox txtDb 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   540
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "DatabaseUpdate.frx":0016
      Top             =   1500
      Width           =   6972
   End
   Begin VB.CheckBox chkDb 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Update Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1260
      Value           =   1  'Checked
      Width           =   2172
   End
   Begin VB.TextBox txtHeader 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "DatabaseUpdate.frx":001C
      Top             =   60
      Width           =   7452
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   192
      Left            =   120
      TabIndex        =   13
      Top             =   6060
      Width           =   36
   End
   Begin VB.Label lblOp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   192
      Left            =   120
      TabIndex        =   12
      Top             =   6300
      Width           =   36
   End
End
Attribute VB_Name = "frmDatabaseUpdate_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbNew As Database                                 'Must be visible from more than one procedure
Dim headerText As String, dbText As String, adjustText As String
Dim dontText As String
Dim versionDb As String

Private Sub cmdDb_Click()
   Dim message As String
   
   message = "A new version of the GenMAPP Database was installed when you " _
         & "updated the program. The new database contains more recent versions of the " _
         & "SwissProt and "
'   If VARIATION = "_sc" Then
'      message = message & "SGD"
'   Else
'      message = message & "GenBank"
'   End If
   message = message & " databases. Checking the Update Database box will install the new " _
         & "GenMAPP database. GenMAPP will copy all of the entries you " _
         & "have previously made in the ""Other"" gene category to the new database. Your " _
         & "old database will be renamed GenMAPP_Old.gdb. It is not used " _
         & "by GenMAPP and may be deleted."
   MsgBox message, vbInformation + vbOKOnly, "Update Database"
End Sub
Private Sub cmdAdjust_Click()
   MsgBox "As a result of installing a new GenMAPP Database, it is likely " _
         & "that some gene identifications in your MAPPs and Expression Datasets now refer " _
         & "to accession numbers that are out of date. GenMAPP will go " _
         & "through your MAPP and Expression Dataset files and change any gene " _
         & "identifications that are incorrect." _
         & vbCrLf & vbCrLf _
         & "The update program updates only those files in the GenMAPP" _
         & "\ folder and subfolders of this. If you " _
         & "are storing files elsewhere and want them updated you should temporarily move " _
         & "your files into these folders before clicking OK.", _
         vbInformation + vbOKOnly, "Adjust MAPPs and Expression Datasets"
End Sub
Private Sub cmdDont_Click()
   MsgBox "Selecting this option leaves your current GenMAPP Database, MAPP files, and " _
          & "Expression Datasets alone. However, we recommend that you update your GenMAPP " _
          & "files because the new GenMAPP database contains more recent data and many new " _
          & "gene accession numbers. You will not lose any information by updating to the " _
          & "new database. GenMAPP will copy all of the entries you have made in " _
          & "the ""Other"" gene category to the new database.", _
          vbInformation + vbOKOnly, "Don't Update"
End Sub
Private Sub cmdLater_Click()
   MsgBox "If you wish to perform this update later or come across other MAPPs or " _
          & "Expression Datasets that need to be updated, you can have GenMAPP rerun " _
          & "its update routine by following these steps:" _
          & vbCrLf & vbCrLf _
          & "(1) Be sure GenMAPP is not running." _
          & vbCrLf & vbCrLf _
          & "(2) Be sure the MAPPs and Expression Datasets to be updated are in GenMAPP\MAPPs " _
          & "and GenMAPP\Expression Datasets folders or subfolders of those." _
          & vbCrLf & vbCrLf _
          & "(3) Download GenMAPP_New.gdb from www.GenMAPP.org to " & appPath & "." _
          & vbCrLf & vbCrLf _
          & "(4) Run GenMAPP again.", _
          vbInformation + vbOKOnly, "Updating Later"
End Sub

Private Sub Form_Load()
'   headerText = "GenMAPP has detected a new GenMAPP Database and can update your files to " _
'              & "accommodate new accession numbers. Check which files you would like to " _
'              & "be updated or click on the More Information boxes if you are unsure of " _
'              & "what is appropriate for you."
'   txtHeader = headerText
'   dbText = "GenMAPP will copy all of the entries you have made in the ""Other"" gene " _
'          & "category to the new database. Your old database will be renamed " _
'          & "GenMAPP_Old.gdb."
'   txtDb = dbText
'   adjustText = "GenMAPP will search through your """ & basePath & """ folder and " _
'              & "subfolders, changing all outdated gene accession numbers to current ones. " _
'              & "If you have MAPPs or Expression Datasets in locations that are not in " _
'              & "this folder or its subfolders, they will not be updated. If you want to " _
'              & "update your files, please move them before you click OK."
'   txtAdjust = adjustText
'   dontText = "The new database will be deleted, leaving your current one in place. If " _
'            & "you wish to update your database in the future, you can download the " _
'            & "latest version from www.GenMAPP.org and GenMAPP will present you with this " _
'            & "dialog again."
'   txtDont = dontText
End Sub

Private Sub chkDb_Click()
   If chkDb.value = vbChecked Then
      chkDont.value = vbUnchecked
   End If
End Sub
Private Sub chkAdjust_Click()
   If chkAdjust.value = vbChecked Then
      chkDont.value = vbUnchecked
   End If
End Sub
Private Sub chkDont_Click()
   If chkDont.value = vbChecked Then
      chkAdjust.value = vbUnchecked
      chkDb.value = vbUnchecked
   End If
End Sub

Private Sub cmdOK_Click()
   Dim dbOld As Database, rs As Recordset, mappExceptions As Boolean, gexExceptions As Boolean
   
'GoTo here
   If chkDont.value = vbChecked Then '+++++++++++++++++++++++++++++++++++++++++++ Dump New Database
      Kill appPath & "GenMAPP_New.gdb"
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If

   MousePointer = vbHourglass
   If chkDb.value = vbChecked Then '+++++++++++++++++++++++++++++++++++++++++++++++ Change Database
      lblStatus = "Transfering 'Other' records"
      DoEvents
      Set dbOld = OpenDatabase(appPath & "GenMAPP.gdb")
      Set dbNew = OpenDatabase(appPath & "GenMAPP_New.gdb")
      Set rs = dbOld.OpenRecordset("Other", dbOpenTable)
      Do Until rs.EOF
         lblOp = rs!Other
         DoEvents
         dbNew.Execute "INSERT INTO Other (Other, Remarks, GenMAPP, Local, Notes)" _
                     & "   VALUES ('" & rs!Other & "', '" & rs!remarks & "', '" & rs!GenMAPP _
                     & "', '" & rs!Local & "', '" & rs!notes & "')"
         dbNew.Execute "INSERT INTO GenMAPP (GenMAPP, Local) VALUES ('" & rs!GenMAPP & "', '')"
         rs.MoveNext
      Loop
      lblStatus = ""
      lblOp = ""
      DoEvents
      dbNew.Close
      dbOld.Close
   End If
   If Dir(appPath & "GenMAPP_Old.gdb") <> "" Then
      Kill appPath & "GenMAPP_Old.gdb"
   End If
   Name appPath & "GenMAPP.gdb" _
        As appPath & "GenMAPP_Old.gdb"
   Name appPath & "GenMAPP_New.gdb" _
        As appPath & "GenMAPP.gdb"
   Set dbNew = OpenDatabase(appPath & "GenMAPP.gdb")                      'New current database
   Set rs = dbNew.OpenRecordset("SELECT * FROM Info")
   versionDb = rs!version
   
   If chkAdjust.value = vbChecked Then '++++++++++++++++++++++ Change MAPPs And Expression Datasets
      lblStatus = "Updating MAPPs"
      DoEvents
      Open mruMappPath & "UnmatchedMAPPGenes.txt" For Output As #30
'      NavigateTree basePath, "mapp"
      lblOp = ""
      DoEvents
      If LOF(30) < 10 Then
         Close #30
         Kill mruMappPath & "UnmatchedMAPPGenes.txt"
      Else
         Close #30
         mappExceptions = True
      End If
      lblStatus = "Updating Expression Datasets"
      DoEvents
      Open GetFolder(mruDataSet) & "UnmatchedGEXGenes.txt" For Output As #30
'      NavigateTree basePath, "gex"
      lblOp = ""
      DoEvents
      If LOF(30) < 10 Then
         Close #30
         Kill GetFolder(mruDataSet) & "UnmatchedGEXGenes.txt"
      Else
         Close #30
         gexExceptions = True
      End If
      lblStatus = ""
      DoEvents
   End If
   If mappExceptions Then
      MsgBox "Some genes in your MAPPs have no matches in the new database. See the " _
             & "exception list in " & mruMappPath & "UnmatchedMAPPGenes.txt" & ".", _
             vbExclamation + vbOKOnly, "Updating MAPPs"
   End If
   If gexExceptions Then
         MsgBox "Some genes in your Expression Datasets have no matches in the new database. " _
                & "See the " _
                & "exception list in " & GetFolder(mruDataSet) & "UnmatchedGEXGenes.txt" & ".", _
                vbExclamation + vbOKOnly, "Updating Expression Datasets"
   End If
   
ExitSub:
   MousePointer = vbDefault
   Hide
End Sub

Sub NavigateTree(path As String, ext As String) '************************ Finds All Files In A Tree
'Not currently used
   '  Enter    path  Directory at base of tree to be examined
   '           ext   extension to find in the tree
   '  Process  Searches for appropriate files first and makes changes.
   '           Then searches for directories and recursively calls itself for each directory,
   '              returning when it reaches the end of the directory
   Dim file As String, index As Integer
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Files First
   file = Dir(path) ' & "\")
   Do Until file = ""
'      If Not cmdCancel.Visible Then Exit Sub
      If UCase(Right(file, Len(ext) + 1)) = "." & UCase(ext) Then
         Select Case UCase(ext)
         Case "GEX"
            UpdateGEX path & file '& "\" & file
         Case "MAPP"
            UpdateMAPP path & file '& "\" & file
         End Select
      End If
      file = Dir
   Loop
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Directories Next
'   file = Dir(path & "\")
   file = Dir(path, vbDirectory)
   Do Until file = ""
      index = index + 1                            'Keep track of where we are in current directory
      If file <> "." And file <> ".." Then
         If (GetAttr(path & file) And vbDirectory) = vbDirectory Then
'         If (GetAttr(path & "\" & file) And vbDirectory) = vbDirectory Then
            DoEvents
            NavigateTree path & file & "\", ext
            file = Dir(path, vbDirectory)        'Return to directory entry where we left off
'            file = Dir(path & "\", vbDirectory)        'Return to directory entry where we left off
            For i = 1 To index - 1                     'because calling Dir again in ListFiles will
               file = Dir                              'lose our place
            Next i
         End If
      End If
      If file <> "" Then file = Dir
   Loop
End Sub
Sub UpdateGEX(file As String) '******************* Verifies GenMAPP #s In Single Expression Dataset
   Dim dbGex As Database, rsExpression As Recordset, rs As Recordset
   Dim readOnly As Boolean
   
   lblOp = file
   DoEvents
   If (GetAttr(file) And vbReadOnly) = vbReadOnly Then
      If MsgBox(file & " has been set as read-only. Do you want to update it?", _
                vbExclamation + vbYesNo, "Updating Expression Datasets") = vbNo Then
         Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Else
         SetAttr file, vbNormal                    'This turns off all attributes including archive
         readOnly = True
      End If
   End If
   Set dbGex = OpenDatabase(file)
   Set rsExpression = dbGex.OpenRecordset("SELECT GenMAPP, ID, systemCode FROM Expression")
   Do Until rsExpression.EOF
      Set rs = dbNew.OpenRecordset("SELECT * FROM GenMAPP WHERE GenMAPP = '" & rsExpression!GenMAPP & "'")
      If rs.EOF And Dat(rsExpression!GenMAPP) <> "" Then                         'GenMAPP unmatched
         Select Case rsExpression!systemCode
         Case "G", "D"                                                         'GenBank or SGD type
            Set rs = dbNew.OpenRecordset("SELECT GenBank, GenMAPP FROM GenBank WHERE GenBank = '" & rsExpression!id & "'")
            If rs.EOF Then                                                               'Not found
               If rsExpression!systemCode = "D" Then
                  Print #30, file & ": SGDID " & rsExpression!id & " Not found"
               Else
                  Print #30, file & ": GenBank # " & rsExpression!id & " Not found"
               End If
            Else                                                                  'Change GenMAPP #
               rsExpression.edit
               rsExpression!GenMAPP = rs!GenMAPP
               rsExpression.Update
            End If
         Case "S"                                                                   'SwissProt type
            Set rs = dbNew.OpenRecordset("SELECT SwissNo, SwissName FROM SwissNo WHERE SwissNo = '" & rsExpression!id & "'")
                                                                               'Check SwissNo first
            If Not rs.EOF Then                                                       'SwissNo found
               s = rs!SwissName                                           'Get SwissName from table
            Else
               s = rsExpression!id                                   'Assume ID is SwissName
            End If
            Set rs = dbNew.OpenRecordset("SELECT SwissName, GenMAPP FROM SwissProt WHERE SwissName = '" & s & "'")
                                                                                   'Check SwissName
            If rs.EOF Then                                                     'SwissName not found
               Print #30, file & ": SwissProt ID " & rsExpression!id & " Not found"
            Else                                                                  'Change GenMAPP #
               rsExpression.edit
               rsExpression!GenMAPP = rs!GenMAPP
               rsExpression.Update
            End If
         Case "O"
            Print #30, file & ": Other ID " & rsExpression!id & " Not found"
         Case Else
            Print #30, file & ": Gene ID " & rsExpression!id & ", " & " Cataloging-system code '" & rsExpression!systemCode & "' not recognized"
         End Select
      End If
      rsExpression.MoveNext
   Loop
   rsExpression.Close
   dbGex.Execute "UPDATE Info SET Version = '" & versionDb & "'"
   If readOnly Then SetAttr file, vbReadOnly                                   'Leave archive unset
End Sub
Sub NavigateMAPPTree(path As String) '*********************************** Finds All Mapps In A Tree
'Not currently used
   '  Enter    path  Directory at base of tree to be examined
   '  Process  Searches for MAPP files first and makes changes.
   '           Then searches for directories and recursively calls itself for each directory,
   '              returning when it reaches the end of the directory
   Dim file As String, index As Integer
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Files First
   file = Dir(path & "\")
   Do Until file = ""
'      If Not cmdCancel.Visible Then Exit Sub
      If Right(file, 5) = ".mapp" Then
         UpdateMAPP path & "\" & file
      End If
      file = Dir
   Loop
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Directories Next
   file = Dir(path & "\", vbDirectory)
   Do Until file = ""
      index = index + 1                         'Keep track of where we are in current directory
      If file <> "." And file <> ".." Then
         If (GetAttr(path & "\" & file) And vbDirectory) = vbDirectory Then
            DoEvents
            NavigateMAPPTree path & "\" & file
            file = Dir(path & "\", vbDirectory)     'Return to directory entry where we left off
            For i = 1 To index - 1                  'because calling Dir again in ListFiles will
               file = Dir                           'lose our place
            Next i
         End If
      End If
      If file <> "" Then file = Dir
   Loop
End Sub
Sub UpdateMAPP(file As String) '******************************** Verifies GenMAPP #s In Single MAPP
'Not currently used
   Dim dbMapp As Database, rsObjects As Recordset, rs As Recordset
   Dim readOnly As Boolean
   
   lblOp = file
   DoEvents
   If (GetAttr(file) And vbReadOnly) = vbReadOnly Then
      If MsgBox(file & " has been set as read-only. Do you want to update it?", _
                vbExclamation + vbYesNo, "Updating MAPPs") = vbNo Then
         Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Else
         SetAttr file, vbNormal                    'This turns off all attributes including archive
         readOnly = True
      End If
   End If
   Set dbMapp = OpenDatabase(file)
   Set rsObjects = dbMapp.OpenRecordset("SELECT GenMAPP, ID, systemCode, Label" & _
                                        "   FROM Objects WHERE Type = 'Gene'")
   Do Until rsObjects.EOF
      Set rs = dbNew.OpenRecordset( _
               "SELECT * FROM GenMAPP WHERE GenMAPP = '" & rsObjects!GenMAPP & "'")
      If rs.EOF And Dat(rsObjects!GenMAPP) <> "" Then                            'GenMAPP unmatched
         Select Case rsObjects!systemCode
         Case "G", "D"                                                         'GenBank or SGD type
            Set rs = dbNew.OpenRecordset( _
                  "SELECT GenBank, GenMAPP" & _
                  "   FROM GenBank WHERE GenBank = '" & rsObjects!id & "'")
            If rs.EOF Then                                                               'Not found
               If rsObjects!systemCode = "D" Then
                  Print #30, file & ": SGDID " & rsObjects!id & ", """ & rsObjects!Label & """ Not found"
               Else
                  Print #30, file & ": GenBank # " & rsObjects!id & ", """ & rsObjects!Label & """ Not found"
               End If
            Else                                                                  'Change GenMAPP #
               rsObjects.edit
               rsObjects!GenMAPP = rs!GenMAPP
               rsObjects.Update
            End If
         Case "S"                                                                   'SwissProt type
            Set rs = dbNew.OpenRecordset("SELECT SwissNo, SwissName FROM SwissNo WHERE SwissNo = '" & rsObjects!id & "'")
                                                                               'Check SwissNo first
            If Not rs.EOF Then                                                       'SwissNo found
               s = rs!SwissName                                           'Get SwissName from table
            Else
               s = rsObjects!id                                   'Assume ID is SwissName
            End If
            Set rs = dbNew.OpenRecordset("SELECT SwissName, GenMAPP FROM SwissProt WHERE SwissName = '" & s & "'")
                                                                                   'Check SwissName
            If rs.EOF Then                                                     'SwissName not found
               Print #30, file & ": SwissProt ID " & rsObjects!id & ", """ & rsObjects!Label & """ Not found"
            Else                                                                  'Change GenMAPP #
               rsObjects.edit
               rsObjects!GenMAPP = rs!GenMAPP
               rsObjects.Update
            End If
         Case "O"
            Print #30, file & ": Other ID " & rsObjects!id & ", """ & rsObjects!Label & """ Not found"
         Case Else
            Print #30, file & ": Gene ID " & rsObjects!id & ", """ & rsObjects!Label & """ Cataloging-system code '" & rsObjects!systemCode & "' not recognized"
         End Select
      End If
      rsObjects.MoveNext
   Loop
   rsObjects.Close
   dbMapp.Execute "UPDATE Info SET Version = '" & versionDb & "'"
   If readOnly Then SetAttr file, vbReadOnly                                   'Leave archive unset
End Sub

Private Sub txtHeader_Change()
   txtHeader = headerText
End Sub
Private Sub txtDb_Change()
   txtDb = dbText
End Sub
Private Sub txtAdjust_Change()
   txtAdjust = adjustText
End Sub
Private Sub txtDont_Change()
   txtDont = dontText
End Sub


