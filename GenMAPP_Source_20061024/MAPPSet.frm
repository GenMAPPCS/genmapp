VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMAPPSet 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MAPP Sets"
   ClientHeight    =   5925
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "MAPPSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkConvert 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Convert old-version MAPPs to current version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   180
      TabIndex        =   18
      ToolTipText     =   $"MAPPSet.frx":08CA
      Top             =   5880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   4512
   End
   Begin VB.CheckBox chkOverwrite 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Overwrite existing files without asking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   180
      TabIndex        =   17
      ToolTipText     =   "If unchecked, each time GenMAPP eincounters an existing MAPP Set file it will ask for confirmation to overwrite."
      Top             =   5640
      Value           =   1  'Checked
      Width           =   3792
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4560
      TabIndex        =   16
      Top             =   5160
      Width           =   672
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4560
      TabIndex        =   15
      Top             =   4800
      Width           =   672
   End
   Begin VB.ListBox lstColorSets 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      ItemData        =   "MAPPSet.frx":0963
      Left            =   180
      List            =   "MAPPSet.frx":0965
      MultiSelect     =   2  'Extended
      TabIndex        =   14
      ToolTipText     =   "Hold down Ctrl key to select more than one Color Set."
      Top             =   4620
      Width           =   4332
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose Expression Dataset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   180
      TabIndex        =   13
      Top             =   3660
      Width           =   2772
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   120
      Top             =   6540
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.DirListBox dirDestination 
      Height          =   2016
      Left            =   3660
      TabIndex        =   12
      Top             =   600
      Width           =   3192
   End
   Begin VB.DriveListBox drvDestination 
      Height          =   288
      Left            =   5760
      TabIndex        =   11
      Top             =   300
      Width           =   1092
   End
   Begin VB.DirListBox dirSource 
      Height          =   2016
      Left            =   180
      TabIndex        =   10
      Top             =   600
      Width           =   3192
   End
   Begin VB.DriveListBox drvSource 
      Height          =   288
      Left            =   2280
      TabIndex        =   9
      Top             =   300
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   5580
      TabIndex        =   7
      Top             =   5460
      Width           =   1272
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Set"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   5580
      TabIndex        =   6
      Top             =   4980
      Width           =   1272
   End
   Begin VB.CheckBox chkSubfolders 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Include subfolders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   180
      TabIndex        =   20
      ToolTipText     =   "Will also export MAPPs in subfolders of selected folder."
      Top             =   2580
      Value           =   1  'Checked
      Width           =   2172
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New subfolder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   3960
      TabIndex        =   21
      Top             =   2580
      Width           =   1272
   End
   Begin VB.Image imgNewFolder 
      Height          =   240
      Left            =   3660
      Picture         =   "MAPPSet.frx":0967
      ToolTipText     =   "Create new subfolder"
      Top             =   2580
      Width           =   240
   End
   Begin VB.Label lblDetail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   420
      TabIndex        =   19
      Top             =   3060
      Width           =   48
   End
   Begin VB.Label lblCurrentMAPP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   180
      TabIndex        =   8
      Top             =   2880
      Width           =   48
   End
   Begin VB.Label lblColorSet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color sets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   4380
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expression data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1692
   End
   Begin VB.Label lblExpression 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Click on box for dataset choice dialog."
      Top             =   4020
      Width           =   48
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HTML destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3660
      TabIndex        =   2
      Top             =   360
      Width           =   1584
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MAPP source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   1224
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Root folders:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1308
   End
End
Attribute VB_Name = "frmMAPPSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsColorSets As Recordset
Dim sourceRoot As String, destinationRoot As String
Dim newExpression As String ', newColorSet As String
Dim dbNewExpression As Database, rsNewColorSets As Recordset
Dim expressionDB As String         'Full value of lblExpression, which is shortened by FileAbbrev()
'Dim mappIndex As String                                  'HTML for TOC of MAPPs in folder hierarchy
Dim dbGeneIndex As Database
'Dim colorSetSelect As String       'Beginning of HTML Color Set list with Javascript to open window
Dim colorIndexesSet(MAX_COLORSETS) As String                             'Array for entire MAPP Set
   '  Save this at beginning because MAPP Set routine must change colorSetIndexes for each MAPP

Private Sub Form_Load()
   Dim begCommand As Integer, endCommand As Integer, createMappSet As Boolean
   Dim commandArg As String
   
   cmdCancel.Caption = "Close"
On Error GoTo ErrorHandler
'   If InStr(commandLine, "set:") Then '======================================== Handle Command Line
   commandArg = CommandLineArg(commandLine, "set:")
   If commandArg <> "" Then '==================================================React To Command Line
'      begCommand = InStr(commandLine, """set:") + 5
'      endCommand = InStr(begCommand, commandLine, """") - 1                  'Doesn't include quote
      mruExportSourcePath = commandArg & "\"
'      mruExportSourcePath = Mid(commandLine, begCommand, endCommand - begCommand + 1) & "\"
'      commandLine = Left(commandLine, begCommand - 6) & Mid(commandLine, endCommand + 2)
'      If InStr(commandLine, """dest:") Then
      commandArg = CommandLineArg(commandLine, "dest:")
      If commandArg <> "" Then
'         begCommand = InStr(commandLine, """dest:") + 6
'         endCommand = InStr(begCommand, commandLine, """") + 1               'Doesn't include quote
'         mruExportPath = Mid(commandLine, begCommand, endCommand - begCommand - 1) & "\"
         mruExportPath = commandArg & "\"
'         commandLine = Left(commandLine, begCommand - 7) & Mid(commandLine, endCommand + 2)
         createMappSet = True
      End If
   End If
   drvSource.drive = Left(mruExportSourcePath, InStr(mruExportSourcePath, ":"))
   dirSource.path = Left(mruExportSourcePath, InStrRev(mruExportSourcePath, "\") - 1)
   drvDestination.drive = Left(mruExportPath, InStr(mruExportPath, ":"))
   dirDestination.path = Left(mruExportPath, InStrRev(mruExportPath, "\") - 1)
ExitSub:
   If createMappSet Then
      Form_Activate
      show
      DoEvents
      cmdCreate_Click
      Hide            'Must Hide here because the call from frmDrafter shows modally after the load
   End If
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
ErrorHandler:
   'Avoids setting a Path when the destination directory doesn't exist
End Sub
Private Sub Form_Activate()
   Dim colors(MAX_COLORSETS) As String, color As Integer, lastColor As Integer, index As Integer
   Dim prevPipe As Integer, pipe As Integer
   Dim begCommand As Integer, endCommand As Integer
   Dim commandArg As String
   
   lblCurrentMAPP = ""
   drvSource.Refresh
   dirSource.Refresh
   drvDestination.Refresh
   dirDestination.Refresh
   If Not mappWindow.dbExpression Is Nothing Then
      expressionDB = mappWindow.dbExpression.name                       'This sets up ColorSet list
      lblExpression = FileAbbrev(expressionDB)                          'This sets up ColorSet list
      commandArg = CommandLineArg(commandLine, "colors:")
      If commandArg <> "" Then '+++++++++++++++++++++++++++++++++++++++++++++ React To Command Line
'         begCommand = InStr(commandLine, """colors:") + 8
'         endCommand = InStr(begCommand, commandLine, """") - 1               'Doesn't include quote
'         If UCase(Mid(commandLine, begCommand, endCommand - begCommand + 1)) = "ALL" Then
         If UCase(commandArg) = "ALL" Then
            For index = 0 To lstColorSets.ListCount - 1
               lstColorSets.selected(index) = True
            Next index
'         ElseIf UCase(Mid(commandLine, begCommand, endCommand - begCommand + 1)) = "NONE" Then
         ElseIf UCase(commandArg) = "NONE" Then
            For index = 0 To lstColorSets.ListCount - 1
               lstColorSets.selected(index) = False
            Next index
         ElseIf Left(commandArg, 1) = "|" Then '===================Individual Colors Pipe Delimited
            prevPipe = 1 '-------------------------------------------------------------Parse Colors
            pipe = InStr(prevPipe + 1, commandArg, "|")
            lastColor = -1
            Do Until pipe > Len(commandArg)
               lastColor = lastColor + 1
               colors(lastColor) = Mid(commandArg, prevPipe + 1, pipe - prevPipe - 1)
               prevPipe = pipe
               pipe = InStr(prevPipe + 1, commandArg, "|")
               If pipe = 0 Then pipe = Len(commandArg) + 1
            Loop
            For index = 0 To lstColorSets.ListCount - 1 '-----------------------------Apply To List
               lstColorSets.selected(index) = False
               For color = 0 To lastColor
                  If colors(color) = lstColorSets.List(index) Then
                     lstColorSets.selected(index) = True
                  End If
               Next color
            Next index
         Else '========================================================================Single Color
            For index = 0 To lstColorSets.ListCount - 1 '-----------------------------Apply To List
               lstColorSets.selected(index) = False
               If commandArg = lstColorSets.List(index) Then
                  lstColorSets.selected(index) = True
               End If
            Next index
         End If
      End If
   End If
   DoEvents
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   cmdCancel_Click
End Sub

Private Sub cmdCreate_Click() '********************************************** Create A New MAPP Set
   Dim index As Integer, slash As Integer, mappSetName As String, mappIndexFile As String
   Dim geneIndexFile As String
   Dim colorSetList As String                                    'Drop-down list box for Color Sets
   Dim optionValue As Integer
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For Valid Web Destination Folder
   creatingMappSet = True
   If ValidHTMLName(dirDestination.path, False) = "" Then
      MsgBox "Your destination path is not universally accepted as a Web address because " _
             & "of characters other than A - Z, a - z, 0 - 9 and underscore. Choose " _
             & "another path or Exit MAPP Sets, create a valid path, and return.", _
             vbExclamation + vbOKOnly
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up Root Paths And Names
   sourceRoot = dirSource.path
   destinationRoot = dirDestination.path
   mappSetName = Mid(sourceRoot, InStrRev(sourceRoot, "\") + 1)
'   geneIndexFile = destinationRoot & "\_GeneIndex_" & ValidHTMLName(mappSetName, False)
   geneIndexFile = "_GeneIndex_" & ValidHTMLName(mappSetName, False)
      '  Without extension because both htm and txt files will be created
'   mappIndexFile = destinationRoot & "\_MAPPIndex_" & ValidHTMLName(mappSetName, False) & ".htm"
   mappIndexFile = "_MAPPIndex_" & ValidHTMLName(mappSetName, False) & ".htm"
   If Dir(mappIndexFile) <> "" And chkOverwrite = vbUnchecked Then
      If MsgBox("A MAPP Set already exists with this name in this location. Do you want to " & _
                "replace it?", vbExclamation + vbYesNo, "Creating MAPP Set") = vbNo Then
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Else
         'Should delete whole folder tree here
      End If
   End If
   MousePointer = vbArrowHourglass
   DoEvents
   mruExportSourcePath = sourceRoot & "\"
   mruExportPath = destinationRoot & "\"
   If expressionDB <> "" Then
      Set mappWindow.dbExpression = OpenDatabase(expressionDB, False, True)
   End If
   cmdCancel.Caption = "Cancel"
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up For Gene Index
   If Dir(appPath & "GeneIndex.$tm") <> "" Then
      Kill appPath & "GeneIndex.$tm"
   End If
   FileCopy appPath & "GeneIndex.gtp", appPath & "GeneIndex.$tm"
   Set dbGeneIndex = OpenDatabase(appPath & "GeneIndex.$tm")
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up For MAPP Index
On Error GoTo InvalidPath
   Open destinationRoot & "\" & mappIndexFile For Output As #32
On Error GoTo 0
   Print #32, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
   Print #32, "<html>"
   Print #32, "<head>"
   Print #32, "   <title>" & mappSetName & " Index</title>"
   Print #32, "   <meta name=""generator"" content=""GenMAPP 2.1"">"
   If lstColorSets.SelCount = 0 Then
      Print #32, "   <script language=""JavaScript"">"
      Print #32, "   <!--"
      Print #32, "      function linkTo(link)"
      Print #32, "      {  window.open(link + '.htm');"
      Print #32, "         return false;"
      Print #32, "      }"
      Print #32, "//-->"
      Print #32, "</script>"
   Else
      Print #32, "   <script language=""JavaScript"">"
      Print #32, "   <!--"
      Print #32, "      function linkTo(link)"
      Print #32, "      {  var selection = document.ColorSets.ColorSet;"
      Print #32, "         window.open(link + '_' + "
      Print #32, "                     selection.options[selection.selectedIndex].value + '.htm');"
      Print #32, "         return false;"
      Print #32, "      }"
      Print #32, "//-->"
      Print #32, "</script>"
   End If
   Print #32, "</head>"
   Print #32, ""
   Print #32, "<body>"
   Print #32, "<font face=""Verdana, Helvetica"">"
   Print #32, "<h1 align=center>" & mappSetName & " Index</h1>"
   Print #32, "<p>Switch to <a href=""" & geneIndexFile & ".htm"">Gene Index</a></p>"
   If Not mappWindow.dbGene Is Nothing And lstColorSets.SelCount > 0 Then
      Print #32, "<p>Gene Database: " & GetFile(mappWindow.dbGene.name) & "</p>"
   End If
   If Not mappWindow.dbExpression Is Nothing And lstColorSets.SelCount > 0 Then
      Print #32, "<p>Expression Dataset: " & GetFile(mappWindow.dbExpression.name) & "</p>"
   End If
   
   htmlSuffix = ""
   
   If lstColorSets.SelCount > 0 Then '++++++++++++++++++++++++++++++ Set Up HTML Color Set List Box
      Print #32, "<form name=""ColorSets"">"
      Print #32, "<p>Choose Color Set&nbsp;"
      Print #32, ColorSetOptions
'      Print #32, "   <select name=""ColorSet"">"
'      optionValue = 0
'      Print #32, "      <option value=""0"">No expression data"
'      If lstColorSets.ListCount > 1 Then
'         optionValue = optionValue + 1
'         Print #32, "      <option value='" & optionValue & "'>Multiple Color Sets"
'      End If
'      For index = 0 To lstColorSets.ListCount - 1
'         If lstColorSets.selected(index) Then
'            optionValue = optionValue + 1
'            Print #32, "      <option value='" & optionValue & "'>" _
'                       & lstColorSets.List(index)
'         End If
'      Next index
'      Print #32, "   </select>"
      Print #32, "</p>"
      Print #32, "</form>"
   End If
   
   ProcessFiles sourceRoot
   
   Print #32, "</body>" '++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Finish MAPP Index
   Print #32, "</html>"
   Close #32
   
   CreateGeneIndex destinationRoot, geneIndexFile, mappIndexFile, mappSetName
   lblCurrentMAPP = "MAPP Set complete"
ExitSub:
   If creatingMappSet Then dbGeneIndex.Close
   htmlSuffix = ""
'   creatingMappSet = False
   mappWindow.FillColorSetList
   MousePointer = vbDefault
   lblCurrentMAPP.foreColor = vbBlack
   lblDetail = ""
   cmdCancel.Caption = "Close"
   creatingMappSet = False
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
InvalidPath:
   MsgBox "Path invalid. May possibly be read-only or no medium in drive.", _
          vbExclamation + vbOKOnly, "Creating MAPP Set"
   Resume ExitSub
End Sub
Sub ProcessFiles(source As String, Optional level As Integer = 0)
   '  Enter    source      Root folder for particular MAPP Set, Eg:
   '                          D:\MAPPSets
   '                       Initially this is the root folder for the entire MAPP set. As this
   '                       procedure is recursively called, source changes to the folder for a
   '                       particular MAPP as determined by the Find Directory section.
   '           level       Depth of recursive call. Also determines indent of MAPP Index entry
   Dim sourceFile As String                                   'May have illegal characters for HTML
   Dim destFile As String                                                 'Illegal characters fixed
   Dim relativeRoot As String                          'Root relative to folder chosen in dirSource
   Dim relativeFolder As String         'Relative destination folder with all legal HTML characters
      '  The relativeFolder is always the relativeRoot plus the MAPP name without extension.
      '  For example
      '     source         D:\MAPPSets
      '     relativeRoot   C:\MAPPSets\Gene_Family_MAPPs\
      '     relativeFolder C:\MAPPSets\Gene_Family_MAPPs\MyMAPP\
      '        This folder will contain MyMAPP_0 - n.Htm and _Support
   Dim mappName As String                               'Name of MAPP without extension. Eg: MyMAPP
   Dim htmlMappPath As String
   Dim dirIndex As Integer, fileIndex As Integer, index As Integer
   Dim slash As Integer, s As String, indent As String, i As Integer
'   Dim relativeFolder As String
   Dim htmlHome As String         'The name of the HTML MAPP file without extension. Backpages will
                                  'be under _Support.
   Dim htmlIndex As Integer
   Dim optionValue As Integer

   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Files First
   '  This finds files in the current directory. When all the files are found, we move down to the
   '  next directory recursively
      
   s = "" '=============================================================MAPP Index Entry For Folder
   For index = 1 To level
      indent = indent & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
   Next index
   s = ValidHTMLName(source, False)
   Print #32, vbCrLf & "<h4>" & indent & Mid(s, InStrRev(source, "\") + 1) & "</h4>"
'   indent = indent & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"          'Extra indent for files under folder
   
   sourceFile = Dir(source & "\", vbReadOnly)                              'Include read-only MAPPs
   Do Until sourceFile = ""
      fileIndex = fileIndex + 1                    'Keep track of where we are in current directory
      If LCase(Right(sourceFile, 5)) = ".mapp" Then '==========================Create The Web Pages
'Debug.Print source & "\" & file
'Debug.Print "  " & destinationRoot & Mid(source & "\" & file, Len(sourceRoot) + 1)
         destFile = ValidHTMLName(sourceFile, False)
         mappName = Left(destFile, InStrRev(destFile, ".") - 1)
         relativeRoot = ValidHTMLName(Mid(source, Len(dirSource.path) + 2), False) & "\"
         '  Eg: source is c:\GenMAPP V2 Data\MAPPs\hu_MAPPArchive\Gene Ontology Lists\GO Component
         '      dirSource.path is c:\GenMAPP V2 Data\MAPPs\hu_MAPPArchive
         '      relativeRoot will be Gene_Ontology_Lists\GO_Component\ with valid HTML and \
'         s = destinationRoot & Mid(source & "\" & destFile, Len(sourceRoot) + 1)
         If relativeRoot = "\" Then
            relativeRoot = ""
         End If
      
'         s = destinationRoot & Mid(source & "\" & destFile, Len(sourceRoot) + 1)
            '  Destination root Eg: C:\GenMAPP\Exports
            '  s Eg: C:\GenMAPP\Exports\Mm_Calcium_Channels.mapp
         htmlHome = dirDestination.path & "\" & relativeRoot & mappName & "\" _
                  & Left(destFile, InStrRev(destFile, ".") - 1)                    'Knock off .mapp
            '  Full path for resultant file without the .mapp
            '  It should be HTML address legal
            '  Eg: "C:\GenMAPP\Exports\My_MAPPs\Mm_Calcium_Channels" for Mm_Calcium_Channels.mapp
         lblCurrentMAPP = FileAbbrev(source & "\" & sourceFile)
         DoEvents
         htmlMappPath = ReverseSlashes(Mid(htmlHome, Len(destinationRoot) + 2))    'Relative folder
         If lstColorSets.SelCount = 0 Then '-------------------------------------------No Color Set
'            Set mappWindow.rsColorSet = Nothing
            colorIndexes(0) = 0                                                    'No colorIndexes
            valueIndex = -1                                                'No value ever displayed
            colorSetHTML = ""
            If Not mappWindow.OpenMAPP(source & "\" & sourceFile) Then
               GoTo NextMAPP                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
            mappWindow.HTMLExport Left(htmlHome, InStrRev(htmlHome, "\")), _
                                       Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
            If Not creatingMappSet Then                                               'Quit clicked
               Exit Sub                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            End If
               '  Relative to main destination root for index. Index is in destinationRoot
               '  so these addresses should all be relative to that
''         ElseIf lstColorSets.SelCount = 1 Then  '----------------------------------Single Color Set
''            s = Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
''               '  Relative to the current MAPP being processed
''               '  Eg: Mm_Calcium Channels
''            s = "'" & s & "_' + options[selectedIndex].value + '.htm'"
''               '  Eg: 'Rn_Translation_Factors_' + options[selectedIndex].value + '.htm'
''            '______________________________________________________Create "No expression data" MAPP
''            Set mappWindow.rsColorSet = Nothing
''            htmlSuffix = "_0"
''            colorSetHTML = _
''               "<select name=""ColorSet"" onChange=""window.open(" & s & ",'_self')"">" _
''               & vbCrLf & "   <option value='0' selected>No expression data" & vbCrLf
''            For index = 0 To lstColorSets.ListCount - 1
''               If lstColorSets.Selected(index) Then
''                  colorSetHTML = colorSetHTML & "   <option value='" & index + 1 & "'>" _
''                                              & lstColorSets.List(index) & vbCrLf
''               End If
''            Next index
''            colorSetHTML = colorSetHTML & "</select>"
''
''            If Not mappWindow.OpenMAPP(source & "\" & sourceFile) Then
''               GoTo NextMAPP                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
''            End If
''            mappWindow.HTMLExport Left(htmlHome, InStrRev(htmlHome, "\")), _
''                                  Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
''                  '  Destination folder, destination file without .mapp
''            htmlIndex = 0
'            s = Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
'               '  Relative to the current MAPP being processed
'               '  Eg: Mm_Calcium Channels
'            s = "'" & s & "_' + options[selectedIndex].value + '.htm'"
'               '  Eg: 'Rn_Translation_Factors_' + options[selectedIndex].value + '.htm'
'            '______________________________________________________Create "No expression data" MAPP
''            Set mappWindow.rsColorSet = Nothing
'            colorIndexes(0) = 0                                                    'No colorIndexes
'            valueIndex = -1                                                'No value ever displayed
'            htmlSuffix = "_0"
'            colorSetHTML = _
'               "<select name=""ColorSet"" onChange=""window.open(" & s & ",'_self')"">" _
'               & vbCrLf & "   <option value='0' selected>No expression data" & vbCrLf
'            For index = 0 To lstColorSets.ListCount - 1
'               If lstColorSets.selected(index) Then
'                  colorSetHTML = colorSetHTML & "   <option value='1'>" _
'                                              & lstColorSets.List(index) & vbCrLf
'               End If
'            Next index
'            colorSetHTML = colorSetHTML & "</select>"
'
'            If Not mappWindow.OpenMAPP(source & "\" & sourceFile) Then
'               GoTo NextMAPP                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'            End If
'            mappWindow.HTMLExport Left(htmlHome, InStrRev(htmlHome, "\")), _
'                                  Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
'                  '  Destination folder, destination file without .mapp
'
'            If Not creatingMappSet Then                                               'Quit clicked
'               Exit Sub                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'            End If
'            For index = 0 To lstColorSets.ListCount - 1 '________________________Process Color Sets
'               If lstColorSets.selected(index) Then
'                  colorIndexes(0) = 1
'                  colorIndexes(1) = index
'                  valueIndex = index
'                  colorSetHTML = _
'                     "<select name=""ColorSet"" onChange=""window.open(" & s & ",'_self')"">" _
'                     & vbCrLf & "   <option value='0'>No expression data" & vbCrLf
'                     colorSetHTML = colorSetHTML & "   <option value='1' selected>" _
'                                  & lstColorSets.List(index) & vbCrLf
'                  colorSetHTML = colorSetHTML & "</select><br>"
'
'                  If Not mappWindow.OpenMAPP(source & "\" & sourceFile) Then
'                     GoTo NextMAPP                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'                  End If
'                  mappWindow.HTMLExport Left(htmlHome, InStrRev(htmlHome, "\")), _
'                                                    Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
'               End If
'            Next index
         Else '--------------------------------------------------------------One Or Many Color Sets
'            s = Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
'               '  Relative to the current MAPP being processed
'               '  Eg: Mm_Calcium Channels
'            s = "'" & s & "_' + options[selectedIndex].value + '.htm'"
'               '  Eg: 'Rn_Translation_Factors_' + options[selectedIndex].value + '.htm'
            '______________________________________________________Create "No expression data" MAPP
'            Set mappWindow.rsColorSet = Nothing
            colorIndexes(0) = 0                                                    'No colorIndexes
            valueIndex = -1                                                'No value ever displayed
            htmlIndex = 0
            htmlSuffix = "_" & htmlIndex
'            colorSetHTML = _
'               "<select name=""ColorSet"" onChange=""window.open(" & s & ",'_self')"">" _
'               & vbCrLf & "   <option value='0' selected>No expression data" & vbCrLf
'            colorSetHTML = colorSetHTML _
'                           & "   <option value='1'>Multiple Color Sets" & vbCrLf
'            For index = 0 To lstColorSets.ListCount - 1
'               If lstColorSets.selected(index) Then
'                  colorSetHTML = colorSetHTML & "   <option value='" & index + 2 & "'>" _
'                                              & lstColorSets.List(index) & vbCrLf
'               End If
'            Next index
'            colorSetHTML = colorSetHTML & "</select>"
            colorSetHTML = ColorSetOptions(0, htmlHome)
            
            If Not mappWindow.OpenMAPP(source & "\" & sourceFile) Then
               GoTo NextMAPP                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
            mappWindow.HTMLExport Left(htmlHome, InStrRev(htmlHome, "\")), _
                                  Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
                  '  Destination folder, destination file without .mapp
            
            If lstColorSets.SelCount > 1 Then '___________________Create "Multiple Color Sets" MAPP
               i = 0 '.........................................................Determine Color Sets
               For index = 0 To lstColorSets.ListCount - 1
                  If lstColorSets.selected(index) Then
                     i = i + 1
                     colorIndexes(i) = index
                  End If
               Next index
               Set mappWindow.rsColorSet = Nothing
               colorIndexes(0) = i                                          'Number of colorIndexes
               valueIndex = -1                                             'No value ever displayed
               htmlIndex = htmlIndex + 1
               htmlSuffix = "_" & htmlIndex
   '            colorSetHTML = _
   '               "<select name=""ColorSet"" onChange=""window.open(" & s & ",'_self')"">" _
   '               & vbCrLf & "   <option value='0'>No expression data" & vbCrLf
   '            colorSetHTML = colorSetHTML & "   <option value='1' selected>Multiple Color Sets" _
   '                           & vbCrLf
   '            For index = 0 To lstColorSets.ListCount - 1
   '               If lstColorSets.selected(index) Then
   '                  colorSetHTML = colorSetHTML & "   <option value='" & index + 2 & "'>" _
   '                                              & lstColorSets.List(index) & vbCrLf
   '               End If
   '            Next index
   '            colorSetHTML = colorSetHTML & "</select>"
               colorSetHTML = ColorSetOptions(htmlIndex, htmlHome)
               
               If Not mappWindow.OpenMAPP(source & "\" & sourceFile) Then
                  GoTo NextMAPP                            'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
               End If
               mappWindow.HTMLExport Left(htmlHome, InStrRev(htmlHome, "\")), _
                                     Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
                     '  Destination folder, destination file without .mapp
            End If
            
            For index = 0 To lstColorSets.ListCount - 1 '__________Make Individual Color Sets MAPPs
               If Not creatingMappSet Then                                            'Quit clicked
                  Exit Sub                                 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
               End If
               If lstColorSets.selected(index) Then
'                  Set mappWindow.rsColorSet = _
'                      mappWindow.dbExpression.OpenRecordset( _
'                      "SELECT * FROM ColorSet WHERE ColorSet = '" & lstColorSets.List(index) & "'")
                  colorIndexes(0) = 1
                  colorIndexes(1) = index
                  valueIndex = index
                  htmlIndex = htmlIndex + 1
                  htmlSuffix = "_" & htmlIndex
'                  optionValue = 0
'                  colorSetHTML = _
'                     "<select name=""ColorSet"" onChange=""window.open(" & s & ",'_self')"">" _
'                     & vbCrLf & "   <option value='0'>No expression data" & vbCrLf
'                  If lstColorSets.ListCount > 1 Then
'                     optionValue = optionValue + 1
'                     colorSetHTML = colorSetHTML & "   <option value='" & optionValue & "'>" _
'                        & "Multiple Color Sets" & vbCrLf
'                  End If
'                  For i = 0 To lstColorSets.ListCount - 1
'                     If lstColorSets.selected(i) Then
'                        optionValue = optionValue + 1
'                        If i = index Then
'                           colorSetHTML = colorSetHTML & "   <option value='" & optionValue _
'                                        & "' selected>" & lstColorSets.List(i) & vbCrLf
'                        Else
'                           colorSetHTML = colorSetHTML & "   <option value='" & optionValue _
'                                        & "'>" & lstColorSets.List(i) & vbCrLf
'                        End If
'                     End If
'                  Next i
'                  colorSetHTML = colorSetHTML & "</select><br>"
                  colorSetHTML = ColorSetOptions(htmlIndex, htmlHome)
                  
                  If Not mappWindow.OpenMAPP(source & "\" & sourceFile) Then
                     GoTo NextMAPP                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                  End If
                  mappWindow.HTMLExport Left(htmlHome, InStrRev(htmlHome, "\")), _
                                                    Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
               End If
            Next index
         End If
         Print #32, indent & "<a href="""" onClick=return(linkTo(""" & htmlMappPath _
                    & """))>" & destFile & "</a><br>"
         AddToIndex source & "\" & sourceFile, htmlMappPath
            '  Index does not depend on the color set
         sourceFile = Dir(source & "\", vbReadOnly)    'Return to directory entry where we left off
         For i = 1 To fileIndex - 1                  'because calling Dir again in other procedures
            sourceFile = Dir                         'will lose our place
         Next i
      End If
NextMAPP:
      If Not creatingMappSet Then Exit Sub                 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      lblCurrentMAPP = ""
      DoEvents
      sourceFile = Dir
   Loop
   
   If chkSubfolders.value = vbChecked Then '+++++++++++++++++++++++++++++++++ Find Directories Next
   '  Find the next directory and recursively call ProcessFiles to find files and subdirectories
      sourceFile = Dir(source & "\", vbDirectory)
      Do Until sourceFile = ""
         dirIndex = dirIndex + 1                   'Keep track of where we are in current directory
         If sourceFile <> "." And sourceFile <> ".." Then
            If (GetAttr(source & "\" & sourceFile) And vbDirectory) = vbDirectory Then
               ProcessFiles source & "\" & sourceFile, level + 1
               sourceFile = Dir(source & "\", vbDirectory)   'Return to dir entry where we left off
               For i = 1 To dirIndex - 1            'because calling Dir again in ProcessFiles will
                  sourceFile = Dir                  'lose our place ("." and ".." are always
                                                    'first 2 directory entries)
               Next i
            End If
         End If
         sourceFile = Dir
      Loop
   End If
   lblCurrentMAPP = ""
   DoEvents
   htmlHome = ""
End Sub
'*************************************************************** Produce HTML Color Set Option List
Function ColorSetOptions(Optional selected As Integer = -1, Optional ByVal path As String = "") _
         As String
   '  Entry    path     For link to options HTML page
   '                    If empty, it is the index list, nothing is selected and there is
   '                    no OnChange.
   '           selected Option number to be set as selected
   '                    If -1, it is the index list, nothing is selected and there is
   '                    no OnChange.
   '  Return   Complete HTML Color Set option box code
   Dim colorSetHTML As String, optionValue As Integer, i As Integer
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Beginning Of Option List
   path = Mid(path, InStrRev(path, "\") + 1)
'               '  Relative to the current MAPP being processed
'               '  Eg: Mm_Calcium Channels
'            s = "'" & s & "_' + options[selectedIndex].value + '.htm'"
'               '  Eg: 'Rn_Translation_Factors_' + options[selectedIndex].value + '.htm'
   If selected = -1 Then '============================================================List On Index
      colorSetHTML = _
         "<select name=""ColorSet"">" & vbCrLf
   Else '==============================================================================List On MAPP
      colorSetHTML = _
         "<select name=""ColorSet""" & vbCrLf & _
         "        onChange=""window.open('" & path & "_' + options[selectedIndex].value + '.htm', '_self')"">" & vbCrLf
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ "No Expression Data" Option
   colorSetHTML = colorSetHTML & "   <option value='0'"
   If selected = 0 Then                                                'Select "No expression data"
      colorSetHTML = colorSetHTML & " selected"
   End If
   colorSetHTML = colorSetHTML & ">No expression data" & vbCrLf
   optionValue = 0
   
   If lstColorSets.SelCount > 1 Then '++++++++++++++++++++++++++++++++ "Multiple Color Sets" Option
      optionValue = optionValue + 1
      colorSetHTML = colorSetHTML & "   <option value='" & optionValue & "'"
      If selected = optionValue Then                                  'Select "Multiple Color Sets"
         colorSetHTML = colorSetHTML & " selected"
      End If
      colorSetHTML = colorSetHTML & ">Multiple Color Sets" & vbCrLf
   End If
   
   For i = 0 To lstColorSets.ListCount - 1 '++++++++++++++++++++++++++++++++++++ Rest Of Color Sets
      If lstColorSets.selected(i) Then
         optionValue = optionValue + 1
         colorSetHTML = colorSetHTML & "   <option value='" & optionValue & "'"
         If selected = optionValue Then                                      'Select This Color Set
            colorSetHTML = colorSetHTML & " selected"
         End If
         colorSetHTML = colorSetHTML & ">" & lstColorSets.List(i) & vbCrLf
      End If
   Next i
   ColorSetOptions = colorSetHTML & "</select><br>"
End Function
Sub CreateGeneIndex(destinationRoot As String, geneIndexFile As String, mappIndexFile As String, _
                    mappSetName As String)
   Dim rsIndex As Recordset, rsInfo As Recordset
   Dim slash As Integer, i As Integer, otherIDs As String, index As Integer
   Dim mappName As String
   'For AllRelatedGenes()
      Dim genes As Integer
      Dim geneIDs(MAX_GENES, 2) As String
      Dim geneFound As Boolean
      'Dim supportedSystem as Boolean                 'System supported in Gene Database [optional]
      Dim systemsList As Variant                                      'Systems to search [optional]
   
   If Not creatingMappSet Then Exit Sub                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   lblDetail = "Creating indexes"
   DoEvents
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up Tab-Delimited Index
   Open destinationRoot & "\" & geneIndexFile & ".txt" For Output As #31
   Print #31, "Type"; vbTab; "Gene ID"; vbTab; "Gene Label"; vbTab; "MAPP"; vbTab; "Other IDs"
      '  Anything beginning with ID, such as "ID Type" in the first column doesn't work
      '  because of the Excel bug
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up HTML Index
   Open destinationRoot & "\" & geneIndexFile & ".htm" For Output As #32
   Print #32, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
   Print #32, "<html>"
   Print #32, "<head>"
   Print #32, "   <title>" & mappSetName & " Gene Index</title>"
   Print #32, "   <meta name=""generator"" content=""GenMAPP 2.1"">"
   If lstColorSets.SelCount = 0 Then
      Print #32, "   <script language=""JavaScript"">"
      Print #32, "   <!--"
      Print #32, "      function linkTo(link)"
      Print #32, "      {  window.open(link + '.htm');"
      Print #32, "         return false;"
      Print #32, "      }"
      Print #32, "//-->"
      Print #32, "</script>"
   Else
      Print #32, "   <script language=""JavaScript"">"
      Print #32, "   <!--"
      Print #32, "      function linkTo(link)"
      Print #32, "      {  var selection = document.ColorSets.ColorSet;"
      Print #32, "         window.open(link + '_' + "
      Print #32, "                     selection.options[selection.selectedIndex].value + '.htm');"
      Print #32, "         return false;"
      Print #32, "      }"
      Print #32, "//-->"
      Print #32, "</script>"
   End If
   Print #32, "</head>"
   Print #32, ""
   Print #32, "<body>"
   Print #32, "<font face=""Verdana, Helvetica"">"
   Print #32, "<h1 align=center>" & mappSetName & " Index</h1>"
   Print #32, "<p>Switch to <a href=""" & mappIndexFile & """>MAPP Index</a></p>"
   If Not mappWindow.dbGene Is Nothing And lstColorSets.SelCount > 0 Then
      Print #32, "<p>Gene Database: " & GetFile(mappWindow.dbGene.name) & "</p>"
   End If
   If Not mappWindow.dbExpression Is Nothing And lstColorSets.SelCount > 0 Then
      Print #32, "<p>Expression Dataset: " & GetFile(mappWindow.dbExpression.name) & "</p>"
   End If
   
   If lstColorSets.SelCount > 0 Then '++++++++++++++++++++++++++++++ Set Up HTML Color Set List Box
      Print #32, "<form name=""ColorSets"">"
      Print #32, "<p>Choose Color Set&nbsp;"
      Print #32, ColorSetOptions
      Print #32, "</p>"
      Print #32, "</form>"
      
      'GeneIndex always sends you to the "No expression data" MAPP
'      Print #32, "<form name=""ColorSets"">"
'      Print #32, "   <p>Choose Color Set&nbsp;"
'      Print #32, "   <select name=""ColorSet"">"
'      Print #32, "      <option value=""0"">No expression data"
'      For index = 0 To lstColorSets.ListCount - 1
'         If lstColorSets.selected(index) Then
'            Print #32, "      <option value=""" & index + 1 & """>" _
'                       & lstColorSets.List(index)
'         End If
'      Next index
'      Print #32, "   </select>"
'      Print #32, "   </p>"
'      Print #32, "</form>"
   End If
      
   Print #32, "<p>Gene IDs are shown as [System]ID, for example [S]CALM_HUMAN where<br>"
   Print #32, "   the System is the Gene ID System code, UniProt [S] in the example.</p>"
   Print #32, "<table border=1 cellspacing=1 cellpadding=1>"
   Print #32, "   <tr>"
   Print #32, "      <td><b>Gene Label</b></td>"
   Print #32, "      <td><b>Gene ID</b></td>"
   Print #32, "      <td><b>MAPP</b></td>"
   Print #32, "      <td><b>Other IDs</b></td>"
   Print #32, "   </tr>"
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Extract Index Recordset
   Set rsIndex = dbGeneIndex.OpenRecordset( _
                 "SELECT * FROM GeneIndex ORDER BY Label, SystemCode, ID, MAPP", dbOpenForwardOnly)
   
   Do Until rsIndex.EOF '++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Each Gene In MAPP
      '=============================================================================Get Related IDs
      AllRelatedGenes rsIndex!id, rsIndex!systemCode, mappWindow.dbGene, genes, geneIDs, geneFound
      otherIDs = ""
      For i = 1 To genes - 1
         otherIDs = otherIDs & "[" & geneIDs(i, 1) & "]" & geneIDs(i, 0) & " "
      Next i
      
      '==========================================================================Write Gene Listing
      If lstColorSets.SelCount = 0 Then
         Print #31, rsIndex!systemCode; vbTab; rsIndex!id; vbTab; rsIndex!Label; vbTab; _
                    rsIndex!mapp & ".htm"; vbTab; otherIDs
      Else
         Print #31, rsIndex!systemCode; vbTab; rsIndex!id; vbTab; rsIndex!Label; vbTab; _
                    rsIndex!mapp & "_0.htm"; vbTab; otherIDs
            '  Always reference the "0" Color Set
      End If
      Print #32, "   <tr>"
      Print #32, "      <td><nobr>" & rsIndex!Label & "</nobr></td>"
      Print #32, "      <td><nobr>" & "[" & rsIndex!systemCode & "]" & rsIndex!id & "</nobr></td>"
      mappName = rsIndex!mapp
      slash = InStrRev(mappName, "/")
      mappName = Mid(mappName, slash + 1, Len(mappName) - slash)
         '  File name of MAPP without .htm
      Print #32, "      <td><nobr><a href="""" onClick=return(linkTo(""" & rsIndex!mapp _
                 & """))>" & mappName & "</a></nobr></td>"
      Print #32, "      <td><nobr>" & otherIDs & "</nobr></td>"
      Print #32, "   </tr>"
      rsIndex.MoveNext
   Loop
   Print #32, "</table>" '+++++++++++++++++++++++++++++++++++++++++++++++++++++++ Finish Gene Index
   Print #32, "</body>"
   Print #32, "</html>"
   
   Close #31, #32
   lblDetail = ""
   DoEvents
End Sub
Sub AddToIndex(mapp As String, ByVal relativePath As String)
   '  Adds entries from an individual MAPP to the temp index database
   '  Entry:   mapp           MAPP to process
   '           relativePath   Where the HTML MAPP is relative to the HTML root
   Dim dbMapp As Database, rsObjects As Recordset
   
   
   If dbGeneIndex Is Nothing Then Exit Sub                 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   If htmlSuffix <> "" Then
'      relativePath = relativePath & "_0"
'   End If
'   relativePath = relativePath & ".htm"
   Set dbMapp = OpenDatabase(mapp, False, True)
   Set rsObjects = dbMapp.OpenRecordset("SELECT * FROM Objects WHERE Type = 'Gene'", _
                                        dbOpenForwardOnly)
   Do Until rsObjects.EOF
      dbGeneIndex.Execute _
                  "INSERT INTO GeneIndex (ID, SystemCode, Label, MAPP)" & _
                  "   VALUES ('" & rsObjects!id & "', '" & rsObjects!systemCode & "', '" & _
                              rsObjects!Label & "', '" & relativePath & "')"
      rsObjects.MoveNext
   Loop
End Sub

Private Sub imgNewFolder_Click()
   Dim newFolder As String, newPath As String
   newFolder = InputBox("New subfolder name.", "Create New Subfolder")
   If newFolder <> "" Then
      If Right(dirDestination, 1) = "\" Then
         '  Just a drive has the \ after it, eg. "C:\"
         newPath = dirDestination & newFolder
      Else
         newPath = dirDestination & "\" & newFolder
      End If
      MkDir newPath
   End If
   dirDestination.Refresh
   dirDestination = newPath
End Sub

Private Sub lblExpression_Change()
   Dim colorSet As String, i As Integer
   
   If lblExpression = "" Then
      lblColorSet.visible = False
      lstColorSets.visible = False
      Set mappWindow.dbExpression = Nothing
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   lstColorSets.visible = True
   lblColorSet.visible = True
   lstColorSets.Clear
On Error GoTo ErrorHandler
   Set dbNewExpression = OpenDatabase(expressionDB, False, True)
   Set rsNewColorSets = dbNewExpression.OpenRecordset("SELECT * FROM ColorSet ORDER BY SetNo")
   If Not rsNewColorSets.EOF Then '================================================Color Sets Exist
      Do Until rsNewColorSets.EOF '--------------------------------------------Get Color Set Titles
         lstColorSets.AddItem rsNewColorSets!colorSet
         For i = 1 To colorIndexes(0) '......................Select Color Sets Active In frmDrafter
            If rsNewColorSets!setNo = colorIndexes(i) Then
               lstColorSets.selected(lstColorSets.ListCount - 1) = True
               Exit For                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
         Next i
'         If Not mappWindow.rsColorSet Is Nothing Then  'There is a current Color Set in MAPP window
'            If rsNewColorSets!colorSet = mappWindow.rsColorSet!colorSet Then
'               '  Set selected Color Set to any color set open in the MAPP window
'               lstColorSets.selected(lstColorSets.ListCount - 1) = True
'            End If
'         End If
         rsNewColorSets.MoveNext
      Loop
   End If
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
ErrorHandler:
   Select Case Err.number
   Case 3051
      MsgBox "Cannot open Expression Dataset. It may be set to Read-only through windows " _
             & "or be in use in some other place.", vbExclamation + vbOKOnly, _
             "Opening Expression Dataset"
      expressionDB = ""
      lblExpression = ""
      lblColorSet.visible = False
      lstColorSets.visible = False
      Set mappWindow.dbExpression = Nothing
   End Select
End Sub

Private Sub cmdChoose_Click()
On Error GoTo OpenError
Retry:
   With dlgDialog
      .CancelError = True
      .DialogTitle = "Choose Expression Dataset"
      .InitDir = GetFolder(mruDataSet)
      .Filter = "newExpression (.gex)|*.gex"
      .FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
      .ShowOpen
      newExpression = .FileName
   End With
   If InStr(newExpression, ".") = 0 Then
      newExpression = newExpression & ".gex"
   End If
   
   If InStr(1, newExpression, "\" & Dir(newExpression), vbTextCompare) = 0 Then
      If MsgBox("newExpression dataset '" & newExpression & " does not exist.", _
               vbExclamation + vbRetryCancel, "Open Expression Dataset") = vbCancel Then
         GoTo ExitSub         'Canceled choosing a dataset >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Else
         GoTo Retry                                        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      End If
   End If
   expressionDB = newExpression
   lblExpression = FileAbbrev(expressionDB)
ExitSub:
   Exit Sub
   
OpenError:
   Select Case Err.number
   Case 32755                                                                       'Cancel clicked
   Case Else
      MsgBox Err.Description, vbCritical, "Open Expression Dataset Error"
   End Select
   On Error GoTo 0
   Resume ExitSub
End Sub
Private Sub cmdCancel_Click()
   If creatingMappSet Then '++++++++++++++++++++++++++++++++++++++++ Quit Clicked During Processing
      Set dbGeneIndex = Nothing
      Kill appPath & "GeneIndex.$tm"
      lblCurrentMAPP.foreColor = vbRed
      lblCurrentMAPP = "Stopping process. Please be patient."
      creatingMappSet = False
      MousePointer = vbDefault
      DoEvents
      Exit Sub
   End If
   If Dir(appPath & "GeneIndex.$tm") <> "" Then
      Set dbGeneIndex = Nothing
      Kill appPath & "GeneIndex.$tm"
   End If
   Hide
   DoEvents
End Sub

Private Sub drvDestination_Change()
On Error GoTo ErrorHandler
   dirDestination.path = drvDestination.drive
   dirDestination.Refresh
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
ErrorHandler:
   MsgBox "Invalid drive; choose another.", vbExclamation + vbOKOnly, "Destination Drive"
   drvDestination.drive = "C:"
   dirDestination.path = drvDestination.drive
   dirDestination.Refresh
   drvDestination.SetFocus
End Sub
Private Sub drvSource_Change()
On Error GoTo ErrorHandler
   dirSource.path = drvSource.drive
   dirSource.Refresh
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
ErrorHandler:
   MsgBox "Invalid drive; choose another.", vbExclamation + vbOKOnly, "Source Drive"
   drvSource.drive = "C:"
   dirSource.path = drvSource.drive
   dirSource.Refresh
   drvSource.SetFocus
End Sub

Private Sub cmdAll_Click()
   Dim i As Integer
   
   For i = 0 To lstColorSets.ListCount - 1
      lstColorSets.selected(i) = True
   Next i
End Sub

Private Sub cmdNone_Click()
   Dim i As Integer
   
   For i = 0 To lstColorSets.ListCount - 1
      lstColorSets.selected(i) = False
   Next i
End Sub


