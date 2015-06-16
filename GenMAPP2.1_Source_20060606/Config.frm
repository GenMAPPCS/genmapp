VERSION 5.00
Begin VB.Form frmConfig 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GenMAPP Data Paths"
   ClientHeight    =   3504
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   7452
   ControlBox      =   0   'False
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3504
   ScaleWidth      =   7452
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInfo 
      Caption         =   "More Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   60
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   1992
   End
   Begin VB.TextBox txtMessage1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   912
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Config.frx":08CA
      Top             =   2100
      Width           =   7332
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   912
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Config.frx":08D8
      Top             =   60
      Width           =   7332
   End
   Begin VB.CommandButton cmdRevert 
      Caption         =   "&Revert to Default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   2700
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1992
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Height          =   372
      Left            =   6300
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3060
      Width           =   1092
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1092
   End
   Begin VB.TextBox txtGenMAPP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1020
      Width           =   4512
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\Gene Databases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   5160
      TabIndex        =   13
      Top             =   1320
      Width           =   1560
   End
   Begin VB.Line Lin 
      BorderWidth     =   2
      Index           =   3
      X1              =   5160
      X2              =   4860
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Lin 
      BorderWidth     =   2
      Index           =   2
      X1              =   4860
      X2              =   5160
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Lin 
      BorderWidth     =   2
      Index           =   1
      X1              =   4860
      X2              =   5160
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Lin 
      BorderWidth     =   2
      Index           =   0
      X1              =   4860
      X2              =   4860
      Y1              =   1320
      Y2              =   1980
   End
   Begin VB.Label lblGenMAPPFolder 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4740
      TabIndex        =   9
      Top             =   1080
      Width           =   48
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GenMAPP folder stored on your computer?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   540
      Width           =   3792
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\Expression Datasets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5160
      TabIndex        =   5
      Top             =   1800
      Width           =   1920
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\MAPPs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   5160
      TabIndex        =   4
      Top             =   1560
      Width           =   732
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "the subfolders MAPPs 1.0 and Expression Datasets.  Where do you want this"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   6768
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "With your GenMAPP Package you received a folder named GenMAPP containing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7104
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const DEFAULT_FOLDER = "C:"
Dim MAPPsPathCreated As String, expressionsPathCreated As String
Dim message As String, message1 As String

Private Sub Form_Load()
   Dim MAPPsRoot As String
   
   Caption = "GenMAPP Data Paths"
   message = "Please select the root folder for the GenMAPP MAPP and Expression Dataset " _
           & "files. GenMAPP needs to know the location of this folder for future events " _
           & "such as updating your files."
   txtMessage = message
'   message1 = "With your GenMAPP Package you received sample MAPPs and Expression Datasets. " _
'            & "These will be put in subfolders titled GenMAPP\MAPPs\MAPPs 9-07-01 " _
'            & "and GenMAPP\Expression Datasets\Expression Datasets 9-07-01"
   txtMessage1 = "" 'message1
   lblGenMAPPFolder = "\GenMAPP 2 Data"
   txtGenMAPP = DEFAULT_FOLDER
   cmdRevert.Visible = False
End Sub
Private Sub Form_Activate()
   If Tag = "New base" Then
      Top = Top - 2700
      MsgBox "GenMAPP has saved your last-used folders for your MAPPs and Expression Datasets, " _
             & "but it also requires an overall folder for making updates, etc. After " _
             & "clicking OK here, select your overall folder and make sure your MAPPs and " _
             & "Expression Datasets are under the subfolders shown before clicking OK in " _
             & "the GenMAPP Data Paths dialog." & vbCrLf & vbCrLf _
             & "For example, you may have chosen the default " & lblGenMAPPFolder & " as your base " _
             & "folder and have MAPPs under " & lblGenMAPPFolder & "\MAPPs. Changing " _
             & "that folder to " & lblGenMAPPFolder & "\MAPPs Other or moving the " _
             & "individual MAPPs to " & lblGenMAPPFolder & "\MAPPs would be acceptable solutions.", _
             vbInformation + vbOKOnly, "GenMAPP Configuration"
      Tag = ""
   End If
End Sub
Private Sub cmdBrowse_Click()
   txtGenMAPP = Dat(txtGenMAPP)
   With frmFolder
      .Caption = "GenMAPP Data Paths"
      .lblMessage = "Choose drive and folder to store the GenMAPP folder containing"
      .lblMessage2 = "the subfolders Gene Databases,  MAPPs, and Expression Datasets."
      .folders.path = txtGenMAPP & "\"
      .drives.drive = "C"
      .folders.Tag = ""
      .show vbModal
      If .folders.Tag <> "Cancel" Then                                               'Not cancelled
         txtGenMAPP = .folders.path
         If Right(txtGenMAPP, 1) = "\" Then                                        'Dump trailing \
            txtGenMAPP = Left(txtGenMAPP, Len(txtGenMAPP) - 1)
         End If
      End If
   End With
End Sub
Private Sub cmdInfo_Click()
   MsgBox "GenMAPP has an updating facility that brings your MAPPs and Expression Datasets " _
          & "in agreement with new versions of the GenMAPP program and/or gene databases. " _
          & "This facility must know where to look for your MAPPs and Expression Datasets " _
          & "in order to operate. You must select a root folder for GenMAPP-related files, " _
          & "under which GenMAPP will create the subfolders MAPPs and Expression Datasets." _
          & vbCrLf & vbCrLf _
          & "For example, if you accept the default C: as the root folder, GenMAPP will " _
          & "create " & lblGenMAPPFolder & "\MAPPs and " & lblGenMAPPFolder _
          & "\Expression Datasets. You may also " _
          & "create further subfolders for your MAPPs and/or Expression Datasets. For example, " _
          & "you may wish to create " & lblGenMAPPFolder & "\MyMAPPs." _
          & vbCrLf & vbCrLf _
          & "You may store your MAPPs and Expression Datasets anywhere on your computer, " _
          & "but the GenMAPP updating program will only see your files if they are in a " _
          & "subfolder of your GenMAPP Data Paths folder.", _
          vbInformation + vbOKOnly, lblGenMAPPFolder & " Data Paths"
End Sub

Private Sub cmdOK_Click()
   Dim basePath As String, geneDBFolder As String, MappFolder As String, expressionFolder As String
   Dim i As Integer
   
   txtGenMAPP = Dat(txtGenMAPP)
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Base Folder
   basePath = txtGenMAPP & lblGenMAPPFolder & "\"
   If Not ValidPathName(basePath) Then
      MsgBox "Path name" & vbCrLf & basePath & vbCrLf & "invalid.", vbExclamation + vbOKOnly, _
             "Configuration"
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
'   If Not PathCheck(basePath) Then Exit Sub                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
On Error GoTo ErrorHandler

   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Add Gene DB Folder
   If AddFolder(basePath & "Gene Databases") = "<ERROR>" Then
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Add MAPPs Folder
   If AddFolder(basePath & "MAPPs") = "<ERROR>" Then
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Add Expressions Folder
   If AddFolder(basePath & "Expression Datasets") = "<ERROR>" Then
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Add Exports Folder
   If AddFolder(basePath & "Exports") = "<ERROR>" Then
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Write Config File
   MousePointer = vbHourglass
   cfgBaseFolder = basePath
   mruGeneDB = basePath & "Gene Databases\"
   mruMappPath = basePath & "MAPPs\"
   mruDataSet = basePath & "Expression Datasets\"
   mruCatalog = basePath & "Catalogs\"
   mruExportPath = basePath & "Exports\"
   WriteConfig
      
   '************************************************************************************ Move Files
   'Should not have to do any of this stuff with new install ?????????????????????????????
   Dim dataFile As String
   dataFile = Dir(appPath & "*.mapp")                                                   'Move MAPPs
   Do While dataFile <> ""
      FileCopy appPath & dataFile, mruMappPath & dataFile
         '  Probably have to write a bunch of stuff in here for specific installation  ????????????
'      FileCopy appPath & dataFile, MappFolder & "\MAPPs 9-07-01\mu_sample_MAPPs\" & dataFile
      Kill appPath & dataFile
      dataFile = Dir
   Loop
   dataFile = Dir(appPath & "*.gex")                                      'Move Expression Datasets
   Do While dataFile <> ""
      FileCopy appPath & dataFile, GetFolder(mruDataSet) & dataFile
'      FileCopy appPath & dataFile, expressionFolder & "\Expression Datasets 9-07-01\" & dataFile
      Kill appPath & dataFile
      dataFile = Dir
   Loop
   dataFile = Dir(appPath & "*Archive*.exe")                       'Move Self-Extracting MAPP Files
   Do While dataFile <> ""
      FileCopy appPath & dataFile, mruMappPath & "Archives\" & dataFile
      Kill appPath & dataFile
      dataFile = Dir
   Loop
   Hide
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
ErrorHandler:
   Select Case Err.number
   Case 52, 75, 3043
      MsgBox "Cannot save to this path. This may be a read-only drive, such as a CD-ROM, " _
             & "or a removable drive with no disk in it.", vbExclamation + vbOKOnly, _
             "Save Configuration Error"
   Case Else
      MsgBox "Problem saving configuration.", vbCritical + vbOKOnly, _
             "Save Configuration Error"
   End Select
End Sub

Private Sub cmdRevert_Click()
   txtGenMAPP = DEFAULT_FOLDER
   cmdRevert.Visible = False
End Sub


Private Sub txtGenMapp_Change()
   Dim slash As Integer
   
'   slash = InStrRev(txtGenMAPP, "\")
'   If slash = 0 Then slash = Len(txtGenMAPP) + 1
'   txtGenMAPP = Left(txtGenMAPP, slash - 1) _
'                  & Mid(DEFAULT_EXPRESSIONS, InStrRev(DEFAULT_EXPRESSIONS, "\"))
   If UCase(txtGenMAPP) <> UCase(DEFAULT_FOLDER) Then
      cmdRevert.Visible = True
   Else
      cmdRevert.Visible = False
   End If
End Sub

Private Sub txtMessage_Change()
   txtMessage = message
End Sub
Private Sub txtMessage1_Change()
   txtMessage1 = message1
End Sub
