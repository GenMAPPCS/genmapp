VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConvert 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert MAPPs and Expression Datasets"
   ClientHeight    =   7785
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "Convert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   519
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "Deselect All"
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
      Left            =   2700
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1332
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
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
      Left            =   1560
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1092
   End
   Begin VB.ListBox lstSwitchTo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   4200
      TabIndex        =   34
      ToolTipText     =   "You may convert only to a single system."
      Top             =   3420
      Width           =   3972
   End
   Begin VB.ListBox lstSwitchFrom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   60
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   32
      ToolTipText     =   "Ctrl- or shift click to select multiple systems. Systems in [brackets] not available."
      Top             =   3420
      Width           =   3972
   End
   Begin VB.Frame fra 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   912
      Left            =   5040
      TabIndex        =   26
      Top             =   1080
      Width           =   3252
      Begin VB.OptionButton optSwitch 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Switch Gene ID system"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   3132
      End
      Begin VB.OptionButton optPrevVersion 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Convert from previous version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   0
         TabIndex        =   29
         Top             =   240
         Width           =   3132
      End
      Begin VB.CheckBox chkConvertSwissProt 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Convert UniProt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   168
         Left            =   240
         TabIndex        =   28
         ToolTipText     =   "Convert GenBank to Model Organism Database IDs."
         Top             =   720
         Width           =   2052
      End
      Begin VB.CheckBox chkConvertGenBank 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Convert GenBank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   168
         Left            =   240
         TabIndex        =   27
         ToolTipText     =   "Convert GenBank to Model Organism Database IDs."
         Top             =   480
         Width           =   2172
      End
   End
   Begin VB.CheckBox chkSubfolders 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Include Subfolders"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   1500
      Value           =   1  'Checked
      Width           =   2172
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   2640
      TabIndex        =   17
      Top             =   1080
      Width           =   2232
      Begin VB.OptionButton optFolder 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Entire folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   0
         TabIndex        =   19
         ToolTipText     =   "Convert MAPPs to latest version."
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
      End
      Begin VB.OptionButton optSingleFile 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Single file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   0
         TabIndex        =   18
         ToolTipText     =   "Convert Expression Datasets to latest version."
         Top             =   0
         Width           =   1512
      End
   End
   Begin VB.Frame fraFileType 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   432
      Left            =   60
      TabIndex        =   13
      Top             =   1080
      Width           =   2412
      Begin VB.OptionButton optExpression 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Expression Datasets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "Convert Expression Datasets to latest version."
         Top             =   0
         Width           =   2352
      End
      Begin VB.OptionButton optMapps 
         BackColor       =   &H00C0FFFF&
         Caption         =   "MAPPs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "Convert MAPPs to latest version."
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
      End
   End
   Begin VB.CommandButton cmdGeneDB 
      Caption         =   "&Gene Database"
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
      Left            =   60
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Width           =   1692
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7320
      Width           =   972
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Height          =   312
      Left            =   60
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   8112
      _ExtentX        =   14314
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdDestination 
      Caption         =   "&Destination"
      Default         =   -1  'True
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
      Left            =   60
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1692
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "&Source"
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
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1692
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
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
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5760
      Width           =   972
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   60
      Top             =   8100
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label lblSwitchTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Switch To"
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
      Left            =   4200
      TabIndex        =   35
      Top             =   3180
      Width           =   876
   End
   Begin VB.Label lblSwitchFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Switch From"
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
      Left            =   60
      TabIndex        =   33
      Top             =   3180
      Width           =   1116
   End
   Begin VB.Label lblBuild 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Build"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   240
      Left            =   6720
      TabIndex        =   25
      Top             =   7560
      Width           =   456
   End
   Begin VB.Label lblErrors 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   840
      TabIndex        =   24
      Top             =   6900
      Width           =   108
   End
   Begin VB.Label lblPrgValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PrgValue"
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
      Left            =   3480
      TabIndex        =   23
      Top             =   6900
      Width           =   816
   End
   Begin VB.Label lblPrgMax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PrgMax"
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
      Left            =   60
      TabIndex        =   22
      Top             =   6900
      Width           =   684
   End
   Begin VB.Label lblDetail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detail"
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
      TabIndex        =   21
      Top             =   6360
      Width           =   528
   End
   Begin VB.Label lblOperation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operation"
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
      Left            =   60
      TabIndex        =   20
      Top             =   6120
      Width           =   876
   End
   Begin VB.Label lblDestinationFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destination file"
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
      Left            =   60
      TabIndex        =   16
      Top             =   7440
      Width           =   1332
   End
   Begin VB.Label lblSourceFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current file"
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
      Left            =   60
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   984
   End
   Begin VB.Label lblMOD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model organism database: "
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
      TabIndex        =   10
      Top             =   660
      Width           =   2388
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose source"
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
      TabIndex        =   3
      Top             =   2160
      Width           =   1344
   End
   Begin VB.Label lblGeneDB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Gene Database"
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
      TabIndex        =   9
      Top             =   420
      Width           =   1692
   End
   Begin VB.Label lblFile 
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
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   6120
      Width           =   48
   End
   Begin VB.Label lblDestination 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose destination"
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
      TabIndex        =   4
      Top             =   2880
      Width           =   1704
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dbGene As Database
Dim commandLine As String
Dim dbExpression As Database, dbMapp As Database
Dim modTable As String, modCode As String, modSpecies As String
Dim convertAll As Boolean
Dim source As String, destination As String                      'Full length for the abbrev labels
   '  Change one and must change the other. "Choose source" in label is "" in variable.
Dim conversionExceptionFile As String                  'Contains exception and warning messages for
                                                       'folder conversions

'Private Sub chkConvertGenBanks_Click()
'   If chkconvertgenbanks = vbChecked Then
'      chkConvertToMOD = vbUnchecked
'   End If
'End Sub
'Private Sub chkConvertToMOD_Click()
'   If chkConvertToMOD = vbChecked Then
'      chkconvertgenbanks = vbUnchecked
'   End If
'End Sub
'
Private Sub cmdExit_Click()
   End
End Sub

Private Sub Form_Load()

'Debug.Print """" & DriveCheck("g:\temp") & """"
'Dim dbExpression As Database
'
'Set dbExpression = OpenDatabase("D:\GenMAPPv2_new\Datasets\Development Data MAS 5v2_short.gex")
'EDToRawData dbExpression
'BuildFileTree "C:\GenMAPP_old\MAPPs\MAPPs 9-07-01\hu_MAPPArchive\", "*.mapp"

'Dim uniqueIDs(MAX_GENES) As String
'Set dbGene = OpenDatabase("D:\GenMAPPv2_new\Gene Databases\Mm-Std_20040411.gdb")
'i = FindUniqueIDs("AA536941", "L", uniqueIDs(), dbGene)
'Stop
   
   ReadConfig
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Initial Defaults For Controls
   Show
   lblBuild = BUILD
   lblMOD = ""
   lblPrgMax.Visible = False
   lblPrgValue.Visible = False
   prgProgress.Visible = False
   lblOperation.Visible = False
   lblDetail.Visible = False
   lblErrors.Visible = False
   cmdDestination.Visible = True
   lblDestination.Visible = True
   optMapps_Click                                                      'Sets source and destination
'   chkConvertToMOD.value = vbChecked
   optFolder_Click                                                     'Sets source and destination
   chkSubfolders.value = vbChecked
   optPrevVersion = True                                                    'Triggers Click() event
   chkConvertGenBank.value = vbChecked
'   fillswitchfrom
'   fillswitchto
   lblDestinationFile = ""
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Open Gene Database
   Set dbGene = Nothing                                                         'Default to Nothing
   OpenGeneDB dbGene, mruGeneDB                                 'May return Nothing if no mruGeneDB
      '  This does not call CheckForGo
   LoadSwitchFrom dbGene
   LoadSwitchTo dbGene
   CheckForGo
End Sub
Sub LoadSwitchFrom(dbGene As Database)
   Static selectedSystems(MAX_SYSTEMS) As String, lastSelectedSystem As Integer          'One-based
   Dim rsSystems As Recordset, tdf As TableDef, tdf1 As TableDef
   Dim systemTo As String, system As String, dash As Integer, hyphen As Integer
   Dim i As Integer, switchPossible As Boolean
   
   lstSwitchFrom.Clear
   If dbGene Is Nothing Then Exit Sub                      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   If lstSwitchTo.ListIndex = -1 Then
      systemTo = ""
   Else
      systemTo = lstSwitchTo.List(lstSwitchTo.ListIndex)
   End If
   
   For Each tdf In dbGene.TableDefs '++++++++++++++++++++++++++++++++++++++++ Each Table In Gene DB
      dash = InStr(tdf.name, "-")
      If dash <> 0 Then '=====================================================Is Relationship Table
         system = Left(tdf.name, dash - 1)                                          'Primary system
         GoSub FillList
         system = Mid(tdf.name, dash + 1)                                           'Related system
         GoSub FillList
      End If
   Next tdf
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         
FillList: '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Put System In List
   If system = "GOCount" Then Return                       '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   For i = 0 To lstSwitchFrom.ListCount - 1 '=========================See If System Already In List
      If system = lstSwitchFrom.List(i) Or "[" & system & "]" = lstSwitchFrom.List(i) Then Exit For
   Next i
   If i > lstSwitchFrom.ListCount - 1 Then '=============================System Not In List, Add It
'      If systemTo = "" Then '------------------------------------------------No systemTo Chosen Yet
'         switchPossible = True
'      Else '------------------------------------------------------------------------systemTo Chosen
'         switchPossible = False
'         For Each tdf1 In dbGene.TableDefs '__________________________Search For Relationship Table
'            If tdf1.name = system & "-" & systemTo Or tdf1.name = systemTo & "-" & system Then
'               switchPossible = True
'               Exit For                                       'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'            End If
'         Next tdf1
'      End If
'      If switchPossible Then '----------------------------------------------------------Add To List
         lstSwitchFrom.AddItem system
'      Else
'         lstSwitchFrom.AddItem "[" & system & "]"
'      End If
   End If
   Return
End Sub
Sub LoadSwitchTo(dbGene As Database)
   Static selectedSystems(MAX_SYSTEMS) As String, lastSelectedSystem As Integer                   'One-based
   Dim rsSystems As Recordset, tdf As TableDef, tdf1 As TableDef
   Dim systemTo As String, system As String, dash As Integer, hyphen As Integer
   Dim i As Integer, switchPossible As Boolean
   
   lstSwitchTo.Clear
   If dbGene Is Nothing Then Exit Sub                      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   For Each tdf In dbGene.TableDefs '++++++++++++++++++++++++++++++++++++++++ Each Table In Gene DB
      dash = InStr(tdf.name, "-")
      If dash <> 0 Then '=====================================================Is Relationship Table
         system = Left(tdf.name, dash - 1)                                          'Primary system
         GoSub FillList
         system = Mid(tdf.name, dash + 1)                                           'Related system
         GoSub FillList
      End If
   Next tdf
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         
FillList: '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Put System In List
   If system = "GOCount" Then Return                       '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   If system = "GeneOntology" Then Return                  '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   If system = "InterPro" Then Return                      '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   For i = 0 To lstSwitchTo.ListCount - 1 '===========================See If System Already In List
      If system = lstSwitchTo.List(i) Then Exit For        'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   Next i
   If i > lstSwitchTo.ListCount - 1 Then '===============================System Not In List, Add It
      For Each tdf1 In dbGene.TableDefs '__________________________Search For Relationship Table
'         If tdf1.name = system & "-" & systemTo Or tdf1.name = systemTo & "-" & system Then
            lstSwitchTo.AddItem system
            Exit For                                       'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'         End If
      Next tdf1
   End If
   Return
End Sub

Private Sub lstSwitchFrom_Click()
   CheckForGo
End Sub

Private Sub lstSwitchTo_Click()
   '  Uses public dbGene database
   Dim tdf As TableDef, system As String, systemTo As String, switchPossible As Boolean
   Dim i As Integer
   
   For i = 0 To lstSwitchFrom.ListCount - 1 '================================See If Switch Possible
      system = lstSwitchFrom.List(i)
      If Left(system, 1) = "[" Then
         system = Mid(system, 2, Len(system) - 2)
      End If
      systemTo = lstSwitchTo.List(lstSwitchTo.ListIndex)
      switchPossible = False
      For Each tdf In dbGene.TableDefs
         If tdf.name = system & "-" & systemTo Or tdf.name = systemTo & "-" & system Then
            switchPossible = True
            Exit For                                       'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      Next tdf
      If switchPossible Then
         lstSwitchFrom.List(i) = system
      Else
         lstSwitchFrom.List(i) = "[" & system & "]"
      End If
   Next i
   CheckForGo
End Sub
Private Sub cmdSelectAll_Click()
   Dim i As Integer
   
   For i = 0 To lstSwitchFrom.ListCount - 1
      lstSwitchFrom.Selected(i) = True
   Next i
   CheckForGo
End Sub
Private Sub cmdDeselectAll_Click()
   Dim i As Integer
   
   For i = 0 To lstSwitchFrom.ListCount - 1
      lstSwitchFrom.Selected(i) = False
   Next i
   CheckForGo
End Sub

Private Sub cmdGeneDB_Click() '************************************************************ Gene DB
   OpenGeneDB dbGene, "**OPEN**"
   LoadSwitchFrom dbGene
   LoadSwitchTo dbGene
   CheckForGo
End Sub

Private Sub cmdSource_Click() '********************************************** Source File Or Folder
   If optSingleFile Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Single File
On Error GoTo OpenError
Retry:
      With dlgDialog
         .CancelError = True
         If optMapps Then
            .Filter = "MAPPs (.mapp)|*.mapp"
         Else
            .Filter = "Expression (.gex)|*.gex"
         End If
         .DialogTitle = "Source for Conversion"
         .InitDir = source
         .FileName = ""
         .FLAGS = cdlOFNExplorer + cdlOFNHideReadOnly
         .ShowOpen
         source = .FileName
         lblSource = FileAbbrev(source, 60)
         lblSource.ToolTipText = source
         convertAll = False
         chkSubfolders.Visible = False
         If InStr(source, ".") = 0 Then
            If optMapps Then
               source = source & ".mapp"
            Else
               source = source & ".gex"
            End If
            lblSource = FileAbbrev(source, 60)
            lblSource.ToolTipText = source
         End If
         If Dir(source) = "" Then
            MsgBox "File " & source & " does not exist.", vbExclamation, "Opening Source"
            Exit Sub                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
         If Right(destination, 1) = "\" Then '-------------------------------Dest File Was A Folder
            destination = destination & GetFile(source)                 'Add source file name to it
         Else '--------------------------------------------------------------Destination Was A File
            destination = GetFolder(destination) & GetFile(source)      'Add source file name to it
         End If
         lblDestination = FileAbbrev(destination, 60)
         lblDestination.ToolTipText = destination
      End With
   Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Folder
On Error GoTo FolderError
      With frmFolder
         .lblNewFolder.Visible = False
         .Caption = "Choose Source Folder"
         .BackColor = vbGray
         .lblMessage = "Choose drive and folder from which to convert."
         .lblMessage2 = ""
         If optMapps Then
            .drives.drive = GetDrive(mruMappPath)
            .folders.path = GetFolder(mruMappPath)
         Else
            .drives.drive = GetDrive(mruDataSet)
            .folders.path = GetFolder(mruDataSet)
         End If
         .folders.Tag = ""
         .Show vbModal
         If .folders.Tag <> "Cancel" Then                                            'Not cancelled
            source = .folders.path & "\"
            lblSource = FileAbbrev(source, 60)
            lblSource.ToolTipText = source
         End If
      End With
      If destination <> "" And destination <> "Same as source" _
            And Right(destination, 1) <> "\" Then
         destination = GetFolder(destination)
         lblDestination = FileAbbrev(destination, 60)
         lblDestination.ToolTipText = destination
      End If
On Error GoTo 0
   End If
   
   If optMapps Then
      mruMappPath = GetFolder(source)
   Else
      mruDataSet = GetFolder(source)
   End If
   lblSource.Visible = True
ExitSub:
   CheckForGo
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

OpenError:
   Select Case Err
   Case 32755                                                                       'Cancel clicked
   Case Else
      FatalError "frmConvert:cmdSource", Err.Description & "  " & source
   End Select
   On Error GoTo 0
   Resume ExitSub

FolderError:
   Select Case Err.number
   Case 68, 76                                                                      'Path not found
      If optMapps Then
         mruMappPath = "C:\"
      Else
         mruDataSet = "C:\"
      End If
      Resume
   Case 32755                                                                       'Cancel clicked
   Case Else
      FatalError "frmConvert:cmdSource", Err.Description & "  " & source
   End Select
   On Error GoTo 0
   Resume ExitSub
End Sub
Private Sub cmdDestination_Click() '************************************ Destination File Or Folder
   Dim expression As String
   
   If optSingleFile Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Single File
On Error GoTo OpenError
Retry:
      With dlgDialog
         .CancelError = True
         If destination = "Same as source" Then
            .FileName = source
         Else
            If Right(destination, 1) = "\" Then
               If GetFile(source) = "" Then                                 'Source not chosen yet
                  .FileName = destination & "*"                    'Dump "\"
'                  .FileName = Left(destination, Len(destination) - 1)                    'Dump "\"
               Else
                  .FileName = destination & GetFile(source)
               End If
            Else
               .FileName = destination
            End If
         End If
         .DialogTitle = "Destination for Conversion"
         .FLAGS = cdlOFNExplorer + cdlOFNHideReadOnly
         .ShowSave
         destination = .FileName
         If InStr(destination, ".") = 0 Then
            If optMapps Then
               destination = destination & ".mapp"
            Else
               destination = destination & ".gex"
            End If
         End If
         lblDestination = FileAbbrev(destination, 60)
         lblDestination.ToolTipText = destination
      End With
   Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Folder
      With frmFolder
         .Tag = "WRITE"
         .lblNewFolder.Visible = True
         .Caption = "Choose Destination Folder"
         .BackColor = vbGray
         .lblMessage = "Choose drive and folder for converted files."
         .lblMessage2 = ""
         If Dir(destination, vbDirectory) = "" Then                     'Invalid destination folder
            destination = "C:\"
         End If
         .drives.drive = GetDrive(destination)
         .folders.path = GetFolder(destination)
         .folders.Tag = ""
         .Show vbModal
         If .folders.Tag <> "Cancel" Then                                            'Not cancelled
            If Right(.folders.path, 1) = "\" Then                        'Don't add extra \ to root
               destination = .folders.path
            Else
               destination = .folders.path & "\"
            End If
            lblDestination = FileAbbrev(destination, 60)
            lblDestination.ToolTipText = destination
         End If
         .Tag = ""
      End With
      If destination <> "" And destination <> "Same as source" _
            And Right(destination, 1) <> "\" Then
         destination = GetFolder(destination)
         lblDestination = FileAbbrev(destination, 60)
         lblDestination.ToolTipText = destination
      End If
   End If
     
   lblDestination.Visible = True
   
ExitSub:
   CheckForGo
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

OpenError:
   Select Case Err.number
   Case 32755                                                                       'Cancel clicked
   Case 20477                                                                        'Bad file name
      Resume ExitSub
   Case Else
      FatalError "frmConvert:cmdDestination", Err.Description & "  " & destination
   End Select
   On Error GoTo 0
   Resume ExitSub
End Sub
Private Sub cmdGo_Click() '********************************************************* Do Conversions
   '  Globals required:
   '     source      Source file or folder for conversion. Must be valid before CheckForGo()
   '                 makes cmdGo visible.
   '     destination Destination file or folder for conversion. Must be valid before CheckForGo()
   '                 makes cmdGo visible.
   Dim dest As String, oldSource As String, dbExpression As Database
   Dim rs As Recordset, i As Integer
   Dim rawFile As String, returnExceptions As String, slash As Integer
   Dim tiles As Boolean        'At least one MAPP in the conversion is tiled (multiple conversions)
   Dim noConversion As Boolean          'Set to true if single-file conversion not made to suppress
                                        '"Conversion finished" message
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Update Configuration File
   If optMapps Then
      '  The mru's are set whether conversion successful or not. Also, setting them here
      '  means that the mru's are set to the base conversion folders for conversion of
      '  whole directory trees.
      mruMappConvertSource = GetFolder(source)
      UpdateConfig "mruMappConvertSource", mruMappConvertSource
   Else
      mruEDConvertSource = source
      UpdateConfig "mruEDConvertSource", mruEDConvertSource
   End If
   mruGeneDB = dbGene.name
   UpdateConfig "mruGeneDB", mruGeneDB
   mruMappPath = destination
   UpdateConfig "mruMAPPPath", mruMappPath
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For Conversion Tables
   If optPrevVersion And (chkConvertGenBank Or chkConvertSwissProt) Then
      Dim noGenBank As Boolean, noSwissProt As Boolean, msg As String
      DetermineMODConversionTables dbGene
      msg = ""
      If chkConvertGenBank = vbChecked And modGB = "" Then                              'No GenBank
         msg = """Convert GenBank"" "
      End If
      If chkConvertSwissProt = vbChecked And modSP = "" Then                          'No SwissProt
         If msg <> "" Then msg = msg & "and "
         msg = msg & """Convert UniProt"" "
      End If
      If msg <> "" Then                                                      'Some table(s) missing
         MsgBox "You have selected " & msg & "but your Gene Database does not have the proper " _
                & "conversion tables. You may have to use a ""-Converter"" " _
                & "Gene Database from the www.GenMAPP.org website.", _
                   vbExclamation + vbOKOnly, "Missing Conversion Tables"
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Else
   End If

   If optFolder Then '++++++++++++++++++++++++++++++++++++++++++++++++++ All MAPPs Or EDs In Folder
      conversionExceptionFile = destination & "ConversionExceptions.txt"
      Open conversionExceptionFile For Output As #99                         'Create new empty file
      Close #99
      ProcessFiles source, 0, tiles
      If FileLen(conversionExceptionFile) > 0 Then                     'Conversion exceptions exist
         MsgBox "GenMAPP has detected exceptions in one or more of your conversions. Look " & _
                "at the file " & vbCrLf & vbCrLf & conversionExceptionFile & vbCrLf & vbCrLf & _
                "for those exceptions.", vbExclamation + vbOKOnly, "Conversions"
      Else
         Kill conversionExceptionFile
      End If
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Single MAPP Or ED
      If Not ProcessAFile(source, destination, tiles) Then
         noConversion = True
      End If
   End If
   
   If tiles Then TileWarning
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Back To Before GO Conditions
   tiles = False
   If Not noConversion Then
      MsgBox "Conversion finished.", vbInformation + vbOKOnly, "Conversions"
   End If
   source = Left(source, InStrRev(source, "\"))
   lblSource = FileAbbrev(source)
   lblSource.ToolTipText = source
   destination = Left(destination, InStrRev(destination, "\"))
   lblDestination = FileAbbrev(destination)
   lblDestination.ToolTipText = destination
   cmdGo.Visible = False
   ReadConfig     'Does something else change config ???????????????
   MousePointer = vbDefault
End Sub
'***************************************************** Convert Folder And Subfolder Of MAPPs Or EDs
Function ProcessFiles(root As String, level As Integer, tiles As Boolean) As String
   '  Enter    root        Root folder for particular file, Eg:
   '                          D:\Datasets\GenMAPP.org MAPPs\Gene Family MAPPs\
   '                       Initially this is the root folder for the entire file set. As this
   '                       procedure is recursively called, root changes to the folder for a
   '                       particular file as determined by the Find Directory section.
   '           tiles       At least one MAPP in the conversion is tiled (multiple conversions)
   '  Return   True if successful
   '           Tiles       Set to true if at least one MAPP tiled
   Dim sourceFile As String, sourcePath As String, oldSource As String
   Dim destFile As String, destPath As String
   Dim destFolder As String                                                        'Relative folder
   Dim dirIndex As Integer, fileIndex As Integer, index As Integer
   Dim relativeFolder As String                     'Folder path relative to source and destination
   Dim ext As String, slash As Integer, i As Integer
   Dim rawFile As String
   Dim conversionCount As Integer

   lblSource.ToolTipText = source   '???????????????
   
   If optMapps Then
      ext = "*.mapp"
   Else
      ext = "*.gex"
   End If
   BuildFileTree root, ext, chkSubfolders
   
   Open App.path & "\TreeFile.$tm" For Input As #FILE_TREE
   
   conversionCount = 0
   tiles = False          'Will be set to true in SingleMAPPToMOD() if multiple substitutions found
   Do Until EOF(FILE_TREE)
      Line Input #FILE_TREE, sourcePath
      If Right(sourcePath, Len(ext) - 2) <> Right(ext, Len(ext) - 2) Then GoTo NextFile
         '  Unfortunately, the Dir function returns files that simply begin with the extension.
         '  For example, the extension "*.gex" will also return the file "whatever.gexz".
         '  Real dumb!
'      If Left(GetFile(sourcePath), 3) = "V1_" Then GoTo NextFile
      relativeFolder = Mid(GetFolder(sourcePath), Len(source) + 1)
      If Dir(destination & relativeFolder, vbDirectory) = "" Then
         AddFolder destination & relativeFolder
      End If
      lblSourceFile = FileAbbrev(sourcePath, 60)
      destPath = destination & relativeFolder & GetFile(sourcePath)
      If ProcessAFile(sourcePath, destPath, tiles) Then
         conversionCount = conversionCount + 1
      End If
NextFile:
   Loop
   Close #FILE_TREE
   Kill App.path & "\TreeFile.$tm"
   lblSourceFile = ""
   lblDestinationFile = conversionCount & " files converted"
   DoEvents
End Function
'************************************************************************ Convert Single MAPP Or ED
Function ProcessAFile(source As String, destination As String, tiles As Boolean) As Boolean
   '  Entry    source      Full path to source file to convert
   '           destination Full path of destination file
   '           tiles       For conversions of folders, this may enter as true if a previous
   '                       MAPP was tiled
   '  Return               True if file converted
   '           tiles       True if MAPP has been tiled (multiple substitutions for genes)
   '  Procedure must be in frmConvert to have access to controls on form
   Dim slash As Integer
   Dim oldSource As String
   Dim rawFile As String
   Dim tempDB As String                                                 'Temporary destination file
   Dim changeLog As String                                                   'Log file with changes
   ReDim fromSystems(MAX_SYSTEMS) As String
   Dim i As Integer, j As Integer, s As String
   
   s = FileWritable(destination)
   If s <> "" Then
      MsgBox "Conversion not made.", vbExclamation + vbOKOnly, "Path Not Available"
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For Existing Destination File
   slash = InStrRev(source, "\")
   If destination = source Then '======================================Converting To Same File Path
      '  Assume user wants to convert and don't ask
   Else '===============================================================Converting To New File Path
      If Dir(destination) <> "" And Not optFolder Then
         '  Don't ask if converting a folder
         If MsgBox("File " & destination & " already exists. Replace it?", _
                   vbExclamation + vbOKCancel, "Opening Destination") = vbCancel Then
            MousePointer = vbDefault
            Exit Function                                  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      End If
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Start Conversion
   tempDB = GetFolder(destination) & "Temp.$tm"
   changeLog = Left(destination, InStrRev(destination, ".")) & "log"
   If Dir(tempDB) <> "" Then Kill tempDB
   FileCopy source, tempDB                                     'Make new copy in destination folder
   lblDestinationFile = FileAbbrev(destination, 50)
   If optMapps Then '==================================================================Convert MAPP
      UpdateSingleMAPP tempDB                                           'Bring up to latest version
'      If optPrevVersion Then '-------------------------------------Convert GenBank And/Or SwissProt
'         If chkConvertGenBank = vbChecked And chkConvertSwissProt = vbChecked Then
'            ConvertGBorSPinFile "GS", dbGene, tempDB, changeLog, tiles
'         ElseIf chkConvertGenBank = vbChecked Then
'            ConvertGBorSPinFile "G", dbGene, tempDB, changeLog, tiles
'         ElseIf chkConvertSwissProt = vbChecked Then
'            ConvertGBorSPinFile "S", dbGene, tempDB, changeLog, tiles
'         End If
'      Else '--------------------------------------------------------------Switch To Specific System
'
''      If chkConvertToMOD Then '------------------------------------------------------Convert To MOD
''         tiles = False
''         SingleMAPPtoMOD destination, dbGene, tiles                                 'Convert to MOD
''      ElseIf chkConvertGenBanks Then '---------------------------------------------Convert GenBanks
''         ConvertGBorSPinFile "G", dbGene, destination, tiles
''            '  No tiles here????????????????
'      End If
   Else '===============================================================Convert Expression Datasets
      UpdateSingleDataset tempDB                                      'Bring up to latest version
'      If chkconvertgenbanks Then
'         '  There has never been a convert to MOD option here
'         ConvertGenBanksInFile dbGene, destination
'      End If
'      Set dbExpression = OpenDatabase(destination)
'      rawFile = EDToRawData(dbExpression)                                'Convert to raw data first
'      If optFolder Then
'         ConvertExpressionData rawFile, dbExpression, dbGene, _
'                              "To file: " & conversionExceptionFile                 'Then re-import
'            '  This will leave a .EX file if exceptions exist
'      Else
'         ConvertExpressionData rawFile, dbExpression, dbGene                        'Then re-import
'            '  This will leave a .EX file if exceptions exist
'      End If
'      Kill rawFile
   End If
   
   If optPrevVersion Then '++++++++++++++++++++++++++++++++++++++++++++ Previous To Current Version
      If chkConvertGenBank = vbChecked And chkConvertSwissProt = vbChecked Then
         ConvertGBorSPinFile "GS", dbGene, tempDB, changeLog, tiles
      ElseIf chkConvertGenBank = vbChecked Then
         ConvertGBorSPinFile "G", dbGene, tempDB, changeLog, tiles
      ElseIf chkConvertSwissProt = vbChecked Then
         ConvertGBorSPinFile "S", dbGene, tempDB, changeLog, tiles
      End If
      oldSource = Left(source, slash) & "V1_" & Mid(source, slash + 1)
         '  The original file (eg MyMAPP.mapp) will be named V1_MyMAPP.mapp
   Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Switch Gene ID Systems
      j = -1
      For i = 0 To lstSwitchFrom.ListCount - 1
         If Left(lstSwitchFrom.List(i), 1) <> "[" And lstSwitchFrom.Selected(i) Then
            j = j + 1
            fromSystems(j) = lstSwitchFrom.List(i)
         End If
      Next i
      ReDim Preserve fromSystems(j) As String
      SwitchIDsInFile fromSystems(), lstSwitchTo.List(lstSwitchTo.ListIndex), dbGene, _
                      tempDB, changeLog
      oldSource = Left(source, slash) & "Old_" & Mid(source, slash + 1)
         '  The original file (eg MyMAPP.mapp) will be named Old_MyMAPP.mapp
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Rename Files
   If Dir(oldSource) <> "" Then Kill oldSource
   Name source As oldSource
   If Dir(destination) <> "" Then Kill destination
   Name tempDB As destination
   ProcessAFile = True
End Function

Sub CheckForGo() '************************************************************* Ready To Show cmdGo
   Dim go As Boolean
   
   cmdGo.Visible = False
   If dbGene Is Nothing Then Exit Sub
   If source = "" Then Exit Sub
   If destination = "" Then Exit Sub
   If optSingleFile Then
On Error GoTo SourcePathError
      If Right(source, 1) = "\" Then Exit Sub
      If Dir(source) = "" Then
         MsgBox "Source file does not exist.", vbExclamation + vbOKOnly, "Single File Conversion"
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
On Error GoTo DestPathError
      If Right(destination, 1) = "\" Then
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Else
On Error GoTo SourcePathError
      If Dir(source, vbDirectory) = "" Then
         MsgBox "Source folder does not exist.", vbExclamation + vbOKOnly, "Folder Conversion"
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
On Error GoTo DestPathError
      If Dir(destination, vbDirectory) = "" Then
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   
   If optSwitch Then
      go = False
      For i = 0 To lstSwitchFrom.ListCount - 1
         If lstSwitchFrom.Selected(i) And Left(lstSwitchFrom.List(i), 1) <> "[" Then
            go = True
            Exit For                                       'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      Next i
      If i > lstSwitchFrom.ListCount - 1 Then Exit Sub     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      If lstSwitchTo.ListIndex = -1 Then Exit Sub          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If

   cmdGo.Visible = True
   cmdGo.Default = True
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
SourcePathError:
   MsgBox "Source file or folder does not exist.", vbExclamation + vbOKOnly, "Conversion"
DestPathError:
   MsgBox "Destination file or folder does not exist.", vbExclamation + vbOKOnly, "Conversion"
End Sub

'**************************************************************************** Opens A Gene Database
Sub OpenGeneDB(dbGene As Database, geneDB As String)
   '  Entry:
   '     dbGene   An open Gene Database or "Nothing"
   '     geneDB   Path and name of Gene Database to open
   '              If blank or ends in / (path but no name), closes any open geneDB and sets it
   '                 to Nothing
   '              If "**OPEN**" then display Open dialog to choose name
   '              If "**CLOSE**" then close dbGene, set to Nothing
   '                 and set GeneDB to "No Gene Database"
   '     frm      The Form making the call. If empty or Nothing, set to active form
   '              Used to call correct dialog box and set statusbar panel for Gene Database
   '  Return:
   '     dbGene   An open Gene Database or "Nothing"
   Dim rs As Recordset, prevGeneDB As String
   
   If Not dbGene Is Nothing Then '+++++++++++++++++++++++++++++++++++ Keep Track Of Current Gene DB
      prevGeneDB = dbGene.name
   End If
   
   If geneDB = "" Or Right(geneDB, 1) = "\" Then '++++++++++++++++++++++++++ No Gene Database Given
      If Not dbGene Is Nothing Then '----------------------------------Close Any Open Gene Database
         dbGene.Close
         Set dbGene = Nothing
         lblGeneDB = "No Gene Database"
         lblMOD = ""
      End If
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If

   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Open Gene Database
   If geneDB = "**OPEN**" Then '----------------------------------------------------Use Open Dialog
      With dlgDialog
On Error GoTo OpenError
         .CancelError = True
         .Filter = "Gene Databases (.gdb)|gdb"
         If mruGeneDB <> "" Then
            .InitDir = Left(mruGeneDB, InStrRev(mruGeneDB, "\") - 1)
         Else
            .InitDir = "C:\"
         End If
         .FileName = "*.gdb"
         .FLAGS = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNFileMustExist
         .ShowOpen
         geneDB = .FileName
      End With
      If InStr(geneDB, ".") = 0 Then
         geneDB = geneDB & ".gdb"
      End If
On Error GoTo 0
   End If

On Error GoTo DatabaseError
   Set dbGene = OpenDatabase(geneDB)
On Error GoTo 0
   Set rs = dbGene.OpenRecordset("SELECT * FROM Info")
   If InStr(rs!Version, "/") Then
      MsgBox "Gene database" & vbCrLf & vbCrLf & geneDB & vbCrLf & vbCrLf & "is an obsolete " _
             & "version and cannot be used with this release of GenMAPP.", _
             vbExclamation + vbOKOnly, "Opening Gene Database"
   End If
   lblGeneDB = geneDB
   
   If Dat(rs!species) = "" Then
      lblMOD = "No species selected"
   Else
      lblMOD = "Species selected: " & Mid(rs!species, 2, Len(rs!species) - 2)
   End If
   
   If rs!species = "|Homo sapiens|" Then
      chkConvertSwissProt.Visible = False
   Else
      chkConvertSwissProt.Visible = True
   End If
      
'   If DetermineMODConversionTables(dbGene) Then
'   Else
'      lblMOD = ""
'      Set dbGene = Nothing
'   End If

ExitSub:
'   CheckForGo
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Error Handlers
OpenError:
   If Err.number = 32755 Then                                                               'Cancel
      
   Else                                                                          'Other than Cancel
      FatalError "frmConvert:OpenGeneDB", Err.Description
   End If
   On Error GoTo 0
   Resume ExitSub
   
DatabaseError:
   MsgBox "Gene database" & vbCrLf & vbCrLf & geneDB & vbCrLf & vbCrLf & "Could not be opened. " _
          & "It may not exist, be set to Read-only, or in use by someone else.", vbExclamation + vbOKOnly, _
          "Opening Gene Database"
   mruGeneDB = ""
'  Don't reset database or statusbar
'   geneDB = "No Gene Database"
'   GoTo ExitSub                                            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
End Sub

Private Sub optExpression_Click()
   optSwitch.Visible = False
   optPrevVersion = True
   source = GetFolder(mruEDConvertSource)
   If source = "" Then
      source = GetFolder(mruDataSet)
   End If
   If source = "" Then
      lblSource = "Choose source"
      lblSource.ToolTipText = ""
   Else
      lblSource = FileAbbrev(source)
      lblSource.ToolTipText = source
   End If
   If destination = "" Then
      lblDestination = "Choose destination"
      lblDestination.ToolTipText = ""
   Else
      destination = GetFolder(mruDataSet)
      lblDestination = FileAbbrev(destination)
      lblDestination.ToolTipText = destination
   End If
'   If optMapps Then
'      chkconverttomod.Visible = True
'   Else
'      chkConvertToMOD.Visible = False
'   End If
   CheckForGo
End Sub
Private Sub optMapps_Click()
'   chkConvertToMOD.Visible = True
   optSwitch.Visible = True
   source = mruMappConvertSource
On Error GoTo SourcePathError
   If source = "" Then
      source = mruMappPath
   End If
   If source = "" Then
      lblSource = "Choose source"
      lblSource.ToolTipText = ""
   Else
      lblSource = FileAbbrev(source)
      lblSource.ToolTipText = source
   End If
   destination = mruMappPath
On Error GoTo DestPathError
   If destination = "" Then                                                 'No mruMappPath given
      lblDestination = "Choose destination"
      lblDestination.ToolTipText = ""
   ElseIf Dir(destination, vbDirectory) = "" Then                            'mruMappPath invalid
      lblDestination = "Choose destination"
      destination = ""
      lblDestination.ToolTipText = ""
   Else
      lblDestination = FileAbbrev(destination)
      lblDestination.ToolTipText = destination
   End If
On Error GoTo 0
'   If optMapps Then
'   Else
'      chkconverttomod.Visible = False
'   End If
   CheckForGo
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
SourcePathError:
   '  Nonexistent disk drive doesn't get picked up above
   lblSource = "Choose source"
   lblSource.ToolTipText = ""
   Exit Sub
DestPathError:
   '  Nonexistent disk drive doesn't get picked up above
   lblDestination = "Choose destination"
   destination = ""
   lblDestination.ToolTipText = ""
End Sub

Private Sub optFolder_Click()
'   If optFolder Then
      chkSubfolders.Visible = True
'   Else
'      chkSubfolders.Visible = False
'   End If
   CheckForGo
End Sub

Private Sub optPrevVersion_Click()
   chkConvertGenBank.Visible = True
   If lblMOD = "Species selected: Homo sapiens" Then
      chkConvertSwissProt.Visible = False
   Else
      chkConvertSwissProt.Visible = True
   End If
   optExpression.Visible = True
   lblSwitchFrom.Visible = False
   lstSwitchFrom.Visible = False
   cmdSelectAll.Visible = False
   cmdDeselectAll.Visible = False
   lblSwitchTo.Visible = False
   lstSwitchTo.Visible = False
   CheckForGo
End Sub
Private Sub optSwitch_Click()
   optMapps = True
   chkConvertGenBank.Visible = False
   chkConvertSwissProt.Visible = False
   optExpression.Visible = False
   lblSwitchFrom.Visible = True
   lstSwitchFrom.Visible = True
   cmdSelectAll.Visible = True
   cmdDeselectAll.Visible = True
   lblSwitchTo.Visible = True
   lstSwitchTo.Visible = True
   CheckForGo
End Sub

Private Sub optSingleFile_Click()
'   If optFolder Then
'      chkSubfolders.Visible = True
'   Else
      chkSubfolders.Visible = False
'   End If
   CheckForGo
End Sub

