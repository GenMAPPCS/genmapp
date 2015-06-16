VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmGeneDBMgr 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gene Database Manager"
   ClientHeight    =   8190
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8940
   Icon            =   "GeneDBMgr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtSpecies 
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
      Left            =   2580
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Type in species for Gene Database or click here and then """"Species"""" list to add species."
      Top             =   300
      Width           =   6312
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   4080
      TabIndex        =   34
      Top             =   5580
      Visible         =   0   'False
      Width           =   1152
   End
   Begin VB.CommandButton cmdAbandon 
      Caption         =   "Abandon"
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
      Left            =   2820
      TabIndex        =   33
      Top             =   5580
      Visible         =   0   'False
      Width           =   1152
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
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
      Left            =   4080
      TabIndex        =   30
      Top             =   5640
      Visible         =   0   'False
      Width           =   1152
   End
   Begin VB.ListBox lstCopyTables 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "GeneDBMgr.frx":08CA
      Left            =   180
      List            =   "GeneDBMgr.frx":08D1
      MultiSelect     =   2  'Extended
      TabIndex        =   29
      ToolTipText     =   "Click on tables to copy. Use Ctrl click to select multiple tables."
      Top             =   2820
      Visible         =   0   'False
      Width           =   5052
   End
   Begin VB.CommandButton cmdPreviousSegment 
      Caption         =   "Previous Segment"
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
      Left            =   4920
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7380
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton cmdNextSegment 
      Caption         =   "Next Segment"
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
      Left            =   6960
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7380
      Visible         =   0   'False
      Width           =   1932
   End
   Begin MSDBGrid.DBGrid dbgGeneDB 
      Bindings        =   "GeneDBMgr.frx":08E1
      Height          =   5952
      Left            =   120
      OleObjectBlob   =   "GeneDBMgr.frx":08F9
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Type species for Gene Database or click on Species list to add species."
      Top             =   1380
      Visible         =   0   'False
      Width           =   8772
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   2940
      Top             =   8520
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.ComboBox cmbSpecies 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      ItemData        =   "GeneDBMgr.frx":12CE
      Left            =   5580
      List            =   "GeneDBMgr.frx":12D0
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "cmbSpecies"
      ToolTipText     =   "Click on species here to add to """"Species"""" column or type directly in column."
      Top             =   960
      Visible         =   0   'False
      Width           =   3312
   End
   Begin VB.ComboBox cmbSystems 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      ItemData        =   "GeneDBMgr.frx":12D2
      Left            =   1260
      List            =   "GeneDBMgr.frx":12D4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click on a Gene Table to edit it."
      Top             =   960
      Width           =   3312
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   7800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7740
      Width           =   1092
   End
   Begin VB.Data dtaGeneDB 
      Caption         =   "Gene Database"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GenBank"
      Top             =   8520
      Width           =   2592
   End
   Begin VB.ComboBox cmbPrimary 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   180
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Choose or enter"
      ToolTipText     =   "Choose or enter a Gene Table."
      Top             =   3780
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.ComboBox cmbRelated 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   4500
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Choose or enter"
      ToolTipText     =   "Choose or enter a Gene Table."
      Top             =   3780
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.TextBox txtPrimaryCode 
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
      Left            =   3000
      MaxLength       =   3
      TabIndex        =   15
      ToolTipText     =   "To add a gene type its ID here and click 'Add Gene'."
      Top             =   3780
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtRelatedCode 
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
      Left            =   7320
      MaxLength       =   3
      TabIndex        =   17
      Top             =   3780
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
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
      Left            =   3300
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   1152
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Height          =   312
      Left            =   180
      TabIndex        =   8
      Top             =   1860
      Visible         =   0   'False
      Width           =   8712
      _ExtentX        =   15372
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtWebLink 
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
      Left            =   1020
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7380
      Visible         =   0   'False
      Width           =   3792
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gene Database species:"
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
      Left            =   300
      TabIndex        =   35
      Top             =   360
      Width           =   2196
   End
   Begin VB.Label lblModSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOD Sys"
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
      Left            =   2640
      TabIndex        =   32
      ToolTipText     =   "Click to designate Model Organism Database table."
      Top             =   660
      Width           =   924
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Organism table:"
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
      Index           =   0
      Left            =   300
      TabIndex        =   31
      Top             =   660
      Width           =   2004
   End
   Begin VB.Label lblCopyGeneDB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy table from: "
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
      Left            =   120
      TabIndex        =   28
      Top             =   2580
      Visible         =   0   'False
      Width           =   1500
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
      Left            =   180
      TabIndex        =   27
      Top             =   2220
      Visible         =   0   'False
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
      Left            =   120
      TabIndex        =   26
      Top             =   7860
      Visible         =   0   'False
      Width           =   528
   End
   Begin VB.Label lblWebLink 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web link"
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
      Left            =   132
      TabIndex        =   23
      ToolTipText     =   "Must be in form described in Help with one and only one ~ to be replaced by the gene ID. Eg: http://Genes.org/gene=~"
      Top             =   7440
      Visible         =   0   'False
      Width           =   756
   End
   Begin VB.Label lblRelated 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Related Gene ID System"
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
      Left            =   4500
      TabIndex        =   20
      Top             =   3540
      Visible         =   0   'False
      Width           =   2184
   End
   Begin VB.Label lblPrimary 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Primary Gene ID System"
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
      TabIndex        =   19
      Top             =   3540
      Visible         =   0   'False
      Width           =   2196
   End
   Begin VB.Label lblRelatedCode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   7320
      TabIndex        =   18
      ToolTipText     =   "Code for Primary Gene Table."
      Top             =   3540
      Visible         =   0   'False
      Width           =   468
   End
   Begin VB.Label lblPrimaryCode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   3000
      TabIndex        =   16
      ToolTipText     =   "Code for Primary Gene Table."
      Top             =   3540
      Visible         =   0   'False
      Width           =   468
   End
   Begin VB.Label Label1 
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
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   48
   End
   Begin VB.Label lblProgressTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Converting "
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
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1008
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
      Left            =   900
      TabIndex        =   10
      Top             =   2220
      Visible         =   0   'False
      Width           =   108
   End
   Begin VB.Label lblProgressMax 
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
      TabIndex        =   9
      Top             =   2220
      Width           =   48
   End
   Begin VB.Label lblSpecies 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Species"
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
      Left            =   4800
      TabIndex        =   7
      Top             =   1020
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblGeneTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gene Table"
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
      Left            =   168
      TabIndex        =   4
      Top             =   1020
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gene Database:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1452
   End
   Begin VB.Label lblGeneDB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gene DB"
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
      Left            =   1740
      TabIndex        =   0
      Top             =   60
      Width           =   828
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Gene Table"
      End
      Begin VB.Menu mnuEditRelations 
         Caption         =   "Edit Relationship Table"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add New Gene Table"
      End
      Begin VB.Menu mnuAddRelations 
         Caption         =   "Add New Relationship Table"
      End
      Begin VB.Menu mnuCreateGOCount 
         Caption         =   "Create GOCount Table"
      End
      Begin VB.Menu mnuCopyTable 
         Caption         =   "Copy Table(s) From . . ."
      End
      Begin VB.Menu mnuDeleteTable 
         Caption         =   "Delete Table"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGeneDBInfo 
         Caption         =   "Gene Database Information"
      End
      Begin VB.Menu mnuUpdateGeneDB 
         Caption         =   "Update Gene Database"
      End
      Begin VB.Menu mnuNewGeneDB 
         Caption         =   "Create New Gene Database"
      End
      Begin VB.Menu mnuModSys 
         Caption         =   "Specify Model Organism Table"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAssignSpecies 
         Caption         =   "Assign Species to Gene Database"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmGeneDBMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SEGMENT_SIZE = 250000                                                'For dgbGeneDB scrolling

Dim loading As Boolean
Dim prevRow As Variant, prevCol As Long, dataChanged As Boolean
Dim systems(MAX_SYSTEMS) As String, systemCodes(MAX_SYSTEMS) As String            'All Gene Tables,
                                                                                  'supported or not
Dim lastSystem As Integer                                                 'Index of last Gene Table
Dim rawDataFile As String
Dim copyDB As String, dbCopy As Database
Dim cancelExit As Boolean
Dim inProcess As String

Private Sub cmdCopy_Click()
   Dim i As Integer, dirty As Boolean
   Dim rsCopySystems As Recordset, rsSystems As Recordset, rsRelations As Recordset
   Dim tdf As TableDef
   
   lblDetail.visible = True
   For i = 0 To lstCopyTables.ListCount - 1 '+++++++++++++++++++++++++++ Go Through List Selections
      If lstCopyTables.selected(i) Then '============================================Selected Table
         '---------------------------------------------------See If Table Already Exists In Systems
         Set rsSystems = mappWindow.dbGene.OpenRecordset( _
               "SELECT System FROM Systems" & _
               "   WHERE System = '" & lstCopyTables.List(i) & "' AND [Date] IS NOT NULL")
         If Not rsSystems.EOF Then
            If MsgBox("Table """ & lstCopyTables.List(i) & """ already exists in """ _
                      & GetFile(mappWindow.dbGene.name) & """. Replace it?", _
                      vbExclamation + vbYesNo, "Copy Table") = vbYes Then
               mappWindow.dbGene.Execute "DROP TABLE [" & lstCopyTables.List(i) & "]"
               If lstCopyTables.List(i) = "GeneOntology" Then
                  mappWindow.dbGene.Execute "DROP TABLE [GeneOntologyCount]"
                  mappWindow.dbGene.Execute "DROP TABLE [GeneOntologyTree]"
               End If
            Else
               GoTo NextTable                              'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
         End If
         mappWindow.dbGene.Execute _
               "DELETE FROM Systems WHERE System = '" & lstCopyTables.List(i) & "'"
         '-------------------------------------------------See If Table Already Exists In Relations
         Set rsRelations = mappWindow.dbGene.OpenRecordset( _
               "SELECT Relation FROM Relations WHERE Relation = '" & lstCopyTables.List(i) & "'")
         If Not rsRelations.EOF Then
            If MsgBox("Table """ & lstCopyTables.List(i) & """ already exists in """ _
                      & GetFile(mappWindow.dbGene.name) & """. Replace it?", _
                      vbExclamation + vbYesNo, "Copy Table") = vbYes Then
               mappWindow.dbGene.Execute "DROP TABLE [" & lstCopyTables.List(i) & "]"
            Else
               GoTo NextTable                              'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
         End If
         dirty = True                                 'Anything beyond here makes change to Gene DB
         mappWindow.dbGene.Execute _
               "DELETE FROM Relations WHERE Relation = '" & lstCopyTables.List(i) & "'"
         '-------------------------------------------------------------------------------Copy Table
         lblDetail = "Copying " & lstCopyTables.List(i) & " table"
         DoEvents
         dbCopy.Execute _
               "SELECT [" & lstCopyTables.List(i) & "].* " & _
               "   INTO [" & lstCopyTables.List(i) & "] IN '" & mappWindow.dbGene.name & "'" & _
               "   FROM [" & lstCopyTables.List(i) & "]"
         If lstCopyTables.List(i) = "GeneOntology" Then '_______________GeneOntology Count And Tree
            lblDetail = "Copying GeneOntologyCount table"
            dbCopy.Execute _
                  "SELECT GeneOntologyCount.* " & _
                  "   INTO GeneOntologyCount IN '" & mappWindow.dbGene.name & "'" & _
                  "   FROM GeneOntologyCount"
            lblDetail = "Copying GeneOntologyTree table"
            dbCopy.Execute _
                  "SELECT GeneOntologyTree.* " & _
                  "   INTO GeneOntologyTree IN '" & mappWindow.dbGene.name & "'" & _
                  "   FROM GeneOntologyTree"
         End If
         If InStr(lstCopyTables.List(i), "-") <> 0 Then '________________________Relationship Table
            dbCopy.Execute _
                  "INSERT INTO Relations IN '" & mappWindow.dbGene.name & "'" & _
                  "   SELECT * FROM Relations WHERE Relation = '" & lstCopyTables.List(i) & "'"
         Else '_______________________________________________________________________Systems Table
            dbCopy.Execute _
                  "INSERT INTO Systems IN '" & mappWindow.dbGene.name & "'" & _
                  "   SELECT * FROM Systems WHERE System = '" & lstCopyTables.List(i) & "'"
         End If
         lblDetail = ""
      End If
NextTable:
   Next i
   If dirty Then
      mappWindow.dbGene.Execute "UPDATE Info SET Modify = '" & Format(Now, "yyyymmdd") & "'"
      ModifyOwner
   End If
   CopyTableVisible False
   CreateGOCountVisible
End Sub

Private Sub cmdDelete_Click() '**************************************************** Delete Table(s)
   Dim i As Integer, dirty As Boolean
   Dim rsInfo As Recordset, modTable As String
   
   Set rsInfo = mappWindow.dbGene.OpenRecordset("SELECT MODSystem FROM Info")
   modTable = Dat(rsInfo!MODSystem)
   
   For i = 0 To lstCopyTables.ListCount - 1 '+++++++++++++++++++++++++++ Go Through List Selections
      If lstCopyTables.selected(i) Then '============================================Selected Table
         If lstCopyTables.List(i) = modTable Then '------------------------------------Table Is MOD
            If MsgBox("The " & lstCopyTables.List(i) & " table is the MOD table for this Gene " _
                      & "Database. Do you wish to delete it? If you answer ""Yes"", your Gene " _
                      & "Database will be without a MOD designation, which means that it will " _
                      & "not be useful in MAPPFinder.", _
                      vbExclamation + vbYesNo, "Deleting MOD Table") = vbYes Then
               mappWindow.dbGene.Execute "UPDATE Info SET MODSystem = ''"
               lblModSys = ""
            Else
               GoTo NextTable                             'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
         End If
'      s = mappWindow.dbGene.name   'Close and open the DB because previous attempt at entering
'      mappWindow.dbGene.Close           'a table may have left it hung up and not DROPable
'      Set mappWindow.dbGene = OpenDatabase(s)
         mappWindow.dbGene.Execute "DROP TABLE [" & lstCopyTables.List(i) & "]"
         If lstCopyTables.List(i) = "GeneOntology" Then
            mappWindow.dbGene.Execute "DROP TABLE [GeneOntologyCount]"
            mappWindow.dbGene.Execute "DROP TABLE [GeneOntologyTree]"
         End If
         mappWindow.dbGene.Execute _
                    "DELETE FROM Systems WHERE System = '" & lstCopyTables.List(i) & "'"
         mappWindow.dbGene.Execute _
                    "DELETE FROM Relations WHERE Relation = '" & lstCopyTables.List(i) & "'"
         dirty = True
      End If
NextTable:
   Next i
   If dirty Then
      mappWindow.dbGene.Execute "UPDATE Info SET Modify = '" & Format(Now, "yyyymmdd") & "'"
      ModifyOwner
   End If
   DeleteTableVisible False
   CreateGOCountVisible
End Sub

Private Sub dbgGeneDB_AfterColUpdate(ByVal ColIndex As Integer)
   'Fired when user clicks off a column if the column value has been changed.
   'Also called from cmbSpecies_Click() because VB doesn't recognize anything but user input
   '  as a change
   'ColIndex is previous column
   
   With dbgGeneDB.columns(ColIndex)
      If .Caption = "Species" Then
         If Left(.text, 1) <> "|" Then .text = "|" & .text
         If Right(.text, 1) <> "|" Then .text = .text & "|"
         UpdateSystemsSpecies dbgGeneDB.columns(ColIndex).text
      End If
   End With
End Sub
Sub UpdateSystemsSpecies(speciesLine As String, Optional system As String = "")
   Dim rsSystems As Recordset, species As String
   Dim lastSpecies As Integer, i As Integer, j As Integer
   Dim speciei(100) As String, noOfSpeciei As Integer
   Dim speciesChanged As Boolean
   
   If Len(Dat(speciesLine)) <= 2 Then Exit Sub             'No species >>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   If system = "" Then system = cmbSystems.text
   
   Set rsSystems = mappWindow.dbGene.OpenRecordset( _
                   "SELECT Species FROM Systems WHERE System = '" & system & "'")
   noOfSpeciei = SeparateValues(speciei, speciesLine, "|")
   species = rsSystems!species
   For i = 0 To noOfSpeciei - 1
      If InStr(species, speciei(i)) = 0 Then
         species = species & speciei(i) & "|"
         speciesChanged = True
      End If
   Next i
   If speciesChanged Then
      If Left(species, 1) <> "|" Then species = "|" & species
      mappWindow.dbGene.Execute "UPDATE Systems SET Species = '" & species & "'" & _
                                "   WHERE System = '" & system & "'"
   End If
End Sub

'/////////////////////////////////////////////////////////////////////// Opening And Closing Window
Private Sub Form_Activate()
   '  Leaves window in the last-used condition.
   
   If Tag = "DontActivate" Then                  'Returning from some called window like frmSystems
      Tag = ""
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   ClearWindow
   
   SetGeneDB mappWindow.dbGene
   
'   If mappWindow.dbGene Is Nothing Then
'      MsgBox "Must choose a Gene Database first.", vbExclamation + vbOKOnly, "Edit Gene Database"
'      Hide
'      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   End If
'
'   lblGeneDB = FileAbbrev(mappWindow.dbGene.name)
'   FillSystemsList
'   mnuEdit_Click
'   dbgGeneDB.ScrollBars = dbgBoth
   
'   dbgGeneDB.Visible = False
'   txtNewGene.Visible = False
'   lblNewGene.Visible = False
'   txtNewGeneCode.Visible = False
'   lblNewGeneCode.Visible = False
'   lblGeneDB = mappWindow.dbGene.name
'   txtLocal = mruLocalSource
'   If Dat(txtLocal) = "" Then
'      MsgBox "State a default local source for for changed or added data.", _
'             vbExclamation + vbOKOnly, "Edit Gene Database"
'      txtLocal.SetFocus
'   End If
'   FillSystemsList
'   If Tag = "Add" Then
'      cmdAddGene_Click
'   Else
'      mnuEdit_Click
'   End If
'   MousePointer = vbDefault
End Sub
Sub SetGeneDB(dbGene As Database) '************************ Sets Menu, Etc For Existence Of Gene DB
   Dim have As Boolean                                                  'True is there is a Gene DB
   Dim rsInfo As Recordset
   Dim tdf As TableDef, goCountExists As Boolean, modGOExists As Boolean
   
   If dbGene Is Nothing Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ No Gene DB
      have = False
      lblGeneDB = "No Gene Database"
      lblModSys = ""
      mnuModSys.visible = False
      txtSpecies = ""
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Most Menus And Labels
      have = True
      lblGeneDB = FileAbbrev(mappWindow.dbGene.name)
      Set rsInfo = mappWindow.dbGene.OpenRecordset( _
                   "SELECT Owner, MODSystem, Species, Notes FROM Info")
      lblModSys = Dat(rsInfo!MODSystem)
      txtSpecies = Dat(rsInfo!species)
      If OfficialGeneDB Then
         mnuModSys.visible = False
         txtSpecies.Enabled = False
      Else
         mnuModSys.visible = True
         txtSpecies.Enabled = True
      End If
      dbgGeneDB.ScrollBars = dbgBoth
   End If
   mnuAdd.Enabled = have
   mnuAddRelations.Enabled = have
   mnuEdit.Enabled = have
   mnuEditRelations.Enabled = have
   mnuGeneDBInfo.Enabled = have
   mnuUpdateGeneDB.Enabled = have
   mnuCopyTable.Enabled = have
   mnuDeleteTable.Enabled = have
   CreateGOCountVisible
End Sub
Function OfficialGeneDB(Optional dbGene As Database = Nothing) As Boolean '*** Is Gene DB Official?
   '  Return   True if Gene DB "Official", limited in modifications
   
   'Official Gene DBs are identified by the owner "GenMAPP.org" or "GenMAPP.org (Modified)".
   'The Gene DB information window now displays the owner. Internally, an official DB has a
   'ANSI code 182 in the Notes field of the Info table. In later GenMAPP versions, this should
   'become the primary identification of an official DB.
   '
   'Any change in an official Gene DB other than in a Remarks field will cause the owner to be
   'changed to "GenMAPP.org (Modified)". The table will still be considered to be official.
   '
   'An official Gene Table is identified by not having a system code beginning in "&". An
   'official Relationship Table has both system codes not beginning with a "&". In the future,
   'official tables will be identified by either a single-character system code or one beginning
   'in "~" (tilde). All program code from now on will conform to this standard, which will also
   'accommodate the current "&" concept.
   '
   'Adding Gene or Relationship Tables: A table may be added to a Gene DB (official or not) if
   'it does not replace an existing official table. In other words, you can add a rat table to
   'a mouse DB. This is implemented by checking the proposed name. If it is the same as an
   'official table, whether the Gene DB is official or not, then it may not be added. Bear in
   'mind that any table added will have a system code that begins with "&" but if it has the
   'same name as an official table, it will not be allowed.
   '
   'Deleting Gene or Relationship Tables: A table may be deleted from an official Gene DB if it
   'is not an official table. This is implemented by limiting the list of tables available for
   'deleting. Anything can be deleted from a nonofficial database.
   '
   'Copying Gene or Relationship Tables: A table may be copied from any other Gene DB to an
   'official Gene DB if it does not replace an official table in the destination DB. This is
   'implemented by limiting the list of tables available for copying.
   '
   'Creating a GOCount Table: One may not be created for an official Gene DB that already has
   'one. This is implemented by the visibility of the menu item.
   '
   'Possible holes in the system:
   '
   'An official table may be modified if it is copied to a nonofficial Gene DB. However, it
   'still may not replace an official table in an official Gene DB. It seems reasonable that
   'a user could make a custom MGI table in a nonofficial Gene DB, for example, but not be able
   'to copy it back to the official "Mm-Std" Gene DB.
   '
   'An official table from another species cannot be copied to an official Gene DB that
   'already has the same-named table. For example, you cannot copy a rat SwissProt table
   'to a mouse Gene DB. I am not sure you would want to do anything like that anyway.
   
   'If you copy an official table to an official DB, it cannot be deleted. In other words,
   'if you copy the SGD table to the Mm-Std Gene DB, the SGI table may not be deleted. I
   'don't see this happening except by mistake but there is not much we can do about it anyway
   'because we do not identify the original source (Mm-Std, Sc-Std, etc) of tables.
   '
   'There are certain, hopefully rare, cases where the owner will show as
   '"GenMAPP.org (Modified)" when the Gene DB has been returned to original condition. One is
   'when a table is added (making it "Modified") and then later deleted.
   
   
   Dim rsInfo As Recordset
   
   If dbGene Is Nothing Then Set dbGene = mappWindow.dbGene              'Which may also be Nothing
   If dbGene Is Nothing Then Exit Function                 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   Set rsInfo = dbGene.OpenRecordset("SELECT Owner, Notes FROM Info")
   If Not rsInfo.EOF Then           'New Gene DBs will have an empty Info table and not be official
      If InStr(rsInfo!notes, Chr(182)) <> 0 _
            Or Dat(rsInfo!owner) = "GenMAPP.org" _
            Or Dat(rsInfo!owner) = "GenMAPP.org (Modified)" Then
         OfficialGeneDB = True
      End If
   End If
End Function
'******************************************************************************* Is Table Official?
Function OfficialTable(systemCode As String, Optional relatedCode As String = "") As Boolean
   '  Return   True if table "Official", not to be modified
   Dim official As Boolean
   
   If Len(systemCode) = 1 Then
      official = True
   ElseIf Left(systemCode, 1) = "~" Then
      official = True
   End If
   If relatedCode <> "" Then                                                 'A relationship table
      If Len(relatedCode) = 1 Or Left(relatedCode, 1) = "~" Then
                                                                 'Official or not remains as it is
      Else                                                                     'Cannot be official
         official = False
      End If
   End If
   OfficialTable = official
End Function
Sub ModifyOwner(Optional dbGene As Database = Nothing)  '******************* Change Official Status
   Dim rsInfo As Recordset
   
   If dbGene Is Nothing Then Set dbGene = mappWindow.dbGene              'Which may also be Nothing
   If dbGene Is Nothing Then Exit Sub                      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   Set rsInfo = dbGene.OpenRecordset("SELECT Owner, Notes FROM Info")
   If Dat(rsInfo!owner) = "GenMAPP.org" Then
      dbGene.Execute "UPDATE Info SET Owner = 'GenMAPP.org (Modified)'"
   End If
End Sub
Sub CreateGOCountVisible() '*************************** See If Possible To Create MOD-GOCount Table
   Dim tdf As TableDef, goCountExists As Boolean, modGOExists As Boolean
   Dim geneDB As String
   
   mnuCreateGOCount.Enabled = False
   If mappWindow.dbGene Is Nothing Then Exit Sub           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   If lblModSys <> "" And Not OfficialGeneDB Then
      mappWindow.dbGene.TableDefs.Refresh
      For Each tdf In mappWindow.dbGene.TableDefs
         If tdf.name = "GeneOntologyCount" Then
            goCountExists = True
         End If
         If tdf.name = lblModSys & "-GeneOntology" Then
            modGOExists = True
         End If
      Next tdf
      If goCountExists And modGOExists Then
         mnuCreateGOCount.Enabled = True
      End If
   End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormCode Then
'      Cancel = True
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   cancelExit = False
   If InvalidRow(dbgGeneDB.Bookmark) Then
      Cancel = True
   Else
      cmdExit_Click
   End If
   If cancelExit Then
      Cancel = True
   End If
End Sub
Private Sub cmdExit_Click()
   Dim index As Integer
   Dim rsInfo As Recordset
   
   cmdExit.SetFocus                                            'To force all other LostFocus events
On Error GoTo ErrorHandler
   If mappWindow.dbGene Is Nothing Then                                             'No Gene DB set
      GoTo EndSub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If dbgGeneDB.visible = True Then
      dbgGeneDB.row = 1                         'Force change in row to make data changes permanent
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For Running Processes
   Select Case inProcess
   Case "GOCount"
      mappWindow.dbGene.Execute "DROP TABLE [" & lblModSys & "-GOCount]"
   Case "AddSystem"
      inProcess = ""
      GoTo EndSub
   End Select
   inProcess = ""
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For MOD
   Set rsInfo = mappWindow.dbGene.OpenRecordset("SELECT MODSystem FROM Info")
   If Dat(rsInfo!MODSystem) = "" Then
      If MsgBox("This Gene Database has no Model Organism table designated. " _
                & "Without one, it will not work in MAPPFinder and will be of limited " _
                & "use in GenMAPP. Do you wish remain in the Gene Database Manager to " _
                & "designate one now?", _
                vbInformation + vbYesNo, "No Model Organism Table") = vbYes Then
         mnuModSys_Click
         cancelExit = True
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   rsInfo.Close
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For Species
   Set rsInfo = mappWindow.dbGene.OpenRecordset("SELECT Species FROM Info")
   If Dat(rsInfo!species) = "" Then
      If MsgBox("This Gene Database has no species designated. Do you wish remain in the " _
                & "Gene Database Manager to designate species now?", _
                vbInformation + vbYesNo, "No Species") = vbYes Then
         cancelExit = True
         txtSpecies.SetFocus
         txtSpecies_Click
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   rsInfo.Close
   
EndSub:
   Tag = ""
'   Hide
   Unload frmGeneDBMgr
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
ErrorHandler:
   If Err.number = 6148 Then
      GoTo EndSub                                          'No row 1, empty table ^^^^^^^^^^^^^^^^^
   Else
      FatalError "frmGeneDBMgr:cmdExit_Click", Err.Description
   End If
End Sub

Sub ClearWindow() '************************************************** Make Unneeded Stuff Invisible
   lblPrimary.visible = False
   cmbPrimary.visible = False
   lblPrimaryCode.visible = False
   txtPrimaryCode.visible = False
   lblRelated.visible = False
   cmbRelated.visible = False
   lblRelatedCode.visible = False
   txtRelatedCode.visible = False
   lblDetail.visible = False
   prgProgress.visible = False
   lblProgressTitle.visible = False
   lblProgressMax.visible = False
   lblPrgMax.visible = False
   lblErrors.visible = False
   cmdProcess.visible = False
   dbgGeneDB.visible = False
   If Not dtaGeneDB.Recordset Is Nothing Then
      dtaGeneDB.RecordSource = "Info"       'Switch to another, existing table to release any table
         'that might be in edit mode. Closing the recordset, for some reason, does not release the
         'connection and so any other action on the table is blocked by a lock.
      dtaGeneDB.Refresh
   End If
   cmdPreviousSegment.visible = False
   cmdNextSegment.visible = False
   lblSpecies.visible = False
   cmbSpecies.visible = False
   dbgGeneDB.visible = False
   lblGeneTable.visible = False
   cmbSystems.visible = False
   lblWebLink.visible = False
   txtWebLink.visible = False
   txtWebLink = ""
   CopyTableVisible False
   DeleteTableVisible False
   If OfficialGeneDB Then
      mnuAdd.Enabled = True
      mnuAddRelations.Enabled = True
      mnuAssignSpecies.Enabled = False
      mnuCopyTable.Enabled = True
      mnuCreateGOCount.Enabled = False
      mnuDeleteTable.Enabled = True
      mnuEditRelations.Enabled = True
      mnuModSys.Enabled = False
   Else
      mnuAdd.Enabled = True
      mnuAddRelations.Enabled = True
      mnuAssignSpecies.Enabled = True
      mnuCopyTable.Enabled = True
      mnuCreateGOCount.Enabled = True
      mnuDeleteTable.Enabled = True
      mnuEditRelations.Enabled = True
      mnuModSys.Enabled = True
   End If
   DoEvents
End Sub
Sub CopyTableVisible(Optional visible As Boolean = True)
'   ClearWindow
   lblCopyGeneDB.visible = visible
   If visible Then
      lblCopyGeneDB = "Copy table from:"
   End If
   lstCopyTables.visible = visible
   cmdCopy.visible = visible
   DoEvents
End Sub
Sub DeleteTableVisible(Optional visible As Boolean = True)
'   ClearWindow
   lblCopyGeneDB.visible = visible
   If visible Then
      lblCopyGeneDB = "Delete table:"
   End If
   lstCopyTables.visible = visible
   cmdDelete.visible = visible
   DoEvents
End Sub

Private Sub lblModSys_Click()
   mnuModSys_Click
End Sub

Private Sub mnuCopyTable_Click() '*********************************** Copy Tables To Active Gene DB
   Dim rsSourceTables As Recordset, rsDestTables As Recordset
   Dim tdf As TableDef, officialDB As Boolean
   
CopyFrom:
   With dlgDialog
      .DialogTitle = "Gene Database To Copy From"
On Error GoTo OpenError
      .CancelError = True
      .Filter = "Gene Databases (.gdb)|gdb"
      If mruGeneDB <> "" Then
         .InitDir = GetFolder(mruGeneDB) 'Left(mruGeneDB, InStrRev(mruGeneDB, "\") - 1)
      End If
      .FileName = "*.gdb"
      .FLAGS = cdlOFNExplorer
      .ShowOpen
      copyDB = .FileName
   End With
On Error GoTo 0
   If InStr(copyDB, ".") = 0 Then
      copyDB = copyDB & ".gdb"
   End If
   
   If UCase(copyDB) = UCase(mappWindow.dbGene.name) Then
      MsgBox "Cannot copy from Gene Database currently open.", vbCritical + vbOKOnly, "Copy Table"
      GoTo CopyFrom
   End If
   
   ClearWindow
   officialDB = OfficialGeneDB                                          'Destination DB is official
   CopyTableVisible
   
   lblCopyGeneDB = "Copy table from: " & FileAbbrev(copyDB)
   lblCopyGeneDB.ToolTipText = copyDB
   
   Set dbCopy = OpenDatabase(copyDB)
   
   lstCopyTables.Clear '++++++++++++++++++++++++++++++++++ Create List Of Tables That Can Be Copied
      '  A table can be copied if the destination Gene DB is not official or if it is not
      '  an official (non-&) table that exists in the destination.
   
   '===========================================================================Copyable Gene Tables
   Set rsSourceTables = dbCopy.OpenRecordset( _
         "SELECT System, SystemCode FROM Systems WHERE [Date] IS NOT NULL ORDER BY System")
   Do Until rsSourceTables.EOF
      If officialDB And OfficialTable(rsSourceTables!systemCode) Then
         Set rsDestTables = mappWindow.dbGene.OpenRecordset( _
               "SELECT System, SystemCode FROM Systems" & _
               "   WHERE [Date] IS NOT NULL AND SystemCode = '" & rsSourceTables!systemCode & "'")
         If rsDestTables.EOF Then                            'Does not exist in destination Gene DB
            lstCopyTables.AddItem rsSourceTables!system
         End If
      Else                                                        'Not an official table or Gene DB
         lstCopyTables.AddItem rsSourceTables!system
      End If
      rsSourceTables.MoveNext
   Loop
   
   '===================================================================Copyable Relationship Tables
   Set rsSourceTables = dbCopy.OpenRecordset( _
         "SELECT Relation, SystemCode, RelatedCode FROM Relations ORDER BY Relation")
   Do Until rsSourceTables.EOF
      If officialDB _
            And OfficialTable(rsSourceTables!systemCode, rsSourceTables!relatedCode) Then
         Set rsDestTables = mappWindow.dbGene.OpenRecordset( _
               "SELECT Relation FROM Relations" & _
               "   WHERE Relation = '" & rsSourceTables!Relation & "'")
         If rsDestTables.EOF Then                            'Does not exist in destination Gene DB
            lstCopyTables.AddItem rsSourceTables!Relation
         End If
      Else                                                        'Not an official table or Gene DB
         lstCopyTables.AddItem rsSourceTables!Relation
      End If
      rsSourceTables.MoveNext
   Loop
   
   If lstCopyTables.ListCount = 0 Then
      MsgBox "There are no tables in the source Gene Database that would not replace an " _
             & "official GenMAPP table in the official GenMAPP destination Gene Database.", _
             vbExclamation + vbOKOnly, "Copying Tables"
      CopyTableVisible False
   End If

ExitSub:
   MousePointer = vbDefault
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Error Handlers
OpenError:
   If Err.number = 32755 Then                                                               'Cancel
      
   Else                                                                          'Other than Cancel
      FatalError "frmGeneDBMgr:mnuCopyTable", Err.Description
   End If
   On Error GoTo 0
   Resume ExitSub                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   
End Sub

Private Sub mnuDeleteTable_Click()
   Dim rsTables As Recordset, officialDB As Boolean
   
   ClearWindow
   DeleteTableVisible
   officialDB = OfficialGeneDB
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++ Create List Of Tables That Can Be Deleted
   lstCopyTables.Clear
      
   '====================================================================================Gene Tables
   Set rsTables = mappWindow.dbGene.OpenRecordset( _
         "SELECT System, SystemCode FROM Systems WHERE [Date] IS NOT NULL ORDER BY System")
   Do Until rsTables.EOF
      If Not (officialDB And OfficialTable(rsTables!systemCode)) Then
         lstCopyTables.AddItem rsTables!system
      End If
      rsTables.MoveNext
   Loop
   
   '============================================================================Relationship Tables
   Set rsTables = mappWindow.dbGene.OpenRecordset( _
         "SELECT Relation, SystemCode, RelatedCode FROM Relations ORDER BY Relation")
   Do Until rsTables.EOF
      If Not (officialDB And OfficialTable(rsTables!systemCode, rsTables!relatedCode)) Then
         lstCopyTables.AddItem rsTables!Relation
      End If
      rsTables.MoveNext
   Loop
   
   If lstCopyTables.ListCount = 0 Then
      MsgBox "There are no tables in the official GenMAPP destination Gene Database that " _
             & "can be deleted.", vbExclamation + vbOKOnly, "Copying Tables"
      DeleteTableVisible False
   End If

   MousePointer = vbDefault
End Sub

Private Sub mnuHelp_Click()
   Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, appPath & "GenMAPP.chm::/GeneDatabaseManager.htm", _
                       HH_DISPLAY_TOPIC, 0)
End Sub

'////////////////////////////////////////////////////////////////////////////// Editing Gene Tables
Private Sub mnuEdit_Click() '********************************************** Edit Current Gene Table
   '  This sets up the window for editing with a Gene Table list (cmbSystems) showing. Clicking a
   '  Gene Table in the list loads that table and starts the editing process.
   Dim rsSystems As Recordset
   
'   If cmbSystems.Text = "Choose Gene Table to edit" Then
'      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   End If
'
   MousePointer = vbHourglass
   DoEvents
   loading = True
   
   ClearWindow                                                       'Make Unneeded Stuff Invisible
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Make Needed Stuff Visible
   FillSystemsList
   lblGeneTable.visible = True
   cmbSystems.RemoveItem 0
   cmbSystems.AddItem "Choose Gene Table to edit", 0
   cmbSystems.ListIndex = 0
   cmbSystems.visible = True
   MousePointer = vbDefault
   DoEvents
   loading = False
End Sub
Private Sub cmbSystems_Click()
   Dim rsSystems As Recordset, species As String, pipe As Integer, nextPipe As Integer
   Dim rsInfo As Recordset, modSys As String
   Dim rsMod As Recordset
   
   If cmbSystems.text = cmbSystems.List(0) Then                               'First item is Choose
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If lblModSys = "Choose table from list" Then '+++++++++++++++++++++++++++++++++ Choose MOD Table
      modSys = cmbSystems.text
      mappWindow.dbGene.Execute "UPDATE Info SET MODSystem = '" & modSys & "'"
      lblModSys = modSys
      For i = 0 To mappWindow.dbGene.TableDefs(modSys).Fields.count - 1
         If mappWindow.dbGene.TableDefs(modSys).Fields(i).name = "Species" Then Exit For
      Next i
      If i <= mappWindow.dbGene.TableDefs(modSys).Fields.count - 1 Then              'Species Found
         Set rsMod = mappWindow.dbGene.OpenRecordset("SELECT Species FROM [" & modSys & "]")
         If Not rsMod.EOF Then                                     'A first row exists in MOD table
            If Dat(rsMod!species) <> txtSpecies _
                  And Dat(rsMod!species) <> "" Then                  'Different from listed species
               If MsgBox("Species found in the first row of your Model Organism table does " _
                         & "not agree with the Gene Database species. Do you wish to change " _
                         & "your Gene Database species to " _
                         & vbCrLf & vbCrLf & rsMod!species & " ?", _
                         vbExclamation + vbYesNo, "MOD Table Change") = vbYes Then
                  txtSpecies = rsMod!species
                  txtSpecies_LostFocus                                              'To write to DB
               End If
            End If
         End If
      End If
      ClearWindow                                                    'Make Unneeded Stuff Invisible
      CreateGOCountVisible
   Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Editing Tables
      ClearWindow                                                    'Make Unneeded Stuff Invisible
      lblGeneTable.visible = True
      cmbSystems.visible = True
      
      MousePointer = vbHourglass
      DoEvents
      
      '============================================================================Set up Data Grid
      dtaGeneDB.DatabaseName = mappWindow.dbGene.name
      dtaGeneDB.RecordSource = "SELECT * FROM " & cmbSystems.text & " ORDER BY ID"
      dtaGeneDB.Refresh
   '   dbgGeneDB.Visible = True
   '   DoEvents
      If Not dtaGeneDB.Recordset.EOF Then
         dtaGeneDB.Recordset.MoveLast                    'Forces VB to put all rows in data control
         dtaGeneDB.Recordset.MoveFirst
      End If
      If dtaGeneDB.Recordset.recordCount > SEGMENT_SIZE Then
         cmdPreviousSegment.visible = True
         cmdPreviousSegment.Enabled = False
         cmdNextSegment.visible = True
         cmdNextSegment.Enabled = True
      Else
         cmdPreviousSegment.visible = False
         cmdNextSegment.visible = False
      End If
      
      '========================================================================Allow Editing Or Not
      Set rsSystems = mappWindow.dbGene.OpenRecordset( _
            "SELECT * FROM Systems WHERE System = '" & cmbSystems.text & "'")
      Select Case rsSystems!systemCode                  'Add only to Other and user-supplied tables
      Case "O"
         dbgGeneDB.AllowAddNew = True
      Case "&" To "&z"
         dbgGeneDB.AllowAddNew = True
         lblWebLink.visible = True
         txtWebLink.visible = True
         If Dat(rsSystems!link) = "" Then
            txtWebLink = "http://"
         Else
            txtWebLink = rsSystems!link
         End If
      Case Else
         dbgGeneDB.AllowAddNew = False
      End Select
      For i = 0 To dbgGeneDB.columns.count - 1
         dbgGeneDB.columns(i).WrapText = True
         If cmbSystems.text = "Other" Or Left(rsSystems!systemCode, 1) = "&" Then
               '  Lock all columns in GenMAPP-supplied tables. Don't lock any columns
               '  in "Other" or user-supplied tables.
            dbgGeneDB.columns(i).Locked = False
         Else
            dbgGeneDB.columns(i).Locked = True
         End If
      Next i
      dbgGeneDB.columns("Remarks").Locked = False                        'Never lock Remarks column
      
      dbgGeneDB.DefColWidth = 0                                      'Let DBGrid autoset the widths
      dbgGeneDB.visible = True
      If dbgGeneDB.AllowAddNew Then                                  'This Gene Table may be edited
         lblWebLink.visible = True
         txtWebLink.visible = True
      End If
      CreateGOCountVisible
      MousePointer = vbDefault
   End If
End Sub
Sub PopulateSpeciesList() '***************************** Gather Species From All Systems Table Rows
   Dim rsSystems As Recordset
   Dim lastSpecies As Integer, i As Integer, j As Integer
   Dim speciei(100) As String, noOfSpeciei As Integer
   
   cmbSpecies.Clear
   
   If mappWindow.dbGene Is Nothing Then
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   Set rsSystems = mappWindow.dbGene.OpenRecordset( _
                   "SELECT Species FROM Systems WHERE [Date] IS NOT NULL OR System = 'Other'")
   Do Until rsSystems.EOF
      If Dat(rsSystems!species) <> "" Then
         noOfSpeciei = SeparateValues(speciei, Dat(rsSystems!species), "|")
         For i = 0 To noOfSpeciei - 1
            For j = 0 To cmbSpecies.ListCount - 1
               If speciei(i) = cmbSpecies.List(j) Then
                  Exit For                                 'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
               End If
            Next j
            If j > cmbSpecies.ListCount - 1 Then
               If Len(Dat(speciei(i))) > 2 Then    'To prevent empty species from appearing in list
                  cmbSpecies.AddItem speciei(i)
               End If
            End If
         Next i
      End If
      rsSystems.MoveNext
   Loop
   If cmbSpecies.ListCount >= 1 Then
      cmbSpecies.text = "Choose species"
   End If
End Sub

Private Sub mnuModSys_Click()
   Dim i As Integer
   
   ClearWindow
   FillSystemsList
   lblGeneTable.visible = True
   For i = 1 To cmbSystems.ListCount - 1
      If cmbSystems.List(i) = "Other" Then
         cmbSystems.RemoveItem i
         Exit For
      End If
   Next i
   cmbSystems.RemoveItem 0
   cmbSystems.AddItem "Choose gene table for MOD", 0
   cmbSystems.ListIndex = 0
   cmbSystems.visible = True
   lblModSys = "Choose table from list"
   cmbSystems.SetFocus
End Sub

Private Sub mnuNewGeneDB_Click()
   Dim geneDB As String, systemCode As String, species As String, dbDate As String, owner As String
   Dim db As Database, rs As Recordset
   
GeneDBName:
   With dlgDialog
      .DialogTitle = "New Gene Database Name"
On Error GoTo OpenError
      .CancelError = True
      .Filter = "Gene Databases (.gdb)|gdb"
      If mruGeneDB <> "" Then
         .InitDir = Left(mruGeneDB, InStrRev(mruGeneDB, "\") - 1)
      End If
      .FileName = "*.gdb"
      .FLAGS = cdlOFNHideReadOnly + cdlOFNExplorer
      .ShowOpen
      geneDB = .FileName
   End With
On Error GoTo 0
   If InStr(geneDB, ".") = 0 Then
      geneDB = geneDB & ".gdb"
   End If
   
'   MousePointer = vbHourglass
   If Dir(geneDB) <> "" Then '++++++++++++++++++++++++++++++++++++++++++ Clear Path For New Gene DB
      Select Case MsgBox("Gene Database already exists. Replace it?", vbExclamation + vbYesNo, _
                         "New Gene Database")
      Case vbNo
         GoTo GeneDBName
      Case vbYes
         If geneDB = mappWindow.dbGene.name Then                      'Same name as current Gene DB
            mappWindow.dbGene.Close
            Set mappWindow.dbGene = Nothing
         End If
On Error GoTo CantKillDB
         Kill geneDB
On Error GoTo 0
      End Select
   End If
         
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up New Gene DB
   owner = InputBox("Organization/person responsible for new Gene Database. You must have " _
                    & "an entry here. Leaving it blank cancels your new Gene Database.", _
                    "New Gene Database Owner", "Example: Gladstone Institutes")
   If Dat(owner) = "" Then
      MsgBox "You have cancelled creation of a new Gene Database.", vbExclamation + vbOKOnly, _
             "New Gene Database Owner"
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   owner = Left(owner, 200)
   dbDate = InputBox("Effective date for new Gene Database. You must have " _
                     & "a valid entry here. Leaving it blank cancels your new Gene Database.", _
                     "New Gene Database Date", Format(Now, "dd-mmm-yyyy"))
   If dbDate = "" Then
      MsgBox "You have cancelled creation of a new Gene Database.", vbExclamation + vbOKOnly, _
             "New Gene Database Date"
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   Do Until IsDate(dbDate)
      dbDate = InputBox("Date not valid. Reenter.", "New Gene Database Date", dbDate)
         If dbDate = "" Then Exit Sub                      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   Loop
'      '  Species entered when entering the MOD for the Gene DB
   FileCopy appPath & "GeneDBTmpl.gtp", geneDB
   Set db = OpenDatabase(geneDB)
   db.Execute "UPDATE Info SET Owner = '" & owner & "'," & _
              "                Version = '" & Format(CDate(dbDate), "yyyymmdd") & "'," & _
              "                Species = ''," & _
              "                Modify = '" & Format(CDate(dbDate), "yyyymmdd") & "'"
   db.Execute "UPDATE Systems SET [Date] = '" & CDate(dbDate) & "'" & _
              "   WHERE System = 'Other'"
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Open New Gene DB
   OpenGeneDB mappWindow.dbGene, geneDB, mappWindow
   SetGeneDB mappWindow.dbGene
'   lblGeneDB = geneDB
'   Form_Activate
   MsgBox "You now have the blank Gene Database:" & vbCrLf & vbCrLf & geneDB & vbCrLf & vbCrLf _
          & "and it is now the active Gene Database. To make it useful you must add at least " _
          & "one Gene Table to it.", _
          vbOKOnly, "New Gene Database"
   mnuAdd_Click
   
ExitSub:
   MousePointer = vbDefault
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Error Handlers
OpenError:
   If Err.number = 32755 Then                                                               'Cancel
      
   Else                                                                          'Other than Cancel
      FatalError "frmGeneDBMgr:mnuNewGeneDB", Err.Description
   End If
   On Error GoTo 0
   Resume ExitSub                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   
CantKillDB:
   MsgBox "You cannot replace the Gene Database because it is either in use in GenMAPP " & _
          "or some other program, or has been set to read-only through Windows. Correct " & _
          "the problem and then return here.", vbExclamation + vbOKOnly, "New Gene Database"
   Resume ExitSub                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
End Sub

Private Sub txtSpecies_Click()
   cmbSpecies.Tag = "Set Gene DB Species"
   ShowSpecies True
End Sub

Private Sub txtSpecies_LostFocus()
   If ActiveControl.name = "cmbSpecies" Then Exit Sub      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   txtSpecies = Dat(txtSpecies)
   If Left(txtSpecies, 1) <> "|" Then txtSpecies = "|" & txtSpecies
   If Right(txtSpecies, 1) <> "|" Then txtSpecies = txtSpecies & "|"
   If Len(txtSpecies) <= 2 Then                                               'Whole species string
      txtSpecies = ""
   ElseIf Len(Trim(Mid(txtSpecies, 2, Len(txtSpecies) - 2))) <= 2 Then      'Characters between | |
      txtSpecies = ""
   End If
   mappWindow.dbGene.Execute "UPDATE Info SET Species = '" & txtSpecies & "'"
   cmbSpecies.Tag = ""
End Sub

Private Sub txtWebLink_Change()
   '  A little inefficient. Might be better to use a form variable and set it once. Must react
   '  to change, though, or check every possible place where editing of a Gene Table might be
   '  finished and react there.
   Dim rsSystems As Recordset
   
   If InStr(txtWebLink, "~") <> 0 Then                               'Only change on valid web link
      '  If entered web link never valid, it remains what it was before
      Set rsSystems = mappWindow.dbGene.OpenRecordset( _
         "SELECT * FROM Systems WHERE System = '" & cmbSystems.text & "'")
      rsSystems.edit
      rsSystems!link = txtWebLink
      rsSystems.Update
   End If
End Sub

'/////////////////////////////////////////////////////////////////////////// Process New Gene Table
Private Sub mnuAdd_Click() '**************************************************** Add New Gene Table
   Dim inLine As String
   Dim columns As Integer                                     'Number of data columns after Gene ID
   Dim errorFile As String
   Dim remarksIndex As Integer, remarks As String
   Dim systemTitle As String, newSystemCode As String
   Dim deleteOldSystemCode As String                                    'Replacing a current system
   Dim geneValue As Integer
   Dim tdfGene As TableDef, idxGene As index
   Dim speciesIndex As Integer
   Dim species As String                                                  'Individual for data line
   Dim speciei As String                                           'All species for this Gene Table
   Dim geneId As String
   Dim rsInfo As Recordset, rsSystems As Recordset, rs As Recordset
   Dim geneDB As String, sql As String
   Dim errorsExists As Integer, errors As String
   Dim prevPipe As Integer, pipe As Integer
   Dim i As Integer, s As String
   
   Do While mappWindow.dbGene Is Nothing '+++++++++++++++++++++++++++++++++++++++ Check For Gene DB
      OpenGeneDB mappWindow.dbGene, "**OPEN**"
      If mappWindow.dbGene Is Nothing Then
         If MsgBox("Must have a Gene Database open.", _
                   vbOKCancel + vbExclamation, "New Gene Table") = vbCancel Then
            GoTo ExitSub                                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      End If
   Loop
   
   ClearWindow                                                       'Make Unneeded Stuff Invisible
'   FillSystemsList
   
Retry:
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Raw Data File
   '  End result of this section is a valid rawDataFile
On Error GoTo OpenError
   With dlgDialog
      .CancelError = True
      .FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
      .DialogTitle = "Gene File to Import"
      .InitDir = mruImportPath
      .FileName = ""
      .Filter = "All files|*.csv;*.tab;*.txt|Comma-separated values (.csv)|*.csv|" _
                       & "Tab-delimited lists (.tab, .txt)|*.tab;*.txt"
      .FilterIndex = 1
      .ShowOpen
On Error GoTo 0
      rawDataFile = .FileName
   End With
   
   MousePointer = vbArrowHourglass
   DoEvents
   
On Error GoTo RawDataFileError
      '  Try to rename the file. If it is in use, read-only, nonexistent, etc., it will error.
   Name rawDataFile As appPath & "rawTemp.$tm"
   Name appPath & "rawTemp.$tm" As rawDataFile
On Error GoTo 0

   mruImportPath = GetFolder(rawDataFile)
   dtaGeneDB.RecordSource = "Info"   'Switch to another, existing table to release the new table if
                                     'it had been used in a previous operation. This occurs when
                                     'entering the same raw data file twice in sucession.
   dtaGeneDB.Refresh
   dbgGeneDB.Refresh
   FillSystemsList
   prgProgress.Max = FileLen(rawDataFile)
   prgProgress.value = 0
   DoEvents
   
EnterSystemTitle: '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Assign System Title
   s = GetFile(rawDataFile)
   s = Left(s, InStrRev(s, ".") - 1)
   s = ValidTableTitle(s)
   systemTitle = InputBox("Enter the title for your Gene Table (max " & _
                          SYSTEM_TITLE_CHAR_LIMIT & " characters). ", "Gene Table Title", s)
CheckSystemTitle:
   If systemTitle = "" Then GoTo ExitSub                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   If Len(systemTitle) > SYSTEM_TITLE_CHAR_LIMIT Then
      systemTitle = InputBox("More than " & SYSTEM_TITLE_CHAR_LIMIT & " characters for Gene " _
                             & "Table. Please shorten.", _
                             "Gene Table Title", systemTitle)
      GoTo CheckSystemTitle                                '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   End If
   If systemTitle <> ValidTableTitle(systemTitle) Then
      systemTitle = InputBox("Gene Table title """ & systemTitle & """ invalid. Suggest " _
                             & "change to:", "Gene Table Title", ValidTableTitle(systemTitle))
      GoTo CheckSystemTitle                                '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   End If
   newSystemCode = "&"
   For i = 0 To lastSystem '=======================================See If Gene Table Already Exists
      If UCase(systemTitle) = UCase(systems(i)) Then '-------------------------------Duplicate Name
         If Not OfficialTable(systemCodes(i)) Then         'User table, can be deleted and replaced
            Select Case MsgBox("Gene Table " & systemCodes(i) & ": " & systems(i) _
                               & " already exists. Delete and replace it?", _
                               vbQuestion + vbYesNoCancel, "Gene Table Name")
            Case vbCancel
               GoTo ExitSub                                'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            Case vbNo
               systemTitle = ""
               GoTo EnterSystemTitle                       '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            Case vbYes
               '  At some point, may want to consider working with a temporary table ???????
               deleteOldSystemCode = systemCodes(i)
               newSystemCode = systemCodes(i)
               Exit For                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End Select
         Else                                     'Not a user table, cannot be deleted and replaced
            MsgBox "Gene Table " & systemCodes(i) & ": " & systems(i) & " already exists. It " _
                   & "is an official GenMAPP.org table and may not be replaced. " _
                   & "Choose another name.", vbExclamation + vbOKOnly, "Gene Table Name"
            systemTitle = ""
            GoTo EnterSystemTitle                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
         End If
      End If
   Next i
   
EnterNewSystemCode: '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Assign System Code
   newSystemCode = Dat(InputBox("Enter the code for your Gene Table. It must consist of two " _
                         & "characters beginning with &.", "Gene Table Code", newSystemCode))
   If newSystemCode = "" Then GoTo ExitSub                 'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   For i = 0 To lastSystem '============================================== See If It Is A Duplicate
      If newSystemCode = systemCodes(i) Then                                            'Dupe found
         If deleteOldSystemCode <> systemCodes(i) Then                   'Not system to be replaced
            MsgBox "Code " & systemCodes(i) & ": " & systems(i) & " already exists. " _
                   & "Choose another.", vbExclamation + vbOKOnly, "Gene Table Code"
            newSystemCode = "&"
            GoTo EnterNewSystemCode                        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
         Else                                                                'System to be replaced
            Exit For         'Dupe found, search no further vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      End If
   Next i
   Mid(newSystemCode, 1, 1) = UCase(Mid(newSystemCode, 1, 1))
   If Left(newSystemCode, 1) <> "&" Or Len(newSystemCode) <> 2 Then
      MsgBox "Gene Table code must have two characters beginning with ""&"".", _
             vbExclamation + vbOKOnly, "Gene Table Code"
      newSystemCode = "&"
      GoTo EnterNewSystemCode                              '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   End If
   
   Open rawDataFile For Binary As #FILE_RAW_DATA 'Cancelling after this has to Close #FILE_RAW_DATA
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Count Data Columns
   Dim delimiter As String * 1
   
   inLine = RemoveQuotes(InputUnixLine(FILE_RAW_DATA))
   If InStr(inLine, vbTab) Then                                                        'Found a tab
      delimiter = vbTab                                                     'Make tab the delimiter
   Else
      delimiter = ","                                                   'Otherwise default to comma
   End If
   speciesIndex = -1
   remarksIndex = -1
   columns = 0           'There is always a Gene ID column. Subsequent columns begin with delimiter
      '  This variable is the number of data columns after the Gene ID.
   i = InStr(inLine, delimiter)
   Do While i
      '  This counts delimiters, ignoring the first field, the Gene ID.
      columns = columns + 1
      If UCase(Mid(inLine, i + 1, 7)) = "REMARKS" Then
         If Mid(inLine, i + 8, 1) = delimiter Or Len(inLine) = i + 7 Then
            remarksIndex = columns
         End If
      End If
      If UCase(Mid(inLine, i + 1, 7)) = "SPECIES" Then
         If Mid(inLine, i + 8, 1) = delimiter Or Len(inLine) = i + 7 Then
            speciesIndex = columns
         End If
      End If
      i = InStr(i + 1, inLine, delimiter)
   Loop
   If Right(inLine, 8) = "~Errors~" Then
      errorsExists = 1
      columns = columns - 1
   End If
   ReDim geneValues(columns) As String                            'Row of values from raw gene file
      '  This is a one-based array; the zero element is not used. If the upper bound is zero,
      '  it indicates that the raw data has only Gene IDs and no data columns.
   Seek #FILE_RAW_DATA, 1
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Titles
   frmDataID.lstTitles.Clear
   frmDataID.lblInstructions = "Check the box if the column should be searchable (max " _
                             & CHAR_DATA_LIMIT & " chrs)."
   frmDataID.lblInstructions.ToolTipText = _
             "Searchable columns may contain no more than " & CHAR_DATA_LIMIT & " characters."
   frmDataID.lstTitles.ToolTipText = _
             "Searchable columns may contain no more than " & CHAR_DATA_LIMIT & " characters."
   s = GetGeneRow(errorsExists, geneId, geneValues, inLine, delimiter)                  'Title line
   For geneValue = 1 To columns '========================================================Each Title
      geneValues(geneValue) = Trim(geneValues(geneValue))
      If geneValues(geneValue) = "" Then '----------------------------------------------Blank Title
         s = ""                      'Print out all the titles and ask for a name for the blank one
         For i = 1 To columns                                                  'Assemble all titles
            If i >= 2 Then s = s & "|"
            If i = geneValue Then
               s = s & "________"
            Else
               s = s & geneValues(i)
            End If
         Next i
         Do
            geneValues(geneValue) = InputBox("The column heading (_________) in" _
                     & vbCrLf & vbCrLf & s & vbCrLf & vbCrLf _
                     & "is blank. Enter a " _
                     & "valid column heading or abort the conversion by clicking Cancel.", _
                     "Invalid Column Heading", "________")
         Loop While geneValues(geneValue) = "________"
      End If
RecheckTitle: '---------------------------------------------------------------Validate Column Title
      geneValues(geneValue) = Trim(TextToSql(geneValues(geneValue)))
      If geneValues(geneValue) = "" Then
         '  The only time a blank title should reach here is if Cancel were entered above or
         '  below and program came back to recheck the title
         Close #FILE_RAW_DATA
         GoTo ExitSub                                      'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
      If InvalidChr(geneValues(geneValue), "column heading") Then
         '  Won't return until all invalid chrs cleaned up or title returns ""
         GoTo RecheckTitle                                 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      End If
      Select Case UCase(geneValues(geneValue)) '_____________________________________Reserved Title
      Case "ID", "SYSTEMCODE", "DATE"
         geneValues(geneValue) = _
               InputBox("Column heading """ & geneValues(geneValue) _
                        & """ is reserved for the system. Please change it.", _
                       "Column Heading Reserved", geneValues(geneValue))
         GoTo RecheckTitle                                 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      End Select
      If Len(geneValues(geneValue)) > TITLE_CHAR_LIMIT Then '______________________________Too Long
         geneValues(geneValue) = _
               InputBox("Column heading """ & geneValues(geneValue) & """ exceeds the " _
                        & TITLE_CHAR_LIMIT & "-character limit. Please shorten it.", _
                       "Column Heading Too Long", Left(geneValues(geneValue), TITLE_CHAR_LIMIT))
         GoTo RecheckTitle                                 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      End If
      For i = 1 To geneValue - 1 '_______________________________________Check For Duplicate Titles
         If geneValues(geneValue) = geneValues(i) Then
            geneValues(geneValue) = _
                  InputBox("You have two columns headed '" & geneValues(geneValue) _
                           & "'. You may change the second one here and click OK, " _
                           & "or abort the conversion by clicking Cancel.", _
                           "Duplicate Column Headings", geneValues(geneValue))
            GoTo RecheckTitle                              '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
         End If
      Next i
      If UCase(geneValues(geneValue)) <> "REMARKS" _
            And geneValues(geneValue) <> "~Errors~" Then
         '  Don't add these columns to the list. Remarks always exists and is always memo.
         frmDataID.lstTitles.AddItem geneValues(geneValue)
         frmDataID.lstTitles.selected(frmDataID.lstTitles.ListCount - 1) = False 'True
      End If
   Next geneValue
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Column Datatypes
   frmDataID.show vbModal                        'All fields default to nonsearchable, memo fields
   If frmDataID.Tag = "Cancel" Then
      Close #FILE_RAW_DATA
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
'   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set MOD System And Species
'   Set rsInfo = mappWindow.dbGene.OpenRecordset("SELECT MODSystem, Species FROM Info")
'   If Dat(rsInfo!MODSystem) = "" Then
'      If MsgBox("Is " & systemTitle & " the Model Organism table for this Gene Database?", _
'                vbInformation + vbYesNo, "Specify Model Organism Table") = vbYes Then
'         modSys = systemTitle
'         newModSys = True
'         lblModSys = modSys
'      End If
'   Else
'      modSys = rsInfo!MODSystem
'      newModSys = False
'      speciesDB = rsInfo!species                                               'Has surrounding | |
'   End If
'   rsInfo.Close
'
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Create Gene Table And Error File
   dbgGeneDB.visible = False
   lblDetail = FileAbbrev(rawDataFile)
   DoEvents
   errorFile = Mid(rawDataFile, InStrRev(rawDataFile, "."))                              'Extension
   errorFile = Left(rawDataFile, InStrRev(rawDataFile, ".") - 1) & ".EX" & errorFile
   Open errorFile For Output As #2
   Set tdfGene = mappWindow.dbGene.CreateTableDef(systemTitle)
   With tdfGene
      .Fields.Append .CreateField("ID", dbText, CHAR_DATA_LIMIT)
      Set idxGene = .CreateIndex("ixID")
      With idxGene                     'Probably faster to index after filling table?????
         .Fields.Append .CreateField("ID")
      End With
      .Indexes.Append idxGene
'      .Fields.Append .CreateField("Species", dbMemo)                  'Species always second column
'      .Fields("Species").AllowZeroLength = True
      For i = 0 To frmDataID.lstTitles.ListCount - 1 '---------------------------Create Data Fields
         If i <> remarksIndex Then                                      'Move Remarks to end of row
            If frmDataID.lstTitles.selected(i) = True Then                        'Searchable field
'               If frmDataID.lstTitles.List(i) <> "Species" Then
                  .Fields.Append .CreateField(frmDataID.lstTitles.List(i), dbText, CHAR_DATA_LIMIT)
                  .Fields(frmDataID.lstTitles.List(i)).AllowZeroLength = True
'               End If
            Else                                                                        'Memo field
               .Fields.Append .CreateField(frmDataID.lstTitles.List(i), dbMemo)
               .Fields(frmDataID.lstTitles.List(i)).AllowZeroLength = True
            End If
         End If
      Next i
      .Fields.Append .CreateField("Date", dbDate)
      .Fields.Append .CreateField("Remarks", dbMemo)
      .Fields("Remarks").AllowZeroLength = True
      If deleteOldSystemCode <> "" Then '-------------------------------Delete Table To Be Replaced
On Error GoTo CantOpenDB
         mappWindow.dbGene.Execute "DROP TABLE " & systemTitle
         mappWindow.dbGene.Execute "DELETE FROM Systems WHERE System = '" & systemTitle & "'"
On Error GoTo 0
      End If
      inProcess = "AddSystem"
      mappWindow.dbGene.TableDefs.Append tdfGene
      Print #2, inLine; delimiter; "~Errors~"; vbLf;
   End With
   DoEvents
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Ready To Import
   lblDetail.visible = True
   lblPrgMax = "Errors"
   lblPrgMax.visible = True
   lblErrors.visible = True
   lblErrors = "0"
   prgProgress.visible = True
   DoEvents
   
   errors = GetGeneRow(errorsExists, geneId, geneValues, inLine, delimiter)        'First data line
'   If newModSys Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++ Get Species For New MOD
'      If speciesIndex <> -1 Then '---------------------------------------------------Species Column
'         '  speciesIndex will be -1 if no species column exists in raw data.
'         speciesDB = geneValues(speciesIndex)
'      End If
'      speciesDB = InputBox("Do you wish to enter the Genus and species for the Model " _
'                           & "Organism (e.g. ""Homo sapiens""). (Cancel leaves the species " _
'                           & "blank.)", "Declare Model Organism Species", speciesDB)
'
'      If speciesDB <> "" Then speciesDB = "|" & speciesDB & "|"
'   End If
''   speciei = speciesDB             'All species represented in this gene table. Start with this one
'      '  Since Gene Databases are now species specific, this should be only one species.
'      '  It will be used to fill the species column of the new Gene Table if no raw species
'      '  column exists (presumably it won't).
   Do While errors <> "**eof**" '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Each Row
      If errors <> "" Then
         lblErrors = lblErrors + 1
      End If
      
      '============================================================================Validate Gene ID
      Dim invalidChrs As String
      If Len(geneId) > CHAR_DATA_LIMIT Then                                                            'Too long
         errors = errors & "Gene ID more than " & CHAR_DATA_LIMIT & " characters. "
         lblErrors = lblErrors + 1
      End If
      invalidChrs = "gene ID"       'Send to InvalidChr function and get return of any invalid chrs
      If InvalidChr(geneId, invalidChrs) Then
         errors = errors & "Invalid character(s) " & invalidChrs & "found in gene ID. "
         lblErrors = lblErrors + 1
      End If
      '---------------------------------------------------------------------Check For Duplicate IDs
         Set rs = mappWindow.dbGene.OpenRecordset( _
                  "SELECT * FROM [" & systemTitle & "] WHERE ID = '" & geneId & "'")
         If Not rs.EOF Then
            errors = errors & "Gene ID """ & geneId & """ duplicated in raw data."
            lblErrors = lblErrors + 1
         End If
      
      '=======================================================================Check For Data Errors
      sql = ""       'Assemble SQL here even though errors may exist because we are testing for
                     'datatypes, etc. anyway and it would take more time to repeat the tests later
                     'than throw away an sql variable here. If the program enters any exception
                     'branch, the sql statement will not be added to because it will just be thrown
                     'away anyhow.
      remarks = ""                                                        'Default Remarks to empty
'      species = speciesDB                                       'Default species to that in Gene DB
         '  This could be added to by a species column in the raw file
      
      j = 1                                       'Column number in Gene Table with Remarks shifted
         '  Gene ID is always 0, Next column is 1
      
      For i = 1 To columns '===================================================Each raw data column
         If i = speciesIndex Then '--------------------------------------------------Species Column
               '  speciesIndex will be -1 if no species column exists in raw data. The species
               '  has already been defaulted to the species from the Gene DB.
               '  Before the decision to allow only one species per Gene DB, this section of
               '  the program would pick up different species and add them to the species list
               '  in the Gene DB's Info table. It still adds all represented species to the
               '  Systems Table for this Gene Table.
            species = geneValues(i)
            If Left(species, 1) <> "|" Then                  'Be sure species surrounded with pipes
               species = "|" & species
            End If
            If Right(species, 1) <> "|" Then
               species = species & "|"
            End If
            '_____________________________________________________Check For Species In speciei List
               prevPipe = 1
               pipe = InStr(prevPipe + 1, species, "|")
               Do While pipe                                        'For each species in raw record
                  If InStr(speciei, Mid(species, prevPipe, pipe - prevPipe + 1)) = 0 Then
                     speciei = speciei + Mid(species, prevPipe + 1, pipe - prevPipe)
                     If Left(speciei, 1) <> "|" Then speciei = "|" & speciei
                  End If
                  prevPipe = pipe
                  pipe = InStr(prevPipe + 1, species, "|")
               Loop
            geneValues(i) = species
         End If
         If i = remarksIndex Then '--------------------------------------------------Remarks Column
            remarks = Dat(geneValues(i))                                 'Handle NULLs the easy way
         ElseIf VarType(geneValues(i)) = vbNull Then '----------------------------------NULL values
            '  NULLs are accepted for any column except Gene ID or Species (if it exists),
            '  which are handled above
            '  By the time it gets here, all the special columns have been pulled out
            sql = sql & ", NULL"
            j = j + 1
         ElseIf tdfGene.Fields(j).Type = dbText Then
            If Len(geneValues(i)) > CHAR_DATA_LIMIT Then
               errors = errors & "Too many characters, '" & geneValues(i) & "', in column '" _
                      & tdfGene.Fields(j).name & "'. "
               geneValues(i) = Left(geneValues(i), CHAR_DATA_LIMIT)
               lblErrors = lblErrors + 1
            End If
            sql = sql & ", '" & geneValues(i) & "'"
            j = j + 1
         Else                                                    'Other datatypes (should be memos)
            sql = sql & ", '" & geneValues(i) & "'"
            j = j + 1
         End If
      Next i
      sql = sql & ", '" & Format(Now, "dd-mmm-yyyy") & "', '" & remarks & "')"

'      If lblErrors = 0 Then '==========================================Attempt To Add To Gene Table
         '  Don't bother to add to Gene Table if any errors exist
         '  If the sql statement fails, the ConvertError routine sets errorNumber
On Error GoTo ConvertError
         sql = "INSERT INTO [" & systemTitle & "]" & _
               "   VALUES ('" & geneId & "'" & sql
         mappWindow.dbGene.Execute sql     'Errors here fall to ConvertError, come back at QuitLine
'      End If

QuitLine:
      If errors <> "" Then '================================================Write To Exception File
         Print #2, inLine;                                               'Original line with errors
         i = 0
         j = InStr(inLine, delimiter) '--------------------------------Count Delimiters And Correct
            '  Microsoft Excel drops delimiters followed by empty cells at the end of a row
            '  beginning with the 16th row. It's a stupid bug but we attempt to get around it
            '  by making sure that the Exception file has the required number of columns before
            '  the ~Error~ column
         Do Until j = 0
            i = i + 1
            j = InStr(j + 1, inLine, delimiter)
         Loop
         For i = i To columns - 1
            Print #2, delimiter; " ";                                      'Add delimiter and space
         Next i
         Print #2, delimiter; errors; " " & vbLf;                'Space at end of error to foil NULL
      End If
      DoEvents
      If inProcess = "" Then '=========================================================Exit Clicked
         '  This could only happen following a DoEvents so the one above should be the only
         '  one between setting the inProcess variable and unsetting it.
         Close #FILE_RAW_DATA, #2
         s = mappWindow.dbGene.name                                   'Close and open to drop table
         mappWindow.dbGene.Close
         Set mappWindow.dbGene = Nothing
         OpenGeneDB mappWindow.dbGene, s, mappWindow
         mappWindow.dbGene.Execute "DROP TABLE " & systemTitle
         mappWindow.dbGene.Execute "DELETE FROM Systems WHERE System = '" & systemTitle & "'"
         GoTo ExitSub
      End If
      errors = GetGeneRow(errorsExists, geneId, geneValues, inLine, delimiter)
   Loop
On Error GoTo 0

   Close #FILE_RAW_DATA, #2 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Finish Up
'   mappWindow.dbGene.Close
'   Set dbExpression = Nothing
   If lblErrors = "0" Then '============================================Write Info & Systems Tables
      mappWindow.dbGene.Execute "UPDATE Info SET Modify = '" & Format(Now, "yyyymmdd") & "'"
      ModifyOwner
      Set rsInfo = mappWindow.dbGene.OpenRecordset("SELECT MODSystem, Species FROM Info")
      If Dat(rsInfo!MODSystem) = "" Then '-------------------------------------------Set MOD System
         If MsgBox("Is " & systemTitle & " the Model Organism table for this Gene Database?", _
                   vbInformation + vbYesNo, "Specify Model Organism Table") = vbYes Then
            lblModSys = systemTitle
            mappWindow.dbGene.Execute "UPDATE Info SET MODSystem = '" & systemTitle & "'"
            If Dat(rsInfo!species) <> speciei And speciei <> "" Then '------------------Set Species
               txtSpecies = InputBox("Enter Gene Database Genus and species (e.g. " _
                                     & """Homo sapiens"" or pipe-delimited for multiple " _
                                     & "species, e.g. ""Mus musculus|Rattus norvegicus"".).", _
                                     "Gene Database Species", speciei)
               txtSpecies = Dat(txtSpecies)
               If Left(txtSpecies, 1) <> "|" Then txtSpecies = "|" & txtSpecies
               If Right(txtSpecies, 1) <> "|" Then txtSpecies = txtSpecies & "|"
               mappWindow.dbGene.Execute _
                  "UPDATE Info SET Species = '" & txtSpecies & "'"
            End If
         End If
      End If
      rsInfo.Close
      
      sql = "ID|" '----------------------------------------------------------Systems Columns Column
      For i = 1 To tdfGene.Fields.count - 4
         If tdfGene.Fields(i).Type = dbText Then
            sql = sql & tdfGene.Fields(i).name & "\sBF|"
         Else
            sql = sql & tdfGene.Fields(i).name & "\BF|"
         End If
      Next i
      sql = "INSERT INTO Systems (System, SystemCode, SystemName, [Date], Columns, Species)" & _
            "   VALUES ('" & systemTitle & "', '" & newSystemCode & "', '" & _
                        systemTitle & "', '" & _
                        Format(Now, "dd-mmm-yyyy") & "', '" & sql & "', '" & _
                        speciei & "')"
      mappWindow.dbGene.Execute sql
      
      Kill errorFile '---------------------------------------------------------------------Clean Up
      FillSystemsList systemTitle
      cmbSystems_Click                                     'Go into edit mode on newly added system
      CreateGOCountVisible
'      geneDB = mappWindow.dbGene.name   'Close and open the DB because previous attempt at entering
'      mappWindow.dbGene.Close           'a table may have left it hung up and not DROPable
'      Set mappWindow.dbGene = OpenDatabase(geneDB)
   Else                                                                               'Errors exist
      MsgBox lblErrors & " errors were detected in your raw data. " _
             & "Correct problems in your raw data file and run the conversion again.", _
             vbExclamation + vbOKOnly, "Gene Table File Conversion"
'      Kill rawDataFile
'      Name errorFile As rawDataFile
      geneDB = mappWindow.dbGene.name   'Close and open the DB because previous attempt at entering
      mappWindow.dbGene.Close           'a table may have left it hung up and not DROPable
      Set mappWindow.dbGene = OpenDatabase(geneDB)
      mappWindow.dbGene.Execute "DROP TABLE " & systemTitle
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
ExitSub:                                         'Must have cancelled or screwed up to be sent here
   MousePointer = vbDefault
   inProcess = ""
   DoEvents
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Error Routines
RawDataFileError:
   MsgBox rawDataFile & " could not be opened. It may be open elsewhere or set to read-only " _
          & "through Windows, or perhaps does not exist.", _
          vbExclamation + vbOKOnly, "Converting Raw Data"
   Resume ExitSub                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   
CantOpenDB:
   MsgBox "The Gene Database could not be opened. It may be open elsewhere or set to read-only " _
          & "through Windows, or perhaps does not exist.", _
          vbExclamation + vbOKOnly, "Converting Raw Data"
   Resume ExitSub                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   
CancelProcedure: '========================================================Cancel Conversion Process
   mappWindow.dbGene.Execute "DROP TABLE " & systemTitle
   systemTitle = ""
   Close #FILE_RAW_DATA
   GoTo ExitSub                                            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

OpenError: '===========================================================Trap Database Opening Errors
   Select Case Err
   Case 32755 '-------------------------------------------------------------------------Cancel Open
      On Error GoTo 0
      Resume ExitSub
   Case Else '------------------------------------------------------------------Unidentified Errors
      Kill errorFile
      FatalError "frmGeneDB:mnuAdd at line:" & vbCrLf & vbCrLf & inLine _
                 & vbCrLf & vbCrLf & "Processing stopped", Err.Description
   End Select
   
ConvertError: '==============================================================Trap Conversion Errors
   errors = errors & "Trapped error '" & Err.Description & "' adding to Gene Table. "
   lblErrors = lblErrors + 1
   Resume QuitLine
End Sub

'/////////////////////////////////////////////////////////////////// Process New Relationship Table
Private Sub mnuAddRelations_Click() '*********************************** Add New Relationship Table
   Dim inLine As String, columns As Integer
   Dim errorFile As String
   Dim systemTitle As String, newSystemCode As String
   Dim geneValue As Integer
   Dim tdfGene As TableDef, idxGene As index
   Dim speciesIndex As Integer
   Dim species As String                                                  'Individual for data line
   Dim speciei As String                                                'All species for Gene Table
   Dim speciess As String                                                  'All species for gene DB
   Dim geneId As String
   Dim rsInfo As Recordset, rsSystems As Recordset
   Dim geneDB As String, sql As String
   Dim errorsExists As Integer, errors As String
   Dim prevPipe As Integer, pipe As Integer
   Dim i As Integer, s As String
   
   Do While mappWindow.dbGene Is Nothing '+++++++++++++++++++++++++++++++++++++++ Check For Gene DB
      OpenGeneDB mappWindow.dbGene, "**OPEN**"
      SetGeneDB mappWindow.dbGene
      If mappWindow.dbGene Is Nothing Then
         If MsgBox("Must have a Gene Database open.", _
                   vbOKCancel + vbExclamation, "New Relationship Table") = vbCancel Then
            GoTo ExitSub                                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      End If
   Loop
   
   ClearWindow                                                       'Make Unneeded Stuff Invisible
   
Retry:
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Raw Data File
   '  End result of this section is a valid rawDataFile
On Error GoTo OpenError
   With dlgDialog
      .CancelError = True
      .FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
      .DialogTitle = "Relationship File to Import"
      .InitDir = mruImportPath
      .FileName = ""
      .Filter = "All files|*.csv;*.tab;*.txt|Comma-separated values (.csv)|*.csv|" _
                       & "Tab-delimited lists (.tab, .txt)|*.tab;*.txt"
      .FilterIndex = 1
      .ShowOpen
On Error GoTo 0
      rawDataFile = .FileName
   End With
   
On Error GoTo RawDataFileError
      '  Try to rename the file. If it is in use, read-only, nonexistent, etc., it will error.
   Name rawDataFile As appPath & "rawTemp.$tm"
   Name appPath & "rawTemp.$tm" As rawDataFile
On Error GoTo 0

   mruImportPath = GetFolder(rawDataFile)
'   dbgGeneDB.Visible = False
   
   lblDetail = FileAbbrev(rawDataFile)
   lblDetail.visible = True
   FillSystemsList                                            'Fills Systems and SystemCodes arrays
'   cmbSystems.Visible = False
   SetGeneDB mappWindow.dbGene                                         'Set menus, labels, and such
   lblPrimary.visible = True
   lblRelated.visible = True
   cmbPrimary.visible = True
   cmbRelated.visible = True
   lblPrimaryCode.visible = True
   lblRelatedCode.visible = True
   txtPrimaryCode.visible = True
   txtRelatedCode.visible = True
   txtPrimaryCode.Enabled = False
   txtRelatedCode.Enabled = False
'   lblGeneTable.visible = False
'   cmbSystems.visible = False
   DoEvents
ExitSub:
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Error Routines
OpenError: '===========================================================Trap Database Opening Errors
   Select Case Err
   Case 32755 '-------------------------------------------------------------------------Cancel Open
      On Error GoTo 0
      Resume ExitSub
   Case Else '------------------------------------------------------------------Unidentified Errors
      Kill errorFile
      FatalError "frmGeneDB:mnuAddRelations at line:" & vbCrLf & vbCrLf & inLine _
                 & vbCrLf & vbCrLf & "Processing stopped", Err.Description
   End Select
   
RawDataFileError:
   MsgBox rawDataFile & " could not be opened. It may be open elsewhere or set to read-only " _
          & "through Windows, or perhaps does not exist.", _
          vbExclamation + vbOKOnly, "Adding Relations Table"
   Resume ExitSub                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   
End Sub
Private Sub cmbPrimary_Change()
   Dim i As Integer
   
   For i = 0 To lastSystem
      If cmbPrimary.text = cmbPrimary.List(i) Then
         txtPrimaryCode = systemCodes(i)
         txtPrimaryCode.Enabled = False
         CheckReadyToProcess
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Next i
   txtPrimaryCode = "&" & Mid(txtPrimaryCode, 2)
   txtPrimaryCode.Enabled = True                                  'Not in list, allow entry of code
   CheckReadyToProcess
End Sub
Private Sub cmbPrimary_Click()
   Dim i As Integer
   
   For i = 0 To lastSystem
      If cmbPrimary.text = cmbPrimary.List(i) Then
         txtPrimaryCode = systemCodes(i)
         txtPrimaryCode.Enabled = False
         CheckReadyToProcess
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Next i
   txtPrimaryCode = "&" & Mid(txtPrimaryCode, 2)
   txtPrimaryCode.Enabled = True                                  'Not in list, allow entry of code
   CheckReadyToProcess
End Sub
Private Sub txtPrimaryCode_Change()
   If txtPrimaryCode.Enabled Then
      txtPrimaryCode = "&" & Mid(txtPrimaryCode, 2)
   End If
   CheckReadyToProcess
End Sub

Private Sub cmbRelated_Change()
   Dim i As Integer
   
   For i = 0 To lastSystem
      If cmbRelated.text = cmbRelated.List(i) Then
         txtRelatedCode = systemCodes(i)
         txtRelatedCode.Enabled = False
         CheckReadyToProcess
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Next i
   txtRelatedCode = "&" & Mid(txtRelatedCode, 2)
   txtRelatedCode.Enabled = True                                  'Not in list, allow entry of code
   CheckReadyToProcess
End Sub
Private Sub cmbRelated_Click()
   Dim i As Integer
   
   For i = 0 To lastSystem
      If cmbRelated.text = cmbRelated.List(i) Then
         txtRelatedCode = systemCodes(i)
         txtRelatedCode.Enabled = False
         CheckReadyToProcess
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Next i
End Sub
Private Sub txtRelatedCode_Change()
   If txtRelatedCode.Enabled Then
      txtRelatedCode = "&" & Mid(txtRelatedCode, 2)
   End If
   CheckReadyToProcess
End Sub

Sub CheckReadyToProcess()
   cmdProcess.visible = False
   If cmbPrimary.text = "Choose or enter" Then Exit Sub    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   If cmbRelated.text = "Choose or enter" Then Exit Sub    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   If Dat(txtPrimaryCode) = "" Then Exit Sub
   If Dat(txtRelatedCode) = "" Then Exit Sub
   If txtPrimaryCode.Enabled _
         And (Left(txtPrimaryCode, 1) <> "&" Or Len(Dat(txtPrimaryCode)) <> 2) Then Exit Sub '>>>>>
   If txtRelatedCode.Enabled _
         And (Left(txtRelatedCode, 1) <> "&" Or Len(Dat(txtRelatedCode)) <> 2) Then Exit Sub '>>>>>
   If InvalidChr(cmbPrimary.text, "Gene Table Name") Then
      cmbPrimary.SetFocus
      Exit Sub
   End If
   If InvalidChr(cmbRelated.text, "Gene Table Name") Then
      cmbRelated.SetFocus
      Exit Sub
   End If
   cmdProcess.visible = True
End Sub
Private Sub cmdProcess_Click() '**************************************** Import New Relations Table
   Dim relationTitle As String, systemTitle As String, inLine As String
   Dim errorFile As String, errors As String, bytes As Long
   Dim rsRelations As Recordset, sql As String, geneDB As String
   Dim rsSystems As Recordset
   Dim primary As String, related As String, indexName As String
   Dim prevDelim As Integer, delim As Integer
   Dim deleteOldTable As String
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check Relation Title And Codes
   relationTitle = cmbPrimary.text & "-" & cmbRelated.text
   Set rsRelations = mappWindow.dbGene.OpenRecordset("SELECT * FROM Relations")
   Do Until rsRelations.EOF '======================================Go Through Each Relational Table
      If relationTitle = rsRelations!Relation Then                                     'Check Title
         If MsgBox(relationTitle & " already exists. Delete and replace it?", _
                   vbExclamation + vbOKCancel, "Adding Relational Table") = vbOK Then
            deleteOldTable = relationTitle
            Exit Do                                        'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         Else
            Exit Sub                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      End If
      rsRelations.MoveNext
   Loop
   If txtPrimaryCode.Enabled Then '=============================User-Entered Primary Code. Check it
      Set rsSystems = mappWindow.dbGene.OpenRecordset( _
                      "SELECT * FROM Systems WHERE SystemCode = '" & txtPrimaryCode & "'")
      If Not rsSystems.EOF Then                                       'Code exists in Systems table
         If cmbPrimary.text <> rsSystems!system Then                      'It's not the same system
            MsgBox "Code " & txtPrimaryCode & "already being used for system " _
                   & rsSystems!system & ". Cannot process.", _
                   vbCritical + vbOKOnly, "Adding Relational Table"
            txtPrimaryCode.SetFocus
            Exit Sub                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      End If
   End If
   If txtRelatedCode.Enabled Then '=============================User-Entered Related Code. Check It
      Set rsSystems = mappWindow.dbGene.OpenRecordset( _
                      "SELECT * FROM Systems WHERE SystemCode = '" & txtRelatedCode & "'")
      If Not rsSystems.EOF Then                                       'Code exists in Systems table
         If cmbRelated.text <> rsSystems!system Then                      'It's not the same system
            MsgBox "Code " & txtRelatedCode & "already being used for system " _
                   & rsSystems!system & ". Cannot process.", _
                   vbCritical + vbOKOnly, "Adding Relational Table"
            txtRelatedCode.SetFocus
            Exit Sub                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      End If
   End If
            
   systemTitle = dtaGeneDB.RecordSource
   dtaGeneDB.RecordSource = "Info"   'Switch to another, existing table to release the new table if
                                     'it had been used in a previous operation. This occurs when
                                     'entering the same raw data file twice in succession.
   dtaGeneDB.Refresh
   dbgGeneDB.Refresh
   lblPrgMax = "Errors"
   lblPrgMax.visible = True
   lblErrors = "0"
   lblErrors.visible = True
   prgProgress.Max = FileLen(rawDataFile)
   prgProgress.value = 0
   prgProgress.visible = True
   DoEvents
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine File Delimiter
   Dim delimiter As String * 1
   
   Open rawDataFile For Binary As #FILE_RAW_DATA       'Any cancelling after this has to Close file
   inLine = RemoveQuotes(InputUnixLine(FILE_RAW_DATA, bytes))
   prgProgress.value = bytes
   If InStr(inLine, vbTab) Then
      delimiter = vbTab
   Else
      delimiter = ","
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++ Create Relationship Table And Error File
   If deleteOldTable <> "" Then
On Error GoTo CantOpenDB
      mappWindow.dbGene.Execute "DROP TABLE [" & relationTitle & "]"
      '  If the table already exists, then the codes would have been checked already
      mappWindow.dbGene.Execute _
                 "DELETE FROM Relations WHERE Relation = '" & relationTitle & "'"
On Error GoTo 0
   End If

On Error GoTo ReplaceTable
      '  If the error already exists here it is because there was an incomplete run previously
      '  and the table name was not yet added to the Relations table.
   mappWindow.dbGene.Execute _
         "CREATE TABLE [" & relationTitle & "] ([Primary] Text(" & CHAR_DATA_LIMIT & ")," & _
         "             Related Text(" & CHAR_DATA_LIMIT & "), Bridge Text(3))"
On Error GoTo 0
   DoEvents
   errorFile = Left(rawDataFile, InStrRev(rawDataFile, ".") - 1) & ".$tm"
   Open errorFile For Output As #2
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Fill In Data Rows
   Dim invalidChrs As String
   Do Until inLine = "**eof**"                                    'First line already fetched above
      errors = ""
      prevDelim = 0
      delim = InStr(inLine, delimiter)
      If delim = 0 Then
         errors = errors & "Not enough data in record. "
      Else
         primary = Dat(Left(inLine, delim - 1))
         prevDelim = delim
         delim = InStr(prevDelim + 1, inLine, delimiter)
         If delim = 0 Then delim = Len(inLine) + 1
         related = Dat(Mid(inLine, prevDelim + 1, delim - prevDelim - 1))
      End If
      If primary = "" Then
         errors = errors & "No Primary gene ID. "
      End If
      If Len(primary) > CHAR_DATA_LIMIT Then
         errors = errors & "Primary ID more than " & CHAR_DATA_LIMIT & " characters. "
      End If
      If related = "" Then
         errors = errors & "No related gene ID. "
      End If
      If Len(related) > CHAR_DATA_LIMIT Then
         errors = errors & "Related ID more than " & CHAR_DATA_LIMIT & " characters. "
      End If
      invalidChrs = "gene ID"       'Send to InvalidChr function and get return of any invalid chrs
      If InvalidChr(primary, invalidChrs) Then
         errors = errors & "Invalid character(s) " & invalidChrs & "found in primary ID. "
      End If
      invalidChrs = "gene ID"       'Send to InvalidChr function and get return of any invalid chrs
      If InvalidChr(related, invalidChrs) Then
         errors = errors & "Invalid character(s) " & invalidChrs & "found in related ID. "
      End If
      
On Error GoTo ConvertError
      If cmbRelated.text = "GeneOntology" Then
         related = Right("0000000" & related, 7)
      End If
      sql = "INSERT INTO [" & relationTitle & "] ([Primary], Related)" & _
            "   VALUES ('" & primary & "', '" & related & "')"
      mappWindow.dbGene.Execute sql   'Errors here fall to ConvertError, come back at WriteRawData:

WriteRawData: '==========================================================Write To New Raw Data File
On Error GoTo 0
      Print #2, primary; delimiter; related; delimiter; errors; " " & vbLf;  'Space at end of error
                                                   'to foil NULL conversion in Excel, vbLf for Unix
      If errors <> "" Then
         lblErrors = lblErrors + 1                   'Count lines with errors, not number of errors
      End If
      inLine = RemoveQuotes(InputUnixLine(FILE_RAW_DATA, bytes))
      prgProgress.value = Min(prgProgress.Max, bytes)            'To avoid looking past end of file
      DoEvents
   Loop

   Close #FILE_RAW_DATA, #2 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Finish Up
   If lblErrors = "0" Then '===================================Index, Write Info & Relations Tables
      indexName = txtPrimaryCode & "P"
      For i = 1 To Len(indexName)
         If Mid(indexName, i, 1) = "&" Then Mid(indexName, i, 1) = "_"
      Next i
      mappWindow.dbGene.Execute _
                 "CREATE INDEX idx" & indexName & " ON [" & relationTitle & "] ([Primary])"
      indexName = txtRelatedCode & "R"
      For i = 1 To Len(indexName)
         If Mid(indexName, i, 1) = "&" Then Mid(indexName, i, 1) = "_"
      Next i
      mappWindow.dbGene.Execute _
                 "CREATE INDEX idx" & indexName & " ON [" & relationTitle & "] (Related)"
      mappWindow.dbGene.Execute _
         "UPDATE Info SET Modify = '" & Format(Now, "yyyymmdd") & "'"
      ModifyOwner
      sql = "INSERT INTO Relations (SystemCode, RelatedCode, Relation, [Type])" & _
            "   VALUES ('" & txtPrimaryCode & "', '" & txtRelatedCode & "', '" & _
                        relationTitle & "', 'User')"
      mappWindow.dbGene.Execute sql
'      lblPrgMax = "Compacting database"
'      lblErrors.Visible = False
'      prgProgress.Visible = False
      DoEvents
'If Not TESTING Then
'      geneDB = mappWindow.dbGene.name
'      mappWindow.dbGene.Close
'      dtaGeneDB.DatabaseName = appPath & "MAPPTmpl.gtp"               'Release DB from data control
'         '  There seems to be no direct way to release a DB from a data control, so the
'         '  data control is bound to another existing database temporarily.
'      dtaGeneDB.RecordSource = "Info"
'      dtaGeneDB.Refresh
'      If Dir(Left(geneDB, InStrRev(geneDB, ".")) & "$tm") <> "" Then             'Just to make sure
'         Kill Left(geneDB, InStrRev(geneDB, ".")) & "$tm"
'      End If
'      DBEngine.CompactDatabase geneDB, Left(geneDB, InStrRev(geneDB, ".")) & "$tm"
'      Kill geneDB
'      Name Left(geneDB, InStrRev(geneDB, ".")) & "$tm" As geneDB
'      Set mappWindow.dbGene = OpenDatabase(geneDB)
'      dtaGeneDB.DatabaseName = mappWindow.dbGene.name
'      dtaGeneDB.Refresh
'End If
      Kill errorFile
   Else '==============================================================================Errors exist
      MsgBox lblErrors & " errors were detected in your raw data. " _
             & "Correct problems in your raw data file and run the conversion again.", _
             vbExclamation + vbOKOnly, "Relations Table File Conversion"
      Kill rawDataFile
      Name errorFile As rawDataFile
      mappWindow.dbGene.Execute "DROP TABLE [" & relationTitle & "]"
   End If

ExitSub:
   lblPrimary.visible = False
   lblRelated.visible = False
   cmbPrimary.visible = False
   cmbRelated.visible = False
   lblPrimaryCode.visible = False
   lblRelatedCode.visible = False
   txtPrimaryCode.visible = False
   txtRelatedCode.visible = False
   cmdProcess.visible = False
   lblErrors.visible = False
   lblDetail.visible = False
   lblPrgMax.visible = False
   prgProgress.visible = False
   FillSystemsList systemTitle
'   cmbSystems.Visible = True
'   dtaGeneDB.RecordSource = systemTitle
'   dtaGeneDB.Refresh
'   dbgGeneDB.Refresh
'   dbgGeneDB.Visible = True
   CreateGOCountVisible
   MousePointer = vbDefault
'   mnuEdit_Click
   DoEvents
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Error Routines
CantOpenDB:
   Select Case Err.number
   Case 3376                                                      'Table does not exist, can't drop
      '  This would be the case if there was incomplete processing at a last attempt.
      Resume Next                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   Case Else
      MsgBox "The Gene Database could not be opened. It may be open elsewhere or set to read-only " _
             & "through Windows, or perhaps does not exist.", _
             vbExclamation + vbOKOnly, "Add Relations Table"
   End Select
   Resume ExitSub                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   
ReplaceTable:
      '  If the error already exists here it is because there was an incomplete run previously
      '  and the table name was not yet added to the Relations table.
   mappWindow.dbGene.Execute "DROP TABLE [" & relationTitle & "]"
   Resume
   
ConvertError: '==============================================================Trap Conversion Errors
   errors = errors & "Trapped error '" & Err.Description & "' adding to Relations Table. "
   Resume WriteRawData
End Sub

Private Sub cmbSpecies_Click()
'   If dbgGeneDB.col <> dbgGeneDB.columns("Species").ColIndex Then        'Other than species column
'      '  Filling the combo box and moving "Clear species" to the top causes this click event so
'      '  make sure we are in the Species column
'      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   End If
   If cmbSpecies.Tag = "Set Gene DB Species" Then
      If InStr(txtSpecies, "|" & cmbSpecies.text & "|") = 0 Then            'Species not in textbox
         If txtSpecies = "" Then txtSpecies = "|"
         txtSpecies = txtSpecies & cmbSpecies.text & "|"
      End If
      txtSpecies.SetFocus
   Else
      If InStr(dbgGeneDB.text, "|" & cmbSpecies.text & "|") = 0 Then           'Species not in grid
         If dbgGeneDB.text = "" Then
            dbgGeneDB.text = "|"
         End If
         dbgGeneDB.text = dbgGeneDB.text & cmbSpecies.text & "|"
         dbgGeneDB_AfterColUpdate dbgGeneDB.columns("Species").ColIndex
         dataChanged = True
      End If
      dbgGeneDB.SetFocus
   End If
End Sub

Private Sub cmbSpecies_LostFocus()
'   InvalidRow dbgGeneDB.row
End Sub

'///////////////////////////////////////////////////////////////////////////////////// Grid Actions
Private Sub cmdNextSegment_Click()
   '  The DBGrid supposedly handles as many rows as system resources allow. BS! It handles a
   '  variable number of rows, first 250,001 (0 to 250,000) but as you add segments it handles
   '  multiples of the 250,000 segment. Never, however, does it handle the millions of rows
   '  we need, So, the Previous and Next Segment command buttons will move it one way or
   '  another by 250,000 rows, allowing access to the entire table.
   Dim segment As Long, row As Long

   MousePointer = vbHourglass
   dbgGeneDB.row = dbgGeneDB.VisibleRows - 1                     'Puts dtaGeneDB in current segment
   segment = dtaGeneDB.Recordset.AbsolutePosition \ 250000                              'Zero based
   segment = segment + 1
   row = segment * 250000 + 1                                 'Row number of beginning of segment
   dtaGeneDB.Recordset.AbsolutePosition = row                            'Force refilling of grid
   If segment >= dtaGeneDB.Recordset.recordCount \ 250000 Then                     'In last segment
      '  dtaGeneDB.Recordset.RecordCount \ 250000 is last segment
      cmdNextSegment.Enabled = False                                       'Turn off command button
   End If
   If segment > 0 Then
      cmdPreviousSegment.Enabled = True
   End If
   MousePointer = vbDefault
End Sub

Private Sub cmdPreviousSegment_Click()
   Dim segment As Long, row As Long
   
   MousePointer = vbHourglass
   dbgGeneDB.row = 0                                             'Puts dtaGeneDB in current segment
   segment = dtaGeneDB.Recordset.AbsolutePosition \ 250000                              'Zero based
   segment = segment - 1
   row = segment * 250000                                    'Row number of beginning of segment
   dtaGeneDB.Recordset.AbsolutePosition = row                           'Force refilling of grid
   dtaGeneDB.Recordset.AbsolutePosition = row + 249999                   'Move to end of segment
   If segment = 0 Then                                                            'In first segment
      cmdPreviousSegment.Enabled = False                                   'Turn off command button
   End If
   If segment < dtaGeneDB.Recordset.recordCount \ 250000 Then              'Segment < TotalSegments
      cmdNextSegment.Enabled = True
   End If
   MousePointer = vbDefault
End Sub

Private Sub dtaGeneDB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   dtaGeneDB.Caption = dtaGeneDB.Recordset.AbsolutePosition
End Sub

Private Sub dbgGeneDB_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, _
                                      Cancel As Integer)
   '  This event only fires if data is changed in the grid. It sets the dataChanged variable to
   '  True. This variable is checked by the InvalidRow() function, which is called whenever the
   '  user moves off a row, i.e. by dbgGeneDB_RowColChange() if moving within the grid,
   '  dbgGeneDB_LostFocus() if moving to another control, or Form_QueryUnload() if clicking
   '  the close box.

   dataChanged = True
End Sub
Private Sub dbgGeneDB_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   '  Fired after dbgGeneDB_AfterColUpdate(). Cursor at new row and column at this point.
   '  Not fired if user clicks on another control. See dbgGeneDB_LostFocus()
   If loading Then Exit Sub                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
On Error GoTo NoSpecies
   If dbgGeneDB.col = dbgGeneDB.columns("Species").ColIndex And Not dbgGeneDB.columns("Species").Locked Then
      ShowSpecies True
   Else
      ShowSpecies False
   End If
   If dbgGeneDB.Bookmark <> prevRow Then                          'User changed rows, check old one
      InvalidRow prevRow
   End If
   prevRow = dbgGeneDB.Bookmark
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
NoSpecies:
   If Err.number = 6147 Then
      ShowSpecies False
   Else
      FatalError "frmGeneDBMgr:dbgGeneDB_RowColChange", Err.Description
   End If
End Sub
Private Sub dbgGeneDB_GotFocus()
On Error GoTo NoSpecies
   If dbgGeneDB.col = dbgGeneDB.columns("Species").ColIndex And cmbSystems.text = "Other" Then
      ShowSpecies True
   End If
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

NoSpecies:
   If Err.number = 6147 Then
      Exit Sub
   Else
      FatalError "frmGeneDBMgr:dbgGeneDB_GotFocus", Err.Description
   End If
End Sub
Private Sub dbgGeneDB_LostFocus()
'   If dbgGeneDB.col <> dbgGeneDB.columns("Species").ColIndex Then        'Other than species column
   If ActiveControl.name <> "cmbSpecies" And ActiveControl.name <> "dbgGeneDB" Then
      ShowSpecies False                                                         'Close Species list
      VerifySystemsSpecies
'      InvalidRow dbgGeneDB.Bookmark
   End If
End Sub
Sub VerifySystemsSpecies(Optional system As String = "") '***** Collect All Species From Gene Table
   Dim rsSystems As Recordset, rsSystem As Recordset, species As String
   Dim lastSpecies As Integer, i As Integer, j As Integer
   Dim speciei(100) As String, noOfSpeciei As Integer
   Dim speciesChanged As Boolean
   
   If system = "" Then system = cmbSystems.text
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ See If Species Column Exists
   For i = 0 To mappWindow.dbGene.TableDefs(system).Fields.count - 1
      If mappWindow.dbGene.TableDefs(system).Fields(i).name = "Species" Then Exit For
   Next i
   If i > mappWindow.dbGene.TableDefs(system).Fields.count - 1 Then              'Species Not Found
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   MousePointer = vbHourglass
   lblDetail.visible = True
   lblDetail = "Collecting species data"
   DoEvents
   
   species = "|"
   Set rsSystem = mappWindow.dbGene.OpenRecordset( _
                  "SELECT Species FROM " & system)
   If Dat(rsSystem!species) <> "" Then
      Do Until rsSystem.EOF
         noOfSpeciei = SeparateValues(speciei, rsSystem!species, "|")
         For i = 0 To noOfSpeciei - 1
            If InStr(species, speciei(i)) = 0 Then
               species = species & speciei(i) & "|"
            End If
         Next i
         rsSystem.MoveNext
      Loop
      mappWindow.dbGene.Execute "UPDATE Systems SET Species = '" & species & "'" & _
                                "   WHERE System = '" & system & "'"
   End If
   
   MousePointer = vbDefault
   lblDetail.visible = False
   lblDetail = ""
   DoEvents

End Sub
'//////////////////////////////////////////////////////////////////////////////////// Other Actions
Private Sub mnuGeneDBInfo_Click()
   Tag = "DontActivate"
   mappWindow.mnuGeneDBInfo_Click
End Sub
Private Sub mnuUpdateGeneDB_Click()
   Dim dbGeneOld As Database, dbGeneNew As Database
   Dim tdfNewDB As TableDef, tdf As TableDef, fld As Field
   Dim rsInfo As Recordset, rsOldTable As Recordset
   Dim oldGeneDBName As String
   Dim newGeneDBName As String, newGeneDBExtension As String, geneDBFolder As String
   Dim newTable As Boolean
   Dim version As String
   
'   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ See If New Gene Database Exists
'   Set dbGene = mappWindow.dbGene
'   newGeneDBName = Left(dbGene.name, InStrRev(dbGene.name, ".") - 1)          'Includes entire path
'   newGeneDBExtension = Mid(dbGene.name, InStrRev(dbGene.name, "."))
'      '  This should always be ".gdb" but just in case . . .
'   geneDBFolder = Left(dbGene.name, InStrRev(dbGene.name, "\"))
'   s = Dir(newGeneDBName & "_*" & newGeneDBExtension)
'   If s = "" Then
'      MsgBox "No new Gene Database" & vbCrLf _
'              & "   " & Mid(dbGene.name, InStrRev(newGeneDBName, "\") + 1) & "_[version]" _
'              & newGeneDBExtension & vbCrLf _
'              & "identified in folder" & vbCrLf _
'              & "   " & Left(dbGene.name, InStrRev(dbGene.name, "\")) & ".", _
'              vbExclamation + vbOKOnly, "Update Gene Database"
'      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   End If
'   newGeneDBName = s
   
ChooseNewGeneDB: '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine New Gene DB
On Error GoTo CancelUpdate
   With dlgDialog
      .DialogTitle = "New Gene Database"
      .CancelError = True
      .InitDir = GetFolder(mruGeneDB)
      .Filter = "Gene Databases (.gdb)|gdb"
      .FileName = GetFolder(mruGeneDB) & "*.gdb"
      .FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
      .ShowOpen
      newGeneDBName = .FileName
   End With
On Error GoTo 0
   If InStr(newGeneDBName, ".") = 0 Then
      newGeneDBName = newGeneDBName & ".gdb"
   End If
   If Dir(newGeneDBName) = "" Then
      MsgBox "The Gene Database" & vbCrLf & vbCrLf & newGeneDBName & vbCrLf & vbCrLf & _
             "does not exist.", vbOKOnly + vbExclamation, "Update Gene Database"
      GoTo ChooseNewGeneDB                                 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Open New Gene DB
      '  By this time we are committed to the new Gene DB and it will replace the current one
      '  at the end of the routine
   MousePointer = vbHourglass
   If Not mappWindow.dbGene Is Nothing Then
      mappWindow.dbGene.Close
      Set mappWindow.dbGene = Nothing
   End If
   Set dbGeneNew = OpenDatabase(newGeneDBName)
      
ChooseOldGeneDB: '++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Gene DB To Fold In
On Error GoTo CancelFoldIn
   With dlgDialog
      .DialogTitle = "Old Gene Database to Fold Into New One"
      .CancelError = True
      .InitDir = GetFolder(newGeneDBName)
      .Filter = "Gene Databases (.gdb)|gdb"
      i = InStrRev(newGeneDBName, "_")
         '  Default old name to something like the new with another date.
         '  If new is        C:\GenMAPP\Mm-Std_20030930.xyz
         '  default will be  C:\GenMAPP\Mm-Std_*.xyz
      If i <> 0 Then
         j = InStrRev(newGeneDBName, ".")
         .FileName = Left(newGeneDBName, i) & "*" & Mid(newGeneDBName, j)
      Else
         .FileName = GetFolder(newGeneDBName) & "*.gdb"
      End If
      .FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
      .ShowOpen
      oldGeneDBName = .FileName
   End With
On Error GoTo 0
   If InStr(oldGeneDBName, ".") = 0 Then
      oldGeneDBName = oldGeneDBName & ".gdb"
   End If
   If Dir(oldGeneDBName) = "" Then
      MsgBox "The Gene Database" & vbCrLf & vbCrLf & oldGeneDBName & vbCrLf & vbCrLf & _
             "does not exist.", vbOKOnly + vbExclamation, "Update Gene Database"
      GoTo ChooseOldGeneDB                                 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   End If
   Set dbGeneOld = OpenDatabase(oldGeneDBName)
   
   lblDetail.visible = True
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Transfer "Remarks" To New Tables
   lblDetail = "Transferring Remarks"
   DoEvents
   For Each tdf In dbGeneNew.TableDefs
      If tdf.name <> "Other" Then
         For Each fld In tdf.Fields
            If fld.name = "Remarks" Then
               '  Only Gene Tables have Remarks fields
               Set rsOldTable = dbGeneOld.OpenRecordset( _
                                "SELECT ID, Remarks FROM [" & tdf.name & "]" & _
                                "   WHERE Remarks IS NOT NULL", dbOpenForwardOnly)
                  '  This may copy some empty Remarks but so what?
               Do Until rsOldTable.EOF
                  dbGeneNew.Execute "UPDATE [" & tdf.name & "]" & _
                                    "   SET Remarks = '" & rsOldTable!remarks & "'" & _
                                    "   WHERE ID = '" & rsOldTable!id & "'"
                  rsOldTable.MoveNext
               Loop
               Exit For
            End If
         Next fld
      End If
   Next tdf
      
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Transfer "Other" Table
   lblDetail = "Transferring ""Other"" table"
   DoEvents
   dbGeneNew.Execute "INSERT INTO Other SELECT * FROM Other IN '" & dbGeneOld.name & "'"
      '  This adds to the new Other table without checking for duplicates
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Transfer User-Added Tables
   lblDetail = "Transferring user-added tables"
   DoEvents
   For Each tdf In dbGeneOld.TableDefs
      newTable = True
      For Each tdfNewDB In dbGeneNew.TableDefs
         If UCase(tdf.name) = UCase(tdfNewDB.name) Then
            '  UCase in case any old Gene DBs had wrong case, especially Unigene (UniGene)
            newTable = False
            Exit For                                       'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      Next tdfNewDB
      If newTable Then
         dbGeneOld.Execute "SELECT * INTO [" & tdf.name & "] IN '" & dbGeneNew.name & "'" & _
                           "   FROM [" & tdf.name & "]"
      End If
   Next tdf
   
   '===========================================================================Adjust Systems Table
      '  Pick up all records from old Systems table that are not in the new one.
   lblDetail = "Updating Systems Table"
   DoEvents
   dbGeneNew.Execute _
             "INSERT INTO Systems" & _
             "   SELECT * FROM Systems IN '" & dbGeneOld.name & "'" & _
             "      WHERE System <> ALL (SELECT System FROM Systems IN '" & dbGeneNew.name & "')"
   
   '=========================================================================Adjust Relations Table
      '  Pick up all records from old Relations table that are not in the new one.
   lblDetail = "Updating Relations Table"
   DoEvents
   dbGeneNew.Execute _
             "INSERT INTO Relations" & _
             "   SELECT * FROM Relations IN '" & dbGeneOld.name & "'" & _
             "      WHERE Relation <> ALL" & _
             "            (SELECT Relation FROM Relations IN '" & dbGeneNew.name & "')"
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Change Name On Old Database
   lblDetail = "Finalizing update"
   DoEvents
'   oldGeneDBName = dbGeneOld.name
'
'   Set rsInfo = dbGeneOld.OpenRecordset("SELECT Version FROM Info")
'   version = rsInfo!version
'   oldGeneDBName = dbGene.name
'   newGeneDBName = dbGeneNew.name
'   dbGene.Close
'   dbGeneNew.Close
'   Name oldGeneDBName As Left(oldGeneDBName, InStrRev(oldGeneDBName, ".") - 1) & "-" & version _
'                         & Mid(oldGeneDBName, InStrRev(oldGeneDBName, "."))
'   Name newGeneDBName As oldGeneDBName
'   Set mappWindow.dbGene = OpenDatabase(oldGeneDBName)
   
   dbGeneOld.Close
   MsgBox "You have integrated" & vbCrLf & vbCrLf & oldGeneDBName & vbCrLf & vbCrLf & _
          "into your new Gene Dababase. You may wish to delete this old one or move it to " & _
          "an archive.", vbOKOnly + vbInformation, "Updating Gene Database"
   
ResetCurrentGeneDB:
   dbGeneNew.Close
   OpenGeneDB mappWindow.dbGene, newGeneDBName, mappWindow
   SetGeneDB mappWindow.dbGene                     'Make this the active Gene DB in the Gene DB Mgr
ExitSub:
   lblDetail.visible = False
   MousePointer = vbDefault
   DoEvents
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
CancelUpdate:
   '  Cancelling when specifying the new Gene DB just exits the Update routine, leaving things
   '  as they were before entering
   If Err <> 32755 Then                                                          'Other than Cancel
      FatalError "frmGeneDBMgr:mnuUpdateGeneDB", Err.Description
   End If
   On Error GoTo 0
   Resume ExitSub                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

CancelFoldIn:
   '  Cancelling when specifying the old Gene DB to fold in does not perform any update but
   '  allows the new Gene DB to remain the current one in the MAPP window.
   If Err <> 32755 Then                                                          'Other than Cancel
      FatalError "frmGeneDBMgr:mnuUpdateGeneDB", Err.Description
   End If
   On Error GoTo 0
   Resume ResetCurrentGeneDB                                          '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
End Sub


'///////////////////////////////////////////////////////////////////// Procedures Called From Above
Function InvalidRow(row As Variant) As Boolean '************************************ Check Grid Row
   '  Entry:   row   The row to check
   '  Return:  True if something wrong with row
   Dim newRow As Variant                                               'Row the cursor has moved to
   
Exit Function
'Don't think we need this or dataChaged variable anymore

   If Not dataChanged Then Exit Function                   'Nothing to check >>>>>>>>>>>>>>>>>>>>>>
   
   newRow = dbgGeneDB.Bookmark                         'User has clicked some other row or off grid
   dbgGeneDB.Bookmark = row                                     'Set grid back to row to be checked
   If cmbSpecies.ListCount > 2 Then          'More than one species allowable, plus "Clear species"
      s = dbgGeneDB.columns("Species").text
      If Left(s, 1) <> "|" Or Right(s, 1) <> "|" Then                          'Pipes on either end
         MsgBox "Must choose at least one species.", _
                vbExclamation + vbOKOnly, "Editing Gene Database"
         dbgGeneDB.col = dbgGeneDB.columns("Species").ColIndex             'Focus to Species column
         ShowSpecies True
         InvalidRow = True
         dbgGeneDB.SetFocus
            '  Do not set dataChanged to False. User may just click off row, which would not force
            '  a recheck
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   dbgGeneDB.Bookmark = newRow                           'Set grid back to next row to be processed
   dataChanged = False                                    'Row OK, do not recheck unless new change
End Function
Sub ShowSpecies(show As Boolean)
   If show Then PopulateSpeciesList
   If cmbSpecies.ListCount < 1 Then                                           'No species available
      show = False
   End If
   cmbSpecies.visible = show
   lblSpecies.visible = show
End Sub
'********************************************************************** Fill Systems Drop-Down List
Sub FillSystemsList(Optional systemTitle As String = "")
   '  Entry:      'systemTitle    System to edit. If blank, first system is chosen
   '  Fills cmbSystems with supported Gene Tables, and systems() and systemCodes() arrays with
   '  all systems
   Dim rsSystems As Recordset
   Dim index As Integer
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Systems
   Set rsSystems = mappWindow.dbGene.OpenRecordset( _
                   "SELECT * FROM Systems ORDER BY System", dbOpenForwardOnly)
   cmbSystems.Clear
   cmbPrimary.Clear
   cmbRelated.Clear
   cmbPrimary.text = "Choose or enter"
   cmbRelated.text = "Choose or enter"
   lastSystem = -1
   cmbSystems.AddItem "Choose Gene Table to edit"
   Do Until rsSystems.EOF
      lastSystem = lastSystem + 1
      systems(lastSystem) = rsSystems!system
      systemCodes(lastSystem) = rsSystems!systemCode
      If rsSystems!system = "Other" Or (VarType(rsSystems!Date) <> vbNull _
            And (VarType(rsSystems!Misc) = vbNull Or InStr(rsSystems!Misc, "|E|") = 0)) Then
         '  "Other" always editable, others editable if the date is not null (it exists) and
         '  not designated as an empty table.
         cmbSystems.AddItem rsSystems!system                                       'For Gene Tables
         If rsSystems!system = systemTitle Then
            cmbSystems.ListIndex = cmbSystems.ListCount - 1                   'Triggers Click event
         End If
      End If
      cmbPrimary.AddItem rsSystems!system                        'For Relations Tables, all systems
      cmbRelated.AddItem rsSystems!system
      rsSystems.MoveNext
   Loop
   If cmbSystems.ListIndex = -1 Then cmbSystems.ListIndex = 0              'Default to first system
End Sub
'************************************************************ Gets Next Gene Row From Raw Data File
Function GetGeneRow(errorsExists As Integer, geneId As String, geneValues() As String, _
                    inLine As String, Optional delimiter As String = vbTab) As String
   'Entry:  The next line in file #file_raw_data, which must be open and set correctly
   '        ErrorsExists   1 if ~Error~ column exists in raw data. Last column ignored
   '        delimiter      Tab or comma character
   'Return: Blank if successful. Error message or **EOF** if not.
   '        geneID      First column, which must be the gene ID.
   '        geneValues() All subsequent columns. One-based.
   '        inLine      The unchanged input line. Errors removed.
   '  Checks line first for tab character to determine if tab-delimited list
   
   Dim Lin As String                                                                'Line from file
   Dim prevMark As Integer, mark As Integer, lastExpValue As Integer
   
   Do '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Get Line
      inLine = InputUnixLine(FILE_RAW_DATA)
   Loop While inLine = ""                                                       'Ignore blank lines
   
   If inLine = "**eof**" Then
      GetGeneRow = "**eof**"
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   prgProgress.value = Seek(50) - 2                            'To avoid looking beyond end of file
   
   If errorsExists Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++ Strip Errors Column
      inLine = Left(inLine, InStrRev(inLine, delimiter) - 1)
   End If
   
   For i = 1 To UBound(geneValues) '++++++++++++++++++++++++++++++++++++++++ Start With Empty Array
      geneValues(i) = ""
   Next i
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Parse The Line
   Lin = RemoveQuotes(inLine)
   lastExpValue = 0                                                                'One-based index
   mark = InStr(Lin, delimiter)
   If mark = 0 Then mark = Len(Lin) + 1
   geneId = Mid(Lin, prevMark + 1, mark - prevMark - 1)
   prevMark = mark
   Do Until prevMark > Len(Lin)
      mark = InStr(prevMark + 1, Lin, delimiter)
      If mark = 0 Then mark = Len(Lin) + 1
      lastExpValue = lastExpValue + 1
      If lastExpValue > UBound(geneValues) Then
         '  We are beyond the bounds of the array and must exit the function.
         If errorsExists Then                      'There is an error column that should be ignored
            mark = InStr(mark + 1, Lin, delimiter)
            If mark Then                                                  'There is a column beyond
               GetGeneRow = "Too many columns. "
            End If
         Else                                                    'No error column, too many columns
            GetGeneRow = "Too many columns. "
         End If
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
      geneValues(lastExpValue) = Mid(Lin, prevMark + 1, mark - prevMark - 1)
      geneValues(lastExpValue) = Trim(TextToSql(geneValues(lastExpValue)))
      prevMark = mark
   Loop
   If lastExpValue < UBound(geneValues) - 1 Then          'Allow one empty column, assuming Remarks
      GetGeneRow = "Too few columns. "
   End If
End Function

Private Sub mnuCreateGOCount_Click()
   GOCountTable lblModSys            'If lblModSys is not valid, this menu item will not be enabled
   ClearWindow
End Sub
Sub GOCountTable(system As String) '***************************************************************
   '  This and CountChildren() are adapted from the MasterUpdate program. Important changes are:
   '     1. MasterUpdate works with SQL Server, this routine works with Jet.
   '     2. The system passed in in MasterUpdate "Sub GOCountTable(system As String)" is
   '        always the MOD system currently in the open Gene DB in the Gene DB Mgr.

   '  The Count field is the easy one. For each GO term that exists in the Species_GeneOntology
   '  table, count the number of times it appears in that table. Total is a monster. GO is not
   '  a strictly hierarchical dataset; it is a DAG (directed acyclic graph), meaning that a
   '  child may have multiple parents and not necessarily on the same level. Total is the
   '  number of each parent and all of its children.
   
   '  To generate it, we take a GO term from the GeneOntology table and, if it exists in the
   '  Species_GeneOntology table, use each of its occurrences as a root of a tree -- a parent.
   '  Each of the parents adds one to the Total. We follow each parent to each of its children,
   '  adding one for each child. In a recursive process, we do the same for each child, using
   '  it as a parent, adding one each time, until we reach a terminal node -- a leaf with no
   '  more children. The terminal nodes also add to Total because they are valid occurrences.
   
   '  Then we back up one level in the tree, follow the next child, until there are no more
   '  children at that level and below, then back up one level again, and so forth. When we
   '  back up to the original parent level, we move on to the next parent.

   Dim systemGo As String                        'Name of GO table for system. Eg: MGI-GeneOntology
   Dim GOCount As String                                    'Name of GOCount table. Eg: MGI-GOCount
   Dim rsParent As Recordset, rs As Recordset
   Dim parent As Long                      'Number of occurrences of just parent in system-GO table
   Dim count As Long      'Number of occurrences of both parent and all children in system-GO table
   Dim genes(100000) As String               'Keep track of all gene names in root through branches
   Dim recordCount As Long
      '  For some gene systems, this may not be a big enough array. If so we can use a temp table.
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Extra Stuff For GenMAPP Function
   Const TEXT_FIELD = 30                                              'Must agree with MasterUpdate
   Dim tdf As TableDef
   Dim dbGenMapp As Database                          'To use same variable as used in MasterUpdate
   Set dbGenMapp = mappWindow.dbGene
   ClearWindow
   
'   If system = "SwissProt" Then '======================================= Handle SwissProt Exception
'      '  For SwissProt, the GOCount table should contain data only for Homo sapiens.
'      systemGo = "SwissProt-GeneOntology_Human"
'   Else
      systemGo = system & "-GeneOntology"
'   End If
   
   GOCount = system & "-GOCount"
'   History system & ": Generating " & GOCount & " table"
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Create Table
'   DropTable dbGenMAPP, GOCount
   For Each tdf In dbGenMapp.TableDefs '===================================Check For Existing Table
      If tdf.name = GOCount Then
         If MsgBox("GOCount table already exists. Overwrite it?", vbQuestion + vbOKCancel, _
                   "Creating GOCount Table") = vbCancel Then
            Exit Sub                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
         dbGenMapp.Execute "DROP TABLE [" & GOCount & "]"
         Exit For
      End If
   Next tdf
   
'   dbGenMAPP.Execute "CREATE TABLE [" & GOCount & "]" & _
'                     "   ([GO] NVARCHAR(" & TEXT_FIELD & "), [Count] SMALLINT, Total INT)"
   dbGenMapp.Execute "CREATE TABLE [" & GOCount & "]" & _
                     "   ([GO] text(" & TEXT_FIELD & "), [Count] Integer, Total Long)"
                     
   lblProgressTitle.visible = True
   prgProgress.visible = True
   lblProgressMax.visible = True
   lblProgressTitle = system & ": Generating " & GOCount & " table"
   inProcess = "GOCount"
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Get All DISTINCT GO IDs
'   rsParent.Open "SELECT COUNT(*) AS TotalRows FROM GeneOntologyCount", dbGenMAPP, adOpenDynamic
   Set rsParent = dbGenMapp.OpenRecordset("SELECT COUNT(*) AS TotalRows FROM GeneOntologyCount")
      '  Uses GeneOntologyCount because each ID is unique and the GeneOntology Count is the
      '  denominator in the MAPPFinder fraction
      '  Treat each GO ID as if it were a parent
      '  Must look at all IDs because even if parent not represented in system-GO,
      '     children might be
      '  Starts from just below three roots, biological process, etc.
'   SetProgressBase rsParent!TotalRows, "records"
   prgProgress.value = 0
   prgProgress.Max = rsParent!TotalRows
   rsParent.Close
'   rsParent.Open "SELECT ID FROM GeneOntologyCount", dbGenMAPP, adOpenStatic
   Set rsParent = dbGenMapp.OpenRecordset("SELECT ID FROM GeneOntologyCount")
   
'   recordCount = 0
   Do Until rsParent.EOF '++++++++++++++++++++++++++++++++++++++++++++++++++++ Each GO ID As Parent
'      recordCount = recordCount + 1                  'absolutePosition doesn't work with SQL Server
'      History , recordCount
      prgProgress.value = rsParent.AbsolutePosition
      lblProgressMax = rsParent!id
      DoEvents
      
      '=========================================================Get Count Of Parent GO ID In System
      parent = 0                                             'Number of distinct genes at this root
      count = -1                                     'Last element in the genes() array, zero-based
         '  Number of distinct genes at this root and children less 1
'      Set rs = dbGenMAPP.Execute( _
'               "SELECT [Primary]" & _
'               "   FROM [" & systemGo & "]" & _
'               "   WHERE Related = '" & rsParent!id & "'")                           'Genes at root
      Set rs = dbGenMapp.OpenRecordset( _
               "SELECT [Primary]" & _
               "   FROM [" & systemGo & "]" & _
               "   WHERE Related = '" & rsParent!id & "'")                           'Genes at root
                  '  Assume there are no duplicates at a particular node
      Do Until rs.EOF
         count = count + 1
         genes(count) = rs![primary]                            'Genes are distinct at parent level
         rs.MoveNext
      Loop
      
      parent = count + 1                                               'Because count is zero-based
      
      '=========================================================Add Count Of All Children To Parent
      CountChildren rsParent!id, systemGo, genes, count
      
      If count > -1 Then '========================================Put In GOCount Table If Any Exist
         dbGenMapp.Execute "INSERT INTO [" & GOCount & "] (GO, [Count], [Total])" & _
                           "   VALUES ('" & rsParent!id & "', " & parent & ", " & count + 1 & ")"
      End If
      
'Print #file_raw_data0, 0; " "; "-"; rsParent!ID; "  "; total
      rsParent.MoveNext
   Loop
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Get Overall Totals for Entire GO Tree
'   Set rs = dbGenMAPP.Execute( _
'         "SELECT COUNT(DISTINCT [Primary]) AS RecordCount FROM [" & systemGo & "]")
   Set rs = dbGenMapp.OpenRecordset("SELECT DISTINCT [Primary] FROM [" & systemGo & "]")
      '  COUNT(DISTINCT whatever) doesn't work with DAO
   rs.MoveLast
   l = rs.recordCount
   dbGenMapp.Execute "INSERT INTO [" & GOCount & "] (GO, [Count], [Total])" & _
                     "   VALUES ('GO', 0, " & rs.recordCount & ")"
   rs.Close
                     
'   If systemGo = "Temp" Then
'      dbGenMAPP.Execute "DROP TABLE Temp"
'   End If
   
'   SetProgressBase
   inProcess = ""
End Sub
Sub CountChildren(parentID As String, systemGo As String, genes() As String, count As Long)
   Dim rs As Recordset, rsChild As Recordset, l As Long
   Static level As Integer, total As Long
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Extra For GenMAPP Program
   Dim dbGenMapp As Database
   Set dbGenMapp = mappWindow.dbGene

   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select All Children Of Parent
'   Set rsChild = dbGenMapp.Execute( _
'            "SELECT ID" & _
'            "   FROM GeneOntology" & _
'            "   WHERE Parent = '" & parentID & "'")
   Set rsChild = dbGenMapp.OpenRecordset( _
            "SELECT ID" & _
            "   FROM GeneOntology" & _
            "   WHERE Parent = '" & parentID & "'")

   Do Until rsChild.EOF '+++++++++++++++++++++ Count All Occurrences Of This Child's Tree In System
      prgProgress.visible = False
      DoEvents
      prgProgress.visible = True
      DoEvents
      '========================================================================This Child as Parent
'      Set rs = dbGenMapp.Execute( _
'               "SELECT [Primary]" & _
'               "   FROM [" & systemGo & "]" & _
'               "   WHERE Related = '" & rsChild!id & "'")
      Set rs = dbGenMapp.OpenRecordset( _
               "SELECT [Primary]" & _
               "   FROM [" & systemGo & "]" & _
               "   WHERE Related = '" & rsChild!id & "'")
      Do Until rs.EOF
         For l = 0 To count                                               'Search for gene in array
            If rs![primary] = genes(l) Then Exit For
         Next l
         If l > count Then                                            'Gene not found, add to array
            count = l
            genes(count) = rs![primary]
         End If
         rs.MoveNext
      Loop

      '=========================================Add All Occurrences Of This Child's Child In System
      CountChildren rsChild!id, systemGo, genes, count
      rsChild.MoveNext
   Loop
End Sub

