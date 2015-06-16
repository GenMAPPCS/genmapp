VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDrafter 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "GenMAPP"
   ClientHeight    =   10530
   ClientLeft      =   135
   ClientTop       =   -2415
   ClientWidth     =   9120
   DrawWidth       =   10
   HelpContextID   =   1
   Icon            =   "Drafter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   9120
   Begin MSComctlLib.StatusBar sbrBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   10275
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   132
            MinWidth        =   1
            Key             =   "Instructions"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   132
            MinWidth        =   1
            Key             =   "Gene DB"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbTools 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgTools16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Gene"
            Object.ToolTipText     =   "Gene"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Label"
            Object.ToolTipText     =   "Label"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Lines"
            Object.ToolTipText     =   "Lines - Click to choose"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Solid"
                  Text            =   "Solid"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Broken"
                  Text            =   "Broken"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Arrow"
                  Text            =   "Arrow"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BrokenArrow"
                  Text            =   "Broken arrow"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Key             =   "DoubleArrow"
                  Text            =   "Double arrow"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Key             =   "BrokenDoubleArrow"
                  Text            =   "Broken double arrow"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "AnchoredLines"
            Object.ToolTipText     =   "Anchored lines - move with objects"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Key             =   "Solid"
                  Text            =   "Solid"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Broken"
                  Text            =   "Broken"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Arrow"
                  Text            =   "Arrow"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BrokenArrow"
                  Text            =   "Broken arrow"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DoubleArrow"
                  Text            =   "Double arrow"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BrokenDoubleArrow"
                  Text            =   "Broken double arrow"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Inhibitor"
            Object.ToolTipText     =   "Inhibition symbol"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Receptor"
            Object.ToolTipText     =   "Receptor"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LigandReceptorSq"
            Object.ToolTipText     =   "Ligand/Receptor - Click to choose"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LigandSq"
                  Text            =   "Ligand"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ReceptorSq"
                  Text            =   "Receptor"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Key             =   "LigandReceptorSq"
                  Text            =   "Both"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LigandReceptorRd"
            Object.ToolTipText     =   "Ligand/Receptor - Click to choose"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LigandRd"
                  Text            =   "Ligand"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ReceptorRd"
                  Text            =   "Receptor"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Key             =   "LigandReceptorRd"
                  Text            =   "Both"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Brace"
            Object.ToolTipText     =   "Brace"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rectangle"
            Object.ToolTipText     =   "Rectangle"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Oval"
            Object.ToolTipText     =   "Oval"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Arc"
            Object.ToolTipText     =   "Arc"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Objects"
            Object.ToolTipText     =   "Object Toolbox"
            ImageIndex      =   13
            Style           =   1
         EndProperty
      EndProperty
      Begin VB.CommandButton cmdColorSetsDown 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5940
         Picture         =   "Drafter.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdColorSetsOK 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5400
         Picture         =   "Drafter.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   60
         Width           =   315
      End
      Begin VB.TextBox txtColorSets 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3960
         TabIndex        =   9
         Text            =   "ColorSets"
         ToolTipText     =   "Click to choose Color Set(s)"
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox cmbZoom 
         Height          =   315
         ItemData        =   "Drafter.frx":0FDE
         Left            =   6360
         List            =   "Drafter.frx":0FFA
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "100%"
         ToolTipText     =   "Choose or enter zoom percent."
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.VScrollBar vsbDrafter 
      CausesValidation=   0   'False
      Height          =   1992
      Left            =   8220
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   192
   End
   Begin VB.HScrollBar hsbDrafter 
      CausesValidation=   0   'False
      Height          =   192
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2820
      Width           =   5652
   End
   Begin VB.PictureBox picDrafter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6432
      Left            =   0
      ScaleHeight     =   6405
      ScaleWidth      =   8745
      TabIndex        =   5
      Top             =   420
      Width           =   8772
      Begin VB.ListBox lstColorSets 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   4260
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         ToolTipText     =   "Press Ctrl to choose multiple color sets"
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   160
         Index           =   0
         Left            =   4200
         MousePointer    =   5  'Size
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   160
      End
      Begin VB.Shape shpSelected 
         BorderColor     =   &H00008000&
         BorderStyle     =   4  'Dash-Dot
         Height          =   612
         Left            =   240
         Top             =   180
         Visible         =   0   'False
         Width           =   1632
      End
      Begin VB.Shape shpSelect 
         BorderStyle     =   3  'Dot
         Height          =   492
         Left            =   240
         Top             =   900
         Visible         =   0   'False
         Width           =   1632
      End
   End
   Begin MSComDlg.CommonDialog dlgPrinter 
      Left            =   900
      Top             =   9060
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      ToPage          =   1
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   420
      Top             =   9060
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer Delay 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   9120
   End
   Begin MSComctlLib.ImageList imgTools16 
      Left            =   0
      Top             =   9600
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":102F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":13C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":1523
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":167D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":17D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":1931
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":1A8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":1BE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":1D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":1E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":1FF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":214D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Drafter.frx":22A7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblGenMAPPVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   192
      Left            =   0
      TabIndex        =   2
      Top             =   10320
      Width           =   36
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Begin VB.Menu mnuExportMAPP 
            Caption         =   "MAPP"
            Begin VB.Menu mnuBMP 
               Caption         =   "To BMP"
            End
            Begin VB.Menu mnuJPEG 
               Caption         =   "To JPEG"
            End
            Begin VB.Menu mnuHTML 
               Caption         =   "To HTML"
            End
         End
         Begin VB.Menu mnuMAPPSet 
            Caption         =   "MAPP Set"
         End
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuInfo 
         Caption         =   "MAPP &Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuMAPPFinder 
         Caption         =   "MAPPFinder"
      End
      Begin VB.Menu mnuMAPPBuilder 
         Caption         =   "MAPPBuilder"
      End
      Begin VB.Menu mnuObjects 
         Caption         =   "O&bject Toolbox"
      End
      Begin VB.Menu mnuConverter 
         Caption         =   "&Converter"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options ..."
      End
      Begin VB.Menu mnuUpdater 
         Caption         =   "&Updater"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuHorizAlign 
         Caption         =   "&Horizontally Align Objects"
         Enabled         =   0   'False
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuVertAlign 
         Caption         =   "&Vertically Align Objects"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSize 
         Caption         =   "Si&ze Genes"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBlock 
         Caption         =   "Bloc&k Genes"
         Enabled         =   0   'False
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuBoardSize 
         Caption         =   "Drafting &Board Size"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRedraw 
         Caption         =   "&Redraw"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom ..."
      End
      Begin VB.Menu mnuZoomToScreen 
         Caption         =   "Zoom To Screen"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Data"
      Begin VB.Menu mnuChoose 
         Caption         =   "Choose Expression &Dataset"
      End
      Begin VB.Menu mnuApply 
         Caption         =   "&Apply Expression Data"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuManager 
         Caption         =   "&Expression Dataset Manager"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGeneDB 
         Caption         =   "Choose &Gene Database"
      End
      Begin VB.Menu mnuGeneDBMgr 
         Caption         =   "Gene Database Manager"
      End
      Begin VB.Menu mnuGeneDBInfo 
         Caption         =   "Gene Database Information"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "Download Data from GenMAPP.org"
      End
      Begin VB.Menu mnuRemoveLocalGenes 
         Caption         =   "Remove Local Genes"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuGenMAPPHelp 
         Caption         =   "&GenMAPP Help"
      End
      Begin VB.Menu mnuAboutGenMAPP 
         Caption         =   "&About GenMAPP"
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmDrafter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim debugOn As Boolean

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\mdiDrafter Stuff
Public statusBarPlace As New Collection, statusBarSelect As New Collection
Public statusBarFinish As New Collection

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\frmDrafter stuff
'****************************************************************** Control And Communication Items
Public mouseIsDown As Boolean
Public dirty As Boolean                                                  'If change is made in MAPP
Dim cancelExit As Boolean                        'Sends message to QueryUnload if user cancels exit
Dim MAPPSaveError As Boolean                 'Set to true if MAPP save request could not be honored
Dim loading As Boolean

'********************************************************************************** MAPP Data Items
Public dbGene As Database                            'Gene Database for this instance of frmDrafter
Public dbExpression As Database                 'Expression Dataset for this instance of frmDrafter
Public rsColorSet As Recordset                           'Color Set for this instance of frmDrafter
Public mappName As String                                   'Path and name for current open MAPP DB
   '  These four items are all a MAPP needs to know about its underlying data. dbGene should
   '  always have some value. it is checked upon activation. dbExpression can be Nothing if an
   '  Expression Dataset has not been chosen. rsColorSet can be Nothing if a colorSet has not been
   '  chosen. When a MAPP database is opened, its data produces objects on the drawingBoard and
   '  then it is closed. When the MAPP is saved, the database is recreated from scratch.
Public objKey As Long                                  'Unique identifier for each object on a MAPP
   '  For a new MAPP, this starts at zero. Others it is the MAX from the MAPP database. It is
   '  incremented each time an object is placed.
Public resetColorSet As Boolean
            'The ExpressionData procedure gets a color set once for the complete rendering of a
            'color set on the Drafting Board. The only place in the program that does this is
            'frmDrafter!mnuApply. This variable is set to True there, forcing ExpressionData to
            'reread the color set at which time ExpressionData sets the variable to False.
Public displayGeneValues As Boolean                            'Show gene values on drafter if true

'*************************************************************************** MAPP Dimensional Items
Public boardWidth As Single, boardHeight As Single                          'Size of Drafting board
   '  This is the size of the 100% zoomed picDrafter in twips. Previously, these dimensions were
   '  kept by picDrafter.Width and .Height but with zooming and rounding off, these dimensions
   '  changed. Now, picDrafter.?? is always board?? * zoom.
Public zoom As Single                                                       'Zoom factor. 1 is 100%
Dim mouseX As Single           'Position of mouse after MouseUp. Used to locate mouse for Click and
Dim mouseY As Single           'DblClick events
Dim selectX As Single, selectY As Single                           'Beginning corner of select area
Dim prevX As Single, prevY As Single                    'Mouse position at previous MouseMove event
Public lineStarted As Boolean                             'True if beginning of line (red X) placed
Public XStart As Single, YStart As Single                            'For beginning of line objects

Public selectArea As New objSelectArea                    'Limits of select area in multiple select
Public activeObject As Object                                          'Object on frmDrafter in use
Public objLines As New Collection
Public objLumps As New Collection
Public newObject As Object                    'New object on frmObjects to be dropped on frmDrafter
Public info As New objInfo
Public legend As New objLegend
Public MAPPTitle As New objLump
Public selections As New Collection
Public drawTop As New Collection

Private Sub cmbZoom_Click()
   ZoomWindow cmbZoom.text
End Sub

Private Sub cmbZoom_LostFocus()
   ZoomWindow cmbZoom.text
End Sub

Rem \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Form Events
Private Sub Form_Load()
   Dim i As Integer
   Dim begCommand As Integer, endCommand As Integer, pipe As Integer

   loading = True
'mdiDrafter Stuff
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Status Bar Text
      statusBarPlace.Add "Click to place center of object", "Poly"
      statusBarSelect.Add _
            "Drag center target to move, edge targets to size or rotate", "Poly"
      statusBarPlace.Add "Click to place center of rectangle", "Rectangle"
      statusBarSelect.Add _
            "Drag center target to move, edge target to size, corner target to rotate", "Rectangle"
      statusBarPlace.Add "Click to place center of oval", "Oval"
      statusBarSelect.Add _
            "Drag center target to move, axis target to size, circle target to rotate", "Oval"
      statusBarPlace.Add "Click to place center of oval", "CellB"
      statusBarSelect.Add _
            "Drag center target to move, axis target to size, circle target to rotate", "CellB"
      statusBarPlace.Add "Click to place center of arc", "Arc"
      statusBarSelect.Add _
            "Drag center target to move, axis target to size, circle target to rotate", "Arc"
      statusBarPlace.Add "Click to place target at center of brace", "Brace"
      statusBarSelect.Add _
            "Drag center target to move, edge target to size, circle target to rotate.", "Brace"
      statusBarPlace.Add "Click to place center of label", "Label"
      statusBarSelect.Add "Drag center target to move", "Label"
      statusBarPlace.Add "Click to place center of gene", "Gene"
      statusBarSelect.Add "Drag center target to move, edge target to size", "Gene"
      statusBarPlace.Add "Click to place center of vesicle", "Vesicle"
      statusBarSelect.Add "Drag center target to move, edge target to size", "Vesicle"
      statusBarPlace.Add "Click to place center of protein", "ProteinA"
      statusBarSelect.Add "Drag center target to move, edge target to size", "ProteinA"
      statusBarPlace.Add "Click to place center of protein", "ProteinB"
      statusBarSelect.Add "Drag center target to move, edge target to size", "ProteinB"
      statusBarPlace.Add "Click to place center of ribosome", "Ribosome"
      statusBarSelect.Add "Drag center target to move", "Ribosome"
      statusBarPlace.Add "Click to place center of organelle", "OrganA"
      statusBarSelect.Add "Drag center target to move", "OrganA"
      statusBarPlace.Add "Click to place center of organelle", "OrganB"
      statusBarSelect.Add "Drag center target to move", "OrganB"
      statusBarPlace.Add "Click to place center of organelle", "OrganC"
      statusBarSelect.Add "Drag center target to move", "OrganC"
      statusBarPlace.Add "Click to place center of cell", "CellA"
      statusBarSelect.Add "Drag center target to move", "CellA"
'      statusBarPlace.Add "Click to place center of nucleus", "CellB"
'      statusBarSelect.Add "Drag center target to move, edge target to rotate", "CellB"
      statusBarPlace.Add "Click to place beginning of line", "Solid"
      statusBarFinish.Add "Click to place end of line", "Solid"
      statusBarSelect.Add "Click and drag targets to move ends", "Solid"
      statusBarPlace.Add "Click to place beginning of line", "Broken"
      statusBarFinish.Add "Click to place end of line", "Broken"
      statusBarSelect.Add "Click and drag targets to move ends", "Broken"
      statusBarPlace.Add "Click to place beginning of line", "DoubleArrow"
      statusBarFinish.Add "Click to place end of line", "DoubleArrow"
      statusBarSelect.Add "Click and drag targets to move ends", "DoubleArrow"
      statusBarPlace.Add "Click to place beginning of line", "BrokenDoubleArrow"
      statusBarFinish.Add "Click to place end of line", "BrokenDoubleArrow"
      statusBarSelect.Add "Click and drag targets to move ends", "BrokenDoubleArrow"
      statusBarPlace.Add "Click to place plain end of line", "Arrow"
      statusBarFinish.Add "Click to place arrow end", "Arrow"
      statusBarSelect.Add "Click and drag targets to move ends", "Arrow"
      statusBarPlace.Add "Click to place plain end of line", "BrokenArrow"
      statusBarFinish.Add "Click to place arrow end", "BrokenArrow"
      statusBarSelect.Add "Click and drag targets to move ends", "BrokenArrow"
      statusBarPlace.Add "Click to place plain end of line", "Receptor"
      statusBarFinish.Add "Click to place receptor end", "Receptor"
      statusBarSelect.Add "Click and drag targets to move ends", "Receptor"
      statusBarPlace.Add "Click to place plain end of line", "ReceptorSq"
      statusBarFinish.Add "Click to place receptor end", "ReceptorSq"
      statusBarSelect.Add "Click and drag targets to move ends", "ReceptorSq"
      statusBarPlace.Add "Click to place plain end of line", "ReceptorRd"
      statusBarFinish.Add "Click to place receptor end", "ReceptorRd"
      statusBarSelect.Add "Click and drag targets to move ends", "ReceptorRd"
      statusBarPlace.Add "Click to place plain end of line", "Inhibitor"
      statusBarFinish.Add "Click to place inhibition (T) end", "Inhibitor"
      statusBarSelect.Add "Click and drag targets to move ends", "Inhibitor"
'      statusBarPlace.Add "Click to place beginning of curve", "Arc"
'      statusBarFinish.Add "Click to place end of curve", "Arc"
'      statusBarSelect.Add "Click and drag targets to move ends", "Arc"
      statusBarPlace.Add "Click to place plain end of ligand", "LigandRd"
      statusBarFinish.Add "Click to place ligand end", "LigandRd"
      statusBarSelect.Add "Click and drag targets to move ends", "LigandRd"
      statusBarPlace.Add "Click to place plain end of ligand", "LigandSq"
      statusBarFinish.Add "Click to place ligand end", "LigandSq"
      statusBarSelect.Add "Click and drag targets to move ends", "LigandSq"
      statusBarSelect.Add "Click and drag targets to place upper-left corner", "InfoBox"
      statusBarSelect.Add "Click and drag targets to place upper-left corner", "Legend"
   
   Caption = PROGRAM_TITLE
   
   Dim btn As Button
   txtColorSets.Left = 0 '++++++++++++++++++++++++++++++++++++++++++++++++++++++ Tool Bar Positions
   For Each btn In tlbTools.Buttons
      txtColorSets.Left = txtColorSets.Left + btn.Width
   Next btn
   txtColorSets.Left = txtColorSets.Left + 20
   txtColorSets.Top = 40 '80
   txtColorSets.Height = 315
   cmdColorSetsDown.Left = txtColorSets.Left + txtColorSets.Width - cmdColorSetsDown.Width - 35
   cmdColorSetsDown.Top = txtColorSets.Top + 35
   cmdColorSetsDown.Height = txtColorSets.Height - 60
   cmdColorSetsOK.Left = cmdColorSetsDown.Left
   cmdColorSetsOK.Top = cmdColorSetsDown.Top
   cmdColorSetsOK.Height = cmdColorSetsDown.Height
   SetColorSetText
   cmbZoom.Left = txtColorSets.Left + txtColorSets.Width + 20
   cmbZoom.Top = 50
   
   Set mappWindow = Me '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ frmDrafter Stuff
   Set drawingBoard = picDrafter
   zoom = 1
   Load frmObjects

   For i = 1 To 3
   Load picPoint(i)
      picPoint(i).Line (0, 0)-(POINT_SIZE - 1, POINT_SIZE - 1), , B
      picPoint(i).Line (0, 0)-(POINT_SIZE - 1, POINT_SIZE - 1)
      picPoint(i).Line (POINT_SIZE - 1, 0)-(0, POINT_SIZE - 1)
   Next i
   Load picPoint(4)                                                                   'Rotate point
   picPoint(4).Circle (POINT_SIZE / 2, POINT_SIZE / 2), _
                      POINT_SIZE / 2 - 5, vbRed, 115 * PI / 180, 75 * PI / 180
   picPoint(4).Line (POINT_SIZE * 0.2, POINT_SIZE * 0.1)-Step(-60, 0)
   picPoint(4).Line (POINT_SIZE * 0.2, POINT_SIZE * 0.1)-Step(0, 60)
   Load picPoint(5)                                                                      'Arc point
   picPoint(5).Circle (POINT_SIZE / 2 - 0, 0), POINT_SIZE / 2 - 0, , , , 2

   Top = 0
   Left = 0
If debugOn Then mnuOpen_Click
   shpSelect.Width = 0                                    'Otherwise screws up MouseUp when loading
   shpSelect.Height = 0                              '(Why is there a MouseUp when loading anyway?)
   zoom = 1                                                          'Put before loading frmDrafter
   Set activeObject = Nothing                                        'Initialize objects in process
   Set newObject = Nothing
   With picDrafter                                              'Establish initial board parameters
      .Left = 0
      .Top = tlbTools.Height
      boardWidth = INITIAL_BOARD_WIDTH
      .Width = boardWidth
      boardHeight = INITIAL_BOARD_HEIGHT
      .Height = boardHeight
   End With
   FormWidth INITIAL_WINDOW_WIDTH                                        'Set visible size of board
   FormHeight INITIAL_WINDOW_HEIGHT
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Open Gene Database
   sbrBar.Panels("Gene DB").text = "No Gene Database"
   Set dbGene = Nothing                                                         'Default to Nothing
   If InStr(commandLine, ".gdb""") Then '=====================================React To Command Line
      endCommand = InStr(commandLine, ".gdb""") + 3                          'Doesn't include quote
      begCommand = InStrRev(commandLine, """", endCommand) + 1               'Doesn't include quote
      mruGeneDB = Mid(commandLine, begCommand, endCommand - begCommand + 1)
      If begCommand - 1 = 0 Then                                       'Must be first thing on line
         commandLine = Left(commandLine, begCommand - 2) & Mid(commandLine, endCommand + 2)
      Else                                                 'Not first, so must be preceded by space
         commandLine = Left(commandLine, begCommand - 3) & Mid(commandLine, endCommand + 2)
      End If
   End If
   OpenGeneDB dbGene, mruGeneDB, Me                             'May return Nothing if no mruGeneDB
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Expression Dataset And Color Set
   Set dbExpression = Nothing                                                   'Default to nothing
   Set rsColorSet = Nothing
   If InStr(commandLine, ".gex""") Then '=================Open Expression Dataset From Command Line
      '  Command line sets the mruDataSet and mruColorSet
      '  It also overrides cfgOptions in assuming that the Color Set should be shown
      endCommand = InStr(commandLine, ".gex""") + 3                          'Doesn't include quote
      begCommand = InStrRev(commandLine, """", endCommand) + 1               'Doesn't include quote
      mruDataSet = Mid(commandLine, begCommand, endCommand - begCommand + 1)
      If begCommand - 4 = 0 Then                                       'Must be first thing on line
         commandLine = Left(commandLine, begCommand - 2) & Mid(commandLine, endCommand + 2)
      Else                                                 'Not first, so must be preceded by space
         commandLine = Left(commandLine, begCommand - 3) & Mid(commandLine, endCommand + 2)
      End If
      If InStr(commandLine, """colors:") Then '-------------------Apply Color Set From Command Line
         If InStr(commandLine, """set:") = 0 Then                                'If not a MAPP set
            '  If it is a MAPP set, there are likely to be many colors or it will default to ALL,
            '  so that processing is done on frmMAPPSet
            begCommand = InStr(commandLine, """colors:") + 8
            endCommand = InStr(begCommand, commandLine, """") - 1            'Doesn't include quote
            s = Mid(commandLine, begCommand, endCommand - begCommand + 1)
            If Left(s, 1) = "|" Then s = Mid(s, 2)
            pipe = InStr(s, "|")
            If pipe = 0 Then pipe = Len(s) + 1
            mruColorSet = Left(s, pipe - 1)
            If InStr(mruColorSet, "\") = 0 And mruColorSet <> "" Then   'Old style, single colorset
               mruColorSet = mruColorSet & "\" & mruColorSet                 'DisplayValue\ColorSet
            End If
            If begCommand - 9 = 0 Then                                 'Must be first thing on line
               commandLine = Left(commandLine, begCommand - 9) & Mid(commandLine, endCommand + 2)
            Else                                           'Not first, so must be preceded by space
               commandLine = Left(commandLine, begCommand - 10) & Mid(commandLine, endCommand + 2)
            End If
         End If
      End If
      s = SetDataSet
   ElseIf InStr(cfgOptions, "C") Then '=================================Open MRU Expression Dataset
      s = SetDataSet
   Else '=====================================================================No Expression Dataset
      Set dbExpression = Nothing
      colorIndexes(0) = 0
      valueIndex = -1
      SetColorSetText
   End If
   
   legend.Create                                'If no Color Set chosen, then no legend will appear
   MAPPTitle.Create -1, -1, "MAPPTitle"
   dirty = False
   If InStr(commandLine, """set:") Then
      mnuMAPPSet_Click
   End If
   loading = False
End Sub
Private Sub Form_Activate()
   Set mappWindow = Me
   Set drawingBoard = picDrafter

'Dim r As Single, noOfPoints As Single, factor As Single, n As Single, sector As Integer
'factor = 1000
'noOfPoints = 6
'r = 370
'n = 0
'sector = -1
'Do While r >= n
'   n = n + 360 / noOfPoints
'   sector = sector + 1
'Loop
''r = r - (n - 360 / noOfPoints)
''r = r / 360 * 2 * PI
''Do While r > n
''   n = n + 2 * PI / noOfPoints
''Loop
''r = r - n - 2 * PI / noOfPoints
''r = r * 360 / (2 * PI)
'Stop

'Dim aX1 As Single, aY1 As Single, aX2 As Single, aY2 As Single, bX1 As Single, bY1 As Single, bX2 As Single, bY2 As Single, X As Single, Y As Single
'here:
'aX1 = 600: aY1 = 1000: aX2 = 200: aY2 = 8000
'bX1 = 1100: bY1 = 1500: bX2 = 5000: bY2 = 1800
'picDrafter.DrawWidth = 1
'picDrafter.Line (aX1, aY1)-(aX2, aY2), vbRed
'picDrafter.Line (bX1, bY1)-(bX2, bY2), vbGreen
'LineIntersection aX1, aY1, aX2, aY2, bX1, bY1, bX2, bY2, X, Y
'Debug.Print X; Y
'picDrafter.DrawWidth = 5
'picDrafter.PSet (X, Y), vbBlack
'Stop
'GoTo here

   If cfgInitialRun <> "False" Then
      Dim hWndHelp As Long
      'The return value is the window handle of the created help window.
      hWndHelp = HtmlHelp(hWnd, appPath & "GenMAPP.chm::/GenMAPPQuickStart.htm", _
                          HH_DISPLAY_TOPIC, 0)
      cfgInitialRun = "False"
   End If
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Choosing Color Set(s)
Public Function FillColorSetList(Optional frm As Form = Nothing)
   Dim rsColorSets As Recordset, i As Integer, multiple As Integer
   
   If frm Is Nothing Then Set frm = Me
   
   If dbExpression Is Nothing Then '++++++++++++++++++++++++++++++++++ No Expression Dataset Chosen
      txtColorSets = "No expression data"
      colorIndexes(0) = 0
      valueIndex = 0
      Set rsColorSet = Nothing
      legend.Create
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Have Expression Dataset
      Set rsColorSets = dbExpression.OpenRecordset("SELECT * FROM ColorSet ORDER BY SetNo")
      If Not rsColorSets.EOF Then '=======================================Have Color Sets Available
         If frm.name = "frmMultipleColorSets" Then '---------------Set Up Multiple Color Set Window
            With frm
               .lstColorSets.Clear
               .lstDisplayValue.Clear
               Do Until rsColorSets.EOF
                  .lstColorSets.AddItem rsColorSets!colorSet
                  .lstDisplayValue.AddItem rsColorSets!colorSet
                  rsColorSets.MoveNext
               Loop
               .lstDisplayValue.AddItem "Don't display value"
               For i = 1 To colorIndexes(0)
                  .lstColorSets.selected(colorIndexes(i)) = True
               Next i
               If valueIndex = -1 Then
                  .lstDisplayValue.selected(.lstDisplayValue.ListCount - 1) = True
               Else
                  .lstDisplayValue.selected(valueIndex) = True
               End If
            End With
         Else
            With lstColorSets '----------------------------------------Set Up Single Color Set List
               .Clear
               .AddItem "No expression data"
               rsColorSets.MoveLast
               If rsColorSets.recordCount > 1 Then
                  .AddItem "Multiple Color Sets"
                  multiple = 1
               End If
               rsColorSets.MoveFirst
               Do Until rsColorSets.EOF
                  .AddItem rsColorSets!colorSet
                  rsColorSets.MoveNext
               Loop
               If colorIndexes(0) = 0 Then                               'Select No Expression Data
                  lstColorSets.selected(0) = True
               ElseIf colorIndexes(0) > 1 Then                          'Select Multiple Color Sets
                  lstColorSets.selected(1) = True
               Else                                                        'Select Single Color Set
                  For i = 1 To colorIndexes(0)
                     lstColorSets.selected(colorIndexes(i) + 1 + multiple) = True
                                              'No expression data first and multiple might be there
                  Next i
               End If
            End With
         End If
      End If
   End If
   SetColorSetText
End Function
Private Sub cmdColorSetsDown_Click()
   If dbExpression Is Nothing Then
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
'   If cmdColorSetsOK.visible Then
'      cmdColorSetsOK_Click
'      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   End If
'
   FillColorSetList
   lstColorSets.Top = -picDrafter.Top + tlbTools.Height
   lstColorSets.Left = txtColorSets.Left - picDrafter.Left
   LstColorSetsVisible True
   lstColorSets.ZOrder
   cmdColorSetsDown.visible = False
'   cmdColorSetsOK.visible = True
'   cmdColorSetsOK.ZOrder
   lstColorSets.ToolTipText = "Color Set and value to display."
   txtColorSets = "Choose Color Set"
   DoEvents
End Sub
Private Sub cmdColorSetsOK_Click()
   Dim i As Integer
   Dim items(MAX_COLORSETS) As String, noOfItems As Integer
   Dim rsColorSets As Recordset, showAll As Integer
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Assemble Color Sets
   cmdColorSetsOK.visible = False
   colorIndexes(0) = 0
   mruColorSet = ""
   If lstColorSets.List(1) = "Show all" Then
      showAll = 1
   End If
   For i = showAll + 1 To lstColorSets.ListCount - 1                                    'Zero based
      If lstColorSets.selected(i) Then
         colorIndexes(0) = colorIndexes(0) + 1
         colorIndexes(colorIndexes(0)) = i - 1 - showAll     'Zero based. No expression & all first
         mruColorSet = mruColorSet & lstColorSets.List(i) & "\"
      End If
   Next i
   
   If colorIndexes(0) = 0 Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++ None Selected
'      txtColorSets = "No expression data"
      LstColorSetsVisible False
      mruColorSet = ""
      valueIndex = -1
      mnuApply_Click
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Many Selected
'      txtColorSets = ""
      lstColorSets.Clear
      lstColorSets.ToolTipText = "Click on Color Set whose value to display."
      noOfItems = SeparateValues(items, mruColorSet, "\")
      lstColorSets.AddItem "Don't display values"
      Set rsColorSets = dbExpression.OpenRecordset("SELECT * FROM ColorSet ORDER BY SetNo")
      Do Until rsColorSets.EOF
         lstColorSets.AddItem rsColorSets!colorSet
         rsColorSets.MoveNext
      Loop
      txtColorSets.Tag = "DisplayValues"
   End If
   SetColorSetText
End Sub
Private Sub lstColorSets_Click()
   Dim i As Integer, j As Integer, sql As String
   Dim items(MAX_COLORSETS) As String, noOfItems As Integer, multiple As Integer
   
   If loading Then Exit Sub                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   If lstColorSets.visible = False Then Exit Sub           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      'Just filling Color Set list
      
   If lstColorSets.List(1) = "Multiple Color Sets" Then multiple = 1
   
   If lstColorSets.selected(0) Then '+++++++++++++++++++++++++++++++++++++++++++ No Expression Data
      colorIndexes(0) = 0
      valueIndex = -1
      mruColorSet = ""
   ElseIf lstColorSets.selected(1) And lstColorSets.List(1) = "Multiple Color Sets" Then '+++++++++
      FillColorSetList frmMultipleColorSets
      With frmMultipleColorSets
         .show vbModal
         If .Tag <> "Cancel" Then
            colorIndexes(0) = 0
            mruColorSet = ""
            For i = 0 To .lstColorSets.ListCount - 1                                    'Zero based
               If .lstColorSets.selected(i) Then
                  colorIndexes(0) = colorIndexes(0) + 1
                  colorIndexes(colorIndexes(0)) = i                                     'Zero based
                  mruColorSet = mruColorSet & .lstColorSets.List(i) & "\"
               End If
            Next i
            If .lstDisplayValue = "Don't display value" Then
               mruColorSet = "\" & mruColorSet
               valueIndex = -1
            Else
               mruColorSet = .lstDisplayValue & "\" & mruColorSet
               For i = 0 To .lstDisplayValue.ListCount - 1
                  If .lstDisplayValue.selected(i) Then
                     valueIndex = i
                     Exit For                              'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                  End If
               Next i
            End If
         End If
      End With
   Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Single Color Set
      colorIndexes(0) = 1
      For i = multiple + 1 To lstColorSets.ListCount - 1                                'Zero based
         '  0 is No Expression Data. 1, if Multiple Color Sets, caught above
         If lstColorSets.selected(i) Then
            colorIndexes(1) = i - 1 - multiple                     'Zero based. No & multiple first
            valueIndex = colorIndexes(1)
            mruColorSet = lstColorSets.List(i) & "\" & lstColorSets.List(i) & "\"
            Exit For                                       'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      Next i
   End If
   
   LstColorSetsVisible False
   txtColorSets.Tag = ""
   SetColorSetText
   mnuApply_Click
   mnuApply.Enabled = True
   picDrafter.SetFocus
'Exit Sub
'
'
'   If txtColorSets.Tag = "DisplayValues" Then '++++++++++++++++ Show All Sets, Choose Display Value
'      If lstColorSets.selected(0) Then '========================================Don't Display Value
'         valueIndex = -1
'         mruColorSet = "\" & mruColorSet                              'Add nothing to front of list
'      Else '===================================================================Choose Display Value
'         For i = 1 To lstColorSets.ListCount - 1                    'Zero based, 0 is Don't display
'            If lstColorSets.selected(i) Then                                      'Chosen Color Set
'               valueIndex = i - 1
'               mruColorSet = lstColorSets.List(i) & "\" & mruColorSet   'Add value to front of list
'               Exit For                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'            End If
'         Next i
'      End If
'      LstColorSetsVisible False
'      txtColorSets.Tag = ""
'      SetColorSetText
'      mnuApply_Click
'      mnuApply.Enabled = True
'      picDrafter.SetFocus
'   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Show All Sets With Current Chosen
'      If lstColorSets.selected(0) Then '=========================================No expression data
'         For i = 1 To lstColorSets.ListCount - 1                         'Wipe out other selections
'            lstColorSets.selected(i) = False
'         Next i
'      ElseIf lstColorSets.selected(1) And lstColorSets.List(1) = "Show all" Then '=========Show All
'         lstColorSets.selected(1) = False                                        'Deselect Show All
'         For i = 2 To lstColorSets.ListCount - 1                            'Select everything else
'            lstColorSets.selected(i) = True
'         Next i
'      End If
'   End If
''   cmdColorSetsDown.visible = True '==============================================Set For Selection
End Sub
Sub LstColorSetsVisible(visible As Boolean)
   '  When lstColorSets is visible, disable other controls because they would move the list box
   '  relative to the toolbar
   lstColorSets.visible = visible
   vsbDrafter.Enabled = Not visible
   hsbDrafter.Enabled = Not visible
   cmbZoom.Enabled = Not visible
End Sub
Sub SetColorSetText()
   Dim rs As Recordset
   
   If dbExpression Is Nothing Then '+++++++++++++++++++++++++++++++++++++++++ No Expression Dataset
      txtColorSets = "No Expression Dataset"
      cmdColorSetsDown.visible = False
      cmdColorSetsOK.visible = False
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Expression Dataset Exists
      Set rs = dbExpression.OpenRecordset("SELECT Count(*) AS TotalColorSets FROM ColorSet")
      If rs!totalcolorsets = 0 Then '===========================================No Color Sets Exist
         txtColorSets = "No expression data"
         cmdColorSetsDown.visible = False
         cmdColorSetsOK.visible = False
      ElseIf colorIndexes(0) = 0 Then '============================Color Sets Exist But None Chosen
         txtColorSets = "No expression data"
         cmdColorSetsDown.visible = True
      Else '====================================================================At Least One Chosen
         If txtColorSets.Tag = "DisplayValues" Then '--------------------------Choose Display Value
            cmdColorSetsDown.visible = False
            txtColorSets = "Choose display value"
         Else '----------------------------------------------------------Selection Process Finished
            cmdColorSetsDown.visible = True
            If colorIndexes(0) = 1 Then '...................................Assign Single Color Set
               Set rs = dbExpression.OpenRecordset( _
                   "SELECT ColorSet FROM ColorSet WHERE SetNo = " & colorIndexes(1))
               txtColorSets = rs!colorSet
'               txtColorSets = lstColorSets.List(colorIndexes(1) + 1 - (rs!totalcolorsets > 1))
            Else
               If colorIndexes(0) = rs!totalcolorsets Then '....................................All
                  txtColorSets = "All Color Sets"
               Else '..........................................................................Many
                  txtColorSets = "Multiple Color Sets"
               End If
            End If
         End If
      End If
   End If
End Sub
Private Sub txtColorSets_Click()
   picDrafter.SetFocus
End Sub


Private Sub mnuConverter_Click()
   If Dir(appPath & "GenMAPPConvert.exe") <> "" Then
      WriteConfig                              'Convert also uses config, so make sure it's current
      Shell """" & appPath & "GenMAPPConvert.exe""", vbNormalFocus
   Else
      MsgBox "GenMAPPConvert.exe not available. Download from GenMAPP.org", _
             vbInformation + vbOKOnly, "Converter"
   End If

End Sub

Private Sub mnuDownload_Click()
   Dim ptr As Long
   InvokeFullDBDL hWnd, appPath
'   InvokeFullDBDL hwnd, ptr
End Sub

Public Sub mnuGeneDBInfo_Click()
   Dim frm As Form, formFound As Boolean
   
'   For Each frm In Forms
'      If frm.name = "frmSystems" Then
'         frm.WindowState = vbNormal
'         formFound = True
'         Exit For
'      End If
'   Next frm
'   If Not formFound Then
      frmSystems.show vbModal
'   End If
'   frmSystems.SetFocus
End Sub

Private Sub mnuBMP_Click()
   Dim file As String
   
   If mappName <> "" Then
      file = Mid(mappName, InStrRev(mappName, "\") + 1)
   End If
   file = SetFileName("BMP", file, mruExportPath)
   If file = "ERROR" Or file = "CANCEL" Then Exit Sub      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   mruExportPath = Left(file, InStrRev(file, "\"))                   'Change configured export path
   
   SavePicture picDrafter.Image, file
End Sub
Private Sub mnuJPEG_Click()
   Dim file As String
   
   If mappName <> "" Then
      file = Mid(mappName, InStrRev(mappName, "\") + 1)
   End If
   file = SetFileName("JPEG", file, mruExportPath)
   If file = "ERROR" Or file = "CANCEL" Then Exit Sub      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   mruExportPath = Left(file, InStrRev(file, "\"))                   'Change configured export path
   
   CreateJPEG file
End Sub
Sub CreateJPEG(file As String)
   Dim dib As New cDIBSection
   Dim dot As Integer
   Dim Pic As StdPicture
   
   dot = InStrRev(file, ".")
   If dot = 0 Then
      dot = Len(file) + 1
   End If
   file = Trim(Left(file, dot - 1)) & Mid(file, dot)
      '  Intel's SaveJPG() crashes if the file name ends in space, eg. "file .jpg"
   MousePointer = vbHourglass
   DoEvents
'   dib.CreateFromPicture CaptureClient(picDrafter.Picture) 'Picture object works as well as a StdPicture object
'   dib.CreateFromPicture CaptureWindow(picDrafter.hwnd, False, 0, 0, picDrafter.Width / Screen.TwipsPerPixelX, picDrafter.Height / Screen.TwipsPerPixelY)
   '  This is unbelievably clumsy. There has to be a better way. CreateFromPicture needs a StdPicture
   '  object to create the device-independent bitmap from which to write the JPEG file. This was the
   '  only way I could find to convert the contents of the picDrafter control to a StdPicture object.
   SavePicture picDrafter.Image, appPath & "tempbmp.$tm"
   Set Pic = LoadPicture(appPath & "tempbmp.$tm")
   Kill appPath & "tempbmp.$tm"
   dib.CreateFromPicture Pic
   SaveJPG dib, file
   MousePointer = vbDefault
ExitSub:
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
ErrorHandler:
   Select Case Err.number
   End Select
End Sub
Private Sub mnuHTML_Click()
   Dim file As String
   
   If mappName <> "" Then
      file = Mid(mappName, InStrRev(mappName, "\") + 1)
   End If
   file = SetFileName("HTML", file, mruExportPath)
   If file = "ERROR" Or file = "CANCEL" Then Exit Sub      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   file = Left(file, InStrRev(file, ".") - 1)                                            'Dump .htm
   mruExportPath = Left(file, InStrRev(file, "\"))                   'Change configured export path
   HTMLExport mruExportPath, Mid(file, InStrRev(file, "\") + 1)
End Sub
Public Sub HTMLExport(ByVal folder As String, ByVal mapp As String) '******** Create Web Site Pages
   '  Enter:   folder   Folder of main HTML page. Backpages will be in _Support/
   '                    Eg. C:\GenMAPP\Exports\MyMapp\
   '           mapp     Name of MAPP file without extension. HTML page a ValidHTMLName version
   '                    Eg. MyMapp
   '  Process  Completed MAPP in C:\GenMAPP\Exports\MyMapp\MyMAPP.htm
   '           JPEG graphics and HTML backpages in C:\GenMAPP\Exports\MyMapp\_Support\
   Dim JPEGfile As String
   Dim supportPath As String                            'Path for support files: JPGs and backpages
                                                        'Name of MAPP & /_Support/
   Dim element As Object
   Dim backpageFile As String, adjust As Single, file As String, backpageHead As String
   
   MousePointer = vbHourglass
'   ValidHTMLName folder, False
'   ValidHTMLName mapp, False
'   mapp = Left(file, InStrRev(file, ".") - 1)
   
   file = folder & mapp & htmlSuffix & ".htm"
   If Len(file) >= 260 Then
      '  At least the Dir function cannot handle paths longer than 259 characters. It creates
      '  a "File not found" error.
      MsgBox "The length of the destination path" & vbCrLf & vbCrLf & file & vbCrLf & vbCrLf _
             & "is too long for Windows to handle. To fix the problem, either shorten the " _
             & "folder name (C:\GenMAPP 2 Data\Whatever . . .) or the source MAPP name (My " _
             & "MAPP.mapp). You will have to run the MAPP Set again.", _
             vbCritical + vbOKOnly, "MAPP Path Error"
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   If creatingMappSet And Dir(file) <> "" Then '+++++++++++++++++++++++++++++++++++ Check Overwrite
      '  Must be creating MAPP Set for frmMAPPSet to be valid
      If frmMAPPSet.chkOverwrite = vbUnchecked Then
         If MsgBox("Overwrite existing " & file & "? ""No"" means that the existing file " _
                   & "will be part of the MAPP Set and no new one will be created.", _
                   vbQuestion + vbYesNo, "Creating MAPP Set File") = vbNo Then
            GoTo ExitSub                                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      End If
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up HTML File
   If creatingMappSet Then
      frmMAPPSet.lblDetail = "Creating MAPP"
      DoEvents
   End If
   If Dir(folder, vbDirectory) = "" Then
      AddFolder folder
   End If
   If Dir(folder & "_Support", vbDirectory) = "" Then
      MkDir folder & "_Support"
   End If
'   If Dir(folder & "_Support\" & mapp, vbDirectory) = "" Then
'      MkDir folder & "_Support\" & mapp
'   End If
'   supportPath = "_Support/" & mapp & "/"                        '/ is Standard HTML path delimiter
   supportPath = "_Support/"                                     '/ is Standard HTML path delimiter
   Open file For Output As #30
   Print #30, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
   Print #30, "<html>"
   Print #30, "<head>"
   Print #30, "   <title>" & frmMAPPInfo.txtTitle & "</title>"
   Print #30, "   <meta name=""generator"" content=""GenMAPP 2.1"">"
   Print #30, "</head>"
   Print #30, ""
   Print #30, "<body>"
   Print #30, colorSetHTML
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Insert MAPP Graphic
   JPEGfile = folder & "_Support\" & mapp & htmlSuffix & ".jpg"
'   JPEGfile = folder & "_Support\" & mapp & "\" & mapp & htmlSuffix & ".jpg"
   CreateJPEG JPEGfile
   JPEGfile = Mid(JPEGfile, InStrRev(JPEGfile, "\") + 1)
   Print #30, "<img src=""_Support/" & JPEGfile & """ alt = """ & JPEGfile _
              & """ usemap=""#MAPP"" border=0>"
   Print #30, "<map name=""MAPP"">"
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Create Backpages And Image Map
   adjust = zoom / Screen.TwipsPerPixelX
   For Each element In mappWindow.objLumps '--------------------------------------------Each Object
      If element.head <> "" Then
         backpageHead = element.head
      Else
         backpageHead = element.title
      End If

      If creatingMappSet Then
         frmMAPPSet.lblDetail = backpageHead
         DoEvents
         If Not creatingMappSet Then                                            'MAPP Set Cancelled
            GoTo EndFile                                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      End If
      Select Case element.objType
      Case "Gene"
'         backpageFile = CreateBackpage(element.id, element.SystemCode, backpageHead, _
                                       dbGene, dbExpression, _
                                       element, folder & "_Support\" & mapp & "\")
         backpageFile = CreateBackpage(element.id, element.systemCode, backpageHead, _
                                       dbGene, dbExpression, _
                                       element, folder & "_Support\")
         If backpageFile <> "" Then                                  'Backpage successfully created
            backpageFile = Mid(backpageFile, InStrRev(backpageFile, "\") + 1)    'Filename, no path
            Print #30, "   <area href=""_Support/" & backpageFile & """"
            Print #30, "         shape=""rect"""
            Print #30, "         coords=""" & Int((element.centerX - element.wide / 2) * adjust) _
                       & "," & Int((element.centerY - element.high / 2) * adjust) & "," _
                       & Int((element.centerX + element.wide / 2) * adjust) & "," _
                       & Int((element.centerY + element.high / 2) * adjust) & """"
            Print #30, "         alt=""Click for backpage, shift-click for separate window"">"
         End If
      Case "Label"
         backpageFile = CreateObjPage(element, folder & "_Support\")
         If backpageFile <> "" Then                                  'Backpage successfully created
            backpageFile = Mid(backpageFile, InStrRev(backpageFile, "\") + 1)    'Filename, no path
            Print #30, "   <area href=""_Support/" & backpageFile & """"
            Print #30, "         shape=""rect"""
            Print #30, "         coords=""" & Int((element.centerX - element.wide / 2) * adjust) _
                       & "," & Int((element.centerY - element.high / 2) * adjust) & "," _
                       & Int((element.centerX + element.wide / 2) * adjust) & "," _
                       & Int((element.centerY + element.high / 2) * adjust) & """"
            Print #30, "         alt=""Click for backpage, shift-click for separate window"">"
         End If
      End Select
   Next element
   If creatingMappSet Then
      frmMAPPSet.lblDetail = ""
      DoEvents
      If Not creatingMappSet Then                                               'MAPP Set Cancelled
         GoTo EndFile                                      'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
   End If
   
EndFile: '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ End HTML file
   Print #30, "</map>"
   Print #30, "</body>"
   Print #30, "</html>"
   Close #30
   
ExitSub:
   MousePointer = vbDefault
End Sub

Private Sub mnuGeneDBMgr_Click()
   frmGeneDBMgr.show vbModal
End Sub

Private Sub mnuMAPPBuilder_Click()
'   Dim file As String, fileDate As Date
   
'   fileDate = CDate("1-JAN-1970")
'   s = Dir(appPath & "MAPPBuilder*.exe")
'   Do Until s = ""
'      If FileDateTime(appPath & s) > fileDate Then
'         file = s
'         fileDate = FileDateTime(appPath & s)
'      End If
'      s = Dir
'   Loop
   If Dir(appPath & "MAPPBuilder.exe") <> "" Then
      WriteConfig                         'MappBuilder also uses config, so make sure it's current
      Shell """" & appPath & "MAPPBuilder.exe"" """ & mruGeneDB & """", vbNormalFocus
   Else
      MsgBox "MAPPBuilder not available. Download from GenMAPP.org", vbInformation + vbOKOnly, _
      "MAPPBuilder"
   End If
End Sub

Private Sub mnuMAPPFinder_Click()
'   Dim file As String, fileDate As Date
   
'   fileDate = CDate("1-JAN-1970")                          'Find most current version of MAPPFinder
'               'Should not need this in final ????????????????????????
'   s = Dir(appPath & "MAPPFinder*.exe")
'   Do Until s = ""
'      If FileDateTime(appPath & s) > fileDate Then
'         file = s
'         fileDate = FileDateTime(appPath & s)
'      End If
'      s = Dir()
'   Loop
   If Dir(appPath & "MAPPFinder.exe") <> "" Then
      WriteConfig                           'MappFinder also uses config, so make sure it's current
      Shell """" & appPath & "MAPPFinder.exe"" """ & mruGeneDB & """", vbNormalFocus
   Else
      MsgBox "MAPPFinder not available. Download from GenMAPP.org", vbInformation + vbOKOnly, _
             "MAPPFinder"
   End If
End Sub

Private Sub mnuObjects_Click()
   Dim mapp As String
   
   If mappName = "" Then
      frmObjects.lblTitle = ""
   Else
      mapp = GetFile(mappName)
      frmObjects.lblTitle = Left(mapp, InStrRev(mapp, ".") - 1)
   End If
   frmObjects.show
End Sub

Private Sub mnuOptions_Click()
   frmOptions.show vbModal
   mnuRedraw_Click
End Sub

Private Sub mnuTest_Click()
'Dim rs As Recordset
'        Dim genes As Integer
'        Dim geneIDs(MAX_GENES, 2) As String
'        Set rs = dbGene.OpenRecordset("SELECT * FROM Systems WHERE SystemCode = 'I'")
'        AllRelatedGenes "P29360", "S", dbGene, genes, geneIDs
'        AllRelatedGenes "AA034714", "G", dbGene, genes, geneIDs
'   MousePointer = vbHourglass
'        s = CreateBackpage("P29360", "S", "The Backpage", dbGene, dbExpression)
'        s = CreateBackpage("AB000095", "G", "The Backpage", dbGene, dbExpression)
'   MousePointer = vbDefault

'GeneData "AA034714", "G", PURPOSE_BACKPAGE, dbGene, s, dbExpression

'CreateDisplayTable dbExpression

picDrafter.Top = -100
picDrafter.Left = -100
'picDrafter.Width = 16000
'picDrafter.Height = 16000
'picDrafter.DrawWidth = 5
'picDrafter.Line (500, 500)-(14000, 14000), vbBlue, BF
'
'   Dim dib As New cDIBSection, Pic As StdPicture
''   dib.CreateDIB picDrafter.hdc, 12000, 12000, dib.hDib
'   Set Pic = picDrafter.Image
'   dib.CreateFromPicture CaptureClient(picDrafter.Picture)     'Picture object works as well as a StdPicture object
'   SaveJPG dib, "D:\GenMAPP V2\Programs\TEST.jpg"
'
End Sub


Private Sub mnuUpdater_Click()
   InvokeUpdate hWnd, appPath
End Sub

Private Sub mnuZoom_Click()
   z = InputBox("Current image: " & Format(zoom * 100, "0") & "%" & vbCrLf & vbCrLf _
                & "Change to (10 to 250%): ", "Zoom", Format(zoom * 100, "0"))
   If z <> "" Then
'      cmbzoom.Text=val(z)
      ZoomWindow Val(z)
   End If
End Sub

Private Sub mnuZoomToScreen_Click()
   Dim wideZoom As Single, highZoom As Single
   Dim scrWidth As Single, scrHeight As Single
   
'   ZoomWindow "To Screen"
'   Exit Sub
   
   DesktopClientArea scrWidth, scrHeight                             'Only client area, not taskbar
   
   'Figure constraining limits of board and screen
   wideZoom = (scrWidth - vsbDrafter.Width - (frmDrafter.Width - frmDrafter.ScaleWidth)) / boardWidth
      '  frmDrafter.Width - frmDrafter.ScaleWidth is then width of the Drafter window borders
      '  (scrWidth - vsbDrafter.Width - (frmDrafter.Width - frmDrafter.ScaleWidth))
      '  is the available width of screen for picDrafter.
   highZoom = (scrHeight - hsbDrafter.Height - tlbTools.Height - sbrBar.Height - (frmDrafter.Height - frmDrafter.ScaleHeight)) / boardHeight
   If wideZoom < highZoom Then
      ZoomWindow wideZoom * 100
   Else
      ZoomWindow highZoom * 100
   End If
   Top = 0
   Left = 0
   FormWidth picDrafter.Width
   FormHeight picDrafter.Height
End Sub
Sub ZoomWindow(percent As Variant)
   Static inFunction As Boolean
   
   If percent = "To Screen" Then
      mnuZoomToScreen_Click
      inFunction = False
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If inFunction Then Exit Sub                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   inFunction = True                                'To prevent recursion when cmbZoom.Text changes
   
   MultipleObjectDeselectAll
   If Right(percent, 1) = "%" Then
      percent = Left(percent, Len(percent) - 1)
   End If
   If IsNumeric(percent) Then '++++++++++++++++++++++++++++++++++++++++++++++++++ Valid Zoom Number
      If percent < 10 Then
         percent = 10
      ElseIf percent > 250 Then
         percent = 250
      End If
      zoom = percent / 100
      If boardWidth * zoom < MIN_BOARD_WIDTH Then '=====================Less Than Min Drawing Board
         '  If zooming down would make the drawing board less than its minimum size, zoom only to
         '  minimum size
         zoom = MIN_BOARD_WIDTH / boardWidth
      End If
      If boardHeight * zoom < MIN_BOARD_HEIGHT Then
         zoom = MIN_BOARD_HEIGHT / boardHeight
      End If
      If boardWidth * zoom > MAX_BOARD_WIDTH Then '=====================More Than Max Drawing Board
         zoom = MAX_BOARD_WIDTH / boardWidth
      End If
      If boardHeight * zoom > MAX_BOARD_HEIGHT Then
         zoom = MAX_BOARD_HEIGHT / boardHeight
      End If
      picDrafter.Width = boardWidth * zoom
      picDrafter.Height = boardHeight * zoom
      FitWindowToBoard
      Form_Resize
      mnuRedraw_Click
   End If
   percent = zoom * 100
   cmbZoom.text = Round(percent) & "%"
   inFunction = False
End Sub
Public Sub FitWindowToBoard() '*********************** Ensures That Window Does Not Go Beyond Board
   Dim resize As Boolean
   
   callingRoutine = "Don't resize"
   If ClientWide() > picDrafter.Width + picDrafter.Left Then          'Window > drafting board edge
      '  If the size of the drafting board is reduced so that there is client window beyond the
      '  right, bottom then the client window is resized.
      WindowState = vbNormal                                             'Can't resize if maximized
      If ClientWide() < picDrafter.Width Then            'Drafting board will fit in current window
         picDrafter.Left = picDrafter.Width - ClientWide()       'Move drafting board to fit window
      Else
         picDrafter.Left = 0                                         'Start board at left of window
         FormWidth picDrafter.Width + picDrafter.Left                   'Reduce window to fit board
      End If
      resize = True
   End If
   DoEvents
   If ClientHigh() > picDrafter.Height + tlbTools.Height - picDrafter.Top Then
      '  The top of the visible picDrafter viewport is below the tool bar (tlbTools).
      WindowState = vbNormal                                             'Can't resize if maximized                                                  'In case maximized
      If ClientHigh() < picDrafter.Height Then           'Drafting board will fit in current window
         picDrafter.Top = picDrafter.Height - ClientHigh()       'Move drafting board to fit window
      Else
         picDrafter.Top = 0                                           'Start board at top of window
         FormHeight picDrafter.Height + picDrafter.Top - tlbTools.Height
      End If
      resize = True
   End If
   callingRoutine = ""
   If resize Then Form_Resize
End Sub

'//////////////////////////////////////////////////////////////////////////////// picDrafter Events
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Mouse Events
Private Sub picDrafter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mouseIsDown = True
   If selections.count Or MousePointer = vbCrosshair Then
'      If selections.Count Then MousePointer = vbSizeAll
      prevX = GridCoord(X)
      prevY = GridCoord(Y)
   Else
      shpSelect.visible = True
      selectX = X   'Origin of selection
      selectY = Y
      shpSelect.Left = X
      shpSelect.Top = Y
      shpSelect.Width = 0
      shpSelect.Height = 0
   End If
End Sub
Private Sub picDrafter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim element As Object, moveX As Single, moveY As Single

   If Not mouseIsDown Then
      If shpSelected.visible = True Then                            'See if over multiple selection
         moveX = X
         moveY = Y
         If X >= shpSelected.Left And X <= shpSelected.Left + shpSelected.Width And Y >= shpSelected.Top And Y <= shpSelected.Top + shpSelected.Height Then
            MousePointer = vbSizeAll
         Else
            MousePointer = vbDefault
         End If
      End If
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If

   If shpSelect.visible Then '-------------------------------------------Selecting Multiple Objects
      If X < selectX Then
         shpSelect.Left = X
         shpSelect.Width = selectX - X
      Else
         shpSelect.Left = selectX
         shpSelect.Width = X - selectX
      End If
      If Y < selectY Then
         shpSelect.Top = Y
         shpSelect.Height = selectY - Y
      Else
         shpSelect.Top = selectY
         shpSelect.Height = Y - selectY
      End If
   ElseIf shpSelected.visible Or MousePointer = vbCrosshair Then '------Mult Selections Or Drop Obj
               '  Constrain both of these actions to grid intersections
      If Abs(X - prevX) > GRID_SIZE / zoom Or Abs(Y - prevY) > GRID_SIZE / zoom Then
         '  Move in multiples of GRID_SIZE
         moveX = GridCoord(X) - prevX
         moveY = GridCoord(Y) - prevY
         If MousePointer = vbSizeAll Then
            shpSelected.Left = shpSelected.Left + moveX
            shpSelected.Top = shpSelected.Top + moveY
         End If
         prevX = GridCoord(X)
         prevY = GridCoord(Y)
      End If
      If MousePointer = vbCrosshair Or MousePointer = vbSizeAll Then
         '  Any other mouse movement is either not dropping a new object or outside the
         '  selected area so treat it as just a mouse click
         dontClick = True
      End If
   End If
End Sub
Private Sub picDrafter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '  Set mousepointer to default only if this is the last or only action. Eg. lines require a
   '     second action so mouse pointer should be left as crosshair after first action.
   Dim element As Object

'Debug.Print "form mouseup event"
   mouseIsDown = False
   shpSelect.visible = False
'   selecting = False
   mouseX = X
   mouseY = Y
   If picDrafter.MousePointer = vbCrosshair Then '+++++++++++++++++++++++++++++ Dropping New Object
      mouseX = GridCoord(mouseX)
      mouseY = GridCoord(mouseY)
      SetActiveObject Nothing                                         'Deactivate any active object
      Select Case frmDrafter.tlbTools.Tag
      Case "Solid", "Broken", "Arrow", "BrokenArrow", "DoubleArrow", _
           "BrokenDoubleArrow", "Inhibitor", "Receptor", "LigandSq", "LigandRd", "ReceptorRd", _
           "ReceptorSq" '=====================================================================Lines
         If lineStarted Then '___________________________________________________________Place Line
            Dim newLine As New objLine
            picDrafter.DrawWidth = 1
            picDrafter.foreColor = vbWhite                                             'Erase red +
            picDrafter.Line (XStart - 50, YStart)-Step(100, 0)
            picDrafter.Line (XStart, YStart - 50)-Step(0, 100)
            sbrBar.Panels("Instructions").text = ""
            If frmDrafter.tlbTools.Tag = "Arc" Then     'Align second point on 90 degree increments
               If Abs(mouseX - XStart) < Abs(mouseY - YStart) Then
                  mouseX = XStart
               Else
                  mouseY = YStart
               End If
            End If
'            objKey = objKey + 1
            newLine.Create XStart / zoom, YStart / zoom, X / zoom, Y / zoom, _
                           frmDrafter.tlbTools.Tag
            objLines.Add newLine, newLine.objKey
            Set newLine = Nothing                           'Only reference should be in collection
            picDrafter.MousePointer = vbDefault
'            SetNewObject Nothing
'            mnuRedraw_Click
            ToolBarClear
         Else '__________________________________________________________________________Start Line
            XStart = mouseX
            YStart = mouseY
            picDrafter.DrawWidth = 1
            picDrafter.foreColor = vbRed                                                'Draw red +
            picDrafter.Line (XStart - 50, YStart)-Step(100, 0)
            picDrafter.Line (XStart, YStart - 50)-Step(0, 100)
            sbrBar.Panels("Instructions").text = statusBarFinish(frmDrafter.tlbTools.Tag)
            lineStarted = True
         End If
      Case Else '=====================================================================Other Objects
         Dim newLump As New objLump
         If Left(tlbTools.Tag, 4) = "Poly" Then
            newLump.Create X / zoom, Y / zoom, Left(tlbTools.Tag, 4), , , , , Mid(tlbTools.Tag, 5)
         Else
            newLump.Create X / zoom, Y / zoom, tlbTools.Tag
         End If
         objLumps.Add newLump, newLump.objKey
'HitRange newLump
         Set newLump = Nothing                              'Only reference should be in collection
         picDrafter.MousePointer = vbDefault
         ToolBarClear
         frmObjects.DeselectAll
      End Select
      dontClick = True                                                  'Don't treat as click event
   ElseIf Button = vbRightButton Then '+++++++++++++++++++++++++++++++++++++++++Handle Right-Clicks
      Dim obj As Object, color As Long

      MousePointer = vbHourglass
      MultipleObjectDeselectAll
      Set obj = ObjClicked
      If Not obj Is Nothing Then
         Select Case obj.objType
         Case "Rectangle", "Oval", "Brace", "objBrace", "Poly" ', "Arc"
            color = PickColor
            DoEvents                               'To show mouse pointer as hourglass after dialog
            If color >= 0 Then
               obj.color(1) = color
               obj.DrawObj
               If obj.color(1) = -1 Then mnuRedraw_Click
            End If
         Case "Label"
            SetActiveObject obj
'            frmbackdata.Tag="Label"
            frmObjData.show vbModal
'            frmLabelData.show vbModal
            obj.DrawObj
            SetActiveObject Nothing
         Case "objLine", "Arc"
            SetActiveObject obj
            frmLineData.show vbModal
            obj.DrawObj
            SetActiveObject Nothing
         Case "Gene"
            SetActiveObject obj
            Set frmGeneFinder.obj = obj
           frmGeneFinder.show vbModal
            Unload frmGeneFinder
            DisplaySingleGene obj
            SetActiveObject Nothing
         Case "Legend"
            frmOptions.Tag = "Legend"
            frmOptions.show vbModal
         Case "InfoBox"
            mnuInfo_Click
         Case Else
            MsgBox "Something other than a right-clickable object clicked, perhaps an object " _
                 & "underneath. If you are on a right-clickable object, try moving the mouse to " _
                 & "another part of the object and right-click again.", _
                 vbExclamation + vbOKOnly, "Right-Click Problem"
         End Select
      End If
      MousePointer = vbDefault
      dontClick = True                                                  'Don't treat as click event
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Select In Area
   ElseIf shpSelect.Width > 50 And shpSelect.Height > 50 And selections.count = 0 Then
      Dim begX As Single, begY As Single, endX As Single, endY As Single    'Corners of select area

'      selecting = True
      SetActiveObject Nothing
      begX = shpSelect.Left
      begY = shpSelect.Top
      endX = shpSelect.Left + shpSelect.Width
      endY = shpSelect.Top + shpSelect.Height
      selectArea.minX = 1E+38                                       'Set to be replaced immediately
      selectArea.minY = 1E+38
      selectArea.maxX = 0
      selectArea.maxY = 0
      For Each element In objLines
'         If (element.XStart >= begX And element.XStart <= endX _
'                    Or element.xEnd >= begX And element.xEnd <= endX) _
'               And (element.YStart >= begY And element.YStart <= endY _
'                    Or element.YEnd >= begY And element.YEnd <= endY) Then
         If (element.XStart * zoom >= begX And element.XStart * zoom <= endX _
                    Or element.xEnd * zoom >= begX And element.xEnd * zoom <= endX) _
               And (element.YStart * zoom >= begY And element.YStart * zoom <= endY _
                    Or element.YEnd * zoom >= begY And element.YEnd * zoom <= endY) Then
            element.SelectMode = True
            selections.Add element, element.objType & element.objKey
         End If
      Next element
      For Each element In objLumps
         If element.centerX * zoom >= begX And element.centerX * zoom <= endX _
               And element.centerY * zoom >= begY _
               And element.centerY * zoom <= endY Then
'         If element.centerX >= begX And element.centerX <= endX _
'               And element.centerY >= begY _
'               And element.centerY <= endY Then
            element.SelectMode = True
            selections.Add element, element.objType & element.objKey
         End If
      Next element
      MultipleObjectsSelected                                                'Set format menu items
      shpSelect.Width = 0                          'So we don't reselect when just clicking on form
      shpSelect.Height = 0
      dontClick = True                                                  'Don't treat as click event
      SetActiveObject Nothing                                          'Turn off any active objects
'      MousePointer = vbSizeAll    'vbDefault
   ElseIf Shift = vbCtrlMask Then '++++++++++++++++++++++++++ Select Or Deselect Individual Objects
      SetActiveObject Nothing                                          'Turn off any active objects
      Set element = ObjClicked
      If Not element Is Nothing Then                                          'Object under pointer
'         selecting = True
         If element.SelectMode Then                                          'Selected, deselect it
            element.SelectMode = False
            selections.Remove element.objType & element.objKey
         Else                                                              'Not selected, select it
            element.SelectMode = True
            selections.Add element, element.objType & element.objKey
         End If
         MultipleObjectsSelected                                             'Set format menu items
      End If
      dontClick = True                                                  'Don't treat as click event
      MousePointer = vbDefault
   ElseIf MousePointer = vbSizeAll Then '++++++++++++++++++++++++++++++++++++ Move Multiple Objects
      '  shpSelected boundary parameters always in nonzoomed coordinates, so Left and Top
      '  must be adjusted
      MoveSelection shpSelected.Left / zoom - selectArea.minX, _
                    shpSelected.Top / zoom - selectArea.minY
      dontClick = True
   ElseIf MousePointer <> vbSizeAll Then
      MousePointer = vbDefault
   End If
End Sub
Sub MoveSelection(moveX As Single, moveY As Single) '********************* Moves Multiple Selection
   '  Entry:   moveX, moveY   Unzoomed amount of move
   Dim oldMinX As Single, oldMinY As Single
   Dim element As Object

   oldMinX = selectArea.minX                            'Keep track in case move goes outside board
   oldMinY = selectArea.minY
   selectArea.minX = GridCoord(selectArea.minX + moveX)
   selectArea.minY = GridCoord(selectArea.minY + moveY)
   selectArea.maxX = selectArea.maxX + moveX
   selectArea.maxY = selectArea.maxY + moveY
   If OutsideBoard(selectArea) Then
      moveX = selectArea.minX - oldMinX
      moveY = selectArea.minY - oldMinY
   End If
   For Each element In selections
      element.Move moveX, moveY, 0
   Next element
   MultipleObjectsSelected
End Sub
Private Sub picDrafter_Click() '******************************************* Registers Click On Form
   '  See explanation in Delay_Timer()
   Dim element As Variant, index As Integer

   If MousePointer = vbCrosshair Then Exit Sub        'Setting line >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   If dontClick Then                                         'Event handles elsewhere, ignore click
      dontClick = False
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   MousePointer = vbHourglass                       'Any click deactivates and deselects everything
   MultipleObjectDeselectAll
   MousePointer = vbDefault
   sbrBar.Panels(1).text = ""
   SetActiveObject Nothing
   Delay.Enabled = True                   'Starts Delay_Timer, calls SingleClick after time reached
End Sub
Private Sub picDrafter_DblClick() '*************************************************** Double Click
   Dim obj As Object

   MousePointer = vbDefault
   dontClick = True
   Set obj = ObjClicked()
   If obj Is Nothing Then Exit Sub                         '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

   MousePointer = vbHourglass
   If TypeOf obj Is objLump Then
      Select Case obj.objType
      Case "Gene"
         ShowBackpage obj, dbGene, dbExpression
      Case Else
'         ShowObjPage obj
      End Select
   Else
      ShowObjPage obj
   End If
   SetActiveObject Nothing                                             'Turn active object back off
   MousePointer = vbDefault
End Sub
Private Sub Delay_Timer() '************************ Double Click If 2nd Click Before Timer Runs Out
   '  Click starts the Delay timer which calls SingleClick at the end of the interval.
   '  SingleClick is, in essence, the single click event handler.
   '  If a double click occurs, Dbl_Click sets dontClick to true and SingleClick exits without
   '     acting on the single click event.
   '  dontClick must be declared globally (here in Drafter.Bas). Also used for other click ops

   SingleClick                                           'Timer has run out, this is a single click
End Sub
Private Sub SingleClick() '****************************************** Single Click Event for Object
   '  See explanation in Delay_Timer()
'Debug.Print "GeneClick"

   MousePointer = vbDefault
   Delay.Enabled = False
   If Not newObject Is Nothing Then Exit Sub               '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   If dontClick Then
      dontClick = False
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If

   SetActiveObject ObjClicked()                                                        'Select object
End Sub

Rem //////////////////////////////////////////////////////////////////////////////// ToolBar Events
Private Sub tlbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
   frmObjects.DeselectAll
   With drawingBoard                                       'picDrafter. FrmDrafter is the container
   If Button.style = tbrDropdown Then
      tlbTools_ButtonDropDown Button
      Exit Sub                'Figure out how to click the dropdown
   End If
   ToolBarClear Button.key                                               'Unclick all other buttons
   If .container.lineStarted Then              'A line was previously started and user changed mind
      .foreColor = vbWhite                                                         'Delete X marker
      drawingBoard.Line (XStart - 50, YStart)-Step(100, 0)
      drawingBoard.Line (XStart, YStart - 50)-Step(0, 100)
      .foreColor = vbBlack
      .container.lineStarted = False
   End If
   If Button.value = tbrPressed Then
      .container.MultipleObjectDeselectAll
      .container.SetActiveObject Nothing                              'Any active object turned off
      If Button.key = "Objects" Then
         mnuObjects_Click
         Button.value = tbrUnpressed
      Else
         .MousePointer = vbCrosshair
         sbrBar.Panels("Instructions").text = statusBarPlace(Button.key)
         tlbTools.Tag = Button.key
      End If
   Else
      .MousePointer = vbDefault
      tlbTools.Tag = ""
   End If
   End With
End Sub

Private Sub tlbTools_ButtonDropDown(ByVal Button As MSComctlLib.Button)
i = i
End Sub

Private Sub tlbTools_ButtonMenuClick(ByVal buttonMenu As MSComctlLib.buttonMenu)
   With drawingBoard
   ToolBarClear buttonMenu.key                                           'Unclick all other buttons
   If .container.lineStarted Then              'A line was previously started and user changed mind
      .foreColor = vbWhite                                                         'Delete X marker
      drawingBoard.Line (XStart - 50, YStart)-Step(100, 0)
      drawingBoard.Line (XStart, YStart - 50)-Step(0, 100)
      .foreColor = vbBlack
      .container.lineStarted = False
   End If
'   If button.value = tbrPressed Then
      .MousePointer = vbCrosshair
      .container.MultipleObjectDeselectAll
      .container.SetActiveObject Nothing                              'Any active object turned off
      sbrBar.Panels("Instructions").text = statusBarPlace(buttonMenu.key)
      tlbTools.Tag = buttonMenu.key
'   Else
'      .MousePointer = vbDefault
'      .container.sbrBar.Panels("Instructions").Text = ""               'Any status bar entry erased
'      tlbTools.Tag = ""
'   End If
   End With
End Sub
Sub ToolBarClear(Optional remainPressed As String = "") '********************** Unpress All Buttons
   '  Entry    remainPressed  This button will remain pressed if it exists. Pressing a button
   '                          should release all the rest so this is the key of the button pressed.
   Dim Button As MSComctlLib.Button, buttonMenu As MSComctlLib.buttonMenu
   
   For Each Button In tlbTools.Buttons
      If Button.key <> remainPressed Then Button.value = tbrUnpressed
   Next Button
   sbrBar.Panels("Instructions").text = ""                             'Any status bar entry erased
End Sub


Rem /////////////////////////////////////////////////////////////////////////////////// Menu Events
Private Sub mnuGeneDB_Click()
   Dim oldGeneName As String, newGeneName As String

   If Not dbGene Is Nothing Then
      oldGeneName = dbGene.name
   End If
   OpenGeneDB dbGene, "**OPEN**", Me
   If Not dbGene Is Nothing Then
      newGeneName = dbGene.name
   End If
   If oldGeneName <> newGeneName Then
      mnuApply_Click
   End If
End Sub

Private Sub mnuNew_Click()
   Dim element As Variant, index As Integer
   
   If dirty Then
      Select Case MsgBox("Save current MAPP?", vbYesNoCancel + vbQuestion, "New MAPP")
      Case vbYes
         mnuSave_Click
         If MAPPSaveError Then
            MAPPSaveError = False
            Exit Sub                                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      Case vbCancel
         Exit Sub                                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case Else
      End Select
   End If
   
'   loading = True
   Caption = PROGRAM_TITLE
   picDrafter.Cls                                                  'Clear everything drawn on board
'   frmTools.lblTitle = ""
'   frmObjects.lblTitle = ""
'   frmObjects.Hide

'  Be sure to set the scrollbars back to beginning position

   boardWidth = INITIAL_BOARD_WIDTH
   picDrafter.Width = boardWidth
'   MinBoardWidth = picDrafter.Width
   WindowState = vbNormal
   FormWidth INITIAL_WINDOW_WIDTH              'Set here so ReSize event doesn't adjust height also
   boardHeight = INITIAL_BOARD_HEIGHT
   picDrafter.Height = boardHeight
'   MinBoardHeight = picDrafter.Height
   FormHeight INITIAL_WINDOW_HEIGHT
   zoom = 1                                                          'Put before loading frmDrafter
   cmbZoom.text = "100%"
   SetActiveObject Nothing                                    'Remove activeObject reference if set
'   picDrafter.Cls
   frmMAPPInfo.Clear
   For Each element In objLines '++++++++++++++++++++++++++++++++++++++++++++++++++Remove All Lines
      objLines.Remove element.objKey              'Remove from collection. Should be last reference
   Next element
   For Each element In objLumps '++++++++++++++++++++++++++++++++++++++++++++++++++Remove All Lumps
      objLumps.Remove element.objKey              'Remove from collection. Should be last reference
   Next element
'   Set MAPPTitle = Nothing                                           'Remove any previous map title
   Set MAPPTitle = New objLump                                       'Remove any previous map title
   MAPPTitle.Create -1, -1, "MAPPTitle"
   
   MultipleObjectDeselectAll
   mnuHorizAlign.Enabled = False
   mnuVertAlign.Enabled = False
   mnuSize.Enabled = False
   mnuBlock.Enabled = False

   info.Create 0, 0
   legend.Create 0, -2                                                 'So as to not hide Info Area
   picDrafter.Left = 0
   picDrafter.Top = tlbTools.Height
   mnuApply.Enabled = False
   dirty = False
'   If Right(MAPPName, 1) = "$" Then                                    'Temp copy of read-only mapp
'      Kill MAPPName
'   End If
   mappName = ""
'   loading = False
End Sub
Sub mnuOpen_Click()
'   Dim dbMapp As Database, dbTemp As Database
'   Dim rsMAPP As Recordset
'   Dim rsInfo As Recordset, rsColorSet As Recordset
'   Dim sql As String, s As String
'   Dim index As Integer
   Dim newMappName As String                  'Temporary until beyond cancel point and mnuNew_Click
'   Dim errorLocation As String   'Set to part of MAPP being opened -- Expression, etc

   If dirty Then
      Select Case MsgBox("Save current MAPP?", vbYesNoCancel + vbQuestion, "Open MAPP")
      Case vbYes
         mnuSave_Click
         If MAPPSaveError Then
            MAPPSaveError = False
            Exit Sub                                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      Case vbCancel
         Exit Sub                                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case Else
         dirty = False                                              'Declined to save, forget dirty
      End Select
   End If
   
On Error GoTo OpenError
   If InStr(commandLine, ".mapp""") Then                                     'React To Command Line
      Dim begCommand As Integer, endCommand As Integer
      endCommand = InStr(commandLine, ".mapp""") + 4                              'Before end quote
      begCommand = InStrRev(commandLine, """", endCommand) + 1                   'After begin quote
      newMappName = Mid(commandLine, begCommand, endCommand - begCommand + 1)
      If begCommand - 2 = 0 Then                                       'Must be first thing on line
         commandLine = Left(commandLine, begCommand - 2) & Mid(commandLine, endCommand + 2)
      Else                                           'Not first, so must be preceded by space
         '  The 2s take out the quotes also
         commandLine = Left(commandLine, begCommand - 3) & Mid(commandLine, endCommand + 2)
      End If
   Else                                                                         'Open MAPP Manually
      dlgDialog.CancelError = True
      dlgDialog.DialogTitle = "Open MAPP"
      dlgDialog.InitDir = mruMappPath
      dlgDialog.Filter = "MAPPs (.mapp)|mapp"
      dlgDialog.FileName = "*.mapp"
      dlgDialog.FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
      dlgDialog.ShowOpen
      newMappName = dlgDialog.FileName
      If InStr(newMappName, ".") = 0 Then
         newMappName = newMappName & ".mapp"
      End If
   End If
On Error GoTo 0
   
   If Dir(newMappName) = "" Then
      MsgBox newMappName & " doesn't exist.", vbExclamation + vbOKOnly, "Opening MAPP"
      GoTo ExitSub                                         '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   If GetAttr(newMappName) And vbReadOnly Then
      MsgBox "This MAPP has been set to read-only through Windows. You can open and " _
             & "manipulate it but you will not be able to save it under the same name.", _
             vbInformation + vbOKOnly, "Opening MAPP"
   End If
   b = OpenMAPP(newMappName)
ExitSub:
   MousePointer = vbDefault
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
OpenError:
   If Err <> 32755 Then                                                          'Other than Cancel
      MsgBox Err.Description, vbCritical, "Open MAPP Error"
      Resume Next   'Temporary. Only valuable when switching database formats  ????????????????????
   End If
   On Error GoTo 0
   Resume ExitSub
End Sub
Public Function OpenMAPP(newMappName As String) As Boolean '************************* Open The MAPP
   Dim dbMapp As Database, dbTemp As Database
   Dim rsMAPP As Recordset
   Dim rsInfo As Recordset, rsObjects As Recordset ', rsColorSet As Recordset
   Dim sql As String, s As String
   Dim index As Integer
   Dim errorLocation As String   'Set to part of MAPP being opened -- Expression, etc
   Dim ok As Boolean

   If dirty Then
      Select Case MsgBox("Save current MAPP?", vbYesNoCancel + vbQuestion, "Open MAPP")
      Case vbYes
         mnuSave_Click
         If MAPPSaveError Then
            MAPPSaveError = False
            Exit Function                                               '>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      Case vbCancel
         Exit Function                                                  '>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case Else
         dirty = False                                              'Declined to save, forget dirty
      End Select
   End If
   
   MousePointer = vbHourglass
   
   mnuNew_Click                                                                 'Clear out old MAPP
   mappName = newMappName
   mruMappPath = GetFolder(mappName)
   
   Set dbMapp = OpenDatabase(mappName, False, True)                            'Open read-only here
      '  It will be reopened when saved
   
   If Not UpdateMAPP(dbMapp) Then GoTo ExitFunction        'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   
   zoom = 1
   cmbZoom.text = "100%"
   picDrafter.Left = 0
   picDrafter.Top = tlbTools.Height
   Set rsInfo = dbMapp.OpenRecordset("SELECT * FROM Info")
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Open MAPP
   With frmMAPPInfo
      !txtTitle = rsInfo!title
      MAPPTitle.Create -1, -1, "MAPPTitle"
      MAPPTitle.title = Dat(rsInfo!title)
      !lblMAPP = Dat(rsInfo!mapp)
      !txtAuthor = Dat(rsInfo!author)
      !txtMaint = Dat(rsInfo!maint)
      !txtEMail = Left(Dat(rsInfo!email), 50)           'To fix old MAPPs that allowed over 50 chrs
      !txtCopyright = Left(Dat(rsInfo!copyright), 50)
      !txtModify = Dat(rsInfo!modify)
      !txtRemarks = Dat(rsInfo!remarks)
      !txtNotes = Dat(rsInfo!notes)
   End With
   boardWidth = rsInfo!boardWidth
   picDrafter.Width = boardWidth
   boardHeight = rsInfo!boardHeight
   picDrafter.Height = boardHeight
   
   WindowState = vbNormal
   Width = rsInfo!WindowWidth
   Height = rsInfo!WindowHeight
   Set rsObjects = dbMapp.OpenRecordset("SELECT MAX(ObjKey) AS objectKey FROM Objects")
      '  Previous MAPPs do not have object keys, so the return here will be NULL
   If VarType(rsObjects!objectKey) = vbNull Then
      objKey = 0
   Else
      objKey = rsObjects!objectKey                     'Will be incremented for the next new object
   End If
   rsObjects.Close
   Set rsObjects = dbMapp.OpenRecordset("SELECT * FROM Objects")
   With rsObjects                                          '~~~~~~~~~~~~~~~~~~~~~~~~~With rsObjects
   Do Until .EOF
      Select Case !Type
      Case "Custom"
      Case "Line", "DottedLine", "Arrow", "DottedArrow", "Receptor", "ReceptorRd", "ReceptorSq", _
           "LigandRd", "LigandSq", "TBar"
         Dim newLine As New objLine
         newLine.color = !color
         Select Case !Type
         Case "Line"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "Solid", !objKey ', Me
         Case "DottedLine"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "Broken", !objKey ', Me
         Case "Arrow"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "Arrow", !objKey ', Me
         Case "DottedArrow"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "BrokenArrow", !objKey ', Me
         Case "Receptor"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "Receptor", !objKey ', Me
         Case "ReceptorSq"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "ReceptorSq", !objKey ', Me
         Case "ReceptorRd"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "ReceptorRd", !objKey ', Me
         Case "LigandSq"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "LigandSq", !objKey ', Me
         Case "LigandRd"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "LigandRd", !objKey ', Me
         Case "TBar"
            newLine.Create !centerX, !centerY, !SecondX, _
                           !SecondY, "Inhibitor", !objKey ', Me
'         Case "Curve"
'            newLine.Create !centerX, !centerY, !SecondX, _
'                           !SecondY, "Arc", !objKey ', Me
         End Select
         newLine.remarks = Dat(!remarks)
         objLines.Add newLine, newLine.objKey
         Set newLine = Nothing                              'Only reference should be in collection
      Case "Rectangle", "Oval", "Arc", "Brace", "Gene", "Vesicle", "ProteinA", "ProteinB", _
           "Ribosome", "OrganA", "OrganB", "OrganC", "CellA", "CellB", "Label", "Curve", "Poly"
         Dim newLump As New objLump
         
         newLump.id = Dat(!id)
         newLump.systemCode = Dat(!systemCode)
         newLump.title = Dat(!Label)
         newLump.head = Dat(!head)
         newLump.remarks = Dat(!remarks)
         newLump.notes = Dat(!notes)
         newLump.links = Dat(!links)
         newLump.Size = NullZero(!SecondX)                                       'Labels  Font size
         newLump.sides = NullZero(!SecondY)                                             'Poly sides
         newLump.Create !centerX, !centerY, !Type, !Width, _
                        !Height, NullZero(!rotation), NVL(!color, -1), , !objKey ', Me
                                                               'After creation to override defaults
         If !Type = "Gene" Then                     'For genes, color comes from expression dataset
            newLump.color(0) = -1
         End If
         objLumps.Add newLump, newLump.objKey
         Set newLump = Nothing                              'Only reference should be in collection
      Case "InfoBox"
         info.Create !centerX, !centerY
      Case "Legend"
         legend.Create !centerX, !centerY ', Dat(!ID) ', Me
      End Select
      .MoveNext
   Loop
   If colorIndexes(0) > 0 Then                                    'Apply colors and values to genes
      mnuApply.Enabled = True
      mnuApply_Click
   End If
   ScrollBars                                        'Set the scroll bars and the size of the board
'   Set MAPPTitle = Nothing
   MAPPTitle.DrawObj True
   .Close
   End With                                                '~~~~~~~~~~~~~~~~~~~~~End With rsObjects
   dbMapp.Close
   dirty = False
   s = Mid(mappName, InStrRev(mappName, "\") + 1)
   s = Left(s, InStrRev(s, ".") - 1)
'   frmObjects.lblTitle = s
'   frmObjects.Hide
'   frmTools.lblTitle = s
   s = s & " - " & PROGRAM_TITLE
   Caption = s
   picDrafter.Left = 0
   picDrafter.Top = tlbTools.Height
   ScrollBars
   mnuRedraw_Click                                               'To put objects in correct Z order
   OpenMAPP = True
ExitFunction:
   MousePointer = vbDefault
   Exit Function
End Function
   
Private Sub mnuRemoveLocalGenes_Click()
   RemoveLocalGenes
End Sub

Private Sub mnuSave_Click()
   Dim dbMapp As Database, rs As Recordset
   Dim index As Integer, element As Variant
   Dim sql As String
   
   If mappName = "" Then
      mnuSaveAs_Click
      GoTo ExitSub                                              '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If GetAttr(mappName) And vbReadOnly Then
      mnuSaveAs_Click
      GoTo ExitSub                                              '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If MAPPSaveError Then
      MAPPSaveError = False
      GoTo ExitSub                                              '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
'On Error GoTo SaveError
   MousePointer = vbHourglass
   Set dbMapp = OpenDatabase(mappName)
On Error GoTo 0

   dbMapp.Execute "DELETE FROM Objects"
   For Each element In objLines '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Lines
      sql = "INSERT INTO Objects (ObjKey, Type, CenterX, CenterY, SecondX, SecondY, Color, Remarks) " & _
            "   VALUES ('" & element.objKey & "',"
      Select Case element.style
      Case "Solid"
         sql = sql & "'Line', "
      Case "Broken"
         sql = sql & "'DottedLine', "
      Case "Arrow"
         sql = sql & "'Arrow', "
      Case "BrokenArrow"
         sql = sql & "'DottedArrow', "
      Case "Receptor"
         sql = sql & "'Receptor', "
      Case "ReceptorSq"
         sql = sql & "'ReceptorSq', "
      Case "ReceptorRd"
         sql = sql & "'ReceptorRd', "
      Case "LigandSq"
         sql = sql & "'LigandSq', "
      Case "LigandRd"
         sql = sql & "'LigandRd', "
      Case "Inhibitor"
         sql = sql & "'TBar', "
      Case "Arc"
         sql = sql & "'Curve', "
      End Select
      sql = sql & element.XStart & ", " & element.YStart & ", " _
          & element.xEnd & ", " & element.YEnd & ", " & element.color & ", '" & element.remarks & "')"
      dbMapp.Execute sql
   Next element
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Info box
   sql = "INSERT INTO Objects (ObjKey, Type, CenterX, CenterY)" _
       & "   VALUES ('" & info.objKey & "', 'InfoBox', " & info.centerX & ", " & _
                     info.centerY & ")"
   dbMapp.Execute sql
   
   If Not legend Is Nothing Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Legend
      sql = "INSERT INTO Objects (ObjKey, Type, ID, CenterX, CenterY)" & _
            "   VALUES ('" & legend.objKey & "', 'Legend', '" & legend.Display & "', " & _
                        legend.centerX & ", " & legend.centerY & ")"
      dbMapp.Execute sql
   End If
   
   For Each element In objLumps '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Lumps
      element.title = TextToSql(element.title)                          'Make titles SQL compatible
      '  Save last so that when they open they will appear on top when opened
      sql = "INSERT INTO Objects" & _
            "      (ObjKey, ID, SystemCode, Type, CenterX, CenterY, SecondX, SecondY," & _
            "       Width, Height, Rotation, Color, Label, Head, Remarks, Links)" & _
            "   VALUES" & _
            "      (" & element.objKey & ", '" & element.id & "', '" & _
                    element.systemCode & "', '" & element.objType & "', " & _
                    element.centerX & ", " & element.centerY & ", " & element.Size & ", " & _
                    element.sides & ", " & element.wide & ", " & element.high & ", " & _
                    element.rotation & ", " & element.color(1) & ", '" & element.title & "', '" & _
                    element.head & "', '" & element.remarks & "', '" & element.links & "')"
      dbMapp.Execute sql
   Next element
   
   dbMapp.Execute "DELETE FROM Info" '+++++++++++++++++++++++++++++++++++++++++++++++++++ Info Data
   With frmMAPPInfo
      sql = "VALUES ('" _
         & MAPPTitle.title & "', '" _
         & !lblMAPP & "', '" _
         & BUILD & "', '" _
         & Dat(!txtAuthor) & "', '" _
         & Dat(!txtMaint) & "', '" _
         & Dat(!txtEMail) & "', '" _
         & Dat(!txtCopyright) & "', '" _
         & Dat(!txtModify) & "', '" _
         & Dat(!txtRemarks) & "', " _
         & boardWidth & ", " _
         & boardHeight & ", " _
         & frmDrafter.Width & ", " _
         & frmDrafter.Height & ", "
      sql = sql & "'" & Dat(TextToSql(!txtNotes)) & "')"
      sql = "INSERT INTO Info (Title, MAPP, Version, Author, Maint, Email, Copyright, Modify," _
          & "                  Remarks, BoardWidth, BoardHeight, WindowWidth, WindowHeight," _
          & "                  Notes) " & sql
      dbMapp.Execute sql
   End With
   dbMapp.Close
   dirty = False
ExitSub:
   MousePointer = vbDefault
   Exit Sub
   
SaveError:
   Select Case Err.number
   Case 52, 75, 3043
      MsgBox "Cannot save to this path. This may be a read-only drive, such as a CD-ROM, " _
             & "or a removable drive with no disk in it.", vbExclamation + vbOKOnly, _
             "Save MAPP Error"
   Case 3045
      MsgBox Err.Description & ". MAPP possibly open in some other program.", vbExclamation, _
            "Save MAPP Error"
      MAPPSaveError = True
   Case 32755                                                                            'Cancelled
   Case Else
      MsgBox Err.Description, vbCritical, "Save MAPP Error"
      MAPPSaveError = True
   End Select
   On Error GoTo 0
   Resume ExitSub
End Sub
Private Sub mnuSaveAs_Click()
   Dim mruDataSet As String                          'Hides global. Delete when config file set up
   Dim oldMAPPName As String, s As String
   
On Error GoTo SaveError
   oldMAPPName = mappName                                    'In case of error on new MAPP
ReEnter:
   dlgDialog.CancelError = True
   dlgDialog.DialogTitle = "Save MAPP"
   dlgDialog.Filter = "mapp"
   dlgDialog.InitDir = GetFolder(mruMappPath)
   dlgDialog.FileName = "*.mapp"
   dlgDialog.FLAGS = cdlOFNExplorer + cdlOFNHideReadOnly
   dlgDialog.ShowSave
   mappName = dlgDialog.FileName
   If InStr(mappName, ".") = 0 Then
      mappName = mappName & ".mapp"
   End If
   If Dir(mappName) <> "" Then
      If GetAttr(mappName) And vbReadOnly Then
         MsgBox "This MAPP has been set to read-only through windows. You may not save it.", _
                vbExclamation + vbOKOnly, "Saving MAPP"
         mappName = oldMAPPName
         GoTo ReEnter
      End If
   End If
   If UCase(Dir(mappName)) = UCase(Mid(mappName, InStrRev(mappName, "\") + 1)) Then
      Select Case MsgBox("Do you want to replace the current " & mappName & "?", _
             vbYesNoCancel + vbQuestion, "Saving MAPP")
      Case vbNo
         GoTo ReEnter                                   '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      Case vbCancel
         GoTo ExitSub                                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End Select
   End If
'   If Right(oldMAPPName, 1) = "$" Then
'      If Dir(oldMAPPName) <> "" Then Kill oldMAPPName
'   End If
   FileCopy appPath & "MAPPTmpl.gtp", mappName                        'Put template in new file
                 'This will generate an error if the MAPP in use somewhere else (like in Access)
   mnuSave_Click
   s = Mid(mappName, InStrRev(mappName, "\") + 1)
   s = Left(s, InStrRev(s, ".") - 1)
'   frmTools.lblTitle = s
'   frmObjects.lblTitle = s
'   frmObjects.Hide
   Caption = s & " - " & PROGRAM_TITLE
   frmDrafter.show
   
ExitSub:
   MousePointer = vbDefault
   Exit Sub
   
SaveError:
   Select Case Err.number
   Case 52, 75
      MsgBox "Cannot save to this path. This may be a read-only drive, such as a CD-ROM, " _
             & "or a removable drive with no disk in it.", vbExclamation + vbOKOnly, _
             "Save MAPP Error"
   Case 70, 3045
      MsgBox Err.Description & ". " & mappName & " possibly open in some other program.", _
            vbExclamation, "Save MAPP Error"
      MAPPSaveError = True
   Case 32755                                                       'Not an error if just cancelled
   Case Else
      MsgBox Err.Description, vbCritical, "Save MAPP Error"
      MAPPSaveError = True
   End Select
   mappName = oldMAPPName                                                     'Set back to old MAPP
   On Error GoTo 0
   Resume ExitSub
End Sub
Sub mnuPrint_Click()
   Dim defaultPrinterName As String, defaultPrinterDriver As String, defaultPrinterPort As String
   Dim element As Object, i As Integer
   Dim prevZoom As Single
   
   defaultPrinterName = Printer.DeviceName
   defaultPrinterDriver = Printer.DriverName
   defaultPrinterPort = Printer.Port
   prevZoom = zoom
   Printer.TrackDefault = True
On Error GoTo ErrHandler
   dlgPrinter.CancelError = True
   dlgPrinter.PrinterDefault = True
   dlgPrinter.FromPage = 1
   dlgPrinter.ToPage = 1
   dlgPrinter.ShowPrinter
On Error GoTo 0
   MousePointer = vbHourglass
   SetActiveObject Nothing                                              'Turn off any active object
   For i = 1 To selections.count                                  'Deselect all multiple selections
      selections(1).SelectMode = False
      selections.Remove 1
   Next i
   '  At this point, not objects should be in the select or edit mode
'   mnuRedraw_Click
   For i = 1 To dlgPrinter.Copies
      '  Must set the printer orientation for each copy or printer reverts to default
      If dlgPrinter.Orientation = cdlPortrait Then
         Printer.Orientation = vbPRORPortrait
         zoom = 11520 / boardWidth   '8 in
         If zoom > 14400 / boardHeight Then     '10 in
            zoom = 14400 / boardHeight
         End If
      Else
         Printer.Orientation = vbPRORLandscape
         zoom = 11520 / boardHeight   '8 in
         If zoom > 14400 / boardWidth Then      '10 in
            zoom = 14400 / boardWidth
         End If
      End If
'      TitleMAPP Printer
      For Each element In objLines
         element.DrawObj True, Printer
      Next element
      For Each element In objLumps                                'Lumps last so they appear on top
         element.DrawObj True, Printer
      Next element
      info.DrawObj True, Printer
      legend.DrawObj True, Printer
      MAPPTitle.DrawObj True, Printer
On Error GoTo ErrHandler
      Printer.EndDoc
   Next i
ExitSub:
   SetDefaultPrinter defaultPrinterName, defaultPrinterDriver, defaultPrinterPort
   MousePointer = vbDefault
   zoom = prevZoom
   Exit Sub
   
ErrHandler:
   Select Case Err.number
   Case 32755                                                                      'Cancelled Print
      Resume ExitSub
   Case 482
      If Printer.DeviceName = "Acrobat Distiller" Then
         '  If cancelled at file name dialog, distiller still goes through printing process although nothing actually is done. When VB tries to execute the Printer.EndDoc, this error appears.
         '  Adobe does something really weird here, Distiller doesn't report the cancellation
         '  of a print request because a user cancelled when choosing a file name. Instead they
         '  prepare to print before the file name is chosen. As soon as a program sends the
         '  first little bit --a dot, character, or whatever -- to the Distiller, it then asks
         '  for a file name. This is well into the printing process. If the user clicks Cancel,
         '  Distiller continues with the printing process, even though it is not printing to a
         '  valid object. It is only when the program tries to close this object that it
         '  realizes that Distiller is dropping all the data into the bit bucket.

         Resume ExitSub
      End If
   End Select
   FatalError "frmDrafter:mnuPrint", Err.Description
End Sub
Private Sub mnuExit_Click()
   If dirty Then
      Select Case MsgBox("Save current MAPP?", vbYesNoCancel + vbQuestion, "Exiting Program")
      Case vbYes
         mnuSave_Click
      Case vbCancel
         cancelExit = True
         Exit Sub                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case Else
      End Select
   End If
'   If Right(MAPPName, 1) = "$" Then                                    'Temp copy of read-only mapp
'      Kill MAPPName
'   End If
   WriteConfig                                        'Save all the latest settings before leaving
   Unload frmDrafter
   End
End Sub
Private Sub mnuMAPPSet_Click()
'   Load frmMAPPSet
   frmMAPPSet.show vbModal
End Sub

Private Sub mnuDraftingTools_Click()
'   frmTools.show
'   WindowAlwaysOnTop (frmTools.hWnd)                                                  'Float on top
End Sub
Private Sub mnuInfo_Click()
   frmMAPPInfo.show vbModal
   MAPPTitle.DrawObj False
   MAPPTitle.title = frmMAPPInfo.txtTitle
   MAPPTitle.DrawObj True
   info.Create info.centerX, info.centerY
   OutsideBoard info
   dirty = True
End Sub
Private Sub mnuBoardSize_Click()
'   frmBoardParams!txtWidth = Format(boardWidth / TWIPS_CM, "0.0")            'Convert pixels to cm.
'   frmBoardParams!txtHeight = Format(boardHeight / TWIPS_CM, "0.0")
'   frmBoardParams.Tag = ""
'   Set frmBoardParams.canvas = picDrafter
   frmBoardParams.show vbModal
'   If frmBoardParams.Tag = "OK" Then
'      callingRoutine = "mnuBoardSize_Click"
'         '  Set so that FormWidth() and FormHeight() do not call resize automatically
'      If ClientWide() > picDrafter.Width + picDrafter.Left Then  'Window size > drafting board edge
'         '  If the size of the drafting board is reduced so that there is client window beyond the
'         '  right, bottom then the client window is resized.
'         WindowState = vbNormal                                          'Can't resize if maximized
'         FormWidth picDrafter.Width + picDrafter.Left
'      End If
'      If ClientHigh() > picDrafter.Height + tlbTools.Height - picDrafter.Top Then
'         '  The top of the visible picDrafter viewport is below the tool bar (tlbTools).
'         WindowState = vbNormal                                          'Can't resize if maximized                                                  'In case maximized
'         FormHeight picDrafter.Height + picDrafter.Top - tlbTools.Height
'      End If
'      callingRoutine = ""
'      Form_Resize                                                'Now we actually do want to resize
''      mnuRedraw_Click
'      dirty = True
'   End If
End Sub
Public Sub mnuRedraw_Click()
   '  Redraws screen leaving any selected objects selected in their new positions
   Dim element As Object, index As Integer
   
   MousePointer = vbHourglass
   callingFunction = "mnuRedraw"
   picDrafter.Cls
   For index = 1 To 5                                                   'Cls doesn't clear controls
      picPoint(index).visible = False
   Next index
   For Each element In objLines
      element.DrawObj
      If element.editMode Then element.SetEdit
   Next element
   For Each element In objLumps
      element.DrawObj
      If element.editMode Then element.SetEdit
   Next element
   selectArea.DrawObj
   info.DrawObj
   legend.DrawObj
   MAPPTitle.DrawObj True
'   mnuApply_Click
   CancelUndo
   callingFunction = ""
   MousePointer = vbDefault
End Sub
Private Sub mnuChoose_Click() '****************************************** Choose Expression Dataset
   Dim expression As String
   
'   Set prevExpression = dbExpression                                               'In case of error
'   Set prevColorSet = rsColorSet
On Error GoTo OpenError
Retry:
    With dlgDialog
      .CancelError = True
      .DialogTitle = "Choose Expression Dataset"
      .Filter = "Expression Dataset (.gex)|*.gex"
      .InitDir = GetFolder(mruDataSet)
      .FileName = GetFolder(mruDataSet) & "*.gex"
      .FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
      .ShowOpen
      expression = .FileName
   End With
   If InStr(expression, ".") = 0 Then
      expression = expression & ".gex"
   End If
      
   mruColorSet = ""
   If SetDataSet(expression) = "" Then
      GoTo Retry                                           '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   End If
   
'   If InStr(1, expression, "\" & Dir(expression), vbTextCompare) = 0 Then
'      If MsgBox("Expression dataset '" & expression & " does not exist.", _
'               vbExclamation + vbRetryCancel, "Open Expression Dataset") = vbCancel Then
'         Set dbExpression = prevExpression          'Matches open Expression Dataset name or nothing
'         Set rsColorSet = prevColorSet
'         GoTo ExitSub          'Canceled choosing a dataset vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'      Else
'         GoTo Retry
'      End If
'   End If
'
'   Set dbExpression = OpenDatabase(expression, False, True)
'   Set rsColorSet = Nothing                                                'No Color Set chosen yet
'   If Not DatasetCurrent(dbExpression) Then '+++++++++++++++++++++++++++++++++++++ Do Version Check
'      MsgBox dbExpression.name & vbCrLf & vbCrLf & "was created in a previous version of " _
'             & "GenMAPP and must be converted to the current version before it may be " _
'             & "opened. In the Drafter window, click the ""Tools"" menu, ""Converter"" " _
'             & "option to convert your Expression Dataset.", _
'             vbCritical + vbOKOnly, "Old Expression Dataset"
'      dbExpression.Close
'      Set dbExpression = prevExpression                               'Reopen old Expression Dataset
'      Set rsColorSet = prevColorSet
'      GoTo ExitSub            'Canceled choosing a dataset vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'   End If
'   If Not HasDisplayTable(dbExpression) Then
'      CreateDisplayTable dbExpression
'   End If
'
'On Error GoTo 0
'   If Not dbExpression Is Nothing Then
'      mruDataSet = dbExpression.name                                        'Reset global data path
'   End If
'
'   FillColorSetList
'   cmbColorSets.SetFocus
   If lstColorSets.visible Then lstColorSets.visible = False
   mnuApply_Click
ExitSub:
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

OpenError:
   Select Case Err
   Case 32755                                                                       'Cancel clicked
'      If Not dbExpression Is Nothing Then                                'Expression Dataset exists
'         If MsgBox("Retain current Expression Dataset? (A ""No"" answer makes no Expression " & _
'                   "Dataset active.)", vbInformation + vbYesNo, "Choose Expression Dataset") _
'                   = vbYes Then
'            'Already done in SetDataSet()
''            Set dbExpression = prevExpression
''            Set rsColorSet = prevColorSet
'         Else                                                         'Set to no Expression Dataset
'            Set dbExpression = Nothing
'            Set rsColorSet = Nothing
'            FillColorSetList
'         End If
'      End If
   Case Else
      FatalError "frmDrafter:mnuChoose", Err.Description & "  " & expression
   End Select
   On Error GoTo 0
   Resume ExitSub
End Sub

'************************************************************************** Sets Expression Dataset
Function SetDataSet(Optional dataSet As String = "", Optional colorSet As String = "") _
         As String
   '  Entry    dataSet     Path of Expression Dataset to open.
   '                       If not stated, defaults to mruDataSet.
   '           colorSet    Same form as mruColorSet: DisplayValue\ColorSet\ColorSet\.....
   '                       Defaults to mruColorSet
   '  Return   Path of Expression Dataset opened or "" if unsuccessful.
   '  If SetDataSet() cannot open the new dataset it returns the mappWindow to the previous
   '  Expression Dataset and ColorSet.
   'To open the most recently used, both parameters should be omitted. Eg: X = SetDataset()
   '  This must be in frmDrafter so that it can see the previous data and set the new parameters.
   Dim prevExpression As Database, prevColorSet As Recordset, expression As String
   Dim colorSets(MAX_COLORSETS) As String, noOfColorSets As Integer, sql As String
   
   If dataSet = "" Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Defaults
      dataSet = mruDataSet
      If colorSet = "" Then colorSet = mruColorSet
         '  Only open the Color Set if we are setting to MRU values. Otherwise, a new Expression
         '  Dataset is being opened and the Color Set will have to be chosen.
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Preliminary Checks
   If dataSet = "" Then                                                'No Expression Dataset given
      Set dbExpression = Nothing
      valueIndex = -1
      colorSets(0) = 0
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If Dir(dataSet) = "" Then                                          'Expression Dataset Not Found
      MsgBox "Expression dataset '" & dataSet & " does not exist.", _
             vbExclamation + vbOKOnly, "Open Expression Dataset"
      Set dbExpression = Nothing
      valueIndex = -1
      colorSets(0) = 0
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If

On Error GoTo OpenError '++++++++++++++++++++++++++++++++++++++++++++++ Open New Expression Dataset
   Set prevExpression = dbExpression                                      'Save previous parameters
'   Set prevColorSet = rsColorSet                                       'Both these could be Nothing
   Set dbExpression = OpenDatabase(dataSet, False, True)
On Error GoTo 0
   If Not DatasetCurrent(dbExpression) Then '======================================Do Version Check
      MsgBox dbExpression.name & vbCrLf & vbCrLf & "was created in a previous version of " _
             & "GenMAPP and must be converted to the current version before it may be " _
             & "opened. In the Drafter window, click the ""Tools"" menu, ""Converter"" " _
             & "option to convert your Expression Dataset.", _
             vbCritical + vbOKOnly, "Old Expression Dataset"
      dbExpression.Close
      Set dbExpression = prevExpression            'Matches open Expression Dataset name or Nothing
'      Set rsColorSet = prevColorSet
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If Not HasDisplayTable(dbExpression) Then '=================================Create Display Table
      MsgBox "This Expression Dataset requires updating, which could take several minutes. The " _
             & "mouse pointer over the Drafter window will be changed to an hourglass. When it " _
             & "becomes an arrow again, the process is finished and you may proceed.", _
             vbExclamation + vbOKOnly, "Updating Expresison Dataset"
      MousePointer = vbHourglass
      CreateDisplayTable dbExpression
      MousePointer = vbDefault
   End If
   mruDataSet = dbExpression.name
   
   If colorSet <> "" Then '++++++++++++++++++++++++++++++++++++++++++++ Assign Color Set Parameters
      noOfColorSets = SeparateValues(colorSets, colorSet, "\")
      sql = "'" & colorSets(1) & "'"
      For i = 2 To noOfColorSets - 1
         sql = sql & ", '" & colorSets(i) & "'"
      Next i
         
      Set rsColorSet = dbExpression.OpenRecordset( _
                       "SELECT * FROM ColorSet WHERE ColorSet IN (" & sql & ")")
      If rsColorSet.EOF Then '=======================================Set For User To Open Color Set
         Set rsColorSet = Nothing
         mruColorSet = ""
         valueIndex = -1
         colorIndexes(0) = 0
         mnuApply.Enabled = False
         txtColorSets = "No expression data"
      Else '================================================================Open Existing Color Set
         mruColorSet = colorSet
         i = 0
         Do Until rsColorSet.EOF
            i = i + 1
            colorIndexes(i) = rsColorSet!setNo
            If colorSets(0) = rsColorSet!colorSet Then           'This is colorset value to display
               valueIndex = rsColorSet!setNo
            End If
            rsColorSet.MoveNext
         Loop
         colorIndexes(0) = i
         mnuApply.Enabled = True
      End If
   Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set For User To Open Color Set
      Set rsColorSet = Nothing
      mruColorSet = ""
      valueIndex = -1
      colorIndexes(0) = 0
      txtColorSets = "No expression data"
      mnuApply.Enabled = False
   End If
   FillColorSetList                                   'This also shows currently selected colorsets
   
   SetDataSet = dbExpression.name
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

OpenError:
Stop
   Select Case Err.number
   Case 52, 3024
      MsgBox mruDataSet & " doesn't exist.", vbExclamation + vbOKOnly, "Opening Dataset"
'      Resume RestorePrevious
   Case Else
      FatalError "frmDrafter:SetDataSet", Err.Description & "  " & dataSet
   End Select
End Function
Public Sub mnuApply_Click() '************************************** Apply Expression Data On Screen
   Dim element As Object
   Dim rsColorSets As Recordset
   Dim rsInfo As Recordset
   Dim colorSets(MAX_COLORSETS) As Integer, displayColorSet As String
   Dim tempGene As objLump
   Dim systemCodes As String                                   'Codes to search for related systems
   Dim i As Integer, j As Integer
   
   MousePointer = vbHourglass
   If dbExpression Is Nothing Or colorIndexes(0) = 0 Then '+++++++++++++ Clear All Genes And Values
      For Each element In objLumps
         If element.objType = "Gene" Then
            element.DrawGeneValue drawingBoard, False
            element.value = ""
            element.color(0) = 1
            element.color(1) = vbWhite
            element.rim(0) = 1
            element.rim(1) = vbWhite
            element.lineStyle = LINE_STYLE_SOLID
            element.DrawObj
         End If
      Next element
      legend.Create                                           'Erases legend under these conditions
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Apply Expression Values
      '================================================Determine Color Set Columns In Display Table
''      i = 0
''      Set rsColorSets = dbExpression.OpenRecordset("SELECT * FROM ColorSet")        'All Color Sets
''      SetRsColorSet
''      Do Until rsColorSet.EOF '-----------------------------------------------Each chosen color Set
''         i = i + 1
''         rsColorSets.MoveFirst
''         j = -1
''         Do Until rsColorSets.EOF                                    'Find order in Color Set table
''            j = j + 1
''            If rsColorSet!colorSet = rsColorSets!colorSet Then
''               colorSets(i) = j                                                         'Zero based
''                  '  Array contains indexes for each chosen colorset
''               Exit Do                                     'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
''            End If
''            rsColorSets.MoveNext
''         Loop
''         rsColorSet.MoveNext
''      Loop
''      colorSets(0) = i                   'colorSets(0) contains the number of Color Sets to display
'      If colorSets(0) <= 1 Then
'         valueIndex = 0                                                    'Value0 in Display table
'      Else
'         valueIndex = Val(Left(lstColorSets.Tag, InStr(lstColorSets.Tag, "\") - 1))
'      End If
      
'      expressionIndex = -1
'      Do Until rsColorSets.EOF
'         expressionIndex = expressionIndex + 1
'         If rsColorSet!colorSet = rsColorSets!colorSet Then
'            Exit Do                                        'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'         End If
'         rsColorSets.MoveNext
'      Loop
''      If Not rsColorSets.EOF Then '=================================================Valid Color Set
         '---------------------------------------------------------Determine Value Index To Display
         If cfgColoring <> "S" Then '----------------------------------------Determine Systems List
            Set rsInfo = dbExpression.OpenRecordset("SELECT SystemCodes FROM Info")
            systemCodes = Dat(rsInfo!systemCodes)
         End If
         For Each element In objLumps '---------------------------------------------------Each Gene
            If element.objType = "Gene" Then
               element.DrawGeneValue drawingBoard, False
'If element.title = "Cdkn1a" Then Stop
               ExpressionDisplay element, dbGene, dbExpression, colorIndexes, valueIndex, _
                                 systemCodes
               element.DrawObj
            End If
         Next element
         legend.Create
''      Else
''         legend.DrawObj False, drawingBoard                        'Invalid Color Set, erase legend
''      End If
   End If
   
'      '=========================================Determine Expression Value Column for Display Table
'      Set rsColorSets = dbExpression.OpenRecordset("SELECT * FROM ColorSet")
'      expressionIndex = -1
'      Do Until rsColorSets.EOF
'         expressionIndex = expressionIndex + 1
'         If rsColorSet!colorSet = rsColorSets!colorSet Then
'            Exit Do                                        'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'         End If
'         rsColorSets.MoveNext
'      Loop
'      If Not rsColorSets.EOF Then '=================================================Valid Color Set
'         If cfgColoring <> "S" Then '----------------------------------------Determine Systems List
'            Set rsInfo = dbExpression.OpenRecordset("SELECT SystemCodes FROM Info")
'            systemCodes = Dat(rsInfo!systemCodes)
'         End If
'         For Each element In objLumps '---------------------------------------------------Each Gene
'            If element.objType = "Gene" Then
'               element.DrawGeneValue drawingBoard, False
'               ExpressionDisplay element, dbGene, dbExpression, expressionIndex, systemCodes
'               element.DrawObj
'            End If
'         Next element
'         legend.Create
'      Else
'         legend.DrawObj False, drawingBoard                        'Invalid Color Set, erase legend
'      End If
'   End If
   
'   If callingFunction <> "mnuRedraw" Then mnuRedraw_Click
   MousePointer = vbDefault
End Sub
Sub DisplaySingleGene(obj As objLump) '**************************** Complete Display of Single Gene
   '  Enter    obj   Gene object
   '  Use only for one gene at a time. mnuApply does all genes on a MAPP more efficiently
   '  because it determines the expressionIndex and systemCodes once for the whole set.
   Dim rsColorSets As Recordset, rsInfo As Recordset
   Dim expressionIndex As Integer
   Dim systemCodes As String
   
   If obj.centerSystemCode = "" Then '+++++++++++++++++++++++++++++++++++++++++++ Gene Not Assigned
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If Not dbExpression Is Nothing And colorIndexes(0) <> 0 Then          'Expression Dataset exists
      Set rsColorSets = dbExpression.OpenRecordset("SELECT * FROM ColorSet")
      expressionIndex = -1
      Do Until rsColorSets.EOF
         expressionIndex = expressionIndex + 1
         If rsColorSet!colorSet = rsColorSets!colorSet Then
            Exit Do                                        'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
         rsColorSets.MoveNext
      Loop
      If Not rsColorSets.EOF Then '=================================================Valid Color Set
         If cfgColoring <> "S" Then '----------------------------------------Determine Systems List
            Set rsInfo = dbExpression.OpenRecordset("SELECT SystemCodes FROM Info")
            systemCodes = Dat(rsInfo!systemCodes)
         End If
         ExpressionDisplay obj, dbGene, dbExpression, colorIndexes, valueIndex, systemCodes
      Else '===========================================================================No Color Set
         '  This should also be set by ExpressionDisplay ????????????????????????????
         obj.value = ""
         obj.color(0) = i
         obj.color(1) = vbWhite
         obj.rim(1) = vbWhite
         obj.centerOrderNo = -1
         obj.rimOrderNo = -1
         obj.centerSystemCode = ""
         obj.rimSystemCode = ""
      End If
      obj.DrawObj
   End If
End Sub
Private Sub mnuSize_Click() '*********************************************************** Size Genes
   '  Sizes all selected genes to the width of the largest selected gene
   Dim obj As Object, maxWidth As Single
   
   If selections.count = 0 Then Exit Sub  'Nothing selected  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   For Each obj In selections '====================================================Find Widest Gene
      If obj.objType = "Gene" Then
         If obj.wide > maxWidth Then maxWidth = obj.wide
      End If
   Next obj
   If maxWidth = 0 Then GoTo ExitSub    'No genes actually found  vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   For Each obj In selections '===============================================Set All To That Width
      If obj.objType = "Gene" Then
         obj.wide = maxWidth
      End If
   Next obj
   mnuRedraw_Click
   dirty = True
ExitSub:
   MousePointer = vbDefault
End Sub
Private Sub mnuHorizAlign_Click() '************************************* Horizontally Align Objects
   '  Horizontally aligns object centers with the leftmost object found
   Dim obj As Object, newCenter As Single, leftMost As Single
   
   If selections.count = 0 Then Exit Sub  'Nothing selected  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   MousePointer = vbHourglass
   leftMost = 10000000000#   'Start leftMost at some huge figure, certainly off right side of board
   For Each obj In selections '================================================Find leftMost Object
      If Not TypeOf obj Is objLine Then                                 'Lines have no center point
         If obj.centerX < leftMost Then
            leftMost = obj.centerX
            newCenter = obj.centerY
         End If
      End If
   Next obj
   If leftMost = 10000000000# Then GoTo ExitSub  'No objects actually found  vvvvvvvvvvvvvvvvvvvvvv
   For Each obj In selections '==============================================Center All To leftMost
      If Not TypeOf obj Is objLine Then
         obj.centerY = newCenter
      End If
   Next obj
   mnuRedraw_Click
   dirty = True
ExitSub:
   MousePointer = vbDefault
End Sub

Private Sub mnuUndo_Click()
   newObj.DrawObj False
   newObj.Restore
   newObj.SelectMode = False
   newObj.SetEdit False
   newObj.DrawObj
   MultipleObjectDeselectAll
   CancelUndo
End Sub

Private Sub mnuVertAlign_Click() '**************************************** Vertically Align Objects
   '  Vertically aligns object centers with the topmost object found
   Dim obj As Object, newCenter As Single, Topmost As Single
   
   If selections.count = 0 Then Exit Sub  'Nothing selected  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   MousePointer = vbHourglass
   Topmost = 10000000000#         'Start topMost at some huge figure, certainly off bottom of board
   For Each obj In selections '=================================================Find Topmost Object
      If Not TypeOf obj Is objLine Then                                 'Lines have no center point
         If obj.centerY < Topmost Then
            Topmost = obj.centerY
            newCenter = obj.centerX
         End If
      End If
   Next obj
   If Topmost = 10000000000# Then GoTo ExitSub  'No objects actually found  vvvvvvvvvvvvvvvvvvvvvvv
   For Each obj In selections '===============================================Center All To Topmost
      If Not TypeOf obj Is objLine Then
         obj.centerX = newCenter
      End If
   Next obj
   mnuRedraw_Click
   dirty = True
ExitSub:
   MousePointer = vbDefault
End Sub
Private Sub mnuBlock_Click() '************************************************ Align Genes In Block
   '  Make selected genes size of the largest and align in a block below topmost
   Dim obj As Object, geneObj As objLump, newCenter As Single, Topmost As Single, index As Integer
   Dim genes As New Collection, geneAdded As Boolean
   
   If selections.count < 2 Then GoTo ExitSub  'Less than 2 selected vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   
   MousePointer = vbHourglass
   Topmost = 10000000000#         'Start topMost as some huge figure, certainly off bottom of board
   For Each obj In selections '======================================Fill Genes Collection In Order
      If obj.objType = "Gene" Then
         If obj.centerY < Topmost Then
            Topmost = obj.centerY
            newCenter = obj.centerX
         End If
         geneAdded = False
         index = 1
         Do While index <= genes.count 'Add to temp collection to display them ordered on top below
            If obj.centerY <= genes(index).centerY Then
               genes.Add obj, , index
               geneAdded = True
               Exit Do                                             'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
            index = index + 1
         Loop
         If Not geneAdded Then genes.Add obj                                            'Add to end
      End If
      
   Next obj
   
   If genes.count < 2 Then GoTo ExitSub       'Less than 2 genes found vvvvvvvvvvvvvvvvvvvvvvvvvvvv
   
   mnuSize_Click
   
   For index = 2 To genes.count '===================================Place Starting With Second Gene
      genes(index).centerX = newCenter
      genes(index).centerY = genes(index - 1).centerY + genes(index - 1).high / 2 _
                           + genes(index).high / 2
   Next index
   mnuRedraw_Click
   MousePointer = vbHourglass
   For index = 1 To genes.count '=====================================================Redraw On Top
      genes(index).DrawObj
   Next index
   dirty = True
ExitSub:
   MousePointer = vbDefault
End Sub
Public Sub mnuManager_Click()
   frmExpression.show vbModal
   If Not dbExpression Is Nothing And Not rsColorSet Is Nothing Then
      mnuApply_Click
   End If
End Sub

Private Sub mnuGenMAPPHelp_Click()
   Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, appPath & "\GenMAPP.chm::/GenMAPP.htm", HH_DISPLAY_TOPIC, 0)
End Sub
Private Sub mnuAboutGenMAPP_Click()
   frmAbout.show vbModal
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer) '**************************** Keyboard
   Dim deleteID As String, index As Integer, obj As Object
   Dim newObj As Object, newObjects As New Collection
   
   Select Case KeyCode
   Case vbKeyReturn
      If Screen.ActiveControl.name = "cmbZoom" Then
'         cmbZoom_LostFocus
      End If
   Case vbKeyRight, vbKeyLeft, vbKeyUp, vbKeyDown '======================================Nudge Keys
      If shpSelected.visible Then
         Dim moveX As Single, moveY As Single
         Select Case KeyCode
         Case vbKeyRight
            moveX = GRID_SIZE
         Case vbKeyLeft
            moveX = -GRID_SIZE
         Case vbKeyUp
            moveY = -GRID_SIZE
         Case vbKeyDown
            moveY = GRID_SIZE
         End Select
         shpSelected.Left = shpSelected.Left + moveX
         shpSelected.Top = shpSelected.Top + moveY
         MoveSelection moveX, moveY
      End If
   Case vbKeyDelete, vbKeyBack '=====================================================Delete Objects
      If activeObject Is Nothing Then '+++++++++++++++++++++++++++++++++++++Delete Multiple Objects
         For Each obj In selections
            ObjDelete obj
         Next obj
         MultipleObjectDeselectAll
'         For index = 1 To selections.Count '-----------------Deselect All Multiple Selections
''            selections(1).SelectMode = False
'            selections.Remove 1
'         Next index
      Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Delete Single Objects
         dirty = True                                   'For controls, doesn't involve Place method
         ObjDelete activeObject
      End If
      mnuRedraw_Click
   Case vbKeyInsert '=============================================================Duplicate Objects
      If activeObject Is Nothing Then '++++++++++++++++++++++++++++++++++Duplicate Multiple Objects
         For Each obj In selections
''            selecting = True
            Set newObj = ObjDuplicate(obj)
            newObjects.Add newObj, newObj.objType & newObj.objKey
            Set newObj = Nothing                            'Only reference should be in collection
         Next obj
         For index = 1 To selections.count '-----------------------Deselect All Multiple Selections
            selections(1).SelectMode = False
            selections.Remove 1
         Next index
         For Each obj In newObjects '--------------------------------------------Select New Objects
            obj.SelectMode = True                         'Must draw after deselecting to be on top
            selections.Add obj, obj.objType & obj.objKey
         Next obj
         For index = 1 To newObjects.count '-----------------------------Dump NewObjects Collection
            newObjects.Remove 1
         Next index
      Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Duplicate Single Objects
         Set newObj = ObjDuplicate(activeObject)
         SetActiveObject newObj
         Set newObj = Nothing                               'Only reference should be in collection
      End If
      MousePointer = vbDefault
   End Select
End Sub
Public Sub Form_Resize()
   Static resizing As Boolean, wide As Single, high As Single
   
   If resizing Then Exit Sub
   If WindowState = vbMinimized Then Exit Sub
   If resizing Then Exit Sub
   If loading Then Exit Sub
   If callingRoutine = "Don't resize" Then Exit Sub
   If callingRoutine = "ScrollBars" Then Exit Sub
   If callingRoutine = "mnuBoardSize_Click" Then Exit Sub
   
   resizing = True
   
   If WindowState = vbMaximized Then '++++++++++++++++++++++++++++++++++++++++++++++++++++ Maximize
      wide = ClientWide
      high = ClientHigh
      WindowState = vbNormal
      If picDrafter.Width > wide Then
         If picDrafter.Width + picDrafter.Left < wide Then
            picDrafter.Left = wide - picDrafter.Width
         End If
         FormWidth wide
      Else
         picDrafter.Left = 0
         FormWidth picDrafter.Width
      End If
      If picDrafter.Height > high Then
         If picDrafter.Height + (picDrafter.Top - tlbTools.Height) < high Then
            picDrafter.Top = high - picDrafter.Height + tlbTools.Height
         End If
         FormHeight high
      Else
         picDrafter.Top = tlbTools.Height
         FormHeight picDrafter.Height
      End If
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Form Dimensions To At Least Minimum
   If ClientWide() < MIN_BOARD_WIDTH Then
      FormWidth MIN_BOARD_WIDTH                                 'Set form from desired client width
   End If
   If ClientHigh() < MIN_BOARD_HEIGHT Then
      FormHeight MIN_BOARD_HEIGHT
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Board Dimensions
   '  Drafting board (picDrafter) right and bottom is always a minimum of what it was or the
   '  edge of the resized client window. If the client window is expanded beyond the right, bottom
   '  of the drafting board, the window is reduced to a size that does not increase the drafting
   '  board size.
   If ClientWide > picDrafter.Width + picDrafter.Left Then '======Window size > drafting board edge
      If ClientWide > picDrafter.Width Then '----------------Screen Pulled To Size > Drafting Board
         picDrafter.Left = 0
         FormWidth picDrafter.Width
      Else '--------------------------------------------------Drafting Board Will Fit On New Screen
         picDrafter.Left = ClientWide - picDrafter.Width
      End If
   End If
   If ClientHigh > picDrafter.Height + picDrafter.Top - tlbTools.Height Then
      If ClientHigh > picDrafter.Height Then '---------------Screen Pulled To Size > Drafting Board
         picDrafter.Top = tlbTools.Height
         FormHeight picDrafter.Height
      Else '--------------------------------------------------Drafting Board Will Fit On New Screen
         picDrafter.Top = ClientHigh - picDrafter.Height + tlbTools.Height
      End If
   End If

ExitSub:
   ScrollBars
   resizing = False
End Sub
Public Function ClientWide()  '*********************************************** Usable Width Of Form
   ClientWide = ScaleWidth - vsbDrafter.Width                         'ScaleWidth is interior width
End Function
Public Function ClientHigh() As Single '************************************* Usable Height Of Form
   ClientHigh = ScaleHeight - tlbTools.Height - hsbDrafter.Height _
                - sbrBar.Height
End Function
Public Sub FormWidth(cWidth As Single) '**************** Sets Total Width Of Form From Client Width
   Width = cWidth + vsbDrafter.Width + Width - ScaleWidth
      '  Width - ScaleWidth adjusts for the borders
End Sub
Public Sub FormHeight(cHeight As Single) '************ Sets Total Height Of Form From Client Height
   Height = cHeight + tlbTools.Height + hsbDrafter.Height + sbrBar.Height + Height - ScaleHeight
End Sub
Private Sub Form_Deactivate()
'   SetNewObject Nothing
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '************** Close Button
'MsgBox "Unload mode: " & UnloadMode
   If UnloadMode = vbFormCode Then                                            'Exiting from mnuExit
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   cancelExit = False
   mnuExit_Click
   If cancelExit Then Cancel = True                                'User said Cancel to Save msgbox
End Sub

Rem///////////////////////////////////////////////////////////////////////////// Scroll Bar Actions
'  All scroll bar parameters (Max, SmallChange, LargeChange, and Value) divided by 10. They are
'     integers and thus limited to 32767, too small for huge Drafting Boards

Private Sub hsbDrafter_Change()
    picDrafter.Left = Min(-hsbDrafter.value * 10#, 0)  '10# to force single, not integer, arithmetic
    picDrafter.Left = Min(picDrafter.Left, picDrafter.Width - ClientWide())
End Sub

Private Sub vsbDrafter_Change()
   picDrafter.Top = Min(-vsbDrafter.value * 10# + tlbTools.Height, tlbTools.Height)
                                                       '10# to force single, not integer, arithmetic
   picDrafter.Top = Min(picDrafter.Top, picDrafter.Height - ClientHigh() + tlbTools.Height)
End Sub
Public Sub ScrollBars() '********************************************* Sets the Drafter Scroll Bars
   callingRoutine = "ScrollBars"

   hsbDrafter.Min = 0
   hsbDrafter.Top = ClientHigh() + tlbTools.Height
   
   hsbDrafter.Width = ClientWide()
   '  All this stuff must be divided by 10 because the scrollbars use integer numbers, max 32767
   hsbDrafter.Max = (picDrafter.Width - hsbDrafter.Width) / 10
      '  Max is the difference between the ClientWide (hsbDrafter.Width) and the width of the
      '  drafting board (picDrafter.Width). It is the amount of scroll available.
      '  Eg: If the client width of the ddrafter window is 100 and the drafting board width
      '  is 500, Max is 400.
If picDrafter.Left > 0 Then
   '  For some reason picDrafter.Left is unpredictably coming out positive.
'Stop
   picDrafter.Left = 0
End If
   hsbDrafter.value = -picDrafter.Left / 10
      '  Value is how far into the scroll the drafting board is. It can be guaged by how far
      '  to the left of the drafter window the drafting board begins (-picDrafter.Left).
      '  Eg: if the value is 150, there should be 250 left to scroll.
   hsbDrafter.LargeChange = hsbDrafter.Width / 11                                    'Allow overlap
      '  Must be at least 1
            'hsbDrafter.width is visible width of screen (ClientWide).
   hsbDrafter.SmallChange = Min(5 * GRID_SIZE / 10, hsbDrafter.LargeChange)

   vsbDrafter.Min = 0
   vsbDrafter.Top = tlbTools.Height
   vsbDrafter.Left = ClientWide()
   vsbDrafter.Height = ClientHigh()
   vsbDrafter.Max = (picDrafter.Height - vsbDrafter.Height) / 10
   If vsbDrafter.Max < 0 Then vsbDrafter.Max = 0
'      vsbDrafter.value = 0
   vsbDrafter.value = Min((-picDrafter.Top + tlbTools.Height) / 10, vsbDrafter.Max)
      '  To adjust for any approximation error in picDrafter.top
   vsbDrafter.LargeChange = vsbDrafter.Height / 11                                   'Allow overlap
   vsbDrafter.SmallChange = Min(5 * GRID_SIZE / 10, vsbDrafter.LargeChange)

   callingRoutine = ""
End Sub

'************************************************************************* Edit Point Mouse Actions
Private Sub picPoint_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   picPoint(index).Tag = "Moving"
End Sub
Private Sub picPoint_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If picPoint(index).Tag <> "Moving" Then Exit Sub
   Select Case activeObject.objType
   Case "objBrace"
      If index = 1 Then                                                                 'Move brace
         picPoint(index).Left = GridCoord(picPoint(index).Left + X) - POINT_SIZE / 2
         picPoint(index).Top = GridCoord(picPoint(index).Top + Y) - POINT_SIZE / 2
      Else                   'Change span. Edit point can only move parallel to brace's orientation
         Select Case activeObject.Orientation
         Case TOP_BRACE, BOTTOM_BRACE
            If picPoint(index).Left + X >= activeObject.centerX + 100 Then    'Don't make too small
               picPoint(index).Left = GridCoord(picPoint(index).Left + X) - POINT_SIZE / 2
            End If
         Case Else
            If picPoint(index).Top + Y <= activeObject.centerY - 100 Then     'Don't make too small
               picPoint(index).Top = GridCoord(picPoint(index).Top + Y) - POINT_SIZE / 2
            End If
         End Select
      End If
   Case Else
      picPoint(index).Left = GridCoord(picPoint(index).Left + X) - POINT_SIZE / 2
      picPoint(index).Top = GridCoord(picPoint(index).Top + Y) - POINT_SIZE / 2
   End Select
End Sub
Private Sub picPoint_MouseUp(index As Integer, Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   '  Always at grid coordinate because MouseMove snaps picpoint to coordinate and this event
   '     uses last location of picPoint, not the X and Y that comes in
   '  Sends to object Move method in nonzoomed coordinates
   Dim element As Variant, pointNo As Integer
   
   pointNo = index
   If picPoint(index).Tag = "Moving" Then
      If Shift = vbShiftMask Then pointNo = -pointNo                          'Special move options
'      If activeObject.objType = "objLine" Then
'         If activeObject.style = "Arc" Then                  'Orient curves on 90 degree increments
'            If index = 2 Then                                                            'End point
'               If Abs(picPoint(2).Left - picPoint(1).Left) < Abs(picPoint(2).Top - picPoint(1).Top) Then
'                  picPoint(2).Left = picPoint(1).Left
'               Else
'                  picPoint(2).Top = picPoint(1).Top
'               End If
'            Else
'               If Abs(picPoint(2).Left - picPoint(1).Left) < Abs(picPoint(2).Top - picPoint(1).Top) Then
'                  picPoint(1).Left = picPoint(2).Left
'               Else
'                  picPoint(1).Top = picPoint(2).Top
'               End If
'            End If
'         End If
'      End If
      activeObject.Move picPoint(index).Left / zoom + POINT_SIZE / 2, _
            picPoint(index).Top / zoom + POINT_SIZE / 2, pointNo               'Center of point box
         '  This reports in nonzoomed coordinates since object coordinates are always
         '  stored nonzoomed.
      mnuRedraw_Click
   End If
   picPoint(index).Tag = ""
End Sub

Rem //////////////////////////////////////////////////////////////// Individual Object Mouse Events
'  Events here treated differently for different objects

Private Function ObjClicked() As Object '******************************* Finds Which Object Clicked
   '  mouseX and mouseY are both in board coordinates. They come from MouseUp
   
   Dim element As Variant, index As Integer
   
   Set ObjClicked = Nothing                                           'Deactivate any active object
   
   If legend.CheckClick(mouseX, mouseY) Then '++++++++++++++++++++++++++++++++++++++++++ Check Info
      Set ObjClicked = legend
      Exit Function                                 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If info.CheckClick(mouseX, mouseY) Then '++++++++++++++++++++++++++++++++++++++++++++ Check Info
      Set ObjClicked = info
      Exit Function                                 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   For index = objLumps.count To 1 Step -1 '+++++++++++++++++++++++++++++++++++++++++++ Check Lumps
      ' Go backward thru collection to check topmost first
      If objLumps(index).CheckClick(mouseX, mouseY) Then
         Set ObjClicked = objLumps(index)
         Exit Function                              '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Next index
   
   For Each element In objLines '++++++++++++++++++++++++++++++++++++++++++++++++++ Check All Lines
      If element.CheckClick(mouseX, mouseY) Then
         Set ObjClicked = element
         Exit Function                              '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Next element
   
   Set ObjClicked = Nothing
End Function

Rem /////////////////////////////////////////////////////////////////// Actions On Specific Objects
Public Sub SetActiveObject(obj As Object) '**************************** Set Or Change Active Object
   '  All control of active object should be done through here
   '  Only one object can be active at any one time
   '  Active means moving or sizing
   '  Activated by double clicking object
   '  Inactivate by
   '     Activating another object
   '     Choosing a new object from tlbTools
   '     Clicking or double clicking on frmDrafter
   '     Multiple selection of objects
   '  Each object must have a public objType property
   
   If Not activeObject Is Nothing Then                          'Something else active, turn it off
      If Exists(activeObject) Then                             'Object itself may have been deleted
         activeObject.SetEdit False
         frmDrafter.sbrBar.Panels(1).text = ""
      End If
   End If
   Set activeObject = obj
   If Not activeObject Is Nothing Then                                           'New active object
      If TypeOf activeObject Is objLump Then
         objLumps.Remove activeObject.objKey         'Put at end of collection so it appears on top
         objLumps.Add activeObject, activeObject.objKey
      End If
      activeObject.SetEdit
'      HitRange activeObject
      MultipleObjectDeselectAll
      Select Case activeObject.objType
      Case "objLine"
         frmDrafter.sbrBar.Panels("Instructions").text = statusBarSelect(activeObject.style)
      Case Else
         frmDrafter.sbrBar.Panels("Instructions").text = statusBarSelect(activeObject.objType)
      End Select
   End If
End Sub
Sub ObjDelete(obj As Object) '**************************************************** Delete An Object
   dirty = True
   obj.DrawObj False                                                             'Erase from screen
   If TypeOf obj Is objLine Then
      objLines.Remove obj.objKey                  'Remove from collection. Should be last reference
   ElseIf TypeOf obj Is objLump Then
      objLumps.Remove obj.objKey                  'Remove from collection. Should be last reference
   End If
   SetActiveObject Nothing                                       'Remove any activeObject reference
End Sub
Function ObjDuplicate(obj As Object) As Object '******************************* Duplicate An Object
   dirty = True
   If TypeOf obj Is objLine Then
      Dim newLine As New objLine
      newLine.Duplicate obj
      objLines.Add newLine, newLine.objKey
      Set ObjDuplicate = newLine
   ElseIf TypeOf obj Is objLump Then
      Dim newLump As New objLump
      newLump.Duplicate obj
      objLumps.Add newLump, newLump.objKey
      Set ObjDuplicate = newLump
   End If
End Function
Public Sub MultipleObjectsSelected() '********************************** Multiple Object Selections
   '  Enables formatting menu items if more than 1 object selected
   '  Disables them if not
   '  Only 2 places where multiple selections can be made, both in frmDrafter!MouseUp:
   '     shpSelect has size
   '     ctrl-click
   Dim element As Object, genes As Integer
   
   If selections.count = 0 Then
      frmDrafter.shpSelected.visible = False
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Select Area And Box
   selectArea.minX = 1E+38                                          'Set to be replaced immediately
   selectArea.minY = 1E+38
   selectArea.maxX = 0
   selectArea.maxY = 0
   For Each element In selections
      selectArea.minX = Min(selectArea.minX, element.minX)
      selectArea.minY = Min(selectArea.minY, element.minY)
      selectArea.maxX = Max(selectArea.maxX, element.maxX)
      selectArea.maxY = Max(selectArea.maxY, element.maxY)
      If Left(element.objType, 4) = "Gene" Then                     'Count number of genes selected
         genes = genes + 1
      End If
   Next element
'   selectArea.origX = selectArea.minX
'   selectArea.origY = selectArea.minY
   shpSelected.Left = selectArea.minX * zoom
   shpSelected.Top = selectArea.minY * zoom
   shpSelected.Width = (selectArea.maxX - selectArea.minX) * zoom
   shpSelected.Height = (selectArea.maxY - selectArea.minY) * zoom
   shpSelected.visible = True
   
   If selections.count < 2 Then '+++++++++++++++++++++++++++++++++++++ Less Than 2 Objects Selected
      '  These will all be mdiDrafter items in V3
      frmDrafter.mnuVertAlign.Enabled = False
      frmDrafter.mnuHorizAlign.Enabled = False
      frmDrafter.mnuSize.Enabled = False
      frmDrafter.mnuBlock.Enabled = False
   Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 2 Or More Objects Selected
      frmDrafter.mnuVertAlign.Enabled = True
      frmDrafter.mnuHorizAlign.Enabled = True
      If genes > 1 Then                                                   '2 or More genes selected
         frmDrafter.mnuSize.Enabled = True
         frmDrafter.mnuBlock.Enabled = True
      End If
   End If
   frmDrafter.sbrBar.Panels(1).text = "Click outside of green box to deselect all objects"
End Sub
Public Sub MultipleObjectDeselectAll() '************************* Deselects All Multiple Selections
   '  Deselects all multiple-selected objects (selections collection)
   '  If nothing in multiple selection, exits immediately avoiding redraw
   Dim i As Integer
   
   If selections.count Then
      For i = 1 To selections.count                               'Deselect all multiple selections
         selections(1).SelectMode = False
         selections.Remove 1
      Next i
      shpSelected.visible = False
      mnuRedraw_Click
   End If
   frmDrafter.sbrBar.Panels(1).text = ""
End Sub
Public Function MinBoardWidth() As Single '*********************** Min Width To Accommodate Objects
   Dim element As Object, wide As Single
   
   wide = 0
   For Each element In objLines
      wide = Max(wide, element.maxX)
   Next element
   For Each element In objLumps
      wide = Max(wide, element.maxX)
   Next element
   wide = Max(wide, info.maxX)
   wide = Max(wide, legend.maxX)
   wide = Max(wide, MAPPTitle.wide)
   wide = Round(wide / TWIPS_CM + 0.05, 1) * TWIPS_CM
      '  Raise these to the next higher tenth of a cm so that an exact board width in twips does
      '  not end up smaller than the value in cm.
   MinBoardWidth = Max(wide + 20, MIN_BOARD_WIDTH)                      '+20 to clear lines at edge
End Function
Public Function MinBoardHeight() As Single '********************* Min Height To Accommodate Objects
   Dim element As Object, high As Single
   
   high = 0
   For Each element In objLines
      high = Max(high, element.maxY)
   Next element
   For Each element In objLumps
      high = Max(high, element.maxY)
   Next element
   high = Max(high, info.maxY)
   high = Max(high, legend.maxY)
   high = Round(high / TWIPS_CM + 0.05, 1) * TWIPS_CM
      '  Raise these to the next higher tenth of a cm so that an exact board width in twips does
      '  not end up smaller than the value in cm.
   MinBoardHeight = Max(high + 20, MIN_BOARD_HEIGHT)
End Function

