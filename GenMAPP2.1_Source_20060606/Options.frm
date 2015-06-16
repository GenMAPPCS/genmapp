VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOther 
      BorderStyle     =   0  'None
      Caption         =   "Legend"
      Height          =   4632
      Left            =   60
      TabIndex        =   21
      Top             =   420
      Visible         =   0   'False
      Width           =   8892
      Begin VB.CheckBox chkAutoUpdate 
         Caption         =   "Automatically check for program updates"
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
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Each time you start GenMAPP, it will check for program updates and give you the opportunity to apply them."
         Top             =   120
         Width           =   4272
      End
   End
   Begin VB.Frame fraExport 
      BorderStyle     =   0  'None
      Caption         =   "Legend"
      Height          =   4632
      Left            =   240
      TabIndex        =   20
      Top             =   6060
      Visible         =   0   'False
      Width           =   8892
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
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
      Left            =   60
      TabIndex        =   19
      Top             =   5280
      Width           =   1332
   End
   Begin VB.Frame fraLegend 
      BorderStyle     =   0  'None
      Caption         =   "Legend"
      Height          =   4092
      Left            =   60
      TabIndex        =   6
      Top             =   420
      Visible         =   0   'False
      Width           =   8712
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   435
         Left            =   60
         TabIndex        =   23
         Top             =   2820
         Width           =   2175
         Begin VB.OptionButton optLegendAllColorSets 
            Caption         =   "Show all"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optLegendFirstColorSet 
            Caption         =   "Show only first"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.ComboBox cmbLegendFontSize 
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
         ItemData        =   "Options.frx":08CA
         Left            =   1080
         List            =   "Options.frx":08E9
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Font size for the Legend (heads will be larger)."
         Top             =   2160
         Width           =   672
      End
      Begin VB.CheckBox chkShowInfo 
         Caption         =   "Show Information Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3720
         TabIndex        =   15
         Top             =   120
         Width           =   2772
      End
      Begin VB.CheckBox chkShowLegend 
         Caption         =   "Show Legend"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   60
         Width           =   2052
      End
      Begin VB.CheckBox chkGeneDB 
         Caption         =   "Gene Database name"
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
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   2712
      End
      Begin VB.CheckBox chkExpression 
         Caption         =   "Expression Dataset name"
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
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   2712
      End
      Begin VB.CheckBox chkColorSet 
         Caption         =   "Color Set name"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   2712
      End
      Begin VB.CheckBox chkValue 
         Caption         =   "Name of Gene Value column"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1380
         Width           =   3192
      End
      Begin VB.CheckBox chkRemarks 
         Caption         =   "Remarks"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1620
         Width           =   2712
      End
      Begin VB.CheckBox chkColors 
         Caption         =   "Colors and criteria"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1860
         Width           =   2712
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color Sets on Legend"
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
         Index           =   2
         Left            =   60
         TabIndex        =   26
         Top             =   2580
         Width           =   2250
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Font size"
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
         TabIndex        =   16
         Top             =   2220
         Width           =   816
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display on Legend"
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
         TabIndex        =   14
         Top             =   420
         Width           =   1860
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   6360
      TabIndex        =   5
      Top             =   5280
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Height          =   432
      Left            =   7800
      TabIndex        =   4
      Top             =   5280
      Width           =   1332
   End
   Begin VB.Frame fraColoring 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Coloring"
      ForeColor       =   &H80000008&
      Height          =   1572
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   8112
      Begin VB.CheckBox chkOpenColored 
         Caption         =   "Always open GenMAPP with most-recently-used Expression Dataset loaded"
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
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   7452
      End
      Begin VB.OptionButton optSpecific 
         Caption         =   "Exact Match to Gene ID"
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
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Gene object will only be colored when the gene ID on the MAPP is identical to the gene ID in the Expression Dataset."
         Top             =   420
         Width           =   2532
      End
      Begin VB.OptionButton optRelated 
         Caption         =   "All Related Gene IDs"
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
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Gene object will be colored based on data from all gene IDs related to the gene ID on the MAPP and in the Expression Dataset."
         Top             =   120
         Value           =   -1  'True
         Width           =   2292
      End
   End
   Begin MSComctlLib.TabStrip tabOptions 
      Height          =   5112
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   9022
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Coloring"
            Key             =   "Coloring"
            Object.ToolTipText     =   "Specify the coloring scheme for genes."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Legend"
            Key             =   "Legend"
            Object.ToolTipText     =   "Specify the items you want displayed in your Legend."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Other"
            Key             =   "Other"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim coloringDirty As Boolean
Dim legendDirty As Boolean
Dim infoDirty As Boolean

Private Sub cmdOK_Click()
   Hide
   If coloringDirty Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Coloring Options
      If optSpecific.value Then
         cfgColoring = "S"
      Else
         cfgColoring = "R"
      End If
      mappWindow.mnuApply_Click
   End If
   
   cfgOptions = "" '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Various Options
   If chkOpenColored = vbChecked Then
      cfgOptions = cfgOptions & "C"
   End If
   If chkAutoUpdate = vbChecked Then
      cfgCheckForUpdatesOnStart = "True"
   Else
      cfgCheckForUpdatesOnStart = "False"
   End If
   
   If legendDirty Or infoDirty Then '++++++++++++++++++++++ Legend Display Characteristics Returned
      cfgLegend = ""
      If chkShowLegend = vbChecked Then cfgLegend = cfgLegend & "D"
      If chkGeneDB = vbChecked Then cfgLegend = cfgLegend & "G"
      If chkExpression = vbChecked Then cfgLegend = cfgLegend & "E"
      If chkColorSet = vbChecked Then cfgLegend = cfgLegend & "C"
      If chkValue = vbChecked Then cfgLegend = cfgLegend & "V"
      If chkRemarks = vbChecked Then cfgLegend = cfgLegend & "R"
      If chkColors = vbChecked Then cfgLegend = cfgLegend & "L"
      If chkShowInfo = vbChecked Then cfgLegend = cfgLegend & "I"
      cfgLegend = cfgLegend & "F" & cmbLegendFontSize.List(cmbLegendFontSize.ListIndex) & "|"
   End If
   
   If legendDirty Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Redraw Legend
      If chkShowLegend = vbChecked Then
         mappWindow.legend.DrawObj
      Else
         mappWindow.legend.DrawObj False
      End If
   End If
   If infoDirty Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Redraw Info
      If chkShowInfo = vbChecked Then
         mappWindow.info.DrawObj
      Else
         mappWindow.info.DrawObj False
      End If
   End If
End Sub

Private Sub cmdHelp_Click()
   Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, appPath & "\GenMAPP.chm::/Options.htm", HH_DISPLAY_TOPIC, 0)
End Sub

Private Sub Form_Activate()
   FramesInvisible
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Coloring Options Initialization
   If cfgColoring = "S" Then
      optSpecific.value = True
   Else
      optRelated.value = True
   End If

   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Legend Display
   '  Display characters assigned to lblDisplay before form activated
   '  Existence of characters determines what to display
   '  See global declaration under config items in GenMAPP.bas
   
   If InStr(cfgLegend, "D") Then
      chkShowLegend = vbChecked
   Else
      chkShowLegend = vbUnchecked
   End If
   If InStr(cfgLegend, "G") Then
      chkGeneDB = vbChecked
   Else
      chkGeneDB = vbUnchecked
   End If
   If InStr(cfgLegend, "E") Then
      chkExpression = vbChecked
   Else
      chkExpression = vbUnchecked
   End If
   If InStr(cfgLegend, "C") Then
      chkColorSet = vbChecked
   Else
      chkColorSet = vbUnchecked
   End If
   If InStr(cfgLegend, "V") Then
      chkValue = vbChecked
   Else
      chkValue = vbUnchecked
   End If
   If InStr(cfgLegend, "R") Then
      chkRemarks = vbChecked
   Else
      chkRemarks = vbUnchecked
   End If
   If InStr(cfgLegend, "L") Then
      chkColors = vbChecked
   Else
      chkColors = vbUnchecked
   End If
   If InStr(cfgLegend, "I") Then
      chkShowInfo = vbChecked
   Else
      chkShowInfo = vbUnchecked
   End If
   s = Mid(cfgLegend, InStr(cfgLegend, "F") + 1, InStr(cfgLegend, "|") - InStr(cfgLegend, "F") - 1)
   For i = 0 To cmbLegendFontSize.ListCount - 1
      If cmbLegendFontSize.List(i) = s Then
         cmbLegendFontSize.ListIndex = i
         Exit For
      End If
   Next i
   If i > cmbLegendFontSize - 1 Then cmbLegendFontSize.ListIndex = 2                          '8 pt
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Various Options
   If InStr(cfgOptions, "C") Then chkOpenColored = vbChecked
   If cfgCheckForUpdatesOnStart = "True" Then
      chkAutoUpdate = vbChecked
   Else
      chkAutoUpdate = vbUnchecked
   End If
   If cfgLegendAllColorSets Then
      frmOptions.optLegendAllColorSets = True
   Else
      frmOptions.optLegendFirstColorSet = True
   End If
   
         
   Select Case Tag '++++++++++++++++++++++++++++++++ Form Activated By Right Click Rather Than Menu
   Case "Legend"
      tabOptions.Tabs("Legend").selected = True
      fraLegend.visible = True
   Case Else
'      tabOptions.Tabs("Coloring").Selected = True
'      fraColoring.Visible = True                                             'Start at Coloring tab
      tabOptions.Tabs("Legend").selected = True    'Per Kam, we want to start at Legend in any case
      fraLegend.visible = True
   End Select
   Tag = ""
   
   coloringDirty = False
   legendDirty = False
   infoDirty = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If coloringDirty Or legendDirty Then
      Select Case MsgBox("Save changes?", vbQuestion + vbYesNoCancel)
      Case vbYes
         cmdOK_Click
      Case vbCancel
         Cancel = True
      End Select
   End If
End Sub
Private Sub cmdCancel_Click()
   Hide
End Sub
Sub FramesInvisible()
   fraColoring.visible = False
   fraLegend.visible = False
   fraOther.visible = False
End Sub

Private Sub optLegendAllColorSets_Click()
   cfgLegendAllColorSets = True
   legendDirty = True
End Sub

Private Sub optLegendFirstColorSet_Click()
   cfgLegendAllColorSets = False
   legendDirty = True
End Sub

Private Sub tabOptions_Click()
   FramesInvisible
   Select Case tabOptions.SelectedItem.key
   Case "Coloring"
      fraColoring.visible = True
   Case "Legend"
      fraLegend.visible = True
   Case "Other"
      fraOther.visible = True
   Case Else
   End Select
End Sub

Private Sub chkOpenColored_Click()
   legendDirty = True
End Sub
Private Sub chkColors_Click()
   legendDirty = True
End Sub
Private Sub chkColorSet_Click()
   legendDirty = True
End Sub
Private Sub chkExpression_Click()
   legendDirty = True
End Sub
Private Sub chkGeneDB_Click()
   legendDirty = True
End Sub
Private Sub chkRemarks_Click()
   legendDirty = True
End Sub
Private Sub chkShowLegend_Click()
   legendDirty = True
End Sub
Private Sub chkValue_Click()
   legendDirty = True
End Sub
Private Sub cmbLegendFontSize_Click()
   legendDirty = True
End Sub

Private Sub chkShowInfo_Click()
   infoDirty = True
End Sub

Private Sub optRelated_Click()
   coloringDirty = True
End Sub
Private Sub optSpecific_Click()
   coloringDirty = True
End Sub
