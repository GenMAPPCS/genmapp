VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExpression 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expression Dataset Manager"
   ClientHeight    =   6900
   ClientLeft      =   30
   ClientTop       =   615
   ClientWidth     =   8700
   Icon            =   "Expression.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotes 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   492
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "These notes will appear only in this window."
      Top             =   780
      Width           =   7632
   End
   Begin MSComctlLib.StatusBar sbrBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   6645
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   132
            MinWidth        =   2
            Key             =   "Gene DB"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraColorSets 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Color Sets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5232
      Left            =   60
      TabIndex        =   9
      Top             =   1380
      Width           =   8592
      Begin MSFlexGridLib.MSFlexGrid grdCriteria 
         Height          =   1932
         Left            =   60
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Criteria listed in priority order"
         Top             =   3180
         Width           =   7212
         _ExtentX        =   12726
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   4
         BackColor       =   14737632
         BackColorFixed  =   12648447
         BackColorBkg    =   12648447
         AllowBigSelection=   0   'False
         FocusRect       =   2
         GridLinesFixed  =   3
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
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
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   312
         Left            =   120
         TabIndex        =   40
         Top             =   3780
         Width           =   7032
         _ExtentX        =   12409
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
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
         Left            =   7320
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Remove the criterion from the Color Set"
         Top             =   4860
         Width           =   1212
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
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
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Show the criterion in the Criteria Builder"
         Top             =   4440
         Width           =   1212
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "Move Down"
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
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Move selected criterion down one in list"
         Top             =   3660
         Width           =   1212
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Move Up"
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
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Move selected criterion up one in list"
         Top             =   3240
         Width           =   1212
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Criteria Builder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2472
         Left            =   60
         TabIndex        =   14
         Top             =   660
         Width           =   8472
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
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
            Left            =   6720
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Replace the selected criterion in the Color Set"
            Top             =   240
            Width           =   792
         End
         Begin VB.TextBox txtLabel 
            BackColor       =   &H00E0E0E0&
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
            Left            =   4020
            MaxLength       =   40
            TabIndex        =   5
            ToolTipText     =   "Label for this criterion that is displayed in the MAPP Legend."
            Top             =   660
            Width           =   4392
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
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
            Left            =   7560
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Add the criterion to the color set"
            Top             =   240
            Width           =   792
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "New"
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
            Left            =   5880
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Activate Criteria Builder to accept an additional criterion"
            Top             =   240
            Width           =   792
         End
         Begin VB.TextBox txtCriterion 
            BackColor       =   &H00E0E0E0&
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
            Left            =   960
            TabIndex        =   6
            ToolTipText     =   "Build the criterion by typing or clicking on Columns and Operators"
            Top             =   2040
            Width           =   7452
         End
         Begin VB.ListBox lstColumns 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            ItemData        =   "Expression.frx":08CA
            Left            =   120
            List            =   "Expression.frx":08CC
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Click to include column name in criterion"
            Top             =   480
            Width           =   3132
         End
         Begin VB.Label lblOperator 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   " AND "
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
            Index           =   6
            Left            =   3360
            TabIndex        =   23
            Top             =   1560
            Width           =   552
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Columns"
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
            Index           =   16
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   792
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label in Legend"
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
            Left            =   4020
            TabIndex        =   30
            Top             =   420
            Width           =   1404
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
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
            Index           =   15
            Left            =   4020
            TabIndex        =   29
            Top             =   1020
            Width           =   480
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Criterion"
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
            Index           =   13
            Left            =   120
            TabIndex        =   28
            Top             =   2040
            Width           =   768
         End
         Begin VB.Label lblColor 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   252
            Left            =   4020
            TabIndex        =   25
            ToolTipText     =   "Click to assign color to criterion"
            Top             =   1260
            Width           =   1272
         End
         Begin VB.Label lblOperator 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   " OR "
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
            Index           =   7
            Left            =   3360
            TabIndex        =   24
            Top             =   1740
            Width           =   552
         End
         Begin VB.Label lblOperator 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   " <> "
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
            Index           =   5
            Left            =   3360
            TabIndex        =   22
            ToolTipText     =   "Not equal."
            Top             =   1380
            Width           =   552
         End
         Begin VB.Label lblOperator 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   " <= "
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
            Index           =   4
            Left            =   3360
            TabIndex        =   21
            Top             =   1200
            Width           =   552
         End
         Begin VB.Label lblOperator 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   " < "
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
            Index           =   3
            Left            =   3360
            TabIndex        =   20
            Top             =   840
            Width           =   552
         End
         Begin VB.Label lblOperator 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   " >= "
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
            Index           =   2
            Left            =   3360
            TabIndex        =   19
            Top             =   1020
            Width           =   552
         End
         Begin VB.Label lblOperator 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   " > "
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
            Index           =   1
            Left            =   3360
            TabIndex        =   18
            Top             =   660
            Width           =   552
         End
         Begin VB.Label lblOperator 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   " = "
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
            Index           =   0
            Left            =   3360
            TabIndex        =   17
            Top             =   480
            Width           =   552
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ops"
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
            Index           =   17
            Left            =   3360
            TabIndex        =   16
            Top             =   240
            Width           =   552
         End
      End
      Begin VB.TextBox txtColorSet 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   720
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Color set currently being processed"
         Top             =   240
         Width           =   3072
      End
      Begin VB.ComboBox cmbColorSets 
         Enabled         =   0   'False
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
         Left            =   720
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "$Undefined$"
         ToolTipText     =   "Choose from list or enter name for new parameter set."
         Top             =   240
         Width           =   3312
      End
      Begin VB.ComboBox cmbColumns 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "Expression.frx":08CE
         Left            =   4920
         List            =   "Expression.frx":08D5
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Choose the expression dataset column whose value you wish to display next to the gene."
         Top             =   240
         Width           =   3552
      End
      Begin VB.Label lblOperation 
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
         Left            =   120
         TabIndex        =   45
         Top             =   3240
         Visible         =   0   'False
         Width           =   1008
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
         Left            =   120
         TabIndex        =   44
         Top             =   4140
         Width           =   684
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
         Left            =   960
         TabIndex        =   43
         Top             =   4140
         Visible         =   0   'False
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
         Left            =   3300
         TabIndex        =   42
         Top             =   4140
         Visible         =   0   'False
         Width           =   816
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
         TabIndex        =   41
         Top             =   3480
         Visible         =   0   'False
         Width           =   528
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
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
         Index           =   6
         Left            =   4320
         TabIndex        =   38
         Top             =   360
         Width           =   504
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   13
         Top             =   288
         Width           =   528
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gene"
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
         Index           =   14
         Left            =   4320
         TabIndex        =   12
         Top             =   180
         Width           =   480
      End
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   960
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   7632
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   60
      Top             =   7980
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.TextBox txtRemarks 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   960
      MaxLength       =   50
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Appears on gene backpages and in the MAPP Legend."
      Top             =   420
      Width           =   7632
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
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
      Index           =   4
      Left            =   60
      TabIndex        =   8
      Top             =   780
      Width           =   528
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   240
      Index           =   3
      Left            =   60
      TabIndex        =   7
      Top             =   420
      Width           =   804
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   528
   End
   Begin VB.Menu mnuExp 
      Caption         =   "&Expression Datasets"
      Begin VB.Menu mnuImport 
         Caption         =   "&New Dataset"
      End
      Begin VB.Menu mnuExceptions 
         Caption         =   "&Process Exceptions"
      End
      Begin VB.Menu mnuOpenExp 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSaveExp 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveAsExp 
         Caption         =   "Save &As"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGeneDB 
         Caption         =   "Choose &Gene Database"
      End
      Begin VB.Menu mnuExitExp 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuColorSets 
      Caption         =   "&Color Sets"
      Begin VB.Menu mnuNewColorSet 
         Caption         =   "&New"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveColorSet 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddColorSet 
         Caption         =   "&Add"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDeleteColorset 
         Caption         =   "&Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopyColorSet 
         Caption         =   "&Copy from ..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public callingForm As Form
'  Calling form must be passed before frmExpression is shown
Public dbGene As Database         'Gene DB open in the EDM. Should be same as the active mappWindow
Public dbExpression As Database              'Will always be tempGex until Expression Dataset saved
Public expressionName As String                'Actual name of Expression Dataset opened in tempGex
Public rsColorSet As Recordset
   '  These five items are independent of the drawingBoard because the Expression Dataset
   '  Manager does not necessarily have to be working on the same Expression Dataset as the
   '  drawingBoard. Typically, they originally come from the drawingBoard.
Dim drawingBoardExpression As String
   '  Name of dbExpression open in drawingBoard. The actual dbExpression must be closed in
   '  order for Expression Dataset Manager to work on it. The drawingBoardExpression will be
   '  reopened as whatever Expression Dataset Manager was working on last when the EDM is exited.
'Dim drawingBoardColorSet As String
Dim drawingBoardColorIndexes(MAX_COLORSETS) As Integer
Dim drawingBoardValueIndex As Integer
   '  Same as drawingBoardExpression
Public tempGex As String           'Name of temp .gex file being used by Expression Dataset Manager
            'Name is the gex file with a $ in front of it and resides in the app folder
            'Used in EDM because changes are always temporary until user Saves
            'When a new Expression Dataset is opened for editing, the EDM checks to see if the
            'tempGex file already exists. If it does, then the file is being edited in some other
            'instance or was hung up in a previous crash
Dim rsInfo As Recordset, rsExpression As Recordset, rsColorSets As Recordset
Dim selectionStart As Integer, selectionLength As Integer
Public expressionDirty As Boolean, colorSetDirty As Boolean, criterionDirty As Boolean
Public makeDisplayTable As Boolean                        'True if Display Table needs to be remade
Dim loading As Boolean  'True in some loading process so that the various dirties above are not set
Dim expressionChanged As Boolean      'Anytime there is some kind of a change in the ED in the EDM,
                                      'it must be applied to the drawingBoard upon exit
Dim expressionSaveError As Boolean
Dim addException As Boolean           'True if adding exception data to existing Expression Dataset
Dim cancelExit As Boolean
Dim loadingColorSet As Boolean
Dim process As String          'Keep track of current process in case it interferes with some event
Dim criterionMode As String
Dim formActive As Boolean                                                      'See Form_Activate()
Dim doExceptions As Boolean                                      'True if doExceptionsing raw data
'Change all this ????????????????????????????
'  Expresson Database
'     dbExpression   Global Expression Dataset object
'                       Either open or Nothing
'                       Expression Manager always works with the temporary Expression Dataset
'                          in the appPath & "$" & the same name as the Expression Dataset.
'                          E.g. C:\GenMAPP\Expression Datasets\Cardio.gex has the temp
'                               C:\ProgramFiles\GenMAPP\$Cardio.gex.
'     expression     Global variable with database path
'                       Only one Expression Dataset open at a time for both Expression Manager
'                          and Drafter
'                       Any database processed in Expression Manager becomes the database for
'                          Drafter
'     colorSet       Global variable with color set name
'                       Any color set processed in Expression Manager becomes the color set for
'                          Drafter
'

'////////////////////////////////////////////////////////////////////////////////////// Form Events
Private Sub Form_Load()
   mnuNewColorSet.Enabled = False
   mnuDeleteColorset.Enabled = False
   process = ""
   With grdCriteria
      .ColWidth(0) = 300
      .ColWidth(1) = 2000
      .ColAlignment(1) = flexAlignLeftCenter
      .TextMatrix(0, 1) = "Label"
      .ColWidth(2) = 4100
      .ColAlignment(2) = flexAlignLeftCenter
      .TextMatrix(0, 2) = "Criterion"
      .ColWidth(3) = 550
      .TextMatrix(0, 3) = "Color"
   End With
   cmbColumns.ListIndex = 0                                                 'This calls Click event
End Sub
Private Sub Form_Activate()
   Dim i As Integer
   
   If formActive Then Exit Sub 'Program has only jumped to another subsidiary form and is returning
'   If addException Then Exit Sub         'Don't want to do this if just returning from frmException
   
   doExceptions = False
   ClearExpression
   Set dbGene = mappWindow.dbGene                             'Default to ones active in mappWindow
   If Not mappWindow.dbExpression Is Nothing Then '=======================Save Mapp Window Settings
      '  Will be reset later upon exit of GDM. Don't worry about displayIndex
      drawingBoardExpression = mappWindow.dbExpression.name
      For i = 0 To colorIndexes(0)
         drawingBoardColorIndexes(i) = colorIndexes(i)
      Next i
      drawingBoardValueIndex = valueIndex
      mappWindow.dbExpression.Close                              'Close any active ED in mappWindow
      Set mappWindow.dbExpression = Nothing
      colorIndexes(0) = 0
   End If
   If dbGene Is Nothing Then
      OpenGeneDB dbGene, "**OPEN**"
   Else
      sbrBar.Panels(1) = Mid(dbGene.name, InStrRev(dbGene.name, "\") + 1)
   End If

   mnuCopyColorSet.Enabled = False
   
   formActive = True         'True as long as form has not been hidden. The only place this happens
                             'is at the end of cmdExit. Technique allows the program to move to
                             'sub forms (frmException, frmDataID) without triggering Form_Activate
                             'event again.
   If commandLine <> "" Then                                               'gex file double clicked
      '  If command line were a mapp file, it would have been handled and cleared before this
      mnuOpenExp_Click
   ElseIf drawingBoardExpression <> "" Then        'Default EDM to Expression Dataset in mappWindow
      If colorIndexes(0) > 0 Then                                'Default to first active Color Set
         i = colorIndexes(1)
      Else
         i = -1                                                                'No active Color Set
      End If
      If Not FillExpressionValues(drawingBoardExpression, i) Then
         formActive = False
         Unload frmExpression
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
      EnableForm
   End If
   expressionChanged = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If MousePointer = vbHourglass Then                'Busy doing something, like converting dataset
      Cancel = 1
      Exit Sub                           'Don't close  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   mnuExitExp_Click
   If cancelExit Then
      cancelExit = False
      Cancel = 1
   End If
End Sub
Sub DisableForm()
   '  Used when dataset conversion is in progress
   txtRemarks.Enabled = False
   txtRemarks.BackColor = vbGray
   txtNotes.Enabled = False
   txtNotes.BackColor = vbGray
   cmbColorSets.Enabled = False
   cmbColorSets.BackColor = vbGray
'   chkDisplayName.Enabled = False
'   chkDisplayName.BackColor = vbGray
'   chkDisplayRemarks.Enabled = False
'   chkDisplayRemarks.BackColor = vbGray
'   chkDisplayColorSet.Enabled = False
'   chkDisplayColorSet.BackColor = vbGray
'   chkDisplayValue.Enabled = False
'   chkDisplayValue.BackColor = vbGray
End Sub
Sub EnableForm()
   '  Used when form activated or dataset conversion finishes
   txtRemarks.Enabled = True
   txtRemarks.BackColor = vbWhite
   txtNotes.Enabled = True
   txtNotes.BackColor = vbWhite
   cmbColorSets.Enabled = True
   cmbColorSets.BackColor = vbWhite
'   chkDisplayName.Enabled = True
'   chkDisplayName.BackColor = vbYellow
'   chkDisplayRemarks.Enabled = True
'   chkDisplayRemarks.BackColor = vbYellow
'   chkDisplayColorSet.Enabled = True
'   chkDisplayColorSet.BackColor = vbYellow
'   chkDisplayValue.Enabled = True
'   chkDisplayValue.BackColor = vbYellow
End Sub

Private Sub mnuCopyColorSet_Click()
   Dim tdfED As TableDef                                           'Expression Dataset being edited
   Dim EDName As String
   Dim tdfCopyED As TableDef, dbCopyED As Database 'Expression Dataset from which to copy Color Set
   Dim fld As Field
   Dim copyED As String                                        'Path of .gex to copy Color Set from
   Dim noMatch As Boolean                                'True if an ED field not found in a copyED
   Dim nonMatchedColumns As String
   Dim i As Integer, j As Integer
   Dim rs As Recordset
'  Don't call this from here
'   'For GetColorSet()
'     Dim labels(MAX_CRITERIA) As String, criteria(MAX_CRITERIA) As String, _
'          colors(MAX_CRITERIA) As Long
'     Dim notFoundIndex As Integer                       'Index of 'Not found' criterion (last one)
'   '  Call:
'   '     GetColorSet dbExpression, rsColorSet, labels, criteria, colors, notFoundIndex

On Error GoTo OpenError
   With dlgDialog
      .CancelError = True
      .DialogTitle = "Copy Color Set from File"
      .InitDir = GetFolder(mruDataSet)
      .Filter = "Expression (.gex)|*.gex"
      .FileName = GetFolder(mruDataSet) & "*.gex"
      .FLAGS = cdlOFNExplorer + cdlOFNFileMustExist
      .ShowOpen
      copyED = .FileName
   End With
   If InStr(copyED, ".") = 0 Then
      copyED = copyED & ".gex"
   End If
On Error GoTo 0

   MousePointer = vbHourglass
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Test For Match Of Color Sets
      '  At this point, all columns from ED being processed must have a column of the same name
      '  and datatype as a column from copyED.
      '  Can make it looser by parsing out columns from criteria of copyED and checking for
      '  their existence in the ED being processed.
   Set dbCopyED = OpenDatabase(copyED)
   Set rs = dbCopyED.OpenRecordset("SELECT * FROM ColorSet")
   If rs.EOF Then '================================================No Color Sets In ED To Copy From
      MsgBox "The Expression Dataset you are copying from has no Color Sets.", _
             vbExclamation + vbOKOnly, "Copy Color Set"
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   Set tdfED = dbExpression.TableDefs("Expression")
   Set tdfCopyED = dbCopyED.TableDefs("Expression")
   For i = 3 To tdfCopyED.Fields.count - 2            'Don't check OrderNo, ID, SystemCode, Remarks
      For j = 3 To tdfED.Fields.count - 2
         If tdfED.Fields(j).name = tdfCopyED.Fields(i).name Then                         'Same name
            If tdfED.Fields(j).Type = tdfCopyED.Fields(i).Type Then                  'Same datatype
               Exit For                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
         End If
      Next j
      If j > tdfED.Fields.count - 2 Then                                 'Through loop and no match
         noMatch = True
         nonMatchedColumns = nonMatchedColumns & "   " & tdfCopyED.Fields(i).name & vbCrLf
'         Exit For                                          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
   Next i
   
   If noMatch Then
      MsgBox "Your Expression Dataset has columns:" & vbCrLf _
             & nonMatchedColumns _
             & "that are not matched in either name or data type in the Expression Dataset " _
             & "you are copying from. The Color Set could not be copied.", _
             vbExclamation + vbOKOnly, "Copy Color Set"
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Copy Color Set
   Set tdfED = Nothing
'   ClearColorSet
   dbExpression.Execute "DROP TABLE [TempColorSet]"
   s = "SELECT * INTO TempColorSet IN '" & dbExpression.name & "' FROM ColorSet"
   dbCopyED.Execute s
   FillAllColorSets
   colorSetDirty = True
   makeDisplayTable = True
   
ExitSub:
   MousePointer = vbDefault
   Exit Sub

OpenError:
   If Err <> 32755 Then
      FatalError "frmExpression:mnuCopyColorSet", Err.Description
   End If
   Resume ExitSub
End Sub

Private Sub mnuGeneDB_Click()
'   Dim dbGene As Database
'   Set dbGene = callingform.dbGene                   'Not sure why we have to do this, but it works
   OpenGeneDB dbGene, "**OPEN**"
'   Set callingform.dbGene = dbGene
'   callingform.sbrBar.Panels(2).Text = sbrBar.Panels(1).Text           'Set statusbar on frmDrafter
End Sub

'////////////////////////////////////////////////////////////////////////////////////// Menu Events
Private Sub mnuHelp_Click()
   Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, appPath & "GenMAPP.chm::/ExpressionDatasets.htm", HH_DISPLAY_TOPIC, 0)
End Sub

'////////////////////////////////////////////////////////////////////// Expression Dataset Handling
'######################################################################### Expression Dataset Menus
Private Sub mnuImport_Click() '********************************************** Convert Raw Data File
'   Const MAX_COLUMN_ROWS = 35                    'Max rows of column titles and values on frmDataID
'   Const COLUMN_SEPARATION = 100                             'Twips separating columns on frmDataID
   '  Creates a new empty Expression DB, calls ConvertExpressionData()
   Dim rawDataFile As String
   Dim errorFile As String, exceptionFile As String
   Dim rows As Long, exceptionRows As Long, unidentifiedRows As Long
   Dim columns As Integer                        'Number of data columns, including Remarks if they
                      'exist in the raw data. One based. Does not include ID and SystemCode columns
   Dim tdfExpression As TableDef, idxExpression As index, inLine As String
   Dim quote As Integer, comma As Integer, prevComma As Integer, lines As Long
   Dim sql As String, GenMAPP As String, geneId As String, id As String
   Dim systemCodeAll As String       'Set to system code if adding remaining genes to specific type
   Dim orderNo As Long                                             'Order within Expression Dataset
   Dim exceptionRow As Long
   Dim rsGeneID As Recordset
   Dim lastExpValue As Integer, expValue As Integer                                'One-based index
      '  The expValues() array is one based; the zero element is not used. This array is passed
      '  to other functions. If its upper bound is zero, then there is no expression data.
   Dim exceedsTitleCharLimit As Boolean            'True if a column title exceeds TITLE_CHAR_LIMIT
   Dim i As Integer, delim As Integer
   Dim newDataPath As String                         'Used only until new expression confirmed open
   Dim system As String, systemCode As String
   Dim systems(MAX_SYSTEMS, 2), lastSystem As Integer, rsSystems As Recordset
      '  systems(x, 0)     Name of cataloging system. Eg: GenBank
      '  systems(x, 1)     System code. Eg: G
      '  systems(x, 2)     Additional search columns. Eg: |Gene\sBF|Orf\SBF|
   Dim slash As Integer, pipe As Integer, idColumn As String
   Dim unidentified As Long                                            'Number of unidentified rows
   Dim errors As String
   Dim notesIndex As Integer
      '  Notes can be in any column of the raw dataset. If it exists it is shifted to
      '  the second-to-last column of the ED. Otherwise, a NULL column is added to the ED.
   Dim remarksIndex As Integer
      '  Remarks can be in any column of the raw dataset. If it exists it is
      '  shifted to the last column of the ED. Otherwise, a NULL column is added to the ED.
   Dim errorsExists As Integer                             '1 if ~Errors~ column exists in raw data
   Dim AddToOther As Boolean                    'If true, add all unidentified genes to Other table
   Dim otherGene As Boolean                                          'This gene gets added to Other
   'For AllRelatedGenes()
      Dim genes As Integer
      Dim geneIDs(MAX_GENES, 2) As String
      Dim geneFound As Boolean
      'Dim supportedSystem as Boolean                 'System supported in Gene Database [optional]
      Dim systemsList As Variant                                      'Systems to search [optional]
   Dim rsRelations As Recordset, rsRelational As Recordset
   Dim dbED As Database, EDField As Integer, tdf As TableDef
   Dim GenBankRelations(20, 1) As String, lastGenBankRelation As Integer
      '  Names of all GenBank relational files
      '  GenBankRelations(x, 0)    Name of cataloging system. Eg: GenBank
      '  GenBankRelations(x, 1)    System code. Eg: G
   Dim warningColumns As String                          'Warnings for truncations in these columns
   Dim dot As Integer
   Dim columnsDontMatch As Boolean
   
   If expressionDirty Or colorSetDirty Or criterionDirty Then '+++++++++++++++++++ Save Current ED?
      Select Case MsgBox("Save current expression dataset?", vbYesNoCancel + vbQuestion, _
                         "New Expression Dataset")
      Case vbYes
         mnuSaveExp_Click
         If expressionSaveError Then
            expressionSaveError = False
            Exit Sub                                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      Case vbCancel
         Exit Sub                                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case Else
      End Select
   End If
      
   lastGenBankRelation = -1
   CloseDataset
   
   Do While dbGene Is Nothing '++++++++++++++++++++++++++++++++++++++++++++++++++ Check For Gene DB
      '  Must have Gene DB to check expression data against
      OpenGeneDB dbGene, "**OPEN**"
      If dbGene Is Nothing Then
         If MsgBox("Must have a Gene Database to check expression data against.", _
                   vbOKCancel + vbExclamation, "New Expression Dataset") = vbCancel Then
            GoTo ExitSub                                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      End If
   Loop
   
Retry:
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Raw Data File
   '  End result of this section is a valid rawDataFile
On Error GoTo OpenError
   With dlgDialog
      .CancelError = True
      .FLAGS = cdlOFNExplorer + cdlOFNFileMustExist
      .DialogTitle = "Expression Data File To Convert"
      .InitDir = GetFolder(mruImportPath)
      .FileName = ""
      If doExceptions Then
         .Filter = "All files|*.EX.csv;*.EX.tab;*.EX.txt|Comma-separated values (.csv)|*.EX.csv|" _
                          & "Tab-delimited lists (.tab, .txt)|*.EX.tab;*.EX.txt"
      Else
         .Filter = "All files|*.csv;*.tab;*.txt|Comma-separated values (.csv)|*.csv|" _
                          & "Tab-delimited lists (.tab, .txt)|*.tab;*.txt"
      End If
      .FilterIndex = 1
      .ShowOpen
On Error GoTo 0
      rawDataFile = .FileName
      mruImportPath = GetFolder(rawDataFile)
   End With
   
   If Not doExceptions And Dir(Left(rawDataFile, InStrRev(rawDataFile, ".")) & "gex") <> "" Then
      Select Case MsgBox("Expression Dataset already exists. Replace it?", _
                         vbYesNo + vbExclamation, "New Expression Dataset")
      Case vbYes
'         Don't have to worry about whether the Expression Dataset is active because it
'         is a temporary file in the EDM and not hung up by Windows.
'         If expressionName = Left(rawDataFile, InStrRev(rawDataFile, ".")) & "gex" Then
'            MsgBox "Expression Dataset currently in use. Cannot replace it.", _
'                   vbOKOnly + vbExclamation, "New Expression Dataset"
'            Exit Sub                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'         End If
         Kill Left(rawDataFile, InStrRev(rawDataFile, ".")) & "gex"
      Case Else
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End Select
   End If
      
On Error GoTo RawDataFileError
      '  Try to rename the file. If it is in use, read-only, nonexistent, etc., it will error.
      '  The rawTemp.$tm must be on the same physical disk as the raw data file, otherwise
      '  it cannot be renamed.
   If Dir(GetFolder(rawDataFile) & "rawTemp.$tm") <> "" Then
      Kill GetFolder(rawDataFile) & "rawTemp.$tm"
   End If
   Name rawDataFile As GetFolder(rawDataFile) & "rawTemp.$tm"
   Name GetFolder(rawDataFile) & "rawTemp.$tm" As rawDataFile
On Error GoTo 0

   s = rawDataFile                                            'Path of raw data file without ".EX."
   dot = InStr(s, ".EX.")
   If dot <> 0 Then
      s = Left(s, dot) & Mid(s, dot + 4)
   End If
   If OpenExpressionDataset(rawDataFile) = False Then GoTo Retry  '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'   If OpenExpressionDataset(s) = False Then GoTo Retry     '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      '  This is a temp .$gx in the appPath. It also sets expressionName and tempGex variables
      '  We now have an open raw data file and a blank temp ED
   Open rawDataFile For Binary As #FILE_RAW_DATA
      
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Prepare Screen For Processing
   MousePointer = vbHourglass
   DoEvents
   DisableForm                                                         'Raw data file chosen by now
   ClearExpression
   txtName = Mid(expressionName, InStrRev(expressionName, "\") + 1)  'Calls ChangeExpression, which
   txtName = Left(txtName, InStrRev(txtName, ".") - 1)               'enables following menu items
   mnuImport.Enabled = False
   mnuOpenExp.Enabled = False
   mnuSaveExp.Enabled = False
   mnuSaveAsExp.Enabled = False
   mnuExitExp.Enabled = False
   mnuCopyColorSet.Enabled = False
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Get Valid Gene Systems
   '  This section repeated in ConvertExpressionData
   Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems ORDER BY System", dbOpenForwardOnly)
   lastSystem = -1
   Do Until rsSystems.EOF
      If VarType(rsSystems!Date) <> vbNull Or rsSystems!system = "Other" Then     'Supported system
         '  Other system is always supported, date or not
         lastSystem = lastSystem + 1
         systems(lastSystem, 0) = rsSystems!system
         systems(lastSystem, 1) = rsSystems!systemCode
         slash = InStr(1, rsSystems!columns, "\S", vbTextCompare)
         Do While slash                                              'Get additional search columns
            pipe = InStrRev(rsSystems!columns, "|", slash)
            systems(lastSystem, 2) = systems(lastSystem, 2) _
                                   & Mid(rsSystems!columns, pipe, slash - pipe + 2) & "|"
            slash = InStr(slash + 1, rsSystems!columns, "\S", vbTextCompare)
         Loop
      End If
      rsSystems.MoveNext
   Loop

   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++ See If Valid System Code In Second Row
   Dim delimiter As String * 1
   inLine = InputUnixLine(FILE_RAW_DATA)
   If InStr(inLine, vbTab) Then
      delimiter = vbTab
   Else
      delimiter = ","
   End If
   inLine = RemoveQuotes(InputUnixLine(FILE_RAW_DATA), delimiter)                                                'Second row
   i = InStr(inLine, delimiter)
   delim = InStr(i + 1, inLine, delimiter)                           'Delimiter after second column
   If delim = 0 Then delim = Len(inLine) + 1                'Handles tables with ID and system only
   systemCode = Mid(inLine, i + 1, delim - i - 1)          'Second column
   i = 0 '-------------------------------------------------------See If System Exists And Supported
   Do Until systemCode = systems(i, 1) Or i > lastSystem
      i = i + 1
   Loop
   If i > lastSystem And Not doExceptions Then          'System code doesn't exist or not supported
      If MsgBox("Gene ID system code """ & systemCode & """ in the first data row of your " _
                & "Raw Data File not in your Gene Database. Continue?", vbExclamation + vbYesNo, _
                "New Dataset") = vbNo Then
         GoTo CancelProcedure                              '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Count Data Columns
   Seek #FILE_RAW_DATA, 1
   inLine = InputUnixLine(FILE_RAW_DATA)
   notesIndex = -1                                               'Columns don't exist in raw data
   remarksIndex = -1                                               'Columns don't exist in raw data
   columns = -1                             'There are always at least 2 columns: ID and SystemCode.
                                            'Subsequent columns begin with delimiter
   i = InStr(inLine, delimiter)
   Do While i
      columns = columns + 1
      If UCase(Mid(inLine, i + 1, 5)) = "NOTES" Then                          'Fields not in quotes
         If Mid(inLine, i + 6, 1) = delimiter Or Len(inLine) = i + 5 Then           'Next delimiter
            notesIndex = columns                                                  'Zero-based index
         End If
      ElseIf UCase(Mid(inLine, i + 1, 7)) = """NOTES""" Then                      'Fields in quotes
         '  No need to test for next delimiter because title is "Notes", with quotes
         notesIndex = columns                                                     'Zero-based index
      End If
      If UCase(Mid(inLine, i + 1, 7)) = "REMARKS" Then                        'Fields not in quotes
         If Mid(inLine, i + 8, 1) = delimiter Or Len(inLine) = i + 7 Then           'Next delimiter
            remarksIndex = columns                                                'Zero-based index
         End If
      ElseIf UCase(Mid(inLine, i + 1, 9)) = """REMARKS""" Then                    'Fields in quotes
         '  No need to test for next delimiter because title is "Remarks", with quotes
         remarksIndex = columns                                                   'Zero-based index
      End If
      i = InStr(i + 1, inLine, delimiter)
   Loop
   If Right(inLine, 8) = "~Errors~" Then
      errorsExists = 1
      columns = columns - 1
   End If
   ReDim expValues(columns) As String                      'Row of values from raw expression file
      '  These are data values only, starting from the third column.
      '  expValues() is a one-based array. The zero index element is not used. If this array is
      '  passed to another function and its upper bound is zero, then there is no expression data
      '  to look for
   Seek #FILE_RAW_DATA, 1
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Titles
   frmDataID.lstTitles.Clear
   frmDataID.lblInstructions = "Check the box if the column contains text data (<= " _
                             & CHAR_DATA_LIMIT & " chars)."
   frmDataID.lblInstructions.ToolTipText = ""
   If remarksIndex = -1 And notesIndex = -1 Then
      frmDataID.lstTitles.ToolTipText = ""
   Else
      frmDataID.lstTitles.ToolTipText = """Notes"" and ""Remarks"" are always checked and " _
                                      & "can contain unlimited characters."
   End If
   s = GetExpressionRow(errorsExists, geneId, systemCode, expValues, delimiter)         'Title line
      
   If doExceptions Then '+++++++++++++++++++++++++++++++++++++++++ Set Up For Processing Exceptions
      For expValue = 1 To columns - 2                             'Don't consider Notes and Remarks
         If expValues(expValue) _
               <> dbExpression.TableDefs("Expression").Fields(expValue + 2).name Then
            '  Compare column titles with those in tempGex
            '  expValues do not include OrderNo, ID, and SystemCode
            columnsDontMatch = True
'            MsgBox "Columns of raw data do not match those of current table. Cannot process " _
'                   & "exceptions.", vbExclamation + vbOKOnly, "ExceptionProcessing Raw Data"
'            GoTo CancelProcedure                           'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      Next expValue
      dbExpression.Execute "DELETE FROM Expression"
      dbExpression.Execute "DELETE FROM Info"
      If Not columnsDontMatch Then
         s = GetExpressionRow(errorsExists, geneId, systemCode, expValues, delimiter) '1st data lin
         GoTo ExceptionProcessing                          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
   End If
   
   'This doesn't occur in exception processing ++++++++++++++++++++++++++++++++++++++++++++++++++++
   For expValue = 1 To columns     '=====================================================Each Title
      '  First two titles always ID and SystemCode
      expValues(expValue) = Trim(expValues(expValue))
      If expValues(expValue) = "" Then '------------------------------------------------Blank Title
         s = ""                      'Print out all the titles and ask for a name for the blank one
         For i = 1 To columns - 1                                              'Assemble all titles
            If i >= 1 Then s = s & "|"
            If i = expValue Then
               s = s & "________"
            Else
               s = s & expValues(i)
            End If
         Next i
         If Len(s) > 900 Then                                       'Too many characters to display
            '  Input boxes have a limit of about 1024 characters. This gives some leeway.
            i = InStr(s, "________")
            If i > 100 Then                'Leave at least 100 characters in front of blank heading
               s = Left("..." & Mid(s, i - 100), 897) & "..."
            Else
               s = Left(s, 897) & "..."
            End If
         End If
         Do
            expValues(expValue) = InputBox("The column heading (_________) in" _
                     & vbCrLf & vbCrLf & s & vbCrLf & vbCrLf _
                     & "is blank. Enter a " _
                     & "column heading or abort the conversion by clicking Cancel.", _
                     "Invalid Column Heading", "________")
         Loop While expValues(expValue) = "________"
      End If
RecheckTitle: '---------------------------------------------------------------Validate Column Title
      expValues(expValue) = Trim(TextToSql(expValues(expValue)))
      If expValues(expValue) = "" Then
         '  The only time a blank title should reach here is if Cancel were entered above or
         '  below and program came back to recheck the title
         GoTo CancelProcedure                              'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
      Select Case UCase(expValues(expValue)) '_______________________________________Reserved Title
      Case "ID", "SYSTEMCODE", "DATE"
         expValues(expValue) = _
               InputBox("Column heading """ & expValues(expValue) _
                        & """ is reserved for the system. Please change it.", _
                       "Column Heading Reserved", expValues(expValue))
         GoTo RecheckTitle                                 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      End Select
      If InvalidChr(expValues(expValue), "column heading") Then
         '  Won't return until all invalid chrs cleaned up or title returns ""
         GoTo RecheckTitle                                 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      End If
      If Len(expValues(expValue)) > TITLE_CHAR_LIMIT Then '________________________________Too Long
         expValues(expValue) = _
               InputBox("Column heading """ & expValues(expValue) & """ exceeds the " _
                & TITLE_CHAR_LIMIT & "-character limit. Please shorten it.", _
                "Column Heading Too Long", Left(expValues(expValue), TITLE_CHAR_LIMIT))
         GoTo RecheckTitle                                 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      End If
      For i = 1 To expValue - 1 '________________________________________Check For Duplicate Titles
         If expValues(expValue) = expValues(i) Then
            expValues(expValue) = _
                  InputBox("You have two columns headed '" & expValues(expValue) _
                           & "'. You may change the second one here and click OK, " _
                           & "or abort the conversion by clicking Cancel.", _
                           "Duplicate Column Headings", expValues(expValue))
            GoTo RecheckTitle                              '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
         End If
      Next i
      frmDataID.lstTitles.AddItem expValues(expValue)
   Next expValue
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Column Datatypes
   s = GetExpressionRow(errorsExists, geneId, systemCode, expValues, delimiter)    'First data line
   If s <> "" Then                                                        'Error in first data line
      MsgBox "Your first data line has " & s & " compared to your title line. Expression " _
             & "Dataset Manager uses this line to determine the data types of your raw data. " _
             & "You will have to make this line agree with the titles before proceeding.", _
             vbCritical + vbOKOnly, "Raw Data Conversion Error"
      GoTo CancelProcedure                                  'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   For expValue = 1 To columns
      If expValue = remarksIndex - 1 Or expValue = remarksIndex Then         'These are always Memo
         frmDataID.lstTitles.selected(expValue - 1) = True                              'Zero based
      ElseIf expValues(expValue) = "" Then                                         'Blank first row
         MsgBox "The data in the first row of your Raw Data File for column" _
                & vbCrLf & vbCrLf & frmDataID.lstTitles.List(expValue - 1) & vbCrLf & vbCrLf _
                & "is empty. GenMAPP guessed that it is text data, but check it.", _
                vbExclamation + vbOKOnly, "Determining Data Types"
         frmDataID.lstTitles.selected(expValue - 1) = True
      Else
         If Not IsNumeric(expValues(expValue)) Then
            frmDataID.lstTitles.selected(expValue - 1) = True
         End If
      End If
   Next expValue
   frmDataID.lstTitles.ListIndex = -1
   
   If columns > 0 Then                                        'To allow for ID and code but no data
      frmDataID.show vbModal                                                'Show Column Title List
   End If
   
   If frmDataID.Tag = "Cancel" Then GoTo CancelProcedure    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      '  Must go through this nonsense to return Screen.ActiveForm back to frmExpression
      '  from frmDataID
      txtColorSet.visible = True
      txtColorSet.Enabled = True
      frmExpression.txtColorSet.SetFocus
   
ExceptionProcessing: '++++++++++++++++++++++++++++++++++++++++ Enter Here For Processing Exceptions
   grdCriteria.visible = False
   cmdUp.visible = False
   cmdDown.visible = False
   cmdEdit.visible = False
   cmdDelete.visible = False
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Create Expression Table And Error File
'   lblDetail.Visible = True
'   lblDetail = FileAbbrev(rawDataFile)
''   lblPrgMax = "Exceptions: 0"
''   lblPrgValue = "Unidentified: 0"
'   errorFile = Left(rawDataFile, InStrRev(rawDataFile, ".") - 1) & ".$tm"
'   Open errorFile For Output As #FILE_EXCEPTIONS
'   Print #FILE_EXCEPTIONS, "Gene ID"; delimiter; "System"; delimiter;
   If doExceptions Then '==========================================Set Up For Processing Exceptions
      Set tdfExpression = dbExpression.TableDefs("Expression")       'Use existing Expression table
   Else '============================================================Set Up For Original Conversion
      Set tdfExpression = dbExpression.CreateTableDef("Expression")
      With tdfExpression
         Dim idxOrderNo As index
         .Fields.Append .CreateField("OrderNo", dbLong)
         Set idxOrderNo = .CreateIndex("ixOrder")
         idxOrderNo.Fields.Append .CreateField("OrderNo", dbLong)
         .Indexes.Append idxOrderNo
         .Fields.Append .CreateField("ID", dbText, CHAR_DATA_LIMIT)
         .Fields.Append .CreateField("SystemCode", dbText, 3)
         Set idxExpression = .CreateIndex("ixID")
         With idxExpression                     'Probably faster to index after filling table?????
            .Fields.Append .CreateField("ID")
            .Fields.Append .CreateField("SystemCode")
         End With
         .Indexes.Append idxExpression
         For i = 0 To frmDataID.lstTitles.ListCount - 1 '------------------------Create Data Fields
            If i <> remarksIndex - 1 Then                               'Move Remarks to end of row
               If frmDataID.lstTitles.selected(i) = True Then                         'String field
                  .Fields.Append .CreateField(frmDataID.lstTitles.List(i), dbText, CHAR_DATA_LIMIT)
                  .Fields(frmDataID.lstTitles.List(i)).AllowZeroLength = True
'                  Print #FILE_EXCEPTIONS, frmDataID.lstTitles.List(i); delimiter;
               Else                                                                  'Numeric field
                  .Fields.Append .CreateField(frmDataID.lstTitles.List(i), dbSingle)
'                  Print #FILE_EXCEPTIONS, frmDataID.lstTitles.List(i); delimiter;
               End If
            End If
         Next i
         .Fields.Append .CreateField("Notes", dbMemo)
         .Fields("Notes").AllowZeroLength = True
'         If remarksIndex = -1 Then                                   'Raw data has no Remarks field
            .Fields.Append .CreateField("Remarks", dbMemo)
            .Fields("Remarks").AllowZeroLength = True
'         End If
         dbExpression.TableDefs.Append tdfExpression
'         Print #FILE_EXCEPTIONS, "Remarks"; delimiter; "~Errors~"; vbLf;
      End With
   End If
   Close #FILE_RAW_DATA
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Do The Conversion
   '  At this point, the Expression table has been created and all the data types determined.
   ConvertExpressionData rawDataFile, dbExpression, dbGene
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++ Make Temporary ED Permanent And Clean Up
   If Dir(expressionName) <> "" Then
      Kill expressionName
   End If
   Name tempGex As expressionName
   
   b = FillExpressionValues(expressionName)           'Should always be true here because it is new
                                                      'Expression Dataset
   mnuSaveExp.Enabled = False
   mnuSaveAsExp.Enabled = True
   expressionDirty = False
   makeDisplayTable = False

ExitSub:
   EnableForm
ExitSub1:
   doExceptions = False
   mnuImport.Enabled = True
'   mnuProcess.Enabled = True
   mnuOpenExp.Enabled = True
   mnuExitExp.Enabled = True
   mnuCopyColorSet.Enabled = True
   lblErrors.visible = False
   lblDetail.visible = False
'   lblConversionTitle.Visible = False
'   lblConversion.Visible = False
   grdCriteria.visible = True
   cmdUp.visible = True
   cmdDown.visible = True
   cmdEdit.visible = True
   cmdDelete.visible = True
   expressionChanged = True
   MousePointer = vbDefault
   DoEvents
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Error Routines
RawDataFileError:
   MsgBox rawDataFile & " could not be opened. It may be open elsewhere or set to read-only " _
          & "through Windows, or perhaps does not exist.", _
          vbExclamation + vbOKOnly, "Converting Raw Data"
   Resume ExitSub
   
CancelProcedure: '========================================================Cancel Conversion Process
   dbExpression.Close
   Set dbExpression = Nothing
   expressionName = ""
   Kill tempGex
   tempGex = ""
   ClearExpression
   Close #FILE_RAW_DATA
   GoTo ExitSub                                            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

OpenError: '===========================================================Trap Database Opening Errors
'Need no such MAPP error
   Select Case Err
   Case 32755 '-------------------------------------------------------------------------Cancel Open
      On Error GoTo 0
      Resume ExitSub
   Case Else '------------------------------------------------------------------Unidentified Errors
      Kill errorFile
      FatalError "frmExpression:mnuImport at line:" & vbCrLf & vbCrLf & inLine _
                 & vbCrLf & vbCrLf & "Processing stopped", Err.Description
   End Select
End Sub

Private Sub mnuExceptions_Click()
   doExceptions = True
   mnuImport_Click
End Sub

'***************************************************************** Remove Single Quotes From String
'Private Function RemoveQuotes(ByVal Lin As String) As String
'   Dim newLine As String, quote As Integer
'
'   quote = InStr(Lin, "'")
'   If quote = 0 Then                                                                     'Quick Out
'      RemoveQuotes = Lin
'      Exit Function                                        'No Quotes to remove >>>>>>>>>>>>>>>>>>>
'   End If
'
'   Do While quote
'      Lin = Left(Lin, quote - 1) & Mid(Lin, quote + 1)
'      quote = InStr(Lin, "'")
'   Loop
'      RemoveQuotes = Lin
'End Function
'

Private Sub mnuOpenExp_Click() '********************************** Open Existing Expression Dataset
   Dim cancelOpen As Boolean
   Dim newExpression As String, oldExpression As String
   
   oldExpression = expressionName
   If commandLine = "" Then '===========================================No Command Line, Ask For ED
      cancelOpen = False
      If expressionDirty Or colorSetDirty Or criterionDirty Then
         Select Case MsgBox("Save current Expression Dataset?", vbQuestion + vbYesNoCancel, _
                            "Opening Expression Dataset")
         Case vbCancel
            cancelOpen = True
         Case vbYes
            mnuSaveExp_Click
            If expressionSaveError Then
               expressionSaveError = False
               cancelOpen = True
            End If
         End Select
      End If
      
      If cancelOpen Then Exit Sub                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
Retry:
On Error GoTo OpenError
      With dlgDialog
         .CancelError = True
         .DialogTitle = "Open Expression Dataset"
         .InitDir = GetFolder(mruDataSet)
         .Filter = "Expression (.gex)|*.gex"
         .FileName = GetFolder(mruDataSet) & "*.gex"
         .FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
         .ShowOpen
         newExpression = .FileName
      End With
      If InStr(newExpression, ".") = 0 Then
         newExpression = newExpression & ".gex"
      End If
   Else '======================================================================Process Command Line
      newExpression = commandLine
      If Left(newExpression, 1) = """" Then '----------------------------Strip Quotes If They Exist
         newExpression = Mid(newExpression, 2)
         If Right(newExpression, 1) = """" Then
            newExpression = Left(newExpression, Len(newExpression) - 1)
         End If
      End If
      commandLine = ""
   End If
On Error GoTo 0
   
   CloseDataset
   If FillExpressionValues(newExpression) Then
      If Not expressionDirty Then mnuSaveExp.Enabled = False
      mnuSaveAsExp.Enabled = True
      mnuImport.Enabled = True
      EnableForm
      mnuCopyColorSet.Enabled = True
   End If
ExitSub:
   Exit Sub

OpenError:
   If Err <> 32755 Then
      MsgBox Err.Description, vbCritical, "Open Expression Dataset Error"
   End If
   On Error GoTo 0
   Resume ExitSub
End Sub

Public Sub mnuSaveExp_Click() '******************************************* Save Expression Dataset
   '  There will always be a Temp.gex database in existence and open.
   
   If colorSetDirty Or criterionDirty Then
      If Not SaveColorSet Then                                                'Save color set first
         cancelExit = True
         expressionSaveError = True
         Exit Sub            'Couldn't save color set >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   
   If Not expressionDirty Then
      '  This leaves Temp.gex open but does not copy it into the permanent database and so causes
      '  no write permission errors
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For Validity
   txtRemarks = TextToSql(Dat(txtRemarks))
   If InvalidChr(txtRemarks, "Remarks") Then
      txtRemarks.SetFocus
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   txtNotes = TextToSql(Dat(txtNotes))
   If InvalidChr(txtNotes, "Notes") Then
      txtNotes.SetFocus
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   MousePointer = vbHourglass
'   txtColorSet.Enabled = False
'   cmbColumns.Enabled = False
'   cmdNew.Enabled = False
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Make Temporary Color Set Permanent
'   colorSet = txtColorSet
   Set rsColorSet = dbExpression.OpenRecordset("SELECT * FROM ColorSet WHERE ColorSet = ''")
   rsColorSet.Close                                       'Open it in case it wasn't, then close it
   dbExpression.Execute "DROP TABLE ColorSet"
   dbExpression.Execute "SELECT * INTO ColorSet FROM TempColorSet"
   
   Set rsInfo = dbExpression.OpenRecordset("SELECT * FROM Info")
      'This is open on a normal save but not on a save as. Should clean this up ???????????????????
   With rsInfo '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Save The Info Parameters
      .edit
      !modify = Now
      !remarks = txtRemarks
      !notes = txtNotes
      .Update
   End With
   rsInfo.Close
   
   lblDetail = "Saving"
   If makeDisplayTable Then
      grdCriteria.visible = False                                        'Expose progress indicator
      cmdUp.visible = False
      cmdDown.visible = False
      cmdEdit.visible = False
      cmdDelete.visible = False
      
      CreateDisplayTable dbExpression
      makeDisplayTable = False
      
      grdCriteria.visible = True
      cmdUp.visible = True
      cmdDown.visible = True
      cmdEdit.visible = True
      cmdDelete.visible = True
   End If
   
   dbExpression.Execute "UPDATE Info SET GeneDB = '" & GetFile(mruGeneDB) & "'"
      '  Eg: UPDATE Info SET GeneDB = 'Pathfinder.gdb'
   dbExpression.Execute "UPDATE Info SET Version = '" & BUILD & "'"
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Make Working Database Permanent
   dbExpression.Close
   If Dir(expressionName) <> "" Then              'In Save As, the new expression doesn't exist yet
      Kill expressionName
   End If
   FileCopy tempGex, expressionName
   Set dbExpression = OpenDatabase(tempGex)                             'Reopen in case using again
            '  What about TempColorSet and info???????????????????????????????????
   expressionDirty = False
   expressionChanged = True
   MousePointer = vbDefault
End Sub
Private Sub mnuSaveAsExp_Click() '************************** Save Expression Dataset Under New Name
   '  There will always be a database in existence and open.
   
   Dim newExpression As String
   
   If colorSetDirty Or criterionDirty Then
      If Not SaveColorSet Then                                                'Save color set first
         cancelExit = True
         Exit Sub            'Couldn't save color set >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   
Retry:
On Error GoTo SaveError
   With dlgDialog
      .CancelError = True
      .DialogTitle = "Save Expression Dataset"
      .InitDir = GetFolder(mruDataSet)
      .Filter = "Expression (.gex)|*.gex"
      .FileName = GetFolder(mruDataSet) & "*.gex"
      .FLAGS = cdlOFNExplorer + cdlOFNHideReadOnly
      .ShowSave
      newExpression = .FileName
   End With
   If InStr(newExpression, ".") = 0 Then
      newExpression = newExpression & ".gex"
   End If
   If Dir(newExpression) <> "" Then '---------------------------------------Database already exists
      If MsgBox("Replace existing file?", vbQuestion + vbYesNo, "Saving Expression Dataset") _
            = vbNo Then
         expressionSaveError = True
         Exit Sub                          'Don't overwrite file   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Else
         Kill newExpression
      End If
   End If
On Error GoTo 0
   If PathCheck(GetFolder(newExpression)) <> "" Then                           'Must read and write
      GoTo Retry
   End If

   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Change Name To New Database
   expressionName = newExpression
   txtName = Mid(expressionName, InStrRev(expressionName, "\") + 1)
   txtName = Left(txtName, InStrRev(txtName, ".") - 1)
   
   mnuSaveExp_Click
   
ExitSub:
   Exit Sub                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
SaveError:
   Select Case Err
   Case 52, 75, 3043
      MsgBox "Cannot save to this path. This may be a read-only drive, such as a CD-ROM, " _
             & "or a removable drive with no disk in it.", vbExclamation + vbOKOnly, _
             "Save MAPP Error"
   Case 70, 3045
      MsgBox Err.Description & ". " & newExpression & " possibly open in some other program.", _
            vbExclamation, "Save Expression Dataset Error"
   Case 32755                                                                    'Save As cancelled
   Case Else                                                                      'Some other error
      MsgBox Err.Description, vbCritical, "Save Expression Dataset Error"
      expressionSaveError = True
   End Select
   On Error GoTo 0
   Resume ExitSub
End Sub
Private Sub mnuExitExp_Click() '****************************************** Leave Expression Manager
   Dim cfgItem As String, cfgValue As String, colon As Integer
   Dim i As Integer
   
   cancelExit = False
   If Not dbExpression Is Nothing Then
      If Not HasDisplayTable(dbExpression) Then
         Select Case MsgBox("Save current Expression Dataset?", vbQuestion + vbYesNoCancel, _
                            "Exiting Expression Manager")
'         Select Case MsgBox("This Expression Dataset must be saved to work with the current " _
'                            & "version of GenMAPP. Save it?", vbQuestion + vbYesNoCancel, _
'                            "Old Dataset Version")
         Case vbCancel
            cancelExit = True
         Case vbYes
            expressionDirty = True
            makeDisplayTable = True
            mnuSaveExp_Click
            If expressionSaveError Then
               expressionSaveError = False
               cancelExit = True
            End If
         End Select
      ElseIf expressionDirty Or colorSetDirty Or criterionDirty Then
         Select Case MsgBox("Save current Expression Dataset?", vbQuestion + vbYesNoCancel, _
                            "Exiting Expression Manager")
         Case vbCancel
            cancelExit = True
         Case vbYes
            mnuSaveExp_Click
            If expressionSaveError Then
               expressionSaveError = False
               cancelExit = True
            End If
         End Select
      End If
   End If
         
   If cancelExit Then Exit Sub                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         
   MousePointer = vbHourglass
   Hide
   DoEvents
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Apply Expression Dataset To MAPP(s)
   If dbExpression Is Nothing Then '==============================No Expression Dataset Opened Here
      If drawingBoardExpression <> "" Then                       'Reset previous Expression Dataset
         Set mappWindow.dbExpression = OpenDatabase(drawingBoardExpression, False, True) 'Read-only
         If drawingBoardColorIndexes(0) > 0 Then                                'Reset Color Set(s)
            For i = 0 To drawingBoardColorIndexes(0)
               colorIndexes(i) = drawingBoardColorIndexes(i)
            Next i
            valueIndex = drawingBoardValueIndex
         End If
'         If drawingBoardColorSet <> "" Then
'            Set mappWindow.rsColorSet = mappWindow.dbExpression.OpenRecordset( _
'                     "SELECT * FROM ColorSet WHERE ColorSet = '" & drawingBoardColorSet & "'")
'         End If
      End If
   Else '============================================Set Drafter Expression Data To Last One Edited
      dbExpression.Execute "DROP TABLE TempColorSet"                'Made permanent in SaveColorSet
      dbExpression.Close
      Set dbExpression = Nothing
            'Whatever Expression Dataset and color set are in the manager become the ones used
            'on the drafting board
      Kill expressionName
      Name tempGex As expressionName
      tempGex = ""
      Set mappWindow.dbExpression = OpenDatabase(expressionName, , True)                 'Read-only
      mruDataSet = expressionName                                            'Set global mruDataSet
      colorIndexes(0) = 0                                         'Default no Color Set selected
      If txtColorSet = "" Then '----------------------------------------------No Color Set Selected
         colorIndexes(0) = 0
         frmDrafter.mnuApply.Enabled = False
'         frmDrafter.cmbColorSets.Enabled = False
         mruColorSet = ""
      Else '-----------------------------------------------------------------Color Set Title Exists
         Set mappWindow.rsColorSet = mappWindow.dbExpression.OpenRecordset( _
                           "SELECT * FROM ColorSet WHERE ColorSet = '" & txtColorSet & "'")
         If mappWindow.rsColorSet.EOF Then '........................txtColorSet Doesn't Exist in ED
            colorIndexes(0) = 0
            Set mappWindow.rsColorSet = Nothing
            mappWindow.mnuApply.Enabled = False
'            mappWindow.cmbColorSets.Enabled = False
            mruColorSet = ""
         Else '............................................................txtColorSet Exists in ED
            colorIndexes(0) = 1
            colorIndexes(1) = mappWindow.rsColorSet!setNo
            mappWindow.mnuApply.Enabled = True
'            mappWindow.cmbColorSets.Enabled = True
            mruColorSet = txtColorSet
         End If
      End If
      mappWindow.FillColorSetList                                 'Also highlights active Color Set
      mappWindow.legend.DrawObj
   End If
   
ExitSub:
   CloseDataset
   If mappWindow.dbGene Is Nothing Then                        'Entered EDM without an open Gene DB
      If Not dbGene Is Nothing Then                                                 'One set in EDM
         OpenGeneDB mappWindow.dbGene, dbGene.name, mappWindow
         expressionChanged = True                           'Force reapplication of expression data
      End If
   ElseIf dbGene.name <> mappWindow.dbGene.name Then
      '  If the Gene DB was changed here, it must also be changed in the mappWindow.
      OpenGeneDB mappWindow.dbGene, dbGene.name, mappWindow
      expressionChanged = True                              'Force reapplication of expression data
   End If
   If expressionChanged Then
      mappWindow.mnuApply_Click
   End If
   MousePointer = vbDefault
'   Hide
   formActive = False
End Sub
Sub CloseDataset() '***************************************** Clean Up Expression Manager Variables
'   Set dbGene = Nothing
   If Not dbExpression Is Nothing Then dbExpression.Close            'Otherwise cannot kill tempGex
   Set dbExpression = Nothing
   expressionName = ""
   Set rsColorSet = Nothing
   drawingBoardExpression = ""
   expressionDirty = False                                                       'Just to make sure
   colorSetDirty = False
   criterionDirty = False
   If tempGex <> "" Then
      Kill tempGex
      tempGex = ""
   End If
End Sub

'////////////////////////////////////////////////////////////////////// Expression Dataset Handling
'############################################################################ Expression Procedures
Sub ClearExpression() '***************************************** Clear All Items For Expression Set
   txtName = ""          'All changes in controls must come before menu and command enable settings
   txtRemarks = ""       'Most controls call ChangeExpression, which also sets these
   txtNotes = ""
   cmbColorSets.Clear                           'Names of color sets loaded with expression dataset
   cmbColumns.Clear
   cmbColumns.Enabled = False
   cmbColumns.BackColor = vbGray
   txtColorSet.Enabled = False
   txtColorSet.BackColor = vbGray
   lstColumns.Clear
   lstColumns.BackColor = vbGray
   mnuNewColorSet = False
   mnuDeleteColorset = False
   mnuSaveExp.Enabled = False
   mnuSaveAsExp.Enabled = False
   ClearColorSet
   expressionDirty = False
   makeDisplayTable = False
End Sub
Function OpenExpressionDataset(ByVal newExpressionFile As String) As Boolean
   '  Enter:   newExpressionFile    Full file name and path for either a raw data file or an
   '                                existing expression dataset.
   '  Process: Path and extension are stripped from newExpressionFile and a new file with name.$gx
   '              created in the app path.
   '           newExpressionFile goes through all tests before a new tempGex is created. If any
   '           tests failed, tempGex is left as is.
   '           If all tests passed, newExpressionFile copied into tempGex for editing and
   '           Module-level expressionName becomes actual name of expression dataset (not tempGex).
   '  Return:  True if new tempGex file created.
   Dim newTempGex As String, db As Database, oldExpression As String, writeAccess As Boolean
   Dim expressionFile As String

   expressionDirty = False
   If Dir(newExpressionFile) = "" Then '+++++++++++++++++++++++++++++ See If Expression File Exists
      MsgBox "Expression File" & vbCrLf & vbCrLf & newExpressionFile & vbCrLf & vbCrLf _
             & "doesn't exist.", vbExclamation + vbOKOnly, "Opening Expression Dataset"
      Exit Function          'Expression file doesn't exist >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If GetAttr(newExpressionFile) And vbReadOnly Then '+++++++++++++++++++++++++ Check For Read-Only
      MsgBox "Expression File" & vbCrLf & vbCrLf & newExpressionFile & vbCrLf & vbCrLf _
             & "has been set to read-only through Windows and may not be opened in the " _
             & "Expression Manager.", vbExclamation + vbOKOnly, "Opening Expression Dataset"
      Exit Function                 'Can't open for writing >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
On Error GoTo ErrorHandler
   '---------------------------------------------------Determine Name Of Actual Expression Database
      expressionName = Left(newExpressionFile, InStrRev(newExpressionFile, ".") - 1)
      If Right(expressionName, 3) = ".EX" Then                     'Dump .EX when adding exceptions
         expressionName = Left(expressionName, Len(expressionName) - 3)
      End If
      expressionName = expressionName & ".gex"

RetryTemp:
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ See If Dataset In Use Elsewhere
   '  If database open somewhere else, FileCopy produces error 70. Exactly why is a mystery but
   '  it seems to be the only way to determine the condition.
   
   If Not dbExpression Is Nothing Then '---------------------------------Close Any Existing Dataset
      '  Dataset closed for this instance so that we can see if it is open (FileCopy won't work)
      '  in some other instance. If function ultimately successful, the tempgex is opened.
      '  If not, the original condition is restored.
      writeAccess = dbExpression.Updatable
'      dbExpression.Close
      Set dbExpression = Nothing
'      Delay 5
   Else '---------------------------------------------------------------------See If We Can Copy It
      If Dir(expressionName) <> "" Then
         FileCopy expressionName, GetFolder(expressionName) & "temp.$tm"
'         FileCopy s, appPath & "temp1.$tm"
         '  If it doesn't copy, it is probably open somewhere else and error 70 will send program
         '  to ErrorHandler. If it does copy, the next statement gets rid of the temporary file
         '  produced.
         Kill GetFolder(expressionName) & "temp.$tm"
         Name expressionName As GetFolder(expressionName) & "Temp.$tm"
         Name GetFolder(expressionName) & "Temp.$tm" As expressionName
         '  Rename the expression dataset to see if it is usable and not open elsewhere.
         '  If it doesn't rename it drops to ErrorHandler.
      End If
   End If
'   s = Left(s, Len(s) - 4) & ".ldb"
'   If Dir(s) <> "" Then Kill s
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up Temp File Name
   newTempGex = appPath & Mid(newExpressionFile, InStrRev(newExpressionFile, "\") + 1)
   If InStrRev(newTempGex, ".") >= Len(newTempGex) - 5 Then               'Only look for extensions
      newTempGex = Left(newTempGex, InStrRev(newTempGex, ".") - 1)
      If Right(newTempGex, 3) = ".EX" Then                         'Dump .EX when adding exceptions
         newTempGex = Left(newTempGex, Len(newTempGex) - 3)
         addException = True
      End If
   End If
   newTempGex = newTempGex & ".$gx"
   If Dir(newTempGex) <> "" Then Kill newTempGex
   
   If UCase(Right(newExpressionFile, 4)) = ".GEX" Then '++++++++++++++++++++++++++ Existing Dataset
      Set dbExpression = OpenDatabase(newExpressionFile)                  'Check for latest version
         '  The Expression Dataset is the newExpressionFile. Eg: C:\GenMAPP 2 Data\EDs\MyExp.gex
         '  It will be processed as newTempGex. Eg: C:\Program Files\GenMAPP 2\MyExp.$gx
      If Not UpdateDataset(dbExpression, expressionDirty) Then '=============See if current version
         dbExpression.Close
         Set dbExpression = Nothing
         Exit Function                                     'Didn't update >>>>>>>>>>>>>>>>>>>>>>>>>
      End If
      If expressionDirty Then ChangeExpression                                'Be sure Save enabled
      dbExpression.Close
      Set dbExpression = Nothing
      FileCopy newExpressionFile, newTempGex    'Errors if database in use elsewhere and not caught
'      expressionName = newExpressionFile 'This is the name of the actual ED (not temp) being edited
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ New Dataset Or Exceptions
      expressionFile = Left(newExpressionFile, InStrRev(newExpressionFile, ".") - 1)      'Dump ext
         '  NewExpressionFile if new ED Eg: C:\GenMAPP 2 Data\EDs\MyExp.csv
         '     expressionFile if new ED Eg: C:\GenMAPP 2 Data\EDs\MyExp
         '  NewExpressionFile if exception Eg: C:\GenMAPP 2 Data\EDs\MyExp.EX.csv
         '     expressionFile if exception Eg: C:\GenMAPP 2 Data\EDs\MyExp.EX
      If addException Then '==================================================Processing Exceptions
         expressionFile = Left(expressionFile, InStrRev(expressionFile, ".") - 1)         'Dump .EX
            '  expressionFile Eg: C:\GenMAPP 2 Data\EDs\MyExp
         If Dir(expressionFile & ".gex") = "" Then
            MsgBox "Expression Dataset" & vbCrLf & vbCrLf & expressionFile & ".gex" _
                   & vbCrLf & vbCrLf & " does not exist. Expression Dataset and " _
                   & "Exception File must be in the same folder to process exceptions", _
                   vbExclamation + vbOKOnly, "Error Processing Exception Data"
            GoTo RestoreOriginal               'No gex file vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
         If GetAttr(expressionFile & ".gex") And vbReadOnly Then
            MsgBox "Expression Dataset" & vbCrLf & vbCrLf & expressionFile & ".gex" _
                   & vbCrLf & vbCrLf & " was set to read-only through Windows. This " _
                   & "attribute must be removed before you can process exceptions.", _
                   vbExclamation + vbOKOnly, "Error Processing Exception Data"
            GoTo RestoreOriginal               'gex file R/O vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
         If newTempGex <> tempGex Then                                     'Not the current dataset
            FileCopy expressionFile & ".gex", newTempGex 'Causes error if database in use elsewhere
         End If
      Else '===========================================================================New Database
         If Dir(expressionFile & ".gex") <> "" Then
            If GetAttr(expressionFile & ".gex") And vbReadOnly Then
               MsgBox "Expression Dataset" & vbCrLf & vbCrLf & expressionFile & ".gex" _
                      & vbCrLf & vbCrLf & "was set to read-only through Windows. " _
                      & "This attribute must be removed before you can edit this dataset.", _
                      vbExclamation + vbOKOnly, "Error Creating Expression Dataset"
               GoTo RestoreOriginal            'gex file R/O vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
            FileCopy expressionFile & ".gex", newTempGex   'Forces error if gex file open elsewhere
               Kill newTempGex
            If MsgBox(expressionFile & ".gex" & " already exists. Replace it?", _
                      vbQuestion + vbYesNo, "Creating Expression Dataset") = vbYes Then
'               If doExceptions Then
'                  If Dir(expressionFile & ".gex.old") <> "" Then
'                     Kill expressionFile & ".gex.old"
'                  End If
'                  Name expressionFile & ".gex" As expressionFile & ".gex.old"
'               Else
                  Kill expressionFile & ".gex"
'               End If
            Else
               GoTo RestoreOriginal 'Don't replace existing database vvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
         End If
         If Dir(newTempGex) <> "" Then                                             'Create new temp
            Kill newTempGex
         End If
         FileCopy appPath & "ExpTmpl.gtp", newTempGex           'Error if database in use elsewhere
      End If
'      expressionName = expressionFile & ".gex"
   End If
On Error GoTo 0

   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Open The Database
   '  At this point the opening process has passed all the hurdles and a new dbExpression can be
   '  opened. If the program doesn't reach this point, the old stuff has been restored,
   '  the function has returned false, and it is just like it was never called.
   Set dbExpression = OpenDatabase(newTempGex)
   If tempGex <> "" And newTempGex <> tempGex Then                      'Exists but not current one
      Kill tempGex
   End If
   tempGex = newTempGex
   OpenExpressionDataset = True
'   expressionDirty = False
   makeDisplayTable = False
   addException = False
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

RestoreOriginal:
   expressionDirty = False
   OpenExpressionDataset = False
   makeDisplayTable = False
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

ErrorHandler:
   Select Case Err.number
   Case 70, 75                                                                   'Database already open
      If MsgBox("Cannot open Expression Dataset for editing because it is already open " _
                 & "somewhere else. Choose another Expression Dataset before opening the " _
                 & "Expression Dataset Manager or release the Expression Dataset by closing " _
                 & "the other program, then click Retry. " & vbCrLf & vbCrLf _
                 & "Click Cancel to abort opening the Expression Dataset. If the Expression " _
                 & "Dataset is open in another GenMAPP window, you must Cancel here, return " _
                 & "to that window, release the Expression Dataset (choose another or close " _
                 & "the window), then open Expression Dataset Manager.", _
                vbExclamation + vbRetryCancel, "Opening Expression Dataset") = vbRetry Then
         Resume RetryTemp
      Else                                                                          'Cancel clicked
         Resume RestoreOriginal       'Dataset already open ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      End If
   Case Else
      FatalError "frmExpression:OpenExpressionDataset", Err.Description
      End
   End Select
End Function
Function FillExpressionValues(newExpression As String, Optional newColorIndex As Integer = -1) _
         As Boolean '*************************************************************** Fill In Values
   '  Entry:   newExpression  Full path of Expression Dataset to open
   '           newColorIndex  SetNo of Color Set to open. -1 means don't open.
               '  This doesn't seem to do anything now.
   '  Return:  True if successful
   '           False because Expression Dataset open somewhere else or newExpression is ""
   '  Opens temporary Expression Dataset and fills in all values in Expression Manager
   '  Fills in frmExpression
   '  Title, color set names, etc.
   '  Always works with copy of same expression dataset as Drafter
   '     Expression path in expression
   '     Color sets copied into TempColorSet table
   
   Dim col As Integer, i As Integer
   Dim tdfSearch As New TableDef
   
   If newExpression = "" Then Exit Function 'Nothing to fill >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   If OpenExpressionDataset(newExpression) = False Then
      Exit Function                         'Could not open >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Fill Expression Data
   Set rsInfo = dbExpression.OpenRecordset("SELECT * FROM Info")
   Set rsExpression = dbExpression.OpenRecordset("SELECT * FROM Expression")
   txtName = Mid(newExpression, InStrRev(newExpression, "\") + 1)
   txtName = Left(txtName, InStrRev(txtName, ".") - 1)
   txtRemarks = Dat(rsInfo!remarks)
   txtNotes = Dat(rsInfo!notes)
   cmbColumns.Clear
   For col = 3 To rsExpression.Fields.count - 1  '================================Get Column Titles
      Select Case rsExpression.Fields(col).name
      Case "Notes", "Remarks"
      Case Else
         lstColumns.AddItem rsExpression.Fields(col).name
         cmbColumns.AddItem rsExpression.Fields(col).name
      End Select
   Next col
   cmbColumns.AddItem "[None]"
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up Color Sets
   For Each tdfSearch In dbExpression.TableDefs                    'If TempColorSet exists, dump it
      If tdfSearch.name = "TempColorSet" Then
         dbExpression.Execute "DROP TABLE TempColorSet"
      End If
   Next tdfSearch
   dbExpression.Execute "SELECT * INTO TempColorSet FROM ColorSet"
   FillAllColorSets
   mnuImport.Enabled = True
   mnuSaveAsExp.Enabled = True
   If Not expressionDirty Then
      '  Expression could already be dirty if there was an old Expression Dataset that was
      '  converted on opening
      mnuSaveExp.Enabled = False
   End If
   mnuNewColorSet.Enabled = True
   mnuCopyColorSet.Enabled = True
   expressionDirty = False
   makeDisplayTable = False
   FillExpressionValues = True
End Function
Sub FillAllColorSets(Optional newColorSet As String = "")
   '  Never was called with option. Can probably get rid of it.
   cmbColorSets.Clear
   Set rsColorSets = dbExpression.OpenRecordset("SELECT * FROM TempColorSet")
   If Not rsColorSets.EOF Then '===================================================Color Sets Exist
      Do Until rsColorSets.EOF '-----------------------------------------------Get Color Set Titles
         cmbColorSets.AddItem rsColorSets!colorSet
'         If rsColorSets!colorSet = newColorSet Then                   'Current color set being used
'            cmbColorSets.Tag = cmbColorSets.ListCount - 1
'         End If
         rsColorSets.MoveNext
      Loop
      If colorIndexes(0) > 0 Then '===================================Set To First Active Color Set
         cmbColorSets.Tag = colorIndexes(1)
'      If Val(cmbColorSets.Tag) = -1 Then                      'No color set coming in, set to first
'         cmbColorSets.Tag = 0
''         colorSet = cmbColorSets.List(0)
      End If
   Else '==============================================================================No Color Set
      cmbColorSets.Tag = -1
   End If
   rsColorSets.Close
   FillColorSet                                       'Fill with chosen color set data (or nothing)
End Sub

Private Sub txtName_Change()
   ChangeExpression
End Sub
Private Sub txtRemarks_Change()
   ChangeExpression
End Sub
Private Sub txtNotes_Change()
   ChangeExpression
End Sub
Sub ChangeExpression()
   mnuImport.Enabled = True
   mnuSaveExp.Enabled = True
   mnuSaveAsExp.Enabled = True
   expressionDirty = True
End Sub

'/////////////////////////////////////////////////////////////////////////////// Color Set Handling
   '  Criteria values kept in grdCriteria
   '     Column 0    List number
   '     Column 1    Label
   '     Column 2    Criterion
   '     Column 3    Color (in BackColor)
   '  Last 2 criteria are always
   '     'No criteria met'       default yellow
   '     'Not found'             default white
   '  Criterion selected kept in grdCriteria.Tag to allow return to previous selection if making
   '     a new selection is inappropriate
   '     grdCriteria.Tag = 0 indicates no selection
   
   
   '  Color set titles are kept in cmbColorSets
   '  The parameter currently being operated on is in txtColorSet
   '     A click on cmbColorSets changes cmbColorSets.Text immediately before any processing or checking
   '        can be done
   '  cmbColorSets.Tag is value of cmbColorSets.ListIndex currently being processed
   '     -1 if no param being processed
   '     When cmbColorSets is clicked, ListIndex changes so tag stores ListIndex of the one
   '        currently being processed

'######################################################################## Color Set Command Buttons
Private Sub cmdEdit_Click() '********************************************** Edit Existing Color Set
   Dim i As Integer
   
   If criterionDirty Then
      Select Case MsgBox("Do you want to save the criterion in the Criteria Builder?", _
                         vbQuestion + vbYesNoCancel)
      Case vbYes
         If Not SaveCriterion Then Exit Sub   'Invalid criterion >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case vbCancel
         criterionDirty = False
         Exit Sub                             'Dump criterion >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End Select
   End If
   
   loading = True
   EnableCriteriaBuilder                   'After changes in controls to reverse cmd button enables
   grdCriteria.Tag = grdCriteria.row                                              'Row being edited
   grdCriteria.col = 1
   txtLabel = grdCriteria.text
   grdCriteria.col = 2
   txtCriterion = grdCriteria.text
   grdCriteria.col = 3
   lblColor.BackColor = grdCriteria.CellBackColor
   If txtLabel = "No criteria met" Or txtLabel = "Not found" Then
      txtLabel.Enabled = False
      txtLabel.BackColor = vbGray
      txtCriterion.Enabled = False
      txtCriterion.BackColor = vbGray
      lstColumns.Enabled = False
      lstColumns.BackColor = vbGray
      For i = 0 To lblOperator.count - 1                                        'Turn off operators
         lblOperator(i).BackColor = vbGray
      Next i
      grdCriteria.SetFocus
   Else
      txtLabel.SetFocus
   End If
   cmdNew.Enabled = False                                             'Text changes mess with these
   cmdSave.Enabled = False
   cmdAdd.Enabled = False
   criterionMode = "Edit"
   criterionDirty = False
   loading = False
End Sub
Private Sub cmdDelete_Click() '*************************************************** Delete Criterion
   Dim i As Integer
   
   grdCriteria.RemoveItem grdCriteria.Tag
   For i = 1 To grdCriteria.rows - 1
      grdCriteria.TextMatrix(i, 0) = i
   Next i
   NoCriterionSelected
   ChangeColorSet
   makeDisplayTable = True
End Sub
Private Sub cmdUp_Click() '********************************************** Move Criterion Up In List
   Dim critLabel As String, criterion As String, color As Long, color1 As Long
   
   grdCriteria.col = 1
   critLabel = grdCriteria.text
   grdCriteria.text = grdCriteria.TextMatrix(grdCriteria.row - 1, 1)
   grdCriteria.col = 2
   criterion = grdCriteria.text
   grdCriteria.text = grdCriteria.TextMatrix(grdCriteria.row - 1, 2)
   grdCriteria.col = 3
   color = grdCriteria.CellBackColor
   grdCriteria.row = grdCriteria.row - 1
   color1 = grdCriteria.CellBackColor
   grdCriteria.CellBackColor = color
   grdCriteria.row = grdCriteria.row + 1
   grdCriteria.CellBackColor = color1
   grdCriteria.row = grdCriteria.row - 1
   grdCriteria.col = 2
   grdCriteria.text = criterion
   grdCriteria.col = 1                              'Must put in this order to leave focus on label
   grdCriteria.text = critLabel
   grdCriteria.SetFocus
   grdCriteria.Tag = grdCriteria.row
   If grdCriteria.row <= 1 Then cmdUp.Enabled = False
   cmdDown.Enabled = True
   makeDisplayTable = True
   ChangeColorSet
End Sub
Private Sub cmdDown_Click() '****************************************** Move Criterion Down In List
   Dim critLabel As String, criterion As String, color As Long, color1 As Long
   
   grdCriteria.col = 1
   critLabel = grdCriteria.text
   grdCriteria.text = grdCriteria.TextMatrix(grdCriteria.row + 1, 1)
   grdCriteria.col = 2
   criterion = grdCriteria.text
   grdCriteria.text = grdCriteria.TextMatrix(grdCriteria.row + 1, 2)
   grdCriteria.col = 3
   color = grdCriteria.CellBackColor
   grdCriteria.row = grdCriteria.row + 1
   color1 = grdCriteria.CellBackColor
   grdCriteria.CellBackColor = color
   grdCriteria.row = grdCriteria.row - 1
   grdCriteria.CellBackColor = color1
   grdCriteria.row = grdCriteria.row + 1
   grdCriteria.col = 2
   grdCriteria.text = criterion
   grdCriteria.col = 1                              'Must put in this order to leave focus on label
   grdCriteria.text = critLabel
   grdCriteria.SetFocus
   grdCriteria.Tag = grdCriteria.row
   If grdCriteria.row >= grdCriteria.rows - 3 Then cmdDown.Enabled = False
   cmdUp.Enabled = True
   makeDisplayTable = True
   ChangeColorSet
End Sub

'################################################################################## Color Set Menus
Private Sub mnuNewColorSet_Click() '******************************************* Add A New Color Set
   If colorSetDirty Then
      If Not SaveColorSet Then Exit Sub                    'Couldn't save >>>>>>>>>>>>>>>>>>>>>>>>>
'      Select Case MsgBox("Save current color set?", vbYesNoCancel, "New Color Set")
'      Case vbYes
'         If Not SaveColorSet Then Exit Sub           'Couldn't save >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'      Case vbCancel
'         Exit Sub                                    'Don't save >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'      End Select
   End If
   ClearColorSet
   EnableColorSet
   EnableCriteriaBuilder
   txtColorSet.SetFocus
End Sub
Private Sub mnuSaveColorSet_Click() '************************************** Save Existing Color Set
   '  Should be enabled only when colorSetDirty = True
'   SaveColorSet
End Sub
Private Sub mnuAddColorSet_Click() '********************* Make Existing Color Set An Additional One
   '  Should be enabled only when colorSetDirty = True
   Dim oldTag As String
   
   oldTag = cmbColorSets.Tag
   cmbColorSets.Tag = -1
   If Not SaveColorSet Then cmbColorSets.Tag = oldTag                            'Unsuccessful save
End Sub
Private Sub mnuDeleteColorSet_Click() '********************************** Delete A Single Color Set
   dbExpression.Execute _
            "DELETE FROM TempColorSet" & _
            "   WHERE ColorSet = '" & cmbColorSets.List(cmbColorSets.ListIndex) & "'"
   Set rsColorSets = dbExpression.OpenRecordset("SELECT * FROM TempColorSet")
   cmbColorSets.Clear
   Do Until rsColorSets.EOF '--------------------------------------------------Get Color Set Titles
      cmbColorSets.AddItem rsColorSets!colorSet
      rsColorSets.MoveNext
   Loop
   If rsColorSets.recordCount Then '--------------------------------------------------More Criteria
      cmbColorSets.Tag = 0                                                  'Set to first criterion
   Else '
      cmbColorSets.Tag = -1
   End If
   rsColorSets.Close
   FillColorSet
   ChangeExpression
   makeDisplayTable = True
End Sub

'################################################################################# Color Set Events
Private Sub grdCriteria_Click() '*********************************************** Pick From Criteria
   Static time As Date, d As Date
   Dim dblClick As Boolean
   
   If time = 0 Then time = DateAdd("s", -1, Now) '============================Test For Double Click
   If DateDiff("s", time, Now) < 0.75 Then
      time = Now
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   time = Now
   
   If criterionDirty Then
      Select Case MsgBox("Do you want to save the criterion in the Criteria Builder?", _
                         vbQuestion + vbYesNoCancel, "Changing to Different Criterion")
      Case vbYes
         If Not SaveCriterion Then Exit Sub   'Invalid criterion >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case vbCancel
         Exit Sub                             'Don't pick new row >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case vbNo
         criterionDirty = False               'Dump criterion
      End Select
   End If
   
   criterionDirty = False                          'All this should occur on the first click anyway
   ClearCriteriaBuilder                                                    'Selecting new criterion
   
   If grdCriteria.row > 1 And grdCriteria.row < grdCriteria.rows - 2 Then
      '  Beyond first nonfixed row and before 'No criteria met' and 'Not found'
      cmdUp.Enabled = True
   Else
      cmdUp.Enabled = False
   End If
   If grdCriteria.row < grdCriteria.rows - 3 Then 'Before 1 above 'No criteria met' and 'Not found'
      cmdDown.Enabled = True
   Else
      cmdDown.Enabled = False
   End If
   grdCriteria.Tag = grdCriteria.row
   If grdCriteria.row < grdCriteria.rows - 1 Then                               'Before 'Not found'
      '  No edits allowed on 'Not found'
      cmdEdit.Enabled = True
   Else
      cmdEdit.Enabled = False
   End If
   If grdCriteria.row < grdCriteria.rows - 2 Then         'Before 'No criteria met' and 'Not found'
      ' Can't delete 'No criteria met' and 'Not found'
      cmdDelete.Enabled = True
   Else
      cmdDelete.Enabled = False
   End If
End Sub
Private Sub grdCriteria_DblClick() '************************************** Open Editing Of Criteria
   Dim col As Integer
   
   If grdCriteria.row >= 1 And grdCriteria.row < grdCriteria.rows - 1 Then
      '  Beyond title row and before 'Not found'
      col = grdCriteria.col                                 'Save column because cmdEdit changes it
      cmdEdit_Click
      Select Case col
      Case 1
         If txtLabel.Enabled Then
            txtLabel.SetFocus
         End If
      Case 2
         If txtCriterion.Enabled Then
            txtCriterion.SetFocus
         End If
      Case 3
         lblColor_Click
      End Select
   End If
End Sub
Private Sub grdCriteria_LostFocus()
   Select Case Screen.ActiveControl.name
   Case "cmdEdit", "cmdUp", "cmdDown", "cmdDelete"
   Case Else
      cmdEdit.Enabled = False
      cmdUp.Enabled = False
      cmdDown.Enabled = False
      cmdDelete.Enabled = False
   End Select
End Sub

Private Sub cmbColorSets_Click() '************************************** Choose Different Color Set
   If cmbColorSets.ListIndex = Val(cmbColorSets.Tag) Then                           'Same color set
      Exit Sub                             'Returns to previous title >>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If colorSetDirty Then
      If Not SaveColorSet Then Exit Sub                    'Unsuccessful save >>>>>>>>>>>>>>>>>>>>>
   End If
   
   cmbColorSets.Tag = cmbColorSets.ListIndex                                'Newly chosen parameter
   FillColorSet
   expressionDirty = True
End Sub
Private Sub cmbColumns_Click() '*********************************** Choose Different Display Column
'   If cmbColumns.List(cmbColumns.ListIndex) = "[None]" Then
'      chkDisplayValue = vbUnchecked
'   End If
   ChangeColorSet
   makeDisplayTable = True
End Sub
Private Sub txtColorSet_Change() '**************************************** Change Name Of Color Set
   ChangeColorSet
End Sub

'############################################################################# Color Set Procedures
Sub FillColorSet() '******************************** Fill Controls With Individual Color Set Values
   '  cmbColorSets filled before entry and cmbColorSets.Tag assigned
   '  Sets all status variables and controls for a single color set
   
   Dim row As Integer, i As Integer
   Dim CrLf As Integer, nextCrLf As Integer, pipe1 As Integer, pipe2 As Integer
   Dim criterionLabel As String, criterion As String, criterionColor As Long
   'For GetColorSet()
     Dim labels(MAX_CRITERIA) As String, criteria(MAX_CRITERIA) As String, _
          colors(MAX_CRITERIA) As Long
     Dim notFoundIndex As Integer                       'Index of 'Not found' criterion (last one)
   
   i = Val(cmbColorSets.Tag)
   ClearColorSet                                                               'This sets Tag to -1
   cmbColorSets.Tag = i
   
   EnableColorSet                                 'Even if no color set, need access to add new one
   Set rsColorSet = dbExpression.OpenRecordset( _
         "SELECT * FROM TempColorSet" & _
         "   WHERE ColorSet = '" & cmbColorSets.List(Val(cmbColorSets.Tag)) & "'")
   If rsColorSet.EOF Then
      rsColorSet.Close
      ClearCriteriaBuilder
      NoCriterionSelected
      Exit Sub                         'Nothing to fill  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   loadingColorSet = True
   With rsColorSet
      txtColorSet = !colorSet
      cmbColorSets.ListIndex = Val(cmbColorSets.Tag)                    'Proper place in color sets
      i = 0
      Do Until cmbColumns.List(i) = !column      'Put proper column title at head of drop-down list
         i = i + 1
         If i >= cmbColumns.ListCount Then                      'In case column title isn't in list
            i = 0                                                        'Set to first item in list
            Exit Do
         End If
      Loop
      cmbColumns.ListIndex = i
      GetColorSet dbExpression, rsColorSet, labels, criteria, colors, notFoundIndex
      For i = 0 To notFoundIndex
         grdCriteria.AddItem i + 1 & vbTab & labels(i) & vbTab & criteria(i)
         grdCriteria.row = grdCriteria.rows - 1
         grdCriteria.col = 3
         grdCriteria.CellBackColor = colors(i)
      Next i
   End With
   grdCriteria.RemoveItem 1                      'Remove 'No criteria met' and 'Not found' defaults
   grdCriteria.RemoveItem 1                     'Here to avoid removing last nonfixed row stupidity
   rsColorSet.Close
   Set rsColorSet = Nothing
   ClearCriteriaBuilder
   NoCriterionSelected
   mnuDeleteColorset.Enabled = True
   mnuSaveColorSet.Enabled = False
   mnuAddColorSet.Enabled = False
   criterionDirty = False
   loadingColorSet = False
   colorSetDirty = False
End Sub
Function SaveColorSet() As Boolean '************************************************ Save Color Set
   '  index of parameter being processed in cmbColorSets.Tag.
   '  If cmbColorSets.Tag < 0 then it is a new colorset (either totally new or an additional one
   '  created from an old color set).
   '  cmbColorSets.Text is usually the next parameter set to be processed
   Dim sql As String, criteria As String, i As Integer
   
   If criterionDirty Then                                           'Usually occurs in color change
      If Not SaveCriterion Then
         SaveColorSet = False
         Exit Function                                     'Invalid criterion >>>>>>>>>>>>>>>>>>>>>
      End If
'      Select Case MsgBox("Save criterion in Criteria Builder?", vbQuestion + vbYesNoCancel, _
'                         "Saving Color Set")
'      Case vbYes
'         If Not SaveCriterion Then
'            SaveColorSet = False
'            Exit Function                                'Invalid criterion >>>>>>>>>>>>>>>>>>>>>>>
'         End If
'      Case vbCancel
'         SaveColorSet = False
'         Exit Function                                    'Cancel save >>>>>>>>>>>>>>>>>>>>>>>>>>>>
'      End Select
   End If
   
   If Not colorSetDirty Then
      SaveColorSet = True                                          'Can continue if nothing to save
      Exit Function                                      'No changes >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   txtColorSet = TextToSql(Dat(txtColorSet))
   
   If txtColorSet = "" Then
      MsgBox "Color Set must have a name.", vbExclamation + vbOKOnly, "Saving Color Set"
      txtColorSet.SetFocus
      SaveColorSet = False
      Exit Function                                       'No name >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If InvalidChr(txtColorSet, "color set name") Then
      txtColorSet.SetFocus
      SaveColorSet = False
      Exit Function                           'Invalid chr in name >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   For i = 0 To cmbColorSets.ListCount - 1
      If i <> Val(cmbColorSets.Tag) And txtColorSet = cmbColorSets.List(i) Then
         MsgBox "Color set name already exists.", vbExclamation + vbOKOnly, "Saving Color Set"
         txtColorSet.SetFocus
         SaveColorSet = False
         Exit Function                                    'Name exists >>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Next i
      
   If cmbColumns.ListIndex = -1 Then
      MsgBox "Must choose a Gene Value column.", vbExclamation + vbOKOnly, "Saving Color Set"
      SaveColorSet = False
      Exit Function                                       'No Gene Value >>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If Val(cmbColorSets.Tag) < 0 Then '------------------------------------------------New Color Set
      Set rsColorSets = dbExpression.OpenRecordset("SELECT * FROM TempColorSet")
      Do Until rsColorSets.EOF '................................................Check For Same Name
         If rsColorSets!colorSet = txtColorSet Then '
            rsColorSets.Close
            MsgBox "Cannot add color set. It has the same name as another color set.", _
                   vbExclamation + vbOKOnly, "Saving Color Set"
            SaveColorSet = False
            Exit Function                                 'Duplicate name >>>>>>>>>>>>>>>>>>>>>>>>>
         End If
         rsColorSets.MoveNext
      Loop
      rsColorSets.Close
      cmbColorSets.AddItem txtColorSet                                           'Add to combo list
      cmbColorSets.Tag = cmbColorSets.ListCount - 1                         'Change tag to new item
      sql = txtColorSet                                    'Force return of no rows in SELECT below
   Else '------------------------------------------------------------------------Existing Color Set
      sql = cmbColorSets.List(Val(cmbColorSets.Tag))                'Will SELECT name before change
      cmbColorSets.List(Val(cmbColorSets.Tag)) = txtColorSet                   'Change name in list
   End If
   
   For i = 1 To grdCriteria.rows - 1 '--------------------------------------------Assemble Criteria
      grdCriteria.row = i
      grdCriteria.col = 1
      criteria = criteria & grdCriteria.text & "|"
      grdCriteria.col = 2
      criteria = criteria & grdCriteria.text & "|"
      grdCriteria.col = 3
      criteria = criteria & Format(grdCriteria.CellBackColor) & vbCrLf
   Next i
   Set rsColorSets = dbExpression.OpenRecordset( _
            "SELECT * FROM TempColorSet WHERE ColorSet = '" & sql & "'")
                                                                  'Returns no rows if new color set
   With rsColorSets '---------------------------------------------------------------Update Database
      If .EOF Then
         .AddNew
      Else
         .edit
      End If
      !colorSet = txtColorSet
      !column = cmbColumns.List(cmbColumns.ListIndex)
      !criteria = criteria
      .Update
   End With
   rsColorSets.Close
'   colorSet = txtColorSet
   ChangeExpression
   colorSetDirty = False
   mnuSaveColorSet.Enabled = False
   mnuAddColorSet.Enabled = False
   NoCriterionSelected
   SaveColorSet = True
End Function
Sub ClearColorSet() '******************************************** Clear All Items For One Color Set
   Dim i As Integer
   
   txtColorSet = ""
   cmbColorSets.Tag = -1    'Must do this elsewhere   ????????????????
   colorSetDirty = False
   For i = 1 To grdCriteria.rows - 2   '1 to leave title, -2 because can't remove last nonfixed row
      grdCriteria.RemoveItem (1)
   Next i
   grdCriteria.row = 1
   grdCriteria.TextMatrix(1, 0) = 1                     'First row always exists. Can't use AddItem
   grdCriteria.TextMatrix(1, 1) = "No criteria met"     'First row always exists. Can't use AddItem
   grdCriteria.TextMatrix(1, 2) = ""
   grdCriteria.col = 3
   grdCriteria.CellBackColor = DEFAULT_NOTMET_COLOR
   grdCriteria.AddItem 2 & vbTab & "Not found"
   grdCriteria.row = 2
   grdCriteria.col = 3
   grdCriteria.CellBackColor = vbWhite
   txtCriterion = ""                                   'Sets criterionDirty to true in Change event
   grdCriteria.Enabled = False
   grdCriteria.BackColor = vbGray
   lblColor.BackColor = vbWhite
   NoCriterionSelected
   cmdNew.Enabled = False
   mnuSaveColorSet.Enabled = False
   mnuAddColorSet.Enabled = False
   mnuDeleteColorset = False
   criterionDirty = False
End Sub
Sub EnableColorSet()
   txtColorSet.Enabled = True
   txtColorSet.BackColor = vbWhite
   grdCriteria.Enabled = True
   grdCriteria.BackColor = vbWhite
   cmbColumns.Enabled = True
   cmbColumns.BackColor = vbWhite
   cmdNew.Enabled = True
End Sub
Sub ChangeColorSet() '**************************************************** Change Made To Color Set
   colorSetDirty = True
   mnuAddColorSet.Enabled = True
   mnuSaveExp.Enabled = True
   mnuSaveAsExp.Enabled = True
End Sub
Sub NoCriterionSelected() '**************************************** Turn Off Stuff For No Selection
   grdCriteria.Tag = 0                                         'Nothing selected, this is title row
   txtLabel.Enabled = False
   txtLabel.BackColor = vbGray
   txtCriterion.Enabled = False
   txtCriterion.BackColor = vbGray
   lblColor.Enabled = False
   lblColor.BackColor = vbGray
   cmdUp.Enabled = False
   cmdDown.Enabled = False
   cmdEdit.Enabled = False
   cmdDelete.Enabled = False
   cmdSave.Enabled = False
   cmdAdd.Enabled = False
End Sub

Rem ///////////////////////////////////////////////////////////////////// Criteria Builder Handling
   '  The criterion currently being operated on is in txtCriterion
   '  Text in txtCriterion must always represent grdCriteria.Tag row in grdCriteria
   '     Unless New Criterion
   '  Criterion command buttons
   '     cmdNew      Enabled when valid color set and criterionMode not "Edit" or "New"
   '                 Criteria Builder is inactive
   '     cmdSave     Enabled when editing new or existing criterion (any criterionMode)
   '     cmdAdd      Enabled when editing existing criterion (criterionMode = "Edit")

Rem ############################################################## Criteria Builder Command Buttons
Private Sub cmdNew_Click()
   ClearCriteriaBuilder
   EnableCriteriaBuilder
   criterionMode = "New"
   txtLabel.SetFocus
End Sub
Private Sub cmdSave_Click()
   SaveCriterion
End Sub
Private Sub cmdAdd_Click()
   SaveCriterion
End Sub

Rem ####################################################################### Criteria Builder Events
Private Sub lblColor_Click()
   dlgDialog.CancelError = True
On Error GoTo CancelError
   dlgDialog.FLAGS = cdlCCRGBInit
   dlgDialog.ShowColor
   lblColor.BackColor = dlgDialog.color
   ChangeCriterion
CancelError:
End Sub
Private Sub lstColumns_Click() '****************************************************** Place Column
   Dim selectStart As Integer                  'Changing text in textbox sets SelStart back to zero
   
   If lstColumns.ListIndex = -1 Then Exit Sub       'Click event caused by unhighlighting selection
   selectStart = txtCriterion.SelStart
   txtCriterion = Left(txtCriterion, txtCriterion.SelStart) & "[" & lstColumns.text & "]" _
                & Mid(txtCriterion.text, txtCriterion.SelStart + txtCriterion.SelLength + 1)
   txtCriterion.SelStart = selectStart + Len(lstColumns.text) + 2                   '2 for brackets
   txtCriterion.SetFocus
   lstColumns.ListIndex = -1                                                 'Unhighlight selection
   criterionDirty = True
End Sub
Private Sub lblOperator_Click(index As Integer) '************************ Put Operator in Criterion
   Dim selectStart As Integer                  'Changing text in textbox sets SelStart back to zero
   
   If Not txtCriterion.Enabled Then Exit Sub                   'Disabled >>>>>>>>>>>>>>>>>>>>>>>>>>
      
   selectStart = txtCriterion.SelStart
   txtCriterion.text = Left(txtCriterion, txtCriterion.SelStart) & lblOperator(index).Caption _
                     & Mid(txtCriterion, txtCriterion.SelStart + txtCriterion.SelLength + 1)
   txtCriterion.SelStart = selectStart + Len(lblOperator(index))
   txtCriterion.SetFocus
End Sub
Private Sub txtLabel_Change()
   ChangeCriterion
End Sub
Private Sub txtCriterion_Change()
   ChangeCriterion
End Sub

Rem ################################################################### Criteria Builder Procedures
Sub ChangeCriterion() '*************************************** Some Change Made In Criteria Builder
   criterionDirty = True
   If criterionMode = "Edit" Then
      cmdSave.Enabled = True
   End If
   If txtCriterion.Enabled Then                       'Cannot Add 'No criteria met' and 'Not found'
      If grdCriteria.rows < MAX_CRITERIA Then                 'Don't allow add if too many criteria
         cmdAdd.Enabled = True
      End If
   End If
   mnuSaveExp.Enabled = True
   mnuSaveAsExp.Enabled = True
End Sub
Sub EnableCriteriaBuilder() '*************************************** Clear For New Or Edit Criteria
   Dim i As Integer
   
   lstColumns.Enabled = True
   lstColumns.BackColor = vbWhite
   For i = 0 To lblOperator.count - 1                                            'Turn on operators
      lblOperator(i).BackColor = vbWhite
   Next i
   txtLabel = ""
   txtLabel.Enabled = True
   txtLabel.BackColor = vbWhite
   lblColor.BackColor = vbWhite
   lblColor.Enabled = True
   txtCriterion = ""
   txtCriterion.Enabled = True
   txtCriterion.BackColor = vbWhite
   cmdNew.Enabled = False
   cmdSave.Enabled = False
   cmdAdd.Enabled = False
   cmdUp.Enabled = False                             'All disabled if something in Criteria Builder
   cmdDown.Enabled = False
   cmdEdit.Enabled = False
   cmdDelete.Enabled = False
End Sub
Sub ClearCriteriaBuilder() '************************************************ Clear Criteria Builder
   '  Sets Criteria Builder to disabled condition.
   '  After:
   '     Entering Expression Manager
   '     Changing Expression datasets
   '     Changing color sets
   '     Adding or editing criterion
   '  Leaves cmdNew enabled. Handle this elsewhere if new criterion impossible
   
   Dim i As Integer
   
   If criterionDirty Then
      Select Case MsgBox("Do you want to save the criterion in the Criteria Builder?", _
                         vbQuestion + vbYesNoCancel)
      Case vbYes
         If Not SaveCriterion Then Exit Sub   'Invalid criterion >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case vbCancel
         criterionDirty = False
         Exit Sub                             'Dump criterion >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End Select
   End If
   
   lstColumns.Enabled = False
   lstColumns.BackColor = vbGray
   For i = 0 To lblOperator.count - 1                                           'Turn off operators
      lblOperator(i).BackColor = vbGray
   Next i
   txtLabel = ""
   txtLabel.Enabled = False
   txtLabel.BackColor = vbGray
   txtCriterion = ""
   txtCriterion.Enabled = False
   txtCriterion.BackColor = vbGray
   lblColor.BackColor = vbGray
   lblColor.Enabled = False
   If grdCriteria.rows < MAX_CRITERIA Then
      cmdNew.Enabled = True
   Else
      cmdNew.Enabled = False
   End If
   cmdSave.Enabled = False
   cmdAdd.Enabled = False
   criterionMode = ""
   grdCriteria.Tag = -1
   criterionDirty = False
End Sub
Function SaveCriterion() As Boolean '*************************************** Save Current Criterion
   '  Moves criterion data into grdCriteria
   '  grdCriteria.Tag determines row. New row if Tag = -1 or active control is cmdAdd
   Dim i As Integer
   
   If Not criterionDirty Then
      SaveCriterion = True
      Exit Function                                    'Nothing to save >>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   txtLabel = TextToSql(Dat(txtLabel))
   txtCriterion = Dat(txtCriterion)
   
   If Not CheckCriterion(txtCriterion) Then
      SaveCriterion = False
      Exit Function                                    'Invalid criterion >>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If Val(grdCriteria.Tag) <= 0 Or ActiveControl Is cmdAdd Then '---------------------New Criterion
      grdCriteria.AddItem grdCriteria.rows - 2 & vbTab & txtLabel & vbTab & txtCriterion, _
                          grdCriteria.rows - 2
      grdCriteria.row = grdCriteria.row - 1                     'AddItem sets to row after addition
      grdCriteria.col = 3
      grdCriteria.CellBackColor = lblColor.BackColor
      grdCriteria.Tag = grdCriteria.rows - 3
      For i = Val(grdCriteria.Tag) + 1 To grdCriteria.rows - 1             'Renumber following rows
         grdCriteria.TextMatrix(i, 0) = i
      Next i
   Else '----------------------------------------------------------------Replace Existing Criterion
      grdCriteria.row = Val(grdCriteria.Tag)
      grdCriteria.col = 1
      grdCriteria.text = txtLabel
      grdCriteria.col = 2
      If grdCriteria.text <> txtCriterion Then makeDisplayTable = True
      grdCriteria.text = txtCriterion
      grdCriteria.col = 3
      If grdCriteria.CellBackColor <> lblColor.BackColor Then makeDisplayTable = True
      grdCriteria.CellBackColor = lblColor.BackColor
   End If
   makeDisplayTable = True
   criterionDirty = False
   ClearCriteriaBuilder
   ChangeColorSet
   SaveCriterion = True
End Function
Rem ########################################################################## Criterion Procedures
Function CheckCriterion(criterion As String) As Boolean '******************* See If Criterion Works
   '  Tries an SQL statement with the criterion in the WHERE clause. Returns False if error
   '  Checks colors, label, criterion. Cannot duplicate or be white
   
   Dim i As Integer, sql As String, openBracket As String, closeBracket As String
   Dim duplicateColor As Boolean, duplicateLabel As Boolean, duplicateCriterion As Boolean

   criterion = Dat(criterion)
   If txtLabel = "" Then
      MsgBox "Must have label.", vbExclamation + vbOKOnly
      txtLabel.SetFocus
      CheckCriterion = False
      Exit Function                                    'No label >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   If InvalidChr(txtLabel, "criterion label") Then
      txtLabel.SetFocus
      CheckCriterion = False
      Exit Function                               'Invalid label >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If criterion = "" And txtCriterion.Enabled Then
      MsgBox "No criterion.", vbExclamation + vbOKOnly
      txtCriterion.SetFocus
      CheckCriterion = False
      Exit Function                                'No criterion >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If lblColor.BackColor = vbWhite Then
      duplicateColor = True
   Else
      grdCriteria.col = 3
      For i = 1 To grdCriteria.rows - 1
         grdCriteria.row = i
         If Val(grdCriteria.Tag) <> i Or ActiveControl Is cmdAdd Then        'Not current criterion
            grdCriteria.row = i
            If grdCriteria.CellBackColor = lblColor.BackColor Then
               duplicateColor = True
            End If
            If grdCriteria.TextMatrix(i, 1) = txtLabel Then
               duplicateLabel = True
            End If
            If grdCriteria.TextMatrix(i, 2) = criterion Then
               If grdCriteria.TextMatrix(i, 1) <> "No criteria met" And grdCriteria.TextMatrix(i, 1) <> "Not found" Then
                  duplicateCriterion = True
               End If
            End If
         End If
      Next i
   End If
   If duplicateColor Then
      MsgBox "Pick a different color first.", vbOKOnly + vbExclamation
      CheckCriterion = False
      GoTo ExitFunction         'No new color vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   If duplicateLabel Then
      MsgBox "Same label already exists.", vbOKOnly + vbExclamation
      CheckCriterion = False
      GoTo ExitFunction         'No new label vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   If duplicateCriterion Then
      MsgBox "Same criterion already exists", vbOKOnly + vbExclamation
      CheckCriterion = False
      GoTo ExitFunction         'No new color vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
On Error GoTo ErrHandler
   If Val(grdCriteria.Tag) < grdCriteria.rows - 2 Then                'Last 2 rows have no criteria
      Set rsExpression = dbExpression.OpenRecordset( _
                       "SELECT * FROM Expression WHERE " & criterion)
      sql = criterion
      openBracket = InStr(sql, "[")
      Do While openBracket
         closeBracket = InStr(openBracket + 1, sql, "]")
         If closeBracket = 0 Then GoTo ErrHandler
         Select Case rsExpression(Mid(sql, openBracket, closeBracket - openBracket + 1)).Type
         Case dbSingle, dbDouble
            sql = Left(sql, openBracket - 1) & "1234" & Mid(sql, closeBracket + 1)
         Case Else
            sql = Left(sql, openBracket - 1) & "'Test'" & Mid(sql, closeBracket + 1)
         End Select
         openBracket = InStr(sql, "[")
      Loop
      Set rsExpression = dbExpression.OpenRecordset( _
                       "SELECT * FROM Expression WHERE " & sql)
      rsExpression.Close
   End If
   CheckCriterion = True
ExitFunction:
   Exit Function                   'Normal Exit >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
ErrHandler:
   MsgBox "Error in criterion expression. Correct it before proceeding.", _
             vbCritical + vbOKOnly, "Criterion Check"
   CheckCriterion = False
   Resume ExitFunction
'   beenHere = True                                      'cmbColorSets.Text only change if this happens
End Function
Sub FillSystemsList() '**************************************************** Fills List On frmDataID
   Dim rsSystems As Recordset
   Dim system As Integer                                                          'Zero-based index
   
   frmDataID.lstSystems.Clear
   frmDataID.lstSystemCodes.Clear
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Systems
   Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems ORDER BY System", dbOpenForwardOnly)
   Do Until rsSystems.EOF
      If VarType(rsSystems!Date) <> vbNull Or rsSystems!system = "Other" Then  'Supported system
         '  Other system is always supported, date or not
         frmDataID.lstSystems.AddItem rsSystems!system
         frmDataID.lstSystemCodes.AddItem rsSystems!systemCode
         system = system + 1
      End If
      rsSystems.MoveNext
   Loop
End Sub


