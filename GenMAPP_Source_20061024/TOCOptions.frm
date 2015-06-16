VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmTOCOptions_Old 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table of Contents Options"
   ClientHeight    =   5604
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5604
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   4920
      MaskColor       =   &H0000C0C0&
      TabIndex        =   20
      Top             =   4140
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
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
      Left            =   4740
      TabIndex        =   19
      Top             =   60
      Width           =   972
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
      Height          =   312
      Left            =   4680
      TabIndex        =   18
      Top             =   5220
      Width           =   972
   End
   Begin VB.TextBox txtPercent 
      Height          =   288
      Left            =   2580
      TabIndex        =   17
      Text            =   "0"
      Top             =   5220
      Width           =   372
   End
   Begin VB.ComboBox cmbCriterion 
      Height          =   288
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4860
      Visible         =   0   'False
      Width           =   2472
   End
   Begin VB.ComboBox cmbColorSet 
      Height          =   288
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4500
      Visible         =   0   'False
      Width           =   2472
   End
   Begin VB.TextBox txtExpression 
      Height          =   288
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Double-click to specify an Expression Dataset. "
      Top             =   4140
      Width           =   4752
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   60
      Top             =   5880
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.ListBox lstCatalog 
      Height          =   816
      Left            =   2640
      TabIndex        =   9
      Top             =   2940
      Width           =   1572
   End
   Begin VB.TextBox txtGene 
      Height          =   288
      Left            =   2640
      TabIndex        =   7
      Top             =   2340
      Width           =   1572
   End
   Begin MSFlexGridLib.MSFlexGrid grdGenes 
      Height          =   1632
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "If no genes are selected, Table of Contents will include all genes."
      Top             =   2160
      Width           =   2232
      _ExtentX        =   3937
      _ExtentY        =   2879
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      TextStyleFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkSubFolders 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Include subfolders"
      Height          =   192
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   1812
   End
   Begin VB.DirListBox dirMAPPs 
      Height          =   1152
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   3132
   End
   Begin VB.DriveListBox drvMAPPs 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   1272
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Percent Meeting Criterion"
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
      Index           =   7
      Left            =   60
      TabIndex        =   16
      Top             =   5220
      Width           =   2268
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Criterion"
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
      Index           =   6
      Left            =   60
      TabIndex        =   14
      Top             =   4860
      Width           =   768
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Set"
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
      Left            =   60
      TabIndex        =   12
      Top             =   4500
      Width           =   828
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expression Dataset"
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
      Left            =   60
      TabIndex        =   10
      Top             =   3840
      Width           =   1752
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cataloging system"
      Height          =   192
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   2700
      Width           =   1332
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gene ID"
      Height          =   192
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   2100
      Width           =   588
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Include Genes"
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
      Left            =   60
      TabIndex        =   4
      Top             =   1920
      Width           =   1272
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MAPP Locations"
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
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1488
   End
End
Attribute VB_Name = "frmTOCOptions_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCatalogExpression As Database
Dim rsCatalogColorSet As Recordset
Dim expressionChanged As Boolean, loading As Boolean
Dim CatalogDirty As Boolean
Dim SystemCode(MAX_PRIMARY_TYPES) As String


Private Sub cmdHelp_Click()
'   frmTOC.mnuHelp_Click
End Sub

Private Sub cmdOK_Click()
   Hide
End Sub

Private Sub drvMAPPs_Change()
   dirMAPPs = drvMAPPs
End Sub

Private Sub Form_Load()
   With grdGenes
      .row = 0
      .col = 0
      .CellFontBold = True
      .ColWidth(0) = 1500
      .ColAlignment(0) = flexAlignLeftCenter
      .TextMatrix(0, 0) = "Gene"
      .col = 1
      .CellFontBold = True
      .ColWidth(1) = 1500
      .ColAlignment(1) = flexAlignLeftCenter
      .TextMatrix(0, 1) = "Catalog"
      .col = 2                                                        'Holds invisible primary type
      .ColWidth(2) = 0
   End With
'   cmbColumns.ListIndex = 0                                                'This calls Click event
End Sub
Private Sub grdGenes_Click()
   grdGenes.col = 1
   lstCatalog.ListIndex = -1                          'Default to no selection (causes Click event)
   For i = 0 To lstCatalog.ListCount - 1
      If lstCatalog.List(i) = grdGenes.text Then
         lstCatalog.ListIndex = i
      End If
   Next i
   grdGenes.col = 0
   txtGene = grdGenes.text
   txtGene.SetFocus
End Sub

Private Sub txtGene_Change()
   grdGenes.col = 0
   grdGenes.text = UCase(txtGene)
   EmptyGeneRow
End Sub
Private Sub lstCatalog_Click()
   If lstCatalog.ListIndex = -1 Then Exit Sub              'Nothing selected >>>>>>>>>>>>>>>>>>>>>>
   grdGenes.col = 1
   grdGenes.text = lstCatalog
   grdGenes.col = 2
   grdGenes.text = SystemCode(lstCatalog.ListIndex)
End Sub
Sub EmptyGeneRow()
   If grdGenes.TextMatrix(grdGenes.rows - 1, 0) <> "" Then         'Last row should always be empty
      grdGenes.rows = grdGenes.rows + 1
   ElseIf grdGenes.TextMatrix(grdGenes.rows - 2, 0) = "" Then                  'But not last 2 rows
      grdGenes.rows = grdGenes.rows - 1
   End If
End Sub

Private Sub txtExpression_LostFocus()
   If txtExpression = txtExpression.Tag Then Exit Sub      ' No change >>>>>>>>>>>>>>>>>>>>>>>>>>>>

   SetExpression
End Sub
Private Sub txtExpression_DblClick()
   Dim cancelOpen As Boolean
   Dim newExpression As String, oldExpression As String
   
Retry:
On Error GoTo OpenError
   dlgDialog.CancelError = True
   dlgDialog.DialogTitle = "Specify Expression Dataset"
   dlgDialog.Filter = "Expression (.gex)|*.gex"
'   dlgDialog.FileName = mruDataSet & "*.gex"    'Set in frmtoc. Controls CatalogDataPath
   dlgDialog.FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
   dlgDialog.ShowOpen
   newExpression = dlgDialog.FileName
   If InStr(newExpression, ".") = 0 Then
      newExpression = newExpression & ".gex"
   End If
On Error GoTo 0
   
   If Dir(newExpression) = "" Then
      If MsgBox("Expression dataset '" & newExpression & " does not exist.", _
               vbExclamation + vbRetryCancel, "Specify Expression Dataset") = vbCancel Then
         GoTo ExitSub           'Canceled opening a dataset >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Else
         GoTo Retry
      End If
   End If
   txtExpression = newExpression
   SetExpression
ExitSub:
   Exit Sub

OpenError:
   If Err <> 32755 Then
      MsgBox Err.Description, vbCritical, "Specify Expression Dataset Error"
   End If
   On Error GoTo 0
   Resume ExitSub

End Sub

Private Sub cmbColorSet_Click()
   If Not cmbColorSet.Visible Then Exit Sub                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   SetColorSet
End Sub
Private Sub cmbCriterion_Click()
   SetCriterion
End Sub
Sub ChangeCatalog()
   CatalogDirty = True

End Sub
Sub SetCatalog()
   lstCatalog.Clear
   '  Must come from PrimaryTypes table  ???????????????????????????????????
   lstCatalog.AddItem "SwissProt":  SystemCode(0) = "S"
   lstCatalog.AddItem "GenMAPP":    SystemCode(1) = "G"
   lstCatalog.AddItem "Other":      SystemCode(2) = "O"
End Sub
Sub SetExpression(Optional expression As String = "")
   '  Setting the expression opens up the Color Set choices
   
   If expression <> "" Then txtExpression = expression
   
   If txtExpression = "" Then
      Set dbCatalogExpression = Nothing
      cmbColorSet.Clear
      cmbColorSet.Visible = False
      cmbCriterion.Clear
      cmbCriterion.Visible = False
      Exit Sub                                             'No expression >>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
On Error GoTo ErrorHandler
   txtExpression.Tag = txtExpression
   Set dbCatalogExpression = OpenDatabase(txtExpression, , True)
   cmbColorSet.Visible = True
   cmbColorSet.Clear
   cmbColorSet.AddItem "ANY"
   Set rsCatalogColorSet = dbCatalogExpression.OpenRecordset("ColorSet", dbOpenTable)
   Do Until rsCatalogColorSet.EOF
      cmbColorSet.AddItem rsCatalogColorSet!colorSet
      rsCatalogColorSet.MoveNext
   Loop
   cmbColorSet.ListIndex = 0                                                  'Default to "ANY"
   rsCatalogColorSet.Close
   Set rsCatalogColorSet = Nothing
   
ExitSub:
   Exit Sub

ErrorHandler:
   If Err.number = 3024 Then
      MsgBox "Expression dataset" & vbCrLf & vbCrLf & txtExpression & vbCrLf & vbCrLf _
             & "does not exist.", vbExclamation + vbOKOnly
      Set dbCatalogExpression = Nothing
      cmbColorSet.Visible = False
      cmbCriterion.Visible = False
      txtPercent.Visible = False
      txtExpression.SetFocus
      Resume ExitSub
   Else
      FatalError "Catalog:SetExpression", Err.Description
   End If
End Sub
Sub SetColorSet(Optional colorSet As String = "")
   '  Setting the color set to other than ANY opens up the criterion choices
   Dim labels(MAX_CRITERIA) As String, criteria(MAX_CRITERIA) As String
   Dim colors(MAX_CRITERIA) As Long, notFoundIndex As Integer
   
   If dbCatalogExpression Is Nothing Then
      cmbColorSet.Visible = False
      Exit Sub                                             'No expression >>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If colorSet = "" Then
      colorSet = cmbColorSet.text
   Else
      With cmbColorSet
         .ListIndex = .ListCount - 1                   'Start at last Color Set and move toward ANY
         Do Until .ListIndex = 0 Or cmbColorSet.List(.ListIndex) = colorSet
            .ListIndex = .ListIndex - 1
         Loop                    '.ListIndex will be left at received Color Set or ANY in not found
      End With
   End If
         
   If cmbColorSet.text = "ANY" Then                                            'No Criteria To List
      cmbCriterion.Visible = False
   Else '----------------------------------------------------------List Criteria For That Color Set
      Set rsCatalogColorSet = dbCatalogExpression.OpenRecordset( _
                          "SELECT * FROM ColorSet WHERE ColorSet = '" & cmbColorSet.text & "'")
'      GetColorSet  labels, criteria, colors, notFoundIndex, rsCatalogColorSet!criteria
      cmbCriterion.Clear
      cmbCriterion.AddItem "ANY"
      cmbCriterion.ListIndex = 0
      For i = 1 To notFoundIndex - 2
         cmbCriterion.AddItem labels(i)
      Next i
      cmbCriterion.Visible = True
   End If
End Sub
Sub SetCriterion()
   If cmbCriterion.text = "ANY" Then
      txtPercent.Visible = False
   Else
      txtPercent.Visible = True
   End If
End Sub
