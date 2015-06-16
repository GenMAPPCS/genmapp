VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmTOC_Old 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000C0C0&
   Caption         =   "GenMAPP Table of Contents"
   ClientHeight    =   4644
   ClientLeft      =   132
   ClientTop       =   732
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4644
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   720
      Top             =   4860
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgContents 
      Left            =   120
      Top             =   4860
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TOC.frx":0000
            Key             =   "ClosedFolder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TOC.frx":0452
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TOC.frx":08A4
            Key             =   "MAPP"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treContents 
      Height          =   4512
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6732
      _ExtentX        =   11875
      _ExtentY        =   7959
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imgContents"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmTOC_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CatalogMappPath As String
Dim CatalogDataPath As String  'Might use global mruDataSet only
Dim CatalogExpression As String, CatalogColorSet As String, CatalogCriterion As String
Dim CatalogSubfolders As Boolean, CatalogPercent As Integer
Dim dbCatalogMapp As Database, rsCatalogObject As Recordset, sqlGenes As String
Dim CatalogSaveError As Boolean, CatalogDirty As Boolean

'////////////////////////////////////////////////////////////////////////////// Initial Forms Setup
Private Sub Form_Load()
   '  On first form load, set up frmtocOptions, which stores all the Catalog parameters
   Initialize
End Sub
Sub Initialize()
'   Dim colon As Integer
'
'   With frmtocOptions                                                'Stores current Catalog parameters
''      .Show
'      .SetCatalog
'      colon = InStr(mruMappPath, ":")
'      .drvMAPPs = Left(mruMappPath, colon)
'      .dirMAPPs = mruMappPath
'      .chkSubFolders = vbChecked
'      .dlgDialog.FileName = CatalogDataPath & "*.gex"
'      .SetExpression expression
'      .SetColorSet colorSet
'      .SetCriterion
''      .Hide
'   End With
'Exit Sub
'
'   CatalogMappPath = mruMappPath
'   CatalogDataPath = mruDataSet    'Might not use
'   CatalogExpression = expression
'   CatalogColorSet = colorSet
'   CatalogCriterion = "ANY"
'   CatalogSubfolders = True
'   CatalogPercent = 0
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////// Menus
'//////////////////////////////////////////////////////////////////////////////////////// File Menu
Private Sub mnuOpen_Click()
   Dim newCatalogName As String                   'Temporary until beyond cancel point and mnuNew_Click
   Dim Lin As String, pipe As Integer, prevPipe As Integer

   If CatalogDirty Then
      Select Case MsgBox("Save current Table of Contents?", vbYesNoCancel + vbQuestion, "Open Table of Contents")
      Case vbYes
         mnuSave_Click
         If CatalogSaveError Then
            CatalogSaveError = False
            Exit Sub                                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      Case vbCancel
         Exit Sub                                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>
      Case Else
         CatalogDirty = False                                        'Declined to save, forget Catalogdirty
      End Select
   End If
   
'On Error GoTo OpenError
   If commandLine = "" Then
      dlgDialog.CancelError = True
      dlgDialog.Filter = "Tables of Contents (.gct)|gtc"
      dlgDialog.FileName = mruCatalog & "*.gct"
      dlgDialog.FLAGS = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly
      dlgDialog.ShowOpen
      newCatalogName = dlgDialog.FileName
      If InStr(newCatalogName, ".") = 0 Then
         newCatalogName = newCatalogName & ".gct"
      End If
   Else
      newCatalogName = commandLine
      commandLine = ""
   End If
'On Error GoTo 0
   
   MousePointer = vbHourglass
   Open newCatalogName For Input As #1
   Line Input #1, CatalogMappPath
   Line Input #1, Lin
   If Lin = "" Then
      CatalogSubfolders = False
   Else
      CatalogSubfolders = True
   End If
   Line Input #1, CatalogExpression
   Line Input #1, CatalogColorSet
   Line Input #1, s
   CatalogPercent = Val(s)
'   With frmTOCOptions.grdGenes
'      i = 0
'      .rows = 2
'      Do Until EOF(1)
'         Line Input #1, Lin
'         i = i + 1
'         pipe = InStr(Lin, "|")
'         .TextMatrix(i, 0) = Left(Lin, pipe - 1)                                           'Gene ID
'         prevPipe = pipe
'         pipe = InStr(prevPipe + 1, Lin, "|")
'         .TextMatrix(i, 1) = Mid(Lin, prevPipe + 1, pipe - prevPipe - 1)                 'Type name
'         .TextMatrix(i, 2) = Mid(Lin, pipe + 1)                                          'Type code
'         .rows = .rows + 1
'      Loop
'   End With
   CatalogDirty = False
   catalogName = newCatalogName
   mruCatalog = Left(catalogName, InStrRev(catalogName, "\"))
End Sub
Private Sub mnuSave_Click()
'  gct File Structure by line
'     1  CatalogMappPath
'     2  Include subfolders: Anything on this line means to include
'     3  CatalogExpression
'     4  CatalogColorSet
'     5  CatalogCriterion
'     6  CatalogPercent
'     7 on. This fills up frmtocOptions.grdGenes. For each line:
'        Primary|PrimaryType String|PrimaryType

   If catalogName = "" Then
      mnuSaveAs_Click
      Exit Sub                                                  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   Open catalogName For Output As #1
   Print #1, CatalogMappPath
   If CatalogSubfolders Then
      Print #1, "Include subfolders"
   Else
      Print #1, ""
   End If
   Print #1, CatalogExpression
   Print #1, CatalogColorSet
   Print #1, CatalogCriterion
   Print #1, CatalogPercent
'   With frmTOCOptions.grdGenes
'      For i = 1 To .rows - 2
'         Print #1, .TextMatrix(i, 0) & "|" & .TextMatrix(i, 1) & "|" & .TextMatrix(i, 2)
'      Next i
'   End With
   Close #1
End Sub
Private Sub mnuSaveAs_Click()
   Dim oldCatalogName As String
   
On Error GoTo SaveError
   oldCatalogName = catalogName                                    'In case of error on new Catalog
ReEnter:
   dlgDialog.CancelError = True
   dlgDialog.Filter = "gct"
   dlgDialog.FileName = mruCatalog & "*.gct"
   dlgDialog.FLAGS = cdlOFNExplorer + cdlOFNHideReadOnly
   dlgDialog.ShowSave
   catalogName = dlgDialog.FileName
   If InStr(catalogName, ".") = 0 Then
      catalogName = catalogName & ".gct"
   End If
   If UCase(Dir(catalogName)) = UCase(Mid(catalogName, InStrRev(catalogName, "\") + 1)) Then
      Select Case MsgBox("Do you want to replace the current " & catalogName & "?", _
             vbYesNoCancel + vbQuestion, "Saving MAPP Catalog")
      Case vbNo
         GoTo ReEnter                                   '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      Case vbCancel
         GoTo ExitSub                                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End Select
   End If
   mnuSave_Click
   s = Mid(catalogName, InStrRev(catalogName, "\") + 1)
   s = Left(s, InStrRev(s, ".") - 1)
   Caption = s & " - Table of Contents"
   
ExitSub:
   Exit Sub
   
SaveError:
   If Err = 70 Then
      MsgBox Err.Description & ". " & catalogName & " possibly open in some other program.", _
            vbCritical, "Save MAPP Error"
'      CatalogSaveError = True
   ElseIf Err <> 32755 Then                                         'Not an error if just cancelled
      MsgBox Err.Description, vbCritical, "Save MAPP Error"
'      CatalogSaveError = True
   End If
   catalogName = oldCatalogName                                                       'Set back to old MAPP
   On Error GoTo 0
   Resume ExitSub
End Sub
Private Sub mnuExit_Click()
   Hide
End Sub

'///////////////////////////////////////////////////////////////////////////////////// Options Menu
Private Sub mnuOptions_Click()
   
'   frmTOCOptions.show vbModal
   FillTree
   Exit Sub
   
   
   Dim colon As Integer
   
   
   
   colon = InStr(mruMappPath, ":")
'   With frmTOCOptions
'      .drvMAPPs = Left(mruMappPath, colon)
'      .dirMAPPs = mruMappPath
'      If CatalogSubfolders Then
'         .chkSubfolders = vbChecked
'      Else
'         .chkSubfolders = vbUnchecked
'      End If
'      .dlgDialog.FileName = CatalogDataPath & "*.gex"
'      .txtExpression.Tag = mappWindow.dbExpression.name
''      If colorSet = "" Then
''         .cmbColorSet.Tag = "ANY"
''      Else
''         .cmbColorSet.Tag = CatalogColorSet
''      End If
'      .cmbCriterion.Tag = CatalogCriterion
'      .txtPercent = CatalogPercent
'
'      .show vbModal
'
'      If .Tag = "" Then
'         CatalogDirty = False
'      Else
'         CatalogDirty = True
'      End If
'      CatalogMappPath = .dirMAPPs
'      CatalogDataPath = .txtExpression
'      If .chkSubfolders = vbChecked Then
'         CatalogSubfolders = True
'      Else
'         CatalogSubfolders = False
'      End If
'   End With
   FillTree
End Sub

'//////////////////////////////////////////////////////////////////////////////////////// Help Menu
Public Sub mnuHelp_Click()
   Dim IE As Object
   Set IE = CreateObject("InternetExplorer.Application")
   IE.Visible = True
   IE.Navigate appPath & "\CatalogHelp.htm"
'   IE.StatusText = obj.Head
'   AppActivate "Table of Contents Help"
End Sub

'/////////////////////////////////////////////////////////////////////////////// Processing Catalog
Sub FillTree()
   Dim nodX As Node
   
   Set dbCatalogExpression = Nothing
   Set rsCatalogColorSet = Nothing
'   If CatalogExpression = "" Then
'      treContents.Nodes.Clear
'      Exit Sub                                             'No Expression Dataset >>>>>>>>>>>>>>>>>
'   End If
   
   MousePointer = vbHourglass
On Error GoTo ErrorHandler
   Set dbCatalogExpression = OpenDatabase(CatalogExpression, , True)
   If CatalogColorSet <> "" Then
      Set rsCatalogColorSet = dbCatalogExpression("SELECT * FROM ColorSet WHERE ColorSet = '" & CatalogColorSet & "'")
      If rsCatalogColorSet.EOF Then
         Set dbCatalogExpression = Nothing
         Set rsCatalogColorSet = Nothing
         MsgBox "Color Set """ & CatalogColorSet & """ invalid. Creation of the Table of contents aborted. Choose Options to specify either a valid Color Set or ANY Color Set.", vbOKOnly + vbExclamation, "Creating Table Of Contents"
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Else
      Set rsCatalogColorSet = Nothing
   End If
'   If frmTOCOptions.grdGenes.rows > 2 Then '++++++++++++++++++++++++ Set Up SQL For Genes Specified
'      '  Do we want to follow relationships here??????????????????
'      sqlGenes = "SELECT * FROM Objects WHERE "
'      With frmTOCOptions.grdGenes
'         For i = 1 To .rows - 2
'            sqlGenes = sqlGenes & "(ID = '" & .TextMatrix(i, 0) & "' AND systemCode = '" & .TextMatrix(i, 2) & "') OR "
'         Next i
'      End With
'      sqlGenes = Left(sqlGenes, Len(sqlGenes) - 4)                            'Chop off last " OR "
'   Else
'      sqlGenes = ""
'   End If
   With treContents '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ List The MAPPs
      .Nodes.Clear
      Set nodX = .Nodes.Add(, , CatalogMappPath, CatalogMappPath, "ClosedFolder")
      nodX.Expanded = True
      ListFiles CatalogMappPath
      nodX.Sorted = True
   End With
ExitSub:
   MousePointer = vbDefault
   Exit Sub
   
ErrorHandler:
   If Err.number = 3024 Then
      MsgBox "Expression dataset" & vbCrLf & vbCrLf & CatalogExpression & vbCrLf & vbCrLf _
             & "does not exist. Creation of the Table of contents aborted. Choose Options to " _
             & "specify either a valid Expression Dataset or no Expression Dataset.", _
             vbExclamation + vbOKOnly, "Creating Table of Contents"
      Set dbCatalogExpression = Nothing
      Resume ExitSub
   Else
      FatalError "frmtoc:FillTree", Err.Description
   End If
End Sub
Sub ListFiles(path As String)
   Dim nodX As Node, file As String, index As Integer
   
   With treContents
      '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Files First
      file = Dir(path & "\")
      Do Until file = ""
         If Right(file, 5) = ".mapp" Then
            If QualifyMAPP(path & "\" & file) Then
               Set nodX = .Nodes.Add(path, tvwChild, , Mid(file, InStrRev(file, "\") + 1), "MAPP")
            End If
         End If
         file = Dir
      Loop
      If Not nodX Is Nothing Then nodX.Sorted = True
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Directories Next
      file = Dir(path & "\", vbDirectory)
      Do Until file = ""
         index = index + 1                         'Keep track of where we are in current directory
         If file <> "." And file <> ".." Then
            If (GetAttr(path & "\" & file) And vbDirectory) = vbDirectory Then
               Set nodX = .Nodes.Add(path, tvwChild, path & "\" & file, file, "ClosedFolder")
               ListFiles path & "\" & file
'               If Not nodX Is Nothing Then
               nodX.Expanded = True
               nodX.Sorted = True
               file = Dir(path & "\", vbDirectory)     'Return to directory entry where we left off
               For i = 2 To index - 1                  'because calling Dir again in ListFiles will
                  file = Dir                           'lose our place
               Next i
            End If
         End If
         file = Dir
      Loop
   End With
End Sub
Function QualifyMAPP(mapp) As Boolean
   
   Set dbCatalogMapp = OpenDatabase(mapp, , True)
   If sqlGenes <> "" Then '++++++++++++++++++++++++++++++++++++++++++++++++ Look For Specific Genes
      '  Do we want to follow relationships here??????????????????
      Set rsCatalogObject = dbCatalogMapp.OpenRecordset(sqlGenes)
      If rsCatalogObject.EOF Then
         QualifyMAPP = False
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   QualifyMAPP = True
End Function

