VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmGeneFinder 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Gene Finder"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GeneFinder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CheckBox chkDontShowWarning 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Don't show me this warning again."
      Height          =   252
      Left            =   780
      TabIndex        =   25
      Top             =   3660
      Width           =   3612
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   312
      Left            =   2280
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      Enabled         =   0   'False
      Height          =   312
      Left            =   1200
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   312
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   972
   End
   Begin SHDocVwCtl.WebBrowser brsGeneData 
      Height          =   3492
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   8172
      ExtentX         =   14414
      ExtentY         =   6159
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   6300
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5400
      Width           =   972
   End
   Begin VB.TextBox txtRemarks 
      Height          =   1080
      Left            =   4020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      ToolTipText     =   "Appears on the gene backpage."
      Top             =   720
      Width           =   4272
   End
   Begin VB.TextBox txtHead 
      Height          =   300
      Left            =   4860
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Appears at the top of the gene backpage."
      Top             =   240
      Width           =   3432
   End
   Begin VB.TextBox txtGeneLabel 
      Height          =   300
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   "Appears in the gene box on the MAPP graphic."
      Top             =   1440
      Width           =   2712
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gene Identification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3672
      Begin VB.ComboBox cmbOtherIDs 
         Height          =   336
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   2112
      End
      Begin VB.ComboBox cmbSystems 
         Height          =   336
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   2112
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   312
         Left            =   60
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Searches Gene Database for the gene identification."
         Top             =   780
         Width           =   1212
      End
      Begin VB.TextBox txtGeneID 
         Height          =   312
         Left            =   60
         MaxLength       =   20
         TabIndex        =   0
         Top             =   420
         Width           =   1392
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Gene IDs"
         Height          =   240
         Index           =   6
         Left            =   1500
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   1392
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gene ID System"
         Height          =   240
         Index           =   5
         Left            =   1500
         TabIndex        =   19
         Top             =   180
         Width           =   1452
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gene ID"
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   18
         Top             =   180
         Width           =   732
      End
      Begin VB.Label lblFound 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Gene Found"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   60
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   372
      Left            =   7320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5400
      Width           =   972
   End
   Begin VB.Label lblGB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SwissProt, Locus Link, and Unigene)"
      Height          =   240
      Index           =   4
      Left            =   780
      TabIndex        =   24
      Top             =   3300
      Width           =   3252
   End
   Begin VB.Label lblGB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "that we will be able to support in the future (e.g. Model Organism Databases, "
      Height          =   240
      Index           =   3
      Left            =   780
      TabIndex        =   23
      Top             =   3060
      Width           =   6744
   End
   Begin VB.Label lblGB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "in the future for this organism.  If possible, please use other gene ID systems"
      Height          =   240
      Index           =   2
      Left            =   780
      TabIndex        =   22
      Top             =   2820
      Width           =   6720
   End
   Begin VB.Label lblGB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: GenBank is growing so quickly that we will not be able to support it "
      Height          =   240
      Index           =   1
      Left            =   780
      TabIndex        =   21
      Top             =   2580
      Width           =   6396
   End
   Begin VB.Label lblGB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GenBank is not recommended!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   780
      TabIndex        =   20
      Top             =   2280
      Width           =   3744
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Heading"
      Height          =   240
      Index           =   4
      Left            =   4020
      TabIndex        =   11
      Top             =   240
      Width           =   744
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   240
      Index           =   2
      Left            =   4020
      TabIndex        =   10
      Top             =   480
      Width           =   804
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Backpage:"
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
      Index           =   1
      Left            =   3900
      TabIndex        =   9
      Top             =   0
      Width           =   1068
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gene label"
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   1440
      Width           =   948
   End
End
Attribute VB_Name = "frmGeneFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public obj As objLump                                              'Gene object from calling window
'  obj.container is the form on which the object resides
'  obj.container.dbGene is database
Dim dirty As Boolean        'Change made in gene on this form. Should make calling frmDrafter dirty
Dim systems(MAX_SYSTEMS, 2) As String                           'Systems supported by Gene Database
   '  Systems(x, 0)  Name of cataloging system
   '  Systems(x, 1)  System code
   '  Systems(x, 2)  MOD species
Dim lastSystem As Integer
Dim unsupportedSystems(MAX_SYSTEMS, 1) As String            'Systems not supported by Gene Database
   '  unsupportedSystems(x, 0)  Name of cataloging system
   '  unsupportedSystems(x, 1)  System code
Dim lastUnsupportedSystem As Integer
Dim forwardOnce As Boolean                    'Workaround for wierd behavior of CommandStateChanged

Private Sub brsGeneData_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
'Debug.Print "Command: " & Command & "  Location: " & brsGeneData.LocationName & "  Enable: " & Enable
'   If brsGeneData.ReadyState <> READYSTATE_COMPLETE Then Exit Sub
      If brsGeneData.LocationName = "GeneFinder.htm" Then
         cmdBack.Enabled = False
      Else
         cmdBack.Enabled = True
      End If
'If Not Enable Then Exit Sub
   Select Case Command
   Case CSC_NAVIGATEFORWARD
      cmdForward.Enabled = Enable
'      If brsGeneData.LocationName = "GeneFinder.htm" Then
'         cmdForward.Enabled = False
'      Else
'         cmdForward.Enabled = Enable
'      End If
'      forwardOnce = True                                      'Have navigated forward at least once
   Case CSC_NAVIGATEBACK
'cmdBack.Enabled = Enable
'      If brsGeneData.LocationName = "GeneFinder.htm" Then
'         cmdBack.Enabled = False
'      Else
'         cmdBack.Enabled = Enable
'      End If
   End Select
   '  These states change differently with different configurations. Some configurations
   '  end with a forward state change after navigating back for some reason. The following
   '  If structure works around it.
'   If brsGeneData.LocationName <> "GeneFinder.htm" Then
'      '  If the location is not the home page, then there must be something to go back to.
'      cmdBack.Enabled = True
'   ElseIf forwardOnce Then
'      '  If the location is the home page but we have navigated forward at least once, we must
'      '  be able to navigate forward again.
'      cmdForward.Enabled = True
'   End If
End Sub

Private Sub cmbSystems_Click()
   If cmbSystems.text = "GenBank" Then
      ShowGenBankWarning
   Else
      ShowGenBankWarning False
   End If
   geneFound False
   mruGeneFinderSystem = cmbSystems.text
End Sub
Sub ShowGenBankWarning(Optional showIt As Boolean = True)
   Dim i As Integer
   Static dontShow As Boolean                       'Warning shown only once during GenMAPP session
   
   If chkDontShowWarning = vbChecked Then
      dontShow = True
   End If
   If dontShow Then showIt = False
   If showIt Then
      brsGeneData.visible = False
   End If
   chkDontShowWarning.visible = showIt
   For i = 0 To 4
      lblGB(i).visible = showIt
   Next i
   DoEvents
End Sub
Private Sub cmdBack_Click()
   brsGeneData.GoBack
End Sub

Private Sub cmdForward_Click()
   brsGeneData.GoForward
End Sub

Private Sub cmdHome_Click()
   brsGeneData.Navigate appPath & "Backpages\GeneFinder.htm"
   DoEvents
   cmdBack.Enabled = False
   cmdForward.Enabled = False
End Sub

Private Sub Form_Load()
'   Dim rsGenMAPP As Recordset
'   Dim GenBank(MAX_GENES_PER_SET) As String, blast(MAX_GENES_PER_SET) As String
'   Dim SwissName As String, protein As String, species As String, functions As String
'   Dim primGenBank As String
'   Dim SwissNo(MAX_GENES_PER_SET) As String
   Dim sql As String, i As Integer
   
   If obj.canvas.container.dbGene Is Nothing Then
      MsgBox "Must choose a Gene Database first.", vbCritical + vbOKOnly, "Opening Gene Finder"
      Hide
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   FillSystemsList                                                            'Fill Up Systems List
   
   txtGeneID = ""
   txtRemarks = ""
   txtHead = ""
   txtGeneLabel = ""
   lblFound.visible = False
   brsGeneData.visible = False
   cmdForward.visible = False
   cmdForward.Enabled = False
   cmdBack.visible = False
   cmdHome.visible = False
   DoEvents
   
   If obj.title <> "Gene" Then txtGeneLabel = obj.title
   txtHead = obj.head
   txtRemarks = obj.remarks
End Sub
Private Sub Form_Activate()
   MousePointer = vbHourglass
   brsGeneData.visible = False
   ShowGenBankWarning False
   DoEvents
   If obj.id <> "" Then '+++++++++++++++++++++++++++++++++++++++++++++++ Show Data For Current Gene
      '  A new gene will not have an obj.ID
      txtGeneID = obj.id
      
'      For i = 0 To cmbSystems.ListCount - 1                                    'Unselect everything
''         cmbSystems.Selected(i) = False
'      Next i
      For i = 0 To lastSystem                                             'Find match to systemCode
         If obj.systemCode = systems(i, 1) Then
            cmbSystems.ListIndex = i
            Exit For                                       'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      Next i
      If i > lastSystem Then '=====================================================System Unmatched
         MsgBox "Gene ID system (code " & obj.systemCode & ") not supported in your " _
                & "Gene Database.", vbExclamation + vbOKOnly, "GeneFinder Search"
      Else '=============================================================Valid System, Process Gene
         If CreateBackpage(obj.id, obj.systemCode, obj.head, obj.canvas.container.dbGene, _
                           Nothing, , appPath & "Backpages\GeneFinder.htm", PURPOSE_FINDER) _
               <> "" Then
            brsGeneData.Navigate appPath & "Backpages\GeneFinder.htm"
            brsGeneData.visible = True
            cmdForward.visible = True
            cmdForward.Enabled = False
            cmdBack.visible = True
            cmdBack.Enabled = False
            cmdHome.visible = True
            geneFound
         Else
            MsgBox "Gene [" & obj.title & "] not in your Gene Database.", _
                   vbExclamation + vbOKOnly, "GeneFinder Search"
            txtGeneID.SetFocus
            geneFound False
            cmdSearch.visible = False
         End If
      End If
   Else
      '===========================================Set Gene System List To Most Recently Used System
      For i = 0 To cmbSystems.ListCount - 1
         If cmbSystems.List(i) = mruGeneFinderSystem Then
            cmbSystems.ListIndex = i
            Exit For
         End If
      Next i
      
      txtGeneID.SetFocus
      geneFound False
   End If
   dirty = False
   ShowGenBankWarning False
   forwardOnce = False              'Workaround for CommandStateChanged. Have not navigated forward
   cmdForward.Enabled = False       'with new WebBrowser control instance
   MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
   dirty = False
   Hide
End Sub

Private Sub cmdOK_Click()
   Dim i As Integer
   Dim primaryColumn As String
   Dim rsSystems As Recordset                                                        'Systems table
   'For AllRelatedGenes()
      Dim genes As Integer
      Dim geneIDs(MAX_GENES, 2) As String
      Dim geneFound As Boolean
      'Dim supportedSystem as Boolean                 'System supported in Gene Database [optional]
   
   If InvalidChr(txtGeneLabel, "gene label") Then
      txtGeneLabel.SetFocus
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   If InvalidChr(txtHead, "backpage heading") Then
      txtHead.SetFocus
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   If InvalidChr(txtRemarks, "remarks") Then
      txtRemarks.SetFocus
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
'   If cmdSearch.Visible Then
'      cmdSearch_Click
'      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   End If
   
   If dirty Then
      txtGeneID = Dat(txtGeneID)
      If txtGeneID <> "" And Not lblFound.visible Then '+++++++++++++++++++++++++++++ Identify Gene
         If cmbSystems.text = "" Then                                   'No Gene ID system selected
            MsgBox "You must select a Gene ID system from the list.", _
                   vbExclamation + vbOKOnly, "Saving Gene in GeneFinder"
            GoTo ExitSub                                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
         
         If cmdSearch.visible Then
            cmdSearch_Click
            GoTo ExitSub
         End If
         
'         AllRelatedGenes txtGeneID, systems(cmbSystems.ListIndex, 1), _
'                         obj.canvas.container.dbGene, genes, geneIDs, geneFound ',supportedSystem
'         If Not geneFound Then '===========================================No Gene In Gene Database
'            If MsgBox("Gene [" & txtGeneID & "] not in Gene Table """ _
'                      & systems(cmbSystems.ListIndex, 0) & """. Add to the " _
'                               & """Other"" Gene Table?", vbExclamation + vbYesNo, _
'                               "Saving Gene in GeneFinder") = vbYes Then
'               mappWindow.dbGene.Execute _
'                         "INSERT INTO Other (ID, SystemCode)" & _
'                         "   VALUES ('" & txtGeneID & "', '" & systems(cmbSystems.ListIndex, 1) _
'                             & "')"
'            Else
'               GoTo ExitSub                                'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'            End If
'         End If
         
      End If
      
      obj.id = txtGeneID '========================================== Gene Valid, Fill In Properties
      If cmbSystems.ListIndex <> -1 Then
         '  A user could conceivably type in just a label and leave the gene unidentified.
         '  The section above would not search for a gene and thes would not assign a gene ID type.
         obj.systemCode = systems(cmbSystems.ListIndex, 1)
      End If
      txtGeneLabel = TextToSql(Dat(txtGeneLabel))
      If txtGeneLabel = "" Then
         obj.title = txtGeneID
      Else
         obj.title = txtGeneLabel
      End If
      obj.head = TextToSql(Dat(txtHead))
      obj.remarks = TextToSql(Dat(txtRemarks))
      obj.DrawObj True, drawingBoard
      mappWindow.dirty = True
   End If
   cmdCancel_Click
ExitSub:
End Sub

Private Sub cmdSearch_Click()
   MousePointer = vbHourglass
   
   brsGeneData.visible = False
   lblFound.visible = False
   
   If InvalidChr(txtGeneID, "gene identification", """$,") Then
      txtGeneID.SetFocus
      GoTo ExitSub               'Invalid character in gene ID 'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   If cmbSystems.text = "" Then                                         'No Gene ID system selected
      MsgBox "You must select a Gene ID system from the list.", _
             vbExclamation + vbOKOnly, "GeneFinder Search"
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   txtGeneID = TextToSql(UCase(Dat(txtGeneID)))
   If txtGeneID <> "" Then '+++++++++++++++++++++++++++++++++ Adjust Gene IDs For Different Systems
      Select Case systems(cmbSystems.ListIndex, 0)
      Case "UniGene"
         s = txtGeneID
         If Len(s) >= 2 Then
            Mid(s, 2, 1) = LCase(Mid(s, 2, 1))                    '2nd chr always L/C. Eg: Hs.12345
         End If
         txtGeneID = s
      Case "FlyBase"
         If Left(txtGeneID, 4) <> "FBGN" Then
            txtGeneID = "FBGN" & txtGeneID
         End If
         s = txtGeneID
         Mid(s, 3, 2) = "gn"                                                         'Always "FBgn"
         txtGeneID = s
      Case "InterPro"
         If Left(txtGeneID, 3) <> "IPR" Then
            txtGeneID = "IPR" & txtGeneID
         End If
      Case "MGI"
         If Left(txtGeneID, 4) <> "MGI:" Then
            txtGeneID = "MGI:" & txtGeneID
         End If
      Case "RGD"
         If Left(txtGeneID, 4) <> "RGD:" Then
            txtGeneID = "RGD:" & txtGeneID
         End If
      Case "WormBase"
'         If Left(txtGeneID, 2) <> "CE" Then      'There is a secondary search column, therefore we
'            txtGeneID = "CE" & txtGeneID         'cannot assume a CE prefix
'         End If
      End Select
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Gene In Gene Database
   ShowGenBankWarning False
   If CreateBackpage(txtGeneID, systems(cmbSystems.ListIndex, 1), txtHead, _
                     obj.canvas.container.dbGene, Nothing, , _
                     appPath & "Backpages\GeneFinder.htm", PURPOSE_FINDER) <> "" Then '==Gene Found
      brsGeneData.Navigate appPath & "Backpages\GeneFinder.htm"
      brsGeneData.visible = True
      cmdForward.visible = True
      cmdForward.Enabled = False
      cmdBack.visible = True
      cmdBack.Enabled = False
      cmdHome.visible = True
      geneFound
      If txtGeneLabel = "" Then txtGeneLabel = txtGeneID
   Else '============================================================================Gene Not Found
      brsGeneData.visible = False
      Select Case MsgBox("Gene """ & txtGeneID & """ not found in your " _
                         & systems(cmbSystems.ListIndex, 0) & " data. Add to the " _
                         & """Other"" Gene category? (Click ""Cancel"" to change the " _
                         & "Gene ID and Search again. Click ""No"" to apply the ID " _
                         & "and Gene ID type to your gene without adding it to your " _
                         & "Gene Database.)", vbExclamation + vbYesNoCancel, "GeneFinder Search")
      Case vbYes
         mappWindow.dbGene.Execute _
                   "INSERT INTO Other (ID, SystemCode)" & _
                   "   VALUES ('" & txtGeneID & "', '" & systems(cmbSystems.ListIndex, 1) & "')"
         geneFound True
         dirty = True
      Case vbNo
         cmdSearch.visible = False
      Case Else
         txtGeneID.SetFocus
         geneFound False
'         dirty = True
      End Select
'      If systems(cmbSystems.ListIndex, 0) = "Other" Then '---------------------Other Table Searched
'         Select Case MsgBox("Gene """ & txtGeneID & """ not in ""Other"" Table. Add to the " _
'                            & """Other"" Gene Table?", _
'                            vbExclamation + vbYesNo, "GeneFinder Search")
'         Case vbYes
'            mappWindow.dbGene.Execute _
'                      "INSERT INTO Other (ID, SystemCode)" & _
'                      "   VALUES ('" & txtGeneID & "', '" & systems(cmbSystems.ListIndex, 1) & "')"
'            geneFound True
'            dirty = True
'         Case Else
'            txtGeneID.SetFocus
'            geneFound False
'            dirty = True
'         End Select
'      Else '---------------------------------------------------------------Other Table Not Searched
'         '  Do not allow a gene to be added to Other without searching Other first
'         MsgBox "No match found in the Gene Database for your """ _
'                 & systems(cmbSystems.ListIndex, 0) & """ ID. Either reenter Gene ID, " _
'                 & "download an updated Gene Database from www.GenMAPP.org, or select " _
'                 & """Other"" to search (or add to) that Gene Table.", _
'                 vbExclamation + vbOKOnly, "GeneFinder Search"
'      End If
   End If
ExitSub:
   MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If dirty Then
      Select Case MsgBox("Save changes?", vbYesNoCancel + vbQuestion, "Closing Gene ID Window")
      Case vbYes
         cmdOK_Click
      Case vbNo
         cmdCancel_Click
      Case Else
         Cancel = True                                                                 'Don't close
      End Select
   End If
End Sub

Private Sub Form_Resize()
   Static beenHere As Boolean
   
   If beenHere Then Exit Sub
   If Height < 5000 Then
      beenHere = True
      Height = 5000
   End If
   If Width < 8508 Then
      beenHere = True
      Width = 8508
   End If
   cmdBack.Top = Height - 900
   cmdForward.Top = Height - 900
   cmdHome.Top = Height - 900
   cmdCancel.Top = Height - 900
   cmdOK.Top = Height - 900
   cmdCancel.Left = Width - 2208
   cmdOK.Left = Width - 1188
   brsGeneData.Width = Width - 336
   brsGeneData.Height = Height - 2808
   beenHere = False
End Sub

Private Sub txtGeneLabel_Change()
   dirty = True
End Sub
Private Sub txtGeneLabel_LostFocus()
   If txtHead = "" Then txtHead = txtGeneLabel
End Sub
Private Sub txtHead_Change()
   dirty = True
End Sub
Private Sub txtRemarks_Change()
   dirty = True
End Sub
Private Sub txtGeneID_Change()
   dirty = True
   geneFound False
End Sub
Sub FillSystemsList()
   Dim rsSystems As Recordset, rsInfo As Recordset
   Dim index As Integer, i As Integer, modSpecies As String
   Dim temp(2) As String
      '  temp(0)  Name of cataloging system
      '  temp(1)  System code
      '  temp(2)  MOD species
   Dim mruIndex As Integer
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Systems
   lastSystem = -1                                                               'Zero-based arrays
   lastUnsupportedSystem = -1
   Set rsSystems = obj.canvas.container.dbGene.OpenRecordset( _
         "SELECT * FROM Systems " & _
         "   ORDER BY System = 'SwissProt', System = 'LocusLink', System = 'UniGene', System", dbOpenForwardOnly)
   Do Until rsSystems.EOF
      If (VarType(rsSystems!Date) = vbNull And rsSystems!system <> "Other") _
            Or rsSystems!system = "GeneOntology" _
            Or rsSystems!system = "InterPro" _
            Or InStr(rsSystems!Misc, "|I|") Then                                  'Improper system
            '  Other system is always supported, date or not. GeneOntology and InterPro are never
            '  supported because we do not want users to identify their genes by these systems
            '  (per Kam)
         lastUnsupportedSystem = lastUnsupportedSystem + 1
         unsupportedSystems(lastUnsupportedSystem, 0) = rsSystems!system
         unsupportedSystems(lastUnsupportedSystem, 1) = rsSystems!systemCode
      Else
         lastSystem = lastSystem + 1
         systems(lastSystem, 0) = rsSystems!system
         systems(lastSystem, 1) = rsSystems!systemCode
         systems(lastSystem, 2) = Dat(rsSystems!MOD)
      End If
      rsSystems.MoveNext
   Loop
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Order Systems
   Set rsInfo = obj.canvas.container.dbGene.OpenRecordset("SELECT MODSystem FROM Info")   'Find MOD
   For index = 0 To lastSystem '======================================================Put MOD First
      If systems(index, 0) = rsInfo!MODSystem Then
         temp(0) = systems(index, 0)
         temp(1) = systems(index, 1)
         temp(2) = systems(index, 2)
         For i = index To 1 Step -1
            systems(i, 0) = systems(i - 1, 0)
            systems(i, 1) = systems(i - 1, 1)
            systems(i, 2) = systems(i - 1, 2)
         Next i
         systems(0, 0) = temp(0)
         systems(0, 1) = temp(1)
         systems(0, 2) = temp(2)
         Exit For
      End If
   Next index
   For index = 0 To lastSystem '===================================================Put GenBank Last
      If systems(index, 0) = "GenBank" Then
         temp(0) = systems(index, 0)
         temp(1) = systems(index, 1)
         temp(2) = systems(index, 2)
         For i = index To lastSystem - 1
            systems(i, 0) = systems(i + 1, 0)
            systems(i, 1) = systems(i + 1, 1)
            systems(i, 2) = systems(i + 1, 2)
         Next i
         systems(lastSystem, 0) = temp(0)
         systems(lastSystem, 1) = temp(1)
         systems(lastSystem, 2) = temp(2)
         Exit For
      End If
   Next index
            
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Fill List
   cmbSystems.Clear
   mruIndex = -1
   For index = 0 To lastSystem
      cmbSystems.AddItem systems(index, 0), index
      If systems(index, 0) = mruGeneFinderSystem Then
         mruIndex = index
      End If
   Next index
   
   If mruIndex <> -1 Then '+++++++++++++++++++++++++++++++++++++++ Select Most Recently Used System
      cmbSystems.text = systems(mruIndex, 0)
   End If
   mruGeneFinderSystem = cmbSystems.text
   
'   cmbSystems.AddItem "Gene ID Type"
'   cmbSystems.ListIndex = cmbSystems.ListCount - 1
'   cmbSystems.RemoveItem cmbSystems.ListCount - 1
   geneFound False
End Sub

Sub geneFound(Optional found As Boolean = True)
   If found Then
      lblFound.visible = True
      cmdSearch.visible = False
      cmdOK.Default = True
      cmbSystems.Enabled = False
   Else
      lblFound.visible = False
      cmdSearch.visible = True
      cmdSearch.Default = True
      cmbSystems.Enabled = True
   End If
End Sub


