VERSION 5.00
Begin VB.Form frmGeneList 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Gene Lists"
   ClientHeight    =   7410
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8715
   Icon            =   "frmGeneList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstAllgenes 
      Height          =   5520
      Left            =   4560
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
   End
   Begin VB.ListBox LstChanged 
      Height          =   5520
      ItemData        =   "frmGeneList.frx":08CA
      Left            =   480
      List            =   "frmGeneList.frx":08CC
      TabIndex        =   0
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "All probes in the Expression Dataset that are linked to this MAPP or GO term."
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Probes meeting the critierion in the Expression Dataset that are linked to this MAPP or GO term."
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lblNodeText 
      BackColor       =   &H00C0FFFF&
      Caption         =   "GO:0001234 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8415
   End
   Begin VB.Menu mnuExportList 
      Caption         =   "Export This List as Text"
      Begin VB.Menu exportCriterion 
         Caption         =   "Export Genes Meeting the Criterion"
      End
      Begin VB.Menu exportAll 
         Caption         =   "Export All Genes in This GO Term"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuMAPPFInderHelp 
         Caption         =   "MAPPFinder Help"
      End
   End
End
Attribute VB_Name = "frmGeneList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ExpressionDataset As String

Private Sub lblannot1_Click()

End Sub

Private Sub exportAll_Click()
On Error GoTo error
   Dim FileName As String
   Dim Fsys As New FileSystemObject
   Dim output As TextStream
   Dim dbExpression As Database
   Dim rsData As Recordset
   Dim i As Integer, j As Integer
   Dim temp As String
   frmCriteria.CommonDialog1.FileName = ""
   frmCriteria.CommonDialog1.Filter = "Text Files|*.txt"
   frmCriteria.CommonDialog1.ShowSave
   FileName = frmCriteria.CommonDialog1.FileName
   Set dbExpression = OpenDatabase(ExpressionDataset)
        
   If FileName <> "" Then
      If invalidFileName(FileName) Then
         MsgBox "A filename cannot contain any of the following characters: /\:*?" & Chr(34) & "<>| are not", vbOKOnly
         Exit Sub
      End If
      If Dir(FileName) <> "" Then
         If MsgBox("Overwrite the existing " & FileName & "?", vbOKCancel) = vbCancel Then
            Exit Sub
         End If
      End If
      'export list
      Set output = Fsys.CreateTextFile(FileName)
      output.WriteLine (lblNodeText.Caption)
      temp = "ID"
      For i = 3 To dbExpression.TableDefs("Expression").Fields.count - 1
        temp = temp & Chr(9) & dbExpression.TableDefs("Expression").Fields(i).name
      Next i
      output.WriteLine (temp)
      For i = 0 To lstAllgenes.ListCount - 1
         temp = lstAllgenes.List(i)
         Set rsData = dbExpression.OpenRecordset("Select * FROM Expression WHERE ID = '" & temp & "'")
         For j = 3 To dbExpression.TableDefs("Expression").Fields.count - 1
            temp = temp & Chr(9) & rsData.Fields(j)
         Next j
         output.WriteLine (temp)
         temp = ""
      Next i
    End If
error:
    Select Case Err.Number
        Case 70
            MsgBox "Permission Denied. Is the file " & FileName & " open somewhere else? Close it and try again.", vbOKOnly
     End Select
     dbExpression.Close
End Sub

Private Sub exportCriterion_Click()
   On Error GoTo error
   
   Dim FileName As String
   Dim Fsys As New FileSystemObject
   Dim output As TextStream
   Dim dbExpression As Database
   Dim rsData As Recordset
   Dim i As Integer, j As Integer
   Dim temp As String
   frmCriteria.CommonDialog1.FileName = ""
   frmCriteria.CommonDialog1.Filter = "Text Files|*.txt"
   frmCriteria.CommonDialog1.ShowSave
   FileName = frmCriteria.CommonDialog1.FileName
   Set dbExpression = OpenDatabase(ExpressionDataset)
        
   If FileName <> "" Then
      If invalidFileName(FileName) Then
         MsgBox "A filename cannot contain any of the following characters: /\:*?" & Chr(34) & "<>| are not", vbOKOnly
         Exit Sub
      End If
      If Dir(FileName) <> "" Then
         If MsgBox("Overwrite the existing " & FileName & "?", vbOKCancel) = vbCancel Then
            Exit Sub
         End If
      End If
      'export list
      Set output = Fsys.CreateTextFile(FileName)
      output.WriteLine (lblNodeText.Caption)
      temp = "ID"
      For i = 3 To dbExpression.TableDefs("Expression").Fields.count - 1
        temp = temp & Chr(9) & dbExpression.TableDefs("Expression").Fields(i).name
      Next i
      output.WriteLine (temp)
      For i = 0 To LstChanged.ListCount - 1
         temp = LstChanged.List(i)
         Set rsData = dbExpression.OpenRecordset("Select * FROM Expression WHERE ID = '" & temp & "'")
         For j = 3 To dbExpression.TableDefs("Expression").Fields.count - 1
            temp = temp & Chr(9) & rsData.Fields(j)
         Next j
         output.WriteLine (temp)
         temp = ""
      Next i
    End If
error:
    Select Case Err.Number
        Case 70
            MsgBox "Permission Denied. Is the file " & FileName & " open somewhere else? Close it and try again.", vbOKOnly
     End Select
     dbExpression.Close
End Sub

Private Sub mnuExportList_Click_old()
   Dim FileName As String
   Dim Fsys As New FileSystemObject
   Dim output As TextStream
    Dim dbExpression As Database
    Dim rsData As Recordset
    Dim i As Integer, j As Integer
    Dim temp As String
   frmCriteria.CommonDialog1.FileName = ""
   frmCriteria.CommonDialog1.Filter = "Text Files|*.txt"
   frmCriteria.CommonDialog1.ShowSave
   FileName = frmCriteria.CommonDialog1.FileName
   Set dbExpression = OpenDatabase(ExpressionDataset)
   
    '  Set rsFilter = dbExpressionData.OpenRecordset("SELECT OrderNo, ID, SystemCode FROM" _
     '                & " Expression WHERE (" & sql(criterion) & ")")
    Debug.Print dbExpression.TableDefs("Expression").Fields.count
    'For i = 3 To dbExpression.TableDefs("Expression").fields.count - 1
        
        
   If FileName <> "" Then
      If invalidFileName(FileName) Then
         MsgBox "A filename cannot contain any of the following characters: /\:*?" & Chr(34) & "<>| are not", vbOKOnly
         Exit Sub
      End If
      If Dir(FileName) <> "" Then
         If MsgBox("Overwrite the existing " & FileName & "?", vbOKCancel) = vbCancel Then
            Exit Sub
         End If
      End If
      'export list
      Set output = Fsys.CreateTextFile(FileName)
      output.WriteLine (lblNodeText.Caption)
      output.WriteLine ("Probes meeting the criterion")
      For i = 0 To LstChanged.ListCount - 1
         temp = LstChanged.List(i)
         Set rsData = dbExpression.OpenRecordset("Select * FROM Expression WHERE ID = '" & temp & "'")
         For j = 3 To dbExpression.TableDefs("Expression").Fields.count - 1
            temp = temp & "\t" & rsData.Fields(i)
         Next j
         output.WriteLine (LstChanged.List(i))
      Next i
      output.WriteLine ("All Probes linked to this GO term or MAPP")
      For i = 0 To lstAllgenes.ListCount - 1
         output.WriteLine (lstAllgenes.List(i))
      Next i
      output.Close
   End If
End Sub

Private Sub mnuMAPPFInderHelp_Click()
   Dim hWndHelp As Long
   'The return value is the window handle of the created help window.
   hWndHelp = HtmlHelp(hWnd, programpath & "\GenMAPP.chm::/MAPPFinder.htm", HH_DISPLAY_TOPIC, 0)
End Sub


