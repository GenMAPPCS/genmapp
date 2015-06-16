VERSION 5.00
Begin VB.Form frmRank 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Results Ranked by Z Score and P value"
   ClientHeight    =   10980
   ClientLeft      =   1515
   ClientTop       =   450
   ClientWidth     =   12225
   Icon            =   "frmRank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   12225
   Begin VB.ListBox lstGO 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   11895
   End
   Begin VB.ListBox lstLocal 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      ItemData        =   "frmRank.frx":08CA
      Left            =   120
      List            =   "frmRank.frx":08CC
      TabIndex        =   0
      Top             =   360
      Width           =   11895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmRank.frx":08CE
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   10080
      Width           =   11415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clicking on a specific term will locate that term in the hierarchy."
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   9720
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gene Ontology Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Local Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbMAPPfinder As Database

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Me.Hide
   Cancel = 1
End Sub

Private Sub lstGO_Click()
   Dim space As Integer
   Dim name As String, id As String
   Dim rsGO As DAO.Recordset
   
   Dim currentnode As Node
   space = InStr(1, lstGO.Text, "    ")
   name = Left(lstGO.Text, space - 1)
   'name = fixName(name)
   Set rsGO = dbMAPPfinder.OpenRecordset("SELECT ID FROM GeneOntology WHERE NAME = '" _
               & name & "'")
   
   id = "GO:" & rsGO![id]
   Set currentnode = TreeForm.TView.Nodes.Item(id)
   currentnode.EnsureVisible
   currentnode.Selected = True
   
   TreeForm.Show
   
   
End Sub

Private Sub lstLocal_Click()
   Dim space As Integer
   Dim name As String
   
   Dim currentnode As Node
   space = InStr(1, lstLocal.Text, "    ")
   name = Left(lstLocal.Text, space - 1)
         
   Set currentnode = TreeForm.TView.Nodes.Item(name)
   currentnode.EnsureVisible
   currentnode.Selected = True
   
   TreeForm.Show
   
End Sub

Public Sub setDB(MAPPDB As Database)
   Set dbMAPPfinder = MAPPDB
End Sub


Public Function fixName(oldName As String) As String
   Dim comma As Integer
   Dim i As Integer
   Dim length As Integer
   Dim leftside As String
   Dim rightside As String
   i = 1
   length = Len(oldName)
   
   While i <= length
      comma = InStr(i, oldName, ",")
      If comma > 0 Then  'there's a string
         leftside = Left(oldName, comma - 1)
         rightside = Mid(oldName, comma, length - comma + 1)
         oldName = leftside & "\" & rightside
         i = comma + 2
         length = length + 1
      Else
         i = i + 1
      End If
   Wend
   fixName = oldName
End Function
