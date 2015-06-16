VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSystems 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11736
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   9156
   Icon            =   "Systems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11736
   ScaleWidth      =   9156
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msgSystems 
      Height          =   11880
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   7512
      _ExtentX        =   13250
      _ExtentY        =   20955
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   14737632
      GridColor       =   8421504
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblOwner 
      AutoSize        =   -1  'True
      Caption         =   "Gene Database owner:"
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
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   2040
   End
   Begin VB.Label lblModSys 
      AutoSize        =   -1  'True
      Caption         =   "Model Organism table: "
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
      Left            =   60
      TabIndex        =   4
      Top             =   300
      Width           =   2052
   End
   Begin VB.Label lblModify 
      AutoSize        =   -1  'True
      Caption         =   "Last modification date:"
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
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   2016
   End
   Begin VB.Label lblRelease 
      AutoSize        =   -1  'True
      Caption         =   "Release date: "
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
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   1272
   End
   Begin VB.Label lblSpecies 
      AutoSize        =   -1  'True
      Caption         =   "Species:"
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
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   780
   End
End
Attribute VB_Name = "frmSystems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
   If mappWindow.dbGene Is Nothing Then
      MsgBox "There is no active Gene Database", vbOKOnly + vbExclamation, _
             "Gene Database Information"
      Hide
   Else
      DisplaySystems
   End If
End Sub


Private Sub Form_Load()
''   msgSystems.Clear
'   msgSystems.AddItem "Gene Table" & vbTab & "Code" & vbTab & "Species" & vbTab & "Related Systems"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
End Sub
Sub DisplaySystems(Optional sortBy As String = "System")
'   Const MAX_HEIGHT = 8000
   Const CELL_BACK_COLOR = &H80000000
   Dim i As Integer, totalWidth As Single, totalHeight As Single
   Dim widthFudge As Single, heightFudge As Single
   Dim rsSystems As Recordset, rsRelations As Recordset, rsSystem As Recordset, rsInfo As Recordset
   Dim species As String, speciei(100) As String, lastSpecies As Integer, date1 As Date
   Dim noOfSpecies As Integer
   Dim relationCodes As String, relations As String, lastRelation As Integer
   Dim noOfRelations As Integer
   Dim pipe As Integer

   Dim MAX_HEIGHT As Single
   MAX_HEIGHT = 6000
   
   With msgSystems
   widthFudge = Screen.TwipsPerPixelX * 0.95
   heightFudge = Screen.TwipsPerPixelY * 3.5
      '  The CellWidth and CellHeight properties appear to return the interior dimensions of
      '  the cell. Adding this fudge, determined by experiment, seems to give the total
      '  dimension of a cell. The first parameter should adjust for screen resolution;
      '  we can adjust the second parameter.
   Caption = GetFile(mappWindow.dbGene.name)
   Caption = Left(Caption, InStrRev(Caption, ".") - 1) & " Gene Database"
   Set rsInfo = mappWindow.dbGene.OpenRecordset("SELECT * FROM Info")
   lblOwner = "Gene Database owner: " & rsInfo!owner
   lblModSys = "Model Organism table: " & rsInfo!MODSystem
   If Len(rsInfo!species) >= 2 Then
      lblSpecies = "Species: " & Mid(rsInfo!species, 2, Len(rsInfo!species) - 2)
   End If
   lblRelease = "Release date: " & Format(DateSerial(CInt(Left(rsInfo!version, 4)), CInt(Mid(rsInfo!version, 5, 2)), CInt(Right(rsInfo!version, 2))), "d-Mmm-yyyy")
   lblModify = "Last modification date: " & Format(DateSerial(CInt(Left(rsInfo!modify, 4)), CInt(Mid(rsInfo!modify, 5, 2)), CInt(Right(rsInfo!modify, 2))), "d-Mmm-yyyy")
   .Clear
   .BackColor = vbGray
   .row = 0
   .col = 0
   .CellFontBold = True
   .CellBackColor = CELL_BACK_COLOR
   .ColWidth(.col) = 3000 '2200
   totalWidth = totalWidth + .ColWidth(.col) + widthFudge
   .text = "Gene Table"
   .col = .col + 1
   .CellFontBold = True
   .CellBackColor = CELL_BACK_COLOR
   .ColWidth(.col) = 800
   totalWidth = totalWidth + .ColWidth(.col) + widthFudge
   .text = "Code"
   .col = .col + 1
   .CellFontBold = True
   .CellBackColor = CELL_BACK_COLOR
   .ColWidth(.col) = 3000 '2400
   totalWidth = totalWidth + .ColWidth(.col) + widthFudge
   .text = "Related Systems"
   .Width = totalWidth
   totalHeight = totalHeight + .CellHeight + heightFudge
   Set rsSystems = mappWindow.dbGene.OpenRecordset( _
                   "SELECT * FROM Systems WHERE [Date] IS NOT NULL ORDER BY System")
   Do Until rsSystems.EOF
      relations = vbCrLf
      noOfRelations = 0
      Set rsRelations = mappWindow.dbGene.OpenRecordset( _
                   "SELECT * FROM Relations WHERE SystemCode = '" & rsSystems!systemCode & "'")
      Do Until rsRelations.EOF
         Set rsSystem = mappWindow.dbGene.OpenRecordset( _
                   "SELECT * FROM Systems WHERE SystemCode = '" & rsRelations!relatedCode & "'")
         relations = relations & rsSystem!system & vbCrLf
         noOfRelations = noOfRelations + 1
         rsRelations.MoveNext
      Loop
      Set rsRelations = mappWindow.dbGene.OpenRecordset( _
                   "SELECT * FROM Relations WHERE RelatedCode = '" & rsSystems!systemCode & "'")
      Do Until rsRelations.EOF
         Set rsSystem = mappWindow.dbGene.OpenRecordset( _
                   "SELECT * FROM Systems WHERE SystemCode = '" & rsRelations!systemCode & "'")
         If InStr(relations, vbCrLf & rsSystem!system & vbCrLf) = 0 Then
            relations = relations & rsSystem!system & vbCrLf
            noOfRelations = noOfRelations + 1
         End If
         rsRelations.MoveNext
      Loop
      relations = Mid(relations, 3)
      If Len(relations) Then
         relations = Left(relations, Len(relations) - 2)
      End If
      
'      noOfSpecies = 0
'      pipe = InStr(2, rsSystems!species, "|")
'      Do Until pipe = 0
'         noOfSpecies = noOfSpecies + 1
'         pipe = InStr(pipe + 1, rsSystems!species, "|")
'      Loop
'      species = SeparatePipes(rsSystems!species, vbCrLf)
      .row = .row + 1
      .AddItem rsSystems!system & vbTab & rsSystems!systemCode & vbTab & relations, .row
      .RowHeight(.row) = .CellHeight * Max(Max(noOfSpecies, noOfRelations), 1)
      totalHeight = totalHeight + .RowHeight(.row) '+ heightFudge
      rsSystems.MoveNext
   Loop
   .RemoveItem .row + 1                       'AddItem adds a blank row at the end. This removes it
   If totalHeight > MAX_HEIGHT Then
      .Height = MAX_HEIGHT
      .Width = .Width + 110                                                    'Clear for scrollbar
   Else
      .Height = totalHeight
   End If
   Width = .Width + (Width - ScaleWidth) / 1
   .Top = lblModify.Top + lblModify.Height
   Height = .Top + .Height + (Height - ScaleHeight)
      '  (Height - ScaleHeight) give the difference between the exterior and interior height of
      '  the form. totalHeight is interior but we must set the exterior Height property.
   End With
End Sub

