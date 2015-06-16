VERSION 5.00
Begin VB.Form frmObjects 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFF0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Object Toolbox"
   ClientHeight    =   8580
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   7830
   Icon            =   "Objects.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   8580
   ScaleWidth      =   7830
   Begin VB.Label lblTitle 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9996
   End
End
Attribute VB_Name = "frmObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLines As New Collection                                                 'Local to frmObjects
Dim objLumps As New Collection

Private Sub Form_Load()
'   show                                             'Because object drawing methods use active form
   PlaceObject 600, 850, "ProteinB"
   PlaceObject 600, 1750, "Ribosome"
   PlaceObject 1500, 1200, "OrganA"
   PlaceObject 2000, 1200, "OrganB"
   PlaceObject 2600, 700, "OrganC"
   PlaceObject 2350, 1400, "CellB"
   PlaceObject 2800, 1600, "CellA"
   PlaceObject 3300, 600, "Poly", "3"
   PlaceObject 3300, 1000, "Poly", "5"
   PlaceObject 3300, 1400, "Poly", "6"
   Width = 3700
   Height = 2700
   mappWindow.dirty = False                                          'Creates will set this to true
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim element As Object
   
   DeselectAll
   frmDrafter.ToolBarClear ""                                                  'Unclick all buttons
   mappWindow.SetActiveObject Nothing
   Set mappWindow.newObject = Nothing
   For Each element In objLumps '++++++++++++++++++++++++++++++++++++++++++++++++++ Check All Lumps
      If element.CheckClick(X, Y) Then
         element.SelectMode = True
         With drawingBoard                                 'picDrafter. FrmDrafter is the container
            If .container.lineStarted Then     'A line was previously started and user changed mind
               .foreColor = vbWhite                                                'Delete X marker
               drawingBoard.Line (.container.XStart - 50, .container.YStart)-Step(100, 0)
               drawingBoard.Line (.container.XStart, .container.YStart - 50)-Step(0, 100)
               .foreColor = vbBlack
               .container.lineStarted = False
            End If
            .MousePointer = vbCrosshair
            .container.MultipleObjectDeselectAll
            .container.SetActiveObject Nothing                        'Any active object turned off
            .container.sbrBar.Panels("Instructions").text = .container.statusBarPlace(element.objType)
            If element.objType = "Poly" Then
               .container.tlbTools.Tag = element.objType & element.sides
            Else
               .container.tlbTools.Tag = element.objType
            End If
         End With
         Exit Sub                                   '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Next element
   drawingBoard.MousePointer = vbDefault                                'Blank part of form clicked
   frmDrafter.tlbTools.Tag = ""
End Sub
Private Sub PlaceObject(X As Single, Y As Single, objType As String, _
                        Optional options As String = "")
   Dim newLump As New objLump
   newLump.Create X, Y, objType, , , , , options, , frmObjects
   If newLump.objType = "Oval" Then
      '  CellB is a specialized oval. It is drawn as an oval with a black fill and specific size.
      '  Once it is on the drafting board it is manipulated like any other oval, but here on the
      '  Object Toolbox it must be identified as a CellB so that it will be dropped correctly on
      '  the Drafting Board.
      newLump.objType = "CellB"
   End If
   objLumps.Add newLump, newLump.objType & newLump.sides
   If newLump.objType = "Whatever" Then
      show
      HitRange newLump
   End If
End Sub

Public Sub DeselectAll()
   Dim element As Object, frm As Form
      
   If frmObjects.visible = False Then Exit Sub
   
   Set frm = Screen.ActiveForm
   frmObjects.SetFocus
   For Each element In objLumps '++++++++++++++++++++++++++++++++++++++++++++++++++ Check All Lumps
      element.SelectMode = False
   Next element
   frm.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Hide
   Cancel = -1
End Sub
