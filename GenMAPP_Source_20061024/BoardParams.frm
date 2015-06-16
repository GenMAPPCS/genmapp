VERSION 5.00
Begin VB.Form frmBoardParams 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drafting Board Size"
   ClientHeight    =   2640
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   3612
   Icon            =   "BoardParams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3612
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAccommodate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Size to fit objects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Check to make the drafting board just large enough to contain the current objects on the MAPP graphic"
      Top             =   1920
      Width           =   3312
   End
   Begin VB.TextBox txtHeight 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   300
      TabIndex        =   0
      ToolTipText     =   "Height of the drafting board"
      Top             =   660
      Width           =   612
   End
   Begin VB.TextBox txtWidth 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1500
      TabIndex        =   1
      ToolTipText     =   "Width of the drafting board"
      Top             =   1440
      Width           =   612
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   2520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1032
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   1440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1032
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Centimeters"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1092
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1452
      Left            =   600
      Top             =   120
      Width           =   2292
   End
End
Attribute VB_Name = "frmBoardParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public canvas As Control

Private Sub chkAccommodate_Click() '*************************** Sizes Board To Just Fit All Objects
   Dim element As Object
   
   If chkAccommodate <> vbChecked Then Exit Sub            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   txtWidth = Format((mappWindow.MinBoardWidth) / TWIPS_CM, "0.0")
   txtHeight = Format((mappWindow.MinBoardHeight) / TWIPS_CM, "0.0")
End Sub

Private Sub cmdCancel_Click()
   Tag = ""
   Hide
End Sub

Private Sub cmdOK_Click()
   If Not IsNumeric(Dat(txtHeight)) Then
      MsgBox "Height not a valid numeric value."
      txtHeight.SetFocus
      txtHeight.SelStart = 0
      txtHeight.SelLength = 100
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   If Not IsNumeric(Dat(txtWidth)) Then
      MsgBox "Width not a valid numeric value."
      txtWidth.SetFocus
      txtWidth.SelStart = 0
      txtWidth.SelLength = 100
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Width And Height To Grid
   'Why should these be set to grid???????????????
   'If they are, then 57 set to GridMin will be slightly larger than 57 and warning message pops up
'   boardWidth = GridMin(frmBoardParams!txtWidth * TWIPS_CM)
'   boardHeight = GridMin(frmBoardParams!txtHeight * TWIPS_CM)
'   drawingBoard.Width = frmBoardParams!txtWidth * TWIPS_CM
'   drawingBoard.Height = frmBoardParams!txtHeight * TWIPS_CM
   
   
   With mappWindow '++++++++++++++++++++++++++++++++++++++++++++++ Put Board Within Its Min And Max
      Dim minWidth As Single, minHeight As Single
      
      .boardWidth = frmBoardParams!txtWidth * TWIPS_CM
      .boardHeight = frmBoardParams!txtHeight * TWIPS_CM
      If .boardWidth > MAX_BOARD_WIDTH Then
         .boardWidth = MAX_BOARD_WIDTH
         txtWidth = Format((MAX_BOARD_WIDTH - 1) / TWIPS_CM, "0.0")
      End If
      If .boardHeight > MAX_BOARD_HEIGHT Then
         .boardHeight = MAX_BOARD_HEIGHT
         txtHeight = Format((MAX_BOARD_HEIGHT - 1) / TWIPS_CM, "0.0")
      End If
      minWidth = .MinBoardWidth
      If .boardWidth < MIN_BOARD_WIDTH Or .boardWidth < minWidth Then
         If MIN_BOARD_WIDTH > minWidth Then                                   'Smaller than minimum
            .boardWidth = MIN_BOARD_WIDTH
            txtWidth = Format((MIN_BOARD_WIDTH - 1) / TWIPS_CM, "0.0")
         Else                                                            'Can't accommodate objects
            .boardWidth = minWidth
            txtWidth = Format(drawingBoard.Width / TWIPS_CM, "0.0")
         End If
      End If
      minHeight = .MinBoardHeight
      If .boardHeight < MIN_BOARD_HEIGHT Or .boardHeight < minHeight Then
         If MIN_BOARD_HEIGHT > minHeight Then                                 'Smaller than minimum
            .boardHeight = MIN_BOARD_HEIGHT
            txtHeight = Format((MIN_BOARD_HEIGHT - 1) / TWIPS_CM, "0.0")
         Else                                                            'Can't accommodate objects
            .boardHeight = minHeight
            txtHeight = Format(drawingBoard.Height / TWIPS_CM, "0.0")
         End If
      End If
      
      '=================================================================Reset picDrafter Dimensions
      drawingBoard.Width = .boardWidth * .zoom
      drawingBoard.Height = .boardHeight * .zoom
'      .Width = .boardWidth * .zoom
'      .Height = .boardHeight * .zoom
      
      .FitWindowToBoard
      
'      callingRoutine = "mnuBoardSize_Click"
'         '  Set so that FormWidth() and FormHeight() do not call resize automatically
'      If .ClientWide() > drawingBoard.Width + drawingBoard.Left Then  'Window > drafting board edge
'         '  If the size of the drafting board is reduced so that there is client window beyond the
'         '  right, bottom then the client window is resized.
'         WindowState = vbNormal                                          'Can't resize if maximized
'         .FormWidth drawingBoard.Width + drawingBoard.Left
'      End If
'      If .ClientHigh() > drawingBoard.Height + .tlbTools.Height - drawingBoard.Top Then
'         '  The top of the visible picDrafter viewport is below the tool bar (tlbTools).
'         WindowState = vbNormal                                          'Can't resize if maximized                                                  'In case maximized
'         .FormHeight drawingBoard.Height + drawingBoard.Top - .tlbTools.Height
'      End If
'      callingRoutine = ""
'      .Form_Resize                                               'Now we actually do want to resize
''      mnuRedraw_Click
      .ScrollBars
      .dirty = True
      .MAPPTitle.DrawObj False, drawingBoard           'Title must recenter on changes of MAPP size
      .MAPPTitle.DrawObj True, drawingBoard
   End With
   
   Hide
End Sub

Private Sub Form_Activate()
   txtWidth = Format(mappWindow.boardWidth / TWIPS_CM, "0.0")            'Convert pixels to cm.
   txtHeight = Format(mappWindow.boardHeight / TWIPS_CM, "0.0")
   txtHeight.SetFocus
   chkAccommodate = vbUnchecked
End Sub

Private Sub txtWidth_GotFocus()
   txtWidth.SelStart = 0
   txtWidth.SelLength = 100
End Sub
Private Sub txtHeight_GotFocus()
   txtHeight.SelStart = 0
   txtHeight.SelLength = 100
End Sub

