VERSION 5.00
Begin VB.Form frmMultipleColorSets 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiple Color Sets"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAll 
      Cancel          =   -1  'True
      Caption         =   "&All"
      Height          =   315
      Left            =   1260
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3060
      Width           =   795
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "&None"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3060
      Width           =   795
   End
   Begin VB.ListBox lstDisplayValue 
      Appearance      =   0  'Flat
      Height          =   2565
      ItemData        =   "MultipleColorSets.frx":0000
      Left            =   2220
      List            =   "MultipleColorSets.frx":0007
      TabIndex        =   4
      ToolTipText     =   "Chose single value to display"
      Top             =   420
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   2640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3420
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   3660
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3420
      Width           =   972
   End
   Begin VB.ListBox lstColorSets 
      Appearance      =   0  'Flat
      Height          =   2565
      ItemData        =   "MultipleColorSets.frx":001C
      Left            =   240
      List            =   "MultipleColorSets.frx":0023
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      ToolTipText     =   "Press Ctrl to choose multiple color sets"
      Top             =   420
      Width           =   1815
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Display value"
      Height          =   195
      Index           =   1
      Left            =   2220
      TabIndex        =   5
      Top             =   180
      Width           =   945
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Sets"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frmMultipleColorSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
   Tag = "OK"
   Hide
End Sub
Private Sub cmdCancel_Click()
   Tag = "Cancel"
   Hide
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   cmdCancel_Click
End Sub

Private Sub Form_Load()
   Left = frmDrafter.Left + 1000
   Top = frmDrafter.Top + 1000
   cmdOK.Left = Width - cmdOK.Width - 200
   cmdOK.Top = ScaleHeight - cmdOK.Height - 100
   cmdCancel.Left = cmdOK.Left - cmdCancel.Width - 100
   cmdCancel.Top = cmdOK.Top
   cmdNone.Left = lstColorSets.Left
   cmdNone.Top = lstColorSets.Top + lstColorSets.Height + 50
   cmdAll.Left = lstColorSets.Left + lstColorSets.Width - cmdAll.Width
   cmdAll.Top = cmdNone.Top
End Sub
Private Sub cmdAll_Click()
   Dim i As Integer
   
   For i = 0 To lstColorSets.ListCount - 1
      lstColorSets.selected(i) = True
   Next i
End Sub
Private Sub cmdNone_Click()
   Dim i As Integer
   
   For i = 0 To lstColorSets.ListCount - 1
      lstColorSets.selected(i) = False
   Next i
End Sub



