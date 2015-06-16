VERSION 5.00
Begin VB.Form frmChangeParams_Old 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Color Set"
   ClientHeight    =   3132
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   4560
   Icon            =   "ChangeParams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3132
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstColorSets 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2448
      ItemData        =   "ChangeParams.frx":08CA
      Left            =   120
      List            =   "ChangeParams.frx":08CC
      TabIndex        =   1
      Top             =   540
      Width           =   4332
   End
   Begin VB.Label lblCurrentColorSet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Top             =   60
      Width           =   48
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click on a color set:"
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
      TabIndex        =   0
      Top             =   300
      Width           =   1740
   End
End
Attribute VB_Name = "frmChangeParams_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
   Const indent = 300
   Const ITEM_HEIGHT = 240
   Const ITEM_START = 300
   Dim index As Integer, maxWidth As Integer
   
'   For index = 0 To lblParam.UBound
'      lblParam(index).Visible = True
'      lblParam(index).Left = INDENT
'      lblParam(index).Top = ITEM_START + ITEM_HEIGHT * (index)
'      If lblParam(index).Width > maxWidth Then maxWidth = lblParam(index).Width
'   Next index
'   If maxWidth > 4000 Then
'      Width = maxWidth
'   Else
'      Width = 4000
'   End If
'   Height = ITEM_START + ITEM_HEIGHT * (index + 1) + 200
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Tag = ""
'   Tag = colorSet     'These statement have no effect. If colorset is "XYZ" tag shows "" after
'   Hide               'form is exited, even if Hide is explicitly executed
End Sub

Private Sub lblParam_Click(index As Integer)
'   Tag = lblParam(index)
'   For i = 0 To lblParam.UBound
'      lblParam(i).ForeColor = vbBlack
'   Next i
'   lblParam(index).ForeColor = vbRed
'   DoEvents
'   For i = 1 To lblParam.UBound
'      Unload lblParam(i)
'   Next i
'   lblParam(0).ForeColor = vbBlack           'If left red, makes all others red on next Activate
   Hide
End Sub

Private Sub lstColorSets_Click()
   Tag = lstColorSets.List(lstColorSets.ListIndex)
   Hide
End Sub
