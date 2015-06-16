VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmLabelData 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Label Data"
   ClientHeight    =   2856
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3444
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LabelData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2856
   ScaleWidth      =   3444
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFonts 
      Caption         =   "Fonts"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Click here to change the font, style, size, or color of the label"
      Top             =   2400
      Width           =   972
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   180
      Top             =   3420
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   2400
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   1380
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2400
      Width           =   972
   End
   Begin VB.TextBox txtLinks 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2508
   End
   Begin VB.TextBox txtRemarks 
      Height          =   1140
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "LabelData.frx":08CA
      ToolTipText     =   "These notes will appear only in this window"
      Top             =   1200
      Width           =   3228
   End
   Begin VB.TextBox txtLabel 
      Height          =   300
      Left            =   840
      MaxLength       =   50
      TabIndex        =   0
      ToolTipText     =   "Appears on the MAPP graphic"
      Top             =   60
      Width           =   2508
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link to"
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   588
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   804
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   492
   End
End
Attribute VB_Name = "frmLabelData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim formDataDirty As Boolean

Private Sub cmdFonts_Click()
   dlgDialog.CancelError = True
On Error GoTo ErrorHandler
   Dim fontStyle As Integer
   
   With mappWindow
   dlgDialog.FLAGS = cdlCFBoth Or cdlCFEffects
   dlgDialog.FontName = .activeObject.id
   dlgDialog.fontSize = .activeObject.Size
   dlgDialog.FontBold = Asc(.activeObject.SystemCode) And 1
   dlgDialog.FontItalic = Asc(.activeObject.SystemCode) And 2
   dlgDialog.FontUnderline = Asc(.activeObject.SystemCode) And 4
   dlgDialog.FontStrikethru = Asc(.activeObject.SystemCode) And 8
   dlgDialog.color = .activeObject.color
   dlgDialog.ShowFont
   
   .activeObject.id = dlgDialog.FontName
   .activeObject.Size = dlgDialog.fontSize
   If dlgDialog.FontBold Then fontStyle = fontStyle + 1
   If dlgDialog.FontItalic Then fontStyle = fontStyle + 2
   If dlgDialog.FontUnderline Then fontStyle = fontStyle + 4
   If dlgDialog.FontStrikethru Then fontStyle = fontStyle + 8
   .activeObject.SystemCode = Chr(fontStyle + 16)
   '  This code cannot be a null character or the SQL statement that puts it into the
   '  database does not work. Therefore, regular font is 16 instead of zero.
   .activeObject.color = dlgDialog.color
   End With
Exit Sub

ErrorHandler:
   If Err <> 32755 Then                                                          'Other than Cancel
      MsgBox Err.Description, vbCritical, "Label Font Change Error"
   End If
End Sub

Private Sub Form_Activate()
   txtLabel = mappWindow.activeObject.title
   txtRemarks = mappWindow.activeObject.remarks
   txtLinks = mappWindow.activeObject.links
   If txtLabel = "Label" Then
      txtLabel.SelStart = 0
      txtLabel.SelLength = 100
      txtLabel.SetFocus
   End If
   formDataDirty = False
End Sub

Private Sub Form_Resize()
'   Dim horizAdjust As Single, vertAdjust As Single
   
'   horizAdjust = Width - ScaleWidth
'   vertAdjust = Height - ScaleHeight
   If Width < 3500 Then
      Width = 3500
   End If
   If Height < 3200 Then
      Height = 3200
   End If
   txtLabel.Width = ScaleWidth - 936
   txtLinks.Width = ScaleWidth - 936
   txtRemarks.Width = ScaleWidth - 216
   txtRemarks.Height = ScaleHeight - 1716
   cmdFonts.Top = ScaleHeight - 456
   cmdCancel.Top = ScaleHeight - 456
   cmdOK.Top = ScaleHeight - 456
   cmdCancel.Left = ScaleWidth - 2064
   cmdOK.Left = ScaleWidth - 1044
End Sub

Private Sub txtLabel_Change()
   formDataDirty = True
End Sub
Private Sub txtLinks_Change()
   formDataDirty = True
End Sub
Private Sub txtRemarks_Change()
   formDataDirty = True
End Sub

Private Sub cmdCancel_Click()
   formDataDirty = False
   Hide
End Sub
Private Sub cmdOK_Click()
   If formDataDirty Then
      txtLabel = TextToSql(txtLabel)
      If InvalidChr(txtLabel, "label") Then
         txtLabel.SetFocus
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
      txtRemarks = TextToSql(txtRemarks)
      If InvalidChr(txtRemarks, "notes") Then
         txtRemarks.SetFocus
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
      
      With mappWindow
         .dirty = True
         .activeObject.title = Dat(txtLabel)
         .activeObject.links = Dat(txtLinks)
         .activeObject.remarks = Dat(txtRemarks)
      End With
   End If
   cmdCancel_Click
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If formDataDirty Then
      Select Case MsgBox("Save changes?", vbYesNoCancel + vbQuestion, "Closing Label Data Window")
      Case vbYes
         cmdOK_Click
      Case vbNo
         cmdCancel_Click
      Case Else
         Cancel = True                                                                 'Don't close
      End Select
   End If
End Sub


