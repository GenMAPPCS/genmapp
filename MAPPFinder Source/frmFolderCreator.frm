VERSION 5.00
Begin VB.Form frmFolderCreator 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Enter the name of the new folder"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   Icon            =   "frmFolderCreator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtFolderName 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label lblFolderLoc 
      BackColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Please enter the name of the folder you would like to create to store the exported MAPPs:"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmFolderCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cancel As Integer

Private Sub Text1_Change()

End Sub

Private Sub cmdCancel_Click()
    Cancel = 1
    Me.Hide
End Sub

Private Sub Command1_Click()
   If txtFolderName.Text = "" Then
      MsgBox "You must enter a folder name.", vbOKOnly
      Exit Sub
   End If
   
   If invalidFileName(txtFolderName.Text) Then
      MsgBox "A filename cannot contain any of the following characters: /\:*?" & Chr(34) & "<>| are not", vbOKOnly
      txtFolderName.Text = ""
      Exit Sub
   End If
   
   'txtFolderName.Text = MappBuilderForm_Normal.fixPath(txtFolderName.Text)
   
   If Dir(mapploc & TreeForm.species & "\" & txtFolderName.Text & "\") <> "" Then
      If MsgBox("This folder already exists. Overwrite it?", vbYesNo) = vbNo Then
         Exit Sub
      End If
   End If
   'ok, we want to export the mapps.
   Cancel = 2
   frmFolderCreator.Hide
End Sub

Private Sub Form_Load()
   lblFolderLoc.Caption = "This folder will be created in " & mapploc _
                           & TreeForm.species & "\"
End Sub

Private Sub Form_Exit()
    Cancel = 1
    Me.Hide
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Cancel = 1
   Me.Hide
   
End Sub


