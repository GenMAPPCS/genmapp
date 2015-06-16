VERSION 5.00
Begin VB.Form frmFolder 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Path Information"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "Folder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFolder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   4380
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1380
      Width           =   2352
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4440
      TabIndex        =   4
      Top             =   2460
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5640
      TabIndex        =   3
      Top             =   2460
      Width           =   1092
   End
   Begin VB.DirListBox folders 
      Height          =   2232
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4092
   End
   Begin VB.DriveListBox drives 
      Height          =   288
      Left            =   4380
      TabIndex        =   0
      Top             =   600
      Width           =   1992
   End
   Begin VB.Label lblMessage2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "the subfolders Gene Databases,  MAPPs, and Expression Datasets."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   300
      Width           =   6000
   End
   Begin VB.Label lblCreated 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created under folder at left"
      Height          =   192
      Left            =   4380
      TabIndex        =   6
      Top             =   1680
      Width           =   1872
   End
   Begin VB.Label lblNewFolder 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4380
      TabIndex        =   5
      Top             =   1140
      Width           =   948
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose drive and folder to store the GenMAPP folder containing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   5628
   End
End
Attribute VB_Name = "frmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  Tag property indicates whether we have to write to the drive.
   Dim currentDrive As String, currentFolder As String

Private Sub Form_Activate()
'   lblMessage = "Choose drive and folder to store the GenMAPP folder containing"
   currentDrive = drives.drive
   currentFolder = folders.path
   txtFolder = ""
   If lblNewFolder.Visible Then            'Set lblFolder visible to be able to create a new folder
      txtFolder.Visible = True
      lblCreated.Visible = True
   Else
      txtFolder.Visible = False
      lblCreated.Visible = False
   End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   cmdCancel_Click
End Sub
Private Sub cmdCancel_Click()
   folders.Tag = "Cancel"
   cmdOK_Click
End Sub
Private Sub cmdOK_Click()
'   If Screen.ActiveControl.name = "txtFolder" Then
      '  Calling this by default (hitting the Enter key instead of clicking OK) does not
      '  trigger the Lost_Focus event for txtFolder
      txtFolder_LostFocus
'   End If
   Hide
End Sub

Private Sub drives_Change()
   Dim path As String, slash As Integer, driveStatus As String, invalidDrive As Boolean
   
   If Screen.ActiveForm.name = "frmFolder" Then                             'Form not still loading
      driveStatus = DriveCheck(drives.drive)
      If driveStatus = "MISSING" Then
         invalidDrive = True
         MsgBox "Drive does not exist or has no medium in it.", vbExclamation + vbOKOnly, _
                "Selecting Drive"
      ElseIf driveStatus = "RO" And InStr(Tag, "WRITE") <> 0 Then
         invalidDrive = True
         MsgBox "Drive cannot be written to.", vbExclamation + vbOKOnly, "Selecting Drive"
      End If
      If invalidDrive Then
         drives.drive = Left(currentFolder, 1)
         folders.path = currentFolder
         folders.Refresh
         drives.SetFocus
      Else
         folders.path = drives.drive
         folders.Refresh
      End If
   End If
End Sub

Private Sub txtFolder_LostFocus()
   Dim newFolder As String
   
   If Screen.ActiveControl Is cmdCancel Then
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   txtFolder = Dat(txtFolder)
   If InStr(txtFolder, "\") Then
      txtFolder = Left(txtFolder, InStr(txtFolder, "\") - 1)
   End If
   If Not ValidPathName(txtFolder) Then
      MsgBox "New folder name contains invalid characters.", vbExclamation + vbOKOnly, "New folder"
      txtFolder.SetFocus
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If InStr(txtFolder, "/") Then
      txtFolder = Left(txtFolder, InStr(txtFolder, "/") - 1)
   End If
   If Right(folders.path, 1) = "\" Then
      newFolder = folders.path & txtFolder
   Else
      newFolder = folders.path & "\" & txtFolder
   End If
   If Dir(newFolder, vbDirectory) = "" Then
      MkDir newFolder
   End If
   folders.path = newFolder
   folders.Refresh
   txtFolder = ""
End Sub
