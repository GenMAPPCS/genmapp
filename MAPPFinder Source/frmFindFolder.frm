VERSION 5.00
Begin VB.Form frmFindFolder 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Select the Location of your Local MAPPs"
   ClientHeight    =   4545
   ClientLeft      =   4560
   ClientTop       =   3420
   ClientWidth     =   5340
   Icon            =   "frmFindFolder.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5340
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The folder that is shown as open is the folder you are selecting. You must double-click on your folder of choice."
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Folder:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Drive:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmFindFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MappPath As String
Dim species As String

Private Sub Command1_Click()
    frmLocalMAPPs.txtLocalMAPPs = Dir1.Path
    frmFindFolder.Hide
End Sub

Private Sub Command2_Click()
   frmFindFolder.Hide
End Sub

Private Sub Dir1_Change()
   Dir1.Refresh
End Sub

Private Sub Drive1_Change()
    On Error GoTo catch
    Dir1.Path = Drive1.drive
    Dir1.Refresh
    Exit Sub
catch:
    Select Case Err.Number
    Case 68
        MsgBox "The device you have selected is unavailable. Check it and try again.", vbOKOnly
    Case Else
        MsgBox "An error occured locating that drive."
    End Select
End Sub
Public Sub Load()
   On Error GoTo error
   Drive1.drive = Mid(MAPPFolder, 1, 3)
   Dir1.Path = MAPPFolder

error:
   Select Case Err.Number
      Case 76 'path not found
         MsgBox "MAPPFinder thinks the default MAPP location is " & MAPPFolder & " but " _
            & "this folder no longer exists. Your MAPPFinder.cfg file is out of date. " _
            & "You need to delete this file and reopen MAPPFinder. Sorry for the inconvenience.", vbOKOnly
         End
   End Select


End Sub
Public Sub setSpecies(newspecies As String)
   species = newspecies
End Sub
