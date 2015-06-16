VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfigure 
   BackColor       =   &H00C0FFFF&
   Caption         =   "MAPPFinder Data Path"
   ClientHeight    =   2670
   ClientLeft      =   3855
   ClientTop       =   3525
   ClientWidth     =   6315
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6315
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtGenMAPP 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -120
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmConfig.frx":08CA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "GenMAPP_Sc|*.exe"
   CommonDialog1.ShowOpen
   Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
   On Error GoTo error
   Dim Fsys As Object, config As TextStream
   Dim GenMAPPconfig As String
   Dim GenMAPPConfigFile As TextStream
   Dim line As String, baseMAPP As String, databaseloc As String
   
   If txtGenMAPP.Text = "" Then
      MsgBox "You have not selected the GenMAPP program. You must have GenMAPP installed to run" _
         & " MAPPFinder. GenMAPP can be downloaded from www.GenMAPP.org.", vbOKOnly
   Else 'they selected something so configure.
   Set Fsys = CreateObject("Scripting.FileSystemObject")
   If Right(programpath, 1) <> "\" Then               'Root directory has a backslash, others don't
      programpath = programpath & "\"
   End If
   Set config = Fsys.CreateTextFile(Module1.programpath & "MAPPFinder.cfg")
   config.WriteLine ("MAPPFinder config file. Do not alter or delete.")
   If txtGenMAPP.Text <> "" Then
      config.WriteLine (txtGenMAPP.Text) 'genmapp location
      GenMAPPconfig = Mid(txtGenMAPP.Text, 1, Len(txtGenMAPP.Text) - 3) & "cfg"
      Set GenMAPPConfigFile = Fsys.OpenTextFile(GenMAPPconfig)
      While InStr(1, line, "baseFolder:") = 0
         line = GenMAPPConfigFile.ReadLine
      Wend
      baseMAPP = Mid(line, 13, Len(line) - 13 + 1) & "MAPPs\"
      
      While InStr(1, line, "mruGeneDB:") = 0
         line = GenMAPPConfigFile.ReadLine
      Wend
      databaseloc = Mid(line, 12, Len(line) - 12 + 1)
      config.WriteLine (databaseloc)
      
      config.WriteLine (baseMAPP)
      AddFolder (baseMAPP & "MAPPFinder")
      GenMAPPConfigFile.Close
   End If
   config.Close
   frmConfigure.Hide
   End If
error:
   Select Case Err.Number
      Case 5
         MsgBox "File not found. Please select a new file."
      Case 53
         MsgBox "MAPPFinder cannot find the GenMAPP.cfg file. You need to run the GenMAPP program before using MAPPFinder " _
            & "for the first time.", vbOKOnly
         Exit_Click
      
   End Select
End Sub

Private Sub Close_Click()
    End
End Sub

Private Sub Command3_Click()
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "GenMAPP|*.exe"
   CommonDialog1.ShowOpen
   txtGenMAPP.Text = CommonDialog1.FileName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormCode Then                                            'Exiting from mnuExit
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   
 Exit_Click

End Sub
Private Sub Exit_Click()
   End
End Sub
Public Function AddFolder(Path As String) As String '****************************** Adds Folder To Storage
   'written by Steve Lawlor for GenMAPP .
   ' 11/20/00
   '  Entry:
   '     path  Path to be added to directory structure. May or may not end in \
   '  Return:
   '     Part of path that already exists. To be used to remove added path later if
   '        not needed
   '  For example: Path to be added is
   '        C:\Large\Medium\Small
   '     If C:\Large already existed, C:\Large is returned
   Dim root As String                                            'Part of path that already existed
   Dim partialPath As String, drive As String
   Dim slash As Integer, nextSlash As Integer
         
On Error GoTo errorhandler
   Path = Dat(Path)
   If InStr(Path, ":") = 0 Then                                                           'No drive
      If Left(Path, 1) = "\" Then                                                'Add current drive
         Path = Left(CurDir, InStr(CurDir, ":")) & Path
      Else
         Path = Left(CurDir, InStr(CurDir, ":")) & "\" & Path
      End If
   End If
   slash = InStr(Path, "\")
   Do While slash < Len(Path)
      nextSlash = InStr(slash + 1, Path, "\")
      If nextSlash = 0 Then nextSlash = Len(Path) + 1
      partialPath = Left(Path, nextSlash - 1)
      If Dir(partialPath, vbDirectory) = "" Then
         MkDir partialPath
         If root = "" Then
            root = Left(Path, slash - 1)
         End If
      End If
      slash = nextSlash
   Loop
ExitFunction:
   AddFolder = root
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

errorhandler:
   MsgBox Err.Description & ". Folder not created", vbCritical + vbOKOnly, "Creating folder"
   root = "<ERROR>"
   Resume ExitFunction
End Function

Function Dat(ByVal z As Variant) As String
   Rem************************************************************************
   Rem  CONVERTS VARIANT, PARTICULARLY DATABASE FIELD, TO STRING *************
   Rem************************************************************************
   On Error GoTo DatError
   If VarType(z) <> vbNull Then
      Dat = Trim(z)
   Else
      Dat = ""
   
   End If
DatContinue:
   Exit Function

DatError:
   Dat = ""
   Resume DatContinue
End Function
