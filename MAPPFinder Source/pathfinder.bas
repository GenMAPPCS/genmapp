Attribute VB_Name = "Module1"
Public Const CFG_FILE = "MAPPFinder.cfg"
Public Const MAX_RELATIONS = 30
Public Const MAX_GENES = 2000

Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
         (ByVal hwndCaller As Long, ByVal pszFile As String, _
         ByVal uCommand As Long, ByVal dwData As Long) As Long

Public databaseloc As String
Public genmapploc As String
Public MAPPFolder As String
Public mapploc As String
Public programpath As String
Public mapptmpl As String
Public dbDate As String

Public Sub Main()
  'On Error GoTo error
   Dim Fsys As Object
   Dim config As TextStream
   Dim line As String
   Dim slash As Integer
   Dim commandline As String
   Dim commandfile As TextStream
   Dim dbMAPPfinder As Database
   commandline = Command()
   'command line accepts the GenMAPP gene database. For now this is the only item that is needed.
   
   Set Fsys = CreateObject("Scripting.FileSystemObject")
   MousePointer = vbHourglass
   frmSplash.Show
   frmSplash.Refresh
   programpath = App.Path
   If Right(programpath, 1) <> "\" Then               'Root directory has a backslash, others don't
      programpath = programpath & "\"
   End If
  
   'If UCase(Dir(programpath & CFG_FILE)) <> UCase(CFG_FILE) Then '------------------------Configure
   'frmSplash.Hide
   If CreateConfigFile Then
restart:
        MousePointer = vbHourglass
        'frmSplash.Show
        'frmSplash.Refresh
        Set config = Fsys.OpenTextFile(programpath & CFG_FILE)
        line = config.ReadLine 'Don't mess with this file
        line = config.ReadLine
        genmapploc = line
        slash = InStrRev(line, "\")
        mapptmpl = Left(line, slash) & "MAPPTMPL.gtp"
        line = config.ReadLine
        databaseloc = line
        line = config.ReadLine
        MAPPFolder = line
        mapploc = line
   
   'we're now going to recreate the config file each time the program is loaded. People can change their file
   'locations, so we need to be flexible.
   
   'Else
    '  Set config = Fsys.OpenTextFile(programpath & CFG_FILE)
    '  line = config.ReadLine
    '  line = config.ReadLine
    '  genmapploc = line
    '  slash = InStrRev(line, "\")
    '  mapptmpl = Left(line, slash) & "MAPPTMPL.gtp"
    '  line = config.ReadLine
    '  databaseloc = line
    '  line = config.ReadLine
    '  MAPPFolder = line
    '  mapploc = line & "MAPPFinder\"
   'End If
  
    If commandline <> "" Then
          databaseloc = Mid(commandline, 2, Len(commandline) - 2) '"databaseloc"
    End If
   
    getMRUGEX
    If InStr(1, databaseloc, ".gdb") = 0 Then
      databaseloc = ""
      MsgBox "A GenMAPP database was not found in your settings. You will need to load a" _
            & " database before doing anything else.", vbOKOnly
      frmSplash.Hide
      frmStart.Show
      MousePointer = vbDefault
    Else
      
      Set dbMAPPfinder = OpenDatabase(databaseloc)
      Set rsdate = dbMAPPfinder.OpenRecordset("SELECT Version FROM Info")
      dbDate = rsdate![Version]
   
      dbMAPPfinder.Close
      UpdateDBlabel
  ' TreeForm.FormLoad this has been moved to happen when the user loads existing files or calculates new ones.
   
      frmSplash.Hide
      frmStart.Show
      MousePointer = vbDefault
    End If
    Else
         MsgBox "MAPPFinder could not create a configuration file. Make sure that GenMAPP is installed in the same folder.", vbOKOnly
         End
    End If
    
         
error:
   Select Case Err.Number
      Case 62
         MsgBox "The MAPPFinder configuration file has been corrupted. To fix it," _
            & " you must delete the file MAPPFinder.cfg and rerun the program."
      Case 3024
         MsgBox "The most recently used database can no longer be found. You will " _
            & "need to load a database before doing anything else.", vbOKOnly
            frmSplash.Hide
            frmStart.Show
            MousePointer = vbDefault
      Case 3055
         MsgBox "A GenMAPP database was not found in your settings. You will need to load a" _
            & " database before doing anything else.", vbOKOnly
            frmSplash.Hide
            frmStart.Show
            MousePointer = vbDefault
            
      Case 3044
         MsgBox "The configuration file is looking for the file " & databaseloc & " but this does not seem " _
            & "to exist. If you deleted a folder that previously stored GenMAPP data, you need to open and" _
            & " close GenMAPP to reset your configuration file before MAPPFinder can properly function.", vbOKOnly
          MousePointer = vbDefault
          End
   End Select
   
End Sub

Public Function invalidFileName(File As String) As Boolean
   '?/|<>":* not allowed
   If InStr(1, File, "?") > 0 Then
      invalidFileName = True
   ElseIf InStr(1, File, "/") > 0 Then
      invalidFileName = True
   'ElseIf InStr(1, File, "\") > 0 Then
    '  invalidFileName = True
   ElseIf InStr(1, File, "|") > 0 Then
      invalidFileName = True
   ElseIf InStr(1, File, "<") > 0 Then
      invalidFileName = True
   ElseIf InStr(1, File, ">") > 0 Then
      invalidFileName = True
   ElseIf InStr(1, File, Chr(34)) > 0 Then '"
      invalidFileName = True
   ElseIf InStr(3, File, ":") > 0 Then 'C:\ need to start from 3
      invalidFileName = True
   ElseIf InStr(1, File, "*") > 0 Then
      invalidFileName = True
   Else
      invalidFileName = False
   End If
End Function

Public Sub ClearDirectory(Path As String)
   Dim MyName As String
   Dim Paths As String
   Dim steps As Integer, i As Integer, j As Integer
   MyName = Dir(Path, vbDirectory)    ' Retrieve the first entry.
   steps = 1
   While MyName <> ""                 ' Start the loop.
    ' Ignore the current directory and the encompassing directory.
      If MyName <> "." And MyName <> ".." Then
        ' Use bitwise comparison to see if MyName is a directory.
      
         If (GetAttr(Path & MyName) = vbDirectory) Then ' it represents a directory.
            'Debug.Print MyName ' Display entry only if it
            ClearDirectory Path & MyName & "\"
            'now we need to take steps to get dir back to where it was before the recursion
            MyName = Dir(Path, vbDirectory)
            For i = 2 To steps
               MyName = Dir()
            Next i
         Else 'a file
            Kill Path & MyName
        End If
      End If
      MyName = Dir() ' Get next entry
      steps = steps + 1
   Wend
End Sub

Public Sub UpdateDBlabel()
   frmStart.lblDB.Caption = databaseloc
   frmCriteria.lblDB.Caption = databaseloc
   frmInput.lblDB.Caption = databaseloc
   frmLoadFiles.lblDB.Caption = databaseloc
   
   
   frmLocalMAPPs.lblspecies.Caption = ""
   frmCriteria.lblspecies.Caption = ""
   frmLoadFiles.lblspecies.Caption = ""
End Sub

Private Function CreateConfigFile() As Boolean
  'On Error GoTo error
   Dim Fsys As Object, config As TextStream
   Dim GenMAPPconfig As String
   Dim GenMAPPConfigFile As TextStream
   Dim line As String, baseMAPP As String, databaseloc As String
   Dim File As String, s As String
   
   Set Fsys = CreateObject("Scripting.FileSystemObject")
   Set config = Fsys.CreateTextFile(Module1.programpath & "MAPPFinder.cfg")
   config.WriteLine ("MAPPFinder config file. Do not alter or delete.")
   
  'now we need to find the newest genmapp.exe in this app folder
   CreateConfigFile = False
   filedate = CDate("1-JAN-1970")
   s = Dir(programpath & "GenMAPPv*.exe")
   Do Until s = ""
     If FileDateTime(programpath & s) > filedate Then
        File = s
        filedate = FileDateTime(programpath & s)
     End If
     s = Dir()
    Loop
   If File <> "" Then
     config.WriteLine (programpath & File) 'genmapp location
   Else
      MsgBox "GenMAPP not available. MAPPFinder and GenMAPP must be installed in the same folder.", vbInformation + vbOKOnly, _
             "MAPPFinder"
        CreateConfigFile = False
        Exit Function
            
  End If
   
   'now we need to find the newest genmapp.cfg in this app folder
  
   filedate = CDate("1-JAN-1970")
   s = Dir(programpath & "GenMAPP.cfg")
   Do Until s = ""
      If FileDateTime(programpath & s) > filedate Then
         File = s
         filedate = FileDateTime(programpath & s)
      End If
      s = Dir()
   Loop
   If File <> "" Then
    GenMAPPconfig = programpath & File
    Set GenMAPPConfigFile = Fsys.OpenTextFile(GenMAPPconfig)
   Else
      MsgBox "GenMAPP.cfg not available. You must open GenMAPP before you can use MAPPFinder.", vbInformation + vbOKOnly, _
             "MAPPFinder"
   End If
   
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
   config.Close
   CreateConfigFile = True
error:
   Select Case Err.Number
      Case 5
         MsgBox "File not found. I'm looking for GenMAPPv2.exe and GenMAPPv2.cfg in the applications folder containing MAPPFinder." _
            & " These files should be there. If you changed something please return it to the default configuration.", vbOKOnly
      Case 53
         MsgBox "MAPPFinder cannot find the GenMAPP.cfg file. You need to run the GenMAPP program before using MAPPFinder " _
            & "for the first time.", vbOKOnly
         End
      
   End Select
End Function


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

Private Sub getMRUGEX()
   'open the genmapp.cfg and find the most recently used GEX
   Dim File As String, s As String
   Dim filedate As Date
   Dim GenMAPPconfig As String
   Dim GenMAPPConfigFile As TextStream
   Dim line As String, mruGEX As String
   Dim Fsys As Object
   Set Fsys = CreateObject("Scripting.FileSystemObject")
   filedate = CDate("1-JAN-1970")
   s = Dir(programpath & "GenMAPP.cfg")
   Do Until s = ""
      If FileDateTime(programpath & s) > filedate Then
         File = s
         filedate = FileDateTime(programpath & s)
      End If
      s = Dir()
   Loop
   If File <> "" Then
      GenMAPPconfig = programpath & File
      Set GenMAPPConfigFile = Fsys.OpenTextFile(GenMAPPconfig)
   Else
      MsgBox "GenMAPP.cfg not available. You must open GenMAPP before you can use MAPPFinder.", vbInformation + vbOKOnly, _
             "MAPPFinder"
      End
   End If
   
   While InStr(1, line, "mruDataSet:") = 0
      line = GenMAPPConfigFile.ReadLine
   Wend
   mruGEX = Mid(line, 13, Len(line) - 13 + 1)
   frmInput.txtFileName.Text = mruGEX
   GenMAPPConfigFile.Close
End Sub
