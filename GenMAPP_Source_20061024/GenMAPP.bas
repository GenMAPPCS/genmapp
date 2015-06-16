Attribute VB_Name = "GenMAPP"
'Version 2.1 Changes:
'  20050614
'     GeneDBMgr, Add New Relationship Table. Indexed Primary and Related columns with "P" or "R"
'        at end of index name to allow for both columns to have the same Gene System code, such
'        as having a cross-species EntrezGene relationship.
'  20060126
'     Backpages and GeneFinder order annotation data according to DisplayOrder in Info table.
'     Change TextToSQL to change both single quote and typo quote to 2 single quotes.
'     GeneFinder Gene System in Search defaults to last-chosen system.
'     Implement striped genes. Change Color Set data structures to accommodate.
'     In MAPP Sets, make links between _MAPPIndex and _GeneIndex relative.
'     Allow user to type value into Zoom combo box.
'  20060226
'     GeneFinder lists only those systems deemed proper (no |I| in Systems table Misc column).
'     Option to show only first Color Set legend or all for multiple Color Sets
'  2006????
'     DisplayOrder column added to Info table in GeneDBTmpl.gtp.
'     Program will also work with Gene DBs without DisplayOrder column.
'  20060515
'     Increased MAX_SYSTEMS from 30 t0 100
'  20060522
'     Zooming big MAPPs. Unrepeatable problem in frmDrafter.ScrollBars(). Kludge fix.
'     Zooming big MAPPs small. Set minimum font size floor at 1.5 pts in Utilities.FontSizeFloor()
'     In GDM, system code text fields were only allowing 2 characters.
'  20060525
'     Implemented polygons.
'     Put Color Set selection back on Gene Index per Kristina.
'  20060613
'     Copying Color Sets made ColorSetDirty instead of ExpressionDirty.
'  20060713
'     GenMAPP.cfg now accepts extra items (like [Downloader]) without changing or deleting them.
'  20060809
'     Fix IDWithLink to accommodate multiletter system prefixes (eg: Cel.1234) in BackPages
'  20060831
'     Implement multiple Color Set coice window

'Updates in Version 2:
'  Changes in build 20031230:
'     Late additions through the EDM now add new Gene ID codes to the Info table of the ED.
'     The Color Set drop-down list in the Drafting Board window does not turn gray after exiting
'        from the EDM.
'  Changes in build 20040125:
'     Allowed rather clumsy cancellation of print to Adobe Acrobat Distiller
'     Made GeneFinder resizable
'     Processing exceptions was not making the Display table correctly
'     Processing exceptions was leaving the ".EX" at the end of the Title in the Info table
'     Copying Color Sets: Check to see that there is a Color Set to copy
'  Changes in build 20040201:
'     Exported HTML files were going to mruExport path even when path changed in dialog
'     Dashed line around gene object whenever gene appears > 1 in ED (rather than gene sets > 1)
'     Workaround for weird behavior of WebBrowser_CommandStateChanged. Forward and back
'        navigation from home page
'     Links in Remarks in expression records did not render on backpages
'  Changes in build 20040213:
'     Converter changes the SystemCode column to allow 2 chars rather than one.
'     Info Area default placement at 0, 0 (Upper-left corner).
'     Legend default placement at 0, below fully filled-in Info area.
'     Removed adding of "CE" in front of WormBase IDs because of searchable secondary columns.
'     Allow for a missing column in Gene DBs and Expression Datasets (hopefully Remarks).
'     Changed ObjectClicked order so that Legends and Info noticed before underlying objects.
'     Legend was drawing its opaque background in wrong place when zoomed.
'     Boundaries of Info area was changing when zoomed.
'  Changes in build 20040306:
'     Changed column count in import process.
'     Added a config item for mruExportSourcePath, source for MAPPs, MAPP sets, etc, to be exported
'     Added creation of new subfolder to MAPP Sets destination.
'     New folder organization for MAPP Sets. Each MAPP has its own folder.
'     Fixed bug copying Color Set to MAPP with nonmatching Expression columns
'  Changes in Build 20040327:
'     Selecting objects at zoom levels corrected.
'     Scale adjustments built into Label Data window to keep buttons, etc., in proper place.
'     Conversion process gets rid of extraneous columns in MAPPs and EDs
'  Changes in Build 20040407 and 20040407a:
'     Fixed conversion problem creating additional GeneDB column that was crashing program.
'     Adjusted Line Data and Label Data windows for different Windows font settings.
'     Added more room at bottom of About window.
'  Changes in Build 20040413:
'     In Processing Exceptions, if columns don't match, forces restating of data types.
'  Changes in Build 20040420:
'     UniGene links on backpages corrected.
'  Changes in Build 20040428:
'     Changed MAPP Set for 1 Color Set to also include "No expression data"
'     Changed MAPP Set index to list Color Sets as subs to MAPPs
'     Made MAPP Set drop-down list contain name of Color set
'  Changes in Build 20040510:
'     Error trap for Mapp Set destination folders exceeding 259 characters.
'     MAPP Set Indexes changed for Color Set choice and Gene Indexes produced in HTML
'     Some Save dialogs were not showing files of that type in folder listing
'     Legend was not disappearing when "No expression data" selected
'     MAPP in command line leaving surrounding quotes, causing EDM to come up blank
'  Changes in Build 20040516:
'     In EDM, clicking on No criteria met tried to focus on Criteria or Label, which are not enabled
'     In EDM, empty first column bombing because of quote removal. Fixed.
'     Fixed crashing problem when imported dataset has no data, only gene ID and system.
'     MAPP Set index names changed.
'  Changes in Build 20040529:
'     GeneDB Mgr not allowing addition of a table with only ID and no subsequent columns.
'     Notes changed to Remarks in line and label data popups.
'  Changes in Build 20040606:
'     Added creation of GOCount tables to the Gene DB Mgr.
'  Changes in Build 20040623:
'     In EDM, changed method of verifying raw-data files so that it works on different physical disks.
'     Implemented Switcher and modified Converter.
'  Changes in Build 20040702:
'     Cleaned up invalid property value when importing gene table.
'     ValidTableTitle() function not reading single character at a time.
'     Checked for ED in use elsewhere when opening EDM.
'     In Gene DB Mgr, added Bridge column to imported Relationship Tables.
'     In Gene DB Mgr, added indexes to imported Relationship Tables.
'     Switcher changes LocusLinkSymbol codes back from "Ls" to "L"
'  Changes in Build 20040707:
'     Added routine in UpdateMAPP() to ensure that all MAPPs have a Notes field in Objects.
'     In GDM, fixed problem that would not write a record with no species.
'  Changes in Build 20040714:
'     Added order number sort in expression data for Backpages
'     Fixed command line problem when entering EDM from MAPPFinder
'     Fixed converter problem re Remarks columns and Notes, etc.
'     Changes in choosing Expression Datasets and automatically creating Display tables.
'  Changes in Build 20040729:
'     If GeneFinder or Backpage cannot find a gene table, it moves on rather than erroring.
'     Added ¶ (Unicode 182) to Gene DB, Info, Notes to signify "Official" table.
'     Enabling of GOCount creation cleaned up.
'  Changes in Build 20040801:
'     Auto update option and functionality added.
'     More GDM and species cleanup and dumping GOCount if exit before completion
'  Changes in Build 20040815:
'     In GDM, "Other" wasn't showing up in some lists of editable tables.
'     In GeneFinder, doesn't crash on 1-chr UniGene IDs. Not found instead.
'     Major rewrite of GDM as a result of redefinition of MOD/species.
'     In converter, caught crash if file already open elsewhere.
'  Changes in Build 20040903:
'     Whole issue of adding, deleting, etc tables in official Gene DBs.
'  Changes in Build 20040921:
'     Release data control in GDB so that tables may be deleted, etc.
'  Changes in Build 20041011
'     Proper behavior of Back button in GeneFinder.
'  Changes in Build 20041023
'     GDM warns of reserved column names in imported tables (ID, SystemCode, Date).
'     Click areas for lines now are in proper places when zoomed.
'     Warning instead of crash when more than limit (2000) genes in AllRelatedGenes search.
'  Changes in Build 20041103
'     Confirmation on overwriting an existing GOCount table in Gene DB mgr.
'  Changes in Build 20041117
'     FileWritable() made more sophisticated, working with existing and nonexisting files.
'     MAPP Set creation erroring when trying to create new first-level folder.
'  Changes in Build 20041203
'     GO IDs padded to 7 characters in Relationship table imports in Gene DB Mgr.
'     In OutsideBoard(), removed "If creatingMappSet Then Exit Function"
'  Changes in Build 20041210a
'     Checks in mnuApply_Click() for current dataset.

Option Explicit

Public Declare Function InvokeFullDBDL Lib "GenMAPPDBDL.dll" _
      (handle As Long, ByVal ptr As String) As Long
   '  Calls the data downloader
Public Declare Function InvokeUpdate Lib "GenMAPPDBDL.dll" _
      (handle As Long, ByVal ptr As String) As Long
   '  Calls the Updater
Public Declare Function CheckForGMUpdates Lib "GenMAPPDBDL.dll" _
      (handle As Long, ByVal ptr As String) As Long
   '  Calls the UpdateChecker
Public Const PROGRAM_TITLE = "GenMAPP 2.1"
Public Const BUILD = "20060831"
Public Const TESTING = False

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ GeneData Constants
   '  Constants to send to GeneData() as the purpose for wanting the data
   '  AllRelatedGenes() uses the same constants
   Public Const PURPOSE_BACKPAGE = 0
   Public Const PURPOSE_GENE = 1                                                       'Gene Object
   Public Const PURPOSE_FINDER = 2                                                     'Gene Finder

Public appPath As String                                         'App.Path with consistent \ at end

'********************************************************************************* MAPP Set Globals
Public creatingMappSet As Boolean
Public htmlSuffix As String                        'Color Set suffix for an individual MAPP. Eg. _3
Public colorSetHTML As String                                        'HTML for Color Set Select box

'************************************************************************************ Extra Globals
Public callingRoutine As String
Public mruGeneFinderSystem As String                 'Last chosen in GeneFinder Search Systems list


Sub Main() '*********************************************************************** Start Procedure
   Dim cfgItem As String, cfgValue As String, colon As Integer
   Dim rsGenMAPPInfo As Recordset
   Dim hWnd As Long
   
'   MsgBox "This is a post-release test version of GenMAPP 2.0, build """ & BUILD & """."

'Dim key As String, Lin As String
'here:
'commandLine = """thing:Whatever"" ""Expression.gdb"" ""key:The Next"""
'commandLine = """Expression.gdb"" ""thing:Whatever"" ""key: The Next"""
'commandLine = """thing:Whatever"" ""key: The Next"" ""Expression.gdb"""
'commandLine = """thing:Whatever"" ""key: The Next""""Expression.gdb"""
'commandLine = "thing:Whatever key: The Next Expression.gdb"
'commandLine = "Expression.gdb thing:Whatever key:The Next"""
'commandLine = """thing:Whatever"" Expression.gdb ""key: The Next"""
'commandLine = """thing:Whatever"" ""key: The Next"" Expression.gdb"
'commandLine = "Expression.gdb ""thing:Whatever"" ""key: The Next"""
'commandLine = "Expression.gqb ""thing:Whatever"" key: The Next"
'key = "thing:"
'key = ".gdb"
'key = "key:"
'commandLine = """Expression.gdb"" ""The Next"""
'commandLine = " ""Expression.gdb"" ""The Next"""
'commandLine = """Whatever"" ""The Next"" ""Expression.gdb"""
'commandLine = """Whatever"" Expression.gdb ""The Next"""
'commandLine = "Expression.gdb ""The Next"""
'commandLine = " Expression.gdb ""The Next"""
'commandLine = """Whatever"" ""The Next"" Expression.gdb"
'Lin = CommandLineArg(commandLine, key)
'Stop
'GoTo here




'Dim x As String
'here:
'Debug.Print """" & PathCheck(GetFolder("C:\GenMAPP\Exports\")) & """"
''Debug.Print EmbedLinks("Hello there.")
'Stop
'GoTo here
'Do
'Debug.Print "         1         2         3         4         5         6         7         8         9"
'Debug.Print "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
'Debug.Print FileAbbrev("C:lakshf adsjkhas fdsalkjakjahf jkah fahadlkfewuhfkah jkadskwehrf.sdc")
'Stop
's = "jds http://kjdsf.khd.hdf/kjsdh/ksdhf whatever qwerty@asdf.jhfghj.sdf kjsdh@jhd.dhf."
''s = "qqq.www. abc@zxy.qwe.com"
'Debug.Print s
'Debug.Print EmbedLinks(s)
'Loop

   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Screen Dimensional Items
   '  See frmDrafter global declarations
   boardWidthAdjust = 7.67 * Screen.TwipsPerPixelX
   boardHeightAdjust = 58 * Screen.TwipsPerPixelY     '58
   
   frmSplash.show
   DoEvents
   appPath = App.path
      If Right(appPath, 1) <> "\" Then                   'Root folder has a backslash, others don't
         appPath = appPath & "\"
      End If
   App.HelpFile = appPath & "\GenMAPP.chm"                                        'Locate help file
   commandLine = Command()                               'Look for double-click on mapp or gex file
'   commandLine = """D:\GenMAPPv2_new\Datasets\Calcium Regulation in Cardiac Cell.mapp""" _
         & """D:\GenMAPPv2_new\Datasets\Cardiomyopathy_Model1.gex""" _
         & """colors:|Con vs Exp p<0.05|Con vs Exp, 2&3 fold|""" _
         & """set:D:\GenMAPPv2_new\Datasets\MAPPs\MAPPs 9-07-01\hu_MAPPArchive\Test""" _
         & """dest:C:\GenMAPP\Exports"""
'   commandLine = """D:\GenMAPPv2_New\Gene Databases\Mm-Std_20050206.gdb"" ""D:\GenMAPPv2_New\Datasets\Mm_Cardiomyopathies20040127.gex"" ""colors:|DCM|HCM|"" ""set:D:\GenMAPPv2_new\Test\MAPPSet1"" ""dest:D:\GenMAPPv2_new\Test\Results"""
                
'C:\Program Files\Microsoft Visual Studio\VB98\MAPPFinder 2.0 beta\GenMAPPv2.exe "C:\GenMAPP 2 Data2\MAPPs\MAPPFinder\Mus musculus\common-partner SMAD protein phosphorylation.mapp" "D:\GenMAPP 2 Data\Expression Datasets\BayGenomics022004.gex" "D:\GenMAPP 2 Data\Gene Databases\Mm-Std_20030923.gdb" "colors:|BayGenomics Traps|"
'commandLine = """Whatever"" ""Expression.gdb"" ""The next"""
'commandLine = """D:\GenMAPPv2_new\Test\Mm_testdata_V2.gex"""
'commandLine = """colors:|DCM-Recovery|"" ""D:\GenMAPPv2_new\Datasets\Mm_G Protein Signaling.mapp"" ""D:\GenMAPPv2_new\Datasets\Mm_Cardiomyopathies20040127a.gex"" ""D:\GenMAPPv2_new\Gene Databases\Mm-Std_20040411.gdb"""
'commandLine = """C:\GenMAPP 2 Data2\MAPPs\MAPPFinder\Mus musculus\common-partner SMAD protein phosphorylation.mapp"" ""D:\GenMAPP 2 Data\Expression Datasets\BayGenomics022004.gex"" ""D:\GenMAPP 2 Data\Gene Databases\Mm-Std_20030923.gdb"" ""colors:|BayGenomics Traps|"""
'      If Left(commandLine, 1) = Chr(34) Then                                'Quotes in command line
'         commandLine = Mid(commandLine, 2)                                                'Dump 'em
'      End If
'      If Right(commandLine, 1) = Chr(34) Then
'         commandLine = Left(commandLine, Len(commandLine) - 1)
'      End If
'MsgBox "Command line:" & vbCrLf & vbCrLf & commandLine
   loading = True                         'Let forms know it is an initial load so they don't react
   frmAbout!lblVersion = "Build: " & BUILD                                 'Put BUILD in Help/About
   
   Rem+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Congif File
   If UCase(Dir(appPath & CFG_FILE)) <> UCase(CFG_FILE) Then                        'No config file
      frmConfig.show vbModal
   ElseIf FileLen(appPath & CFG_FILE) < 10 Then                                  'Empty config file
      frmConfig.show vbModal
   End If
   
   ReadConfig
   If mruMappPath = "" Then                                                  'Force new config file
      frmConfig.Tag = "New base"
      frmConfig.show vbModal
   End If
   
   If cfgCheckForUpdatesOnStart = "True" Then
      CheckForGMUpdates hWnd, appPath
   End If
   
   If cfgColoring = "" Then cfgColoring = "R"
   If cfgLegend = "" Then cfgLegend = "DGECVRLIF8|"
   If InStr(cfgLegend, "F") = 0 Then cfgLegend = cfgLegend & "F8|"
   
'   If commandLine = "" And cfgInitialRun <> "False" Then '--------------------------Set Up Initial Run
'      commandLine = """" & mruMappPath & """"           'Should include MAPP file on install config
'      commandLine = commandLine & """" & mruDataSet & """"                  'Should include ED file
'      commandLine = commandLine & """" & mruGeneDB & """"              'Should include Gene DB file
'      commandLine = commandLine & """" & mruColorSet & """"
'   End If
   
''MsgBox "opening database"
'   Set dbGene = OpenDatabase(mruGeneDB)
''MsgBox "database open"
'   Set rsGenMAPPInfo = dbGene.OpenRecordset("SELECT * FROM Info")
   '------------------------------------------------------------------------Set Up Backpages Folder
   If Dir(appPath & "Backpages\", vbDirectory) <> "" Then                            'Folder exists
      If Dir(appPath & "Backpages\*.*") <> "" Then                                      'Dump files
         Kill appPath & "Backpages\*.*"
      End If
   Else
      MkDir appPath & "Backpages"                                                    'Create folder
   End If
'MsgBox "backpage folder finished"
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up Drafter Parameters
   Load frmMAPPInfo
'   With frmDrafter
'      .Width = 9000                                                      'Set visible size of board
'      .Height = 6000
'      .Top = 0
'      .Left = 0
'      .show
'   End With
   
   frmDrafter.show
'   Load frmObjects
   Unload frmSplash
   loading = False
   If commandLine <> "" Then '------------------------------------------------React To Command Line
      '  GenMAPP.exe "C:\What\Fatty.mapp" "D:\Where\UpReg.gex" "\\Lan\E\Genes\Aardvark.gdb" _
      '              "set:C:\MySet" "colors:|1 hour|2 days|14 days|" "dest:C:\Results"
      '  Individual parameters can be in any order.
      '  Extensions .mapp, .gex, .gdb are interpreted as the appropriate parameters.
      '  If no quotes, there can be only one parameter, a .mapp, .gex, or .gdb.
      '  Individual colors must be surrounded by pipes. Eg: |color| (makes parsing more consistent).
      '  All colors is "colors:All" with no pipes.
      '  No colors is "colors:None" with no pipes.
      '  If MAPP set given but no .gex, there will be no coloring.
      '  If MAPP set and .gex given but no color, color:"All" is assumed.
      '  Destination folder for the MAPP set is dest:"C:\Whatever".
      '  If no destination, the set: folder will be used (what a mess). ????????????????
      If Left(commandLine, 1) <> """" Then        'Surround command line with quotes for consistency
         commandLine = """" & commandLine & """"
      End If
      If InStr(commandLine, ".mapp""") Then
         frmDrafter.mnuOpen_Click
      End If
      If InStr(commandLine, ".gex") Then     'If both .mapp and .gex, the .gex will be stripped off
                                             'when the MAPP is opened
'MsgBox "How the Hell did we get here?"
         frmDrafter.mnuManager_Click
      End If
      If InStr(commandLine, """set:") Then                                                'MAPP Set
         If InStr(commandLine, ".gex") Then                           'Expression Dataset specified
            If InStr(commandLine, """colors:") = 0 Then                        'No colors specified
               commandLine = commandLine & " ""colors:All"""                        'Default to All
            End If
         End If
      End If
   End If
End Sub
'**************************************************************************** Opens A Gene Database
Sub OpenGeneDB(dbGene As Database, geneDB As String, Optional frm As Form = Nothing)
   '  20030315 Used in Convert.bas
   '  Entry:
   '     dbGene   An open Gene Database or "Nothing"
   '     geneDB   Path and name of Gene Database to open
   '              If blank or ends in / (path but no name), closes any open geneDB and sets it
   '                 to Nothing
   '              If "**OPEN**" then display Open dialog to choose name
   '              If "**CLOSE**" then close dbGene, set to Nothing
   '                 and set GeneDB to "No Gene Database"
   '     frm      The Form making the call. If empty or Nothing, set to active form
   '              Used to call correct dialog box and set statusbar panel for Gene Database
   '  Return:
   '     dbGene   An open Gene Database or "Nothing"
   Dim rs As Recordset, prevGeneDB As String
   
   If Not dbGene Is Nothing Then '+++++++++++++++++++++++++++++++++++ Keep Track Of Current Gene DB
      prevGeneDB = dbGene.name
   End If
   
   If geneDB = "**CLOSE**" Then '+++++++++++++++++++++++++++++++++++++ Close Any Open Gene Database
      dbGene.Close
      Set dbGene = Nothing
      geneDB = "No Gene Database"
      frm.sbrBar.Panels("Gene DB").text = "No Gene Database"
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   If frm Is Nothing Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Form
      Set frm = Screen.ActiveForm
   End If
   
   If geneDB = "" Or Right(geneDB, 1) = "\" Then '++++++++++++++++++++++++++ No Gene Database Given
      If Not dbGene Is Nothing Then '----------------------------------Close Any Open Gene Database
         dbGene.Close
         Set dbGene = Nothing
         geneDB = "No Gene Database"
         frm.sbrBar.Panels("Gene DB").text = "No Gene Database"
      End If
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If

   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Open Gene Database
   If geneDB = "**OPEN**" Then '----------------------------------------------------Use Open Dialog
      With frm.dlgDialog
On Error GoTo OpenError
         .DialogTitle = "Open Gene Database"
         .CancelError = True
         .Filter = "Gene Databases (.gdb)|gdb"
         If mruGeneDB <> "" Then
            .InitDir = Left(mruGeneDB, InStrRev(mruGeneDB, "\") - 1)
         Else
            .InitDir = "C:\"
         End If
         .FileName = "*.gdb"
         .FLAGS = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNFileMustExist
         .ShowOpen
         geneDB = .FileName
      End With
      If InStr(geneDB, ".") = 0 Then
         geneDB = geneDB & ".gdb"
      End If
On Error GoTo 0
   End If

On Error GoTo DatabaseError
   Set dbGene = OpenDatabase(geneDB)
      '  For some reason the database is not being passed by reference, so the active database
      '  on the calling form must also be set here.
   Set frm.dbGene = dbGene
On Error GoTo 0
   Set rs = dbGene.OpenRecordset("SELECT Version FROM Info")
   If Not rs.EOF Then                                   'New Gene DBs will have an empty Info table
      If InStr(rs!version, "/") Then
         MsgBox "Gene database" & vbCrLf & vbCrLf & geneDB & vbCrLf & vbCrLf & "is an obsolete " _
                & "version and cannot be used with this release of GenMAPP.", _
                vbExclamation + vbOKOnly, "Opening Gene Database"
         geneDB = prevGeneDB
         OpenGeneDB dbGene, geneDB, frm
      End If
   End If
   mruGeneDB = geneDB
   frm.sbrBar.Panels("Gene DB").text = Mid(geneDB, InStrRev(geneDB, "\") + 1)

ExitSub:
   If frm.name = mappWindow.name Then
      If mappWindow.dbGene Is Nothing Then
         mappWindow.mnuGeneDBInfo.Enabled = False
      Else
         mappWindow.mnuGeneDBInfo.Enabled = True
      End If
   End If
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Error Handlers
OpenError:
   If Err.number = 32755 Then                                                               'Cancel
      
   Else                                                                          'Other than Cancel
      FatalError "GenMAPP:OpenGeneDB", Err.Description
   End If
   On Error GoTo 0
   Resume ExitSub
   
DatabaseError:
   MsgBox "Gene database" & vbCrLf & vbCrLf & geneDB & vbCrLf & vbCrLf & "Could not be opened. " _
          & "It may not exist, be set to Read-only or be in use by someone else.", _
          vbExclamation + vbOKOnly, "Opening Gene Database"
'  Don't reset database or statusbar
'   geneDB = "No Gene Database"
'   GoTo ExitSub                                            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
End Sub
Rem ************************************************************* Produces Backpage For Single Gene
Function CreateBackpage(idIn As String, systemIn As String, head As String, dbGene As Database, _
                        Optional dbExpression As Database = Nothing, _
                        Optional obj As Object = Nothing, _
                        Optional path As String = "", _
                        Optional purpose As Integer = PURPOSE_BACKPAGE) As String
   '  Entry:
   '     idIn           Gene identification received (may have to search to find primary)
   '     systemIn       Cataloging system for passed idIn
   '     head           Backpage head
   '     dbGene         The Gene Database for the particular drafter window
   '     dbExpression   An open Expression Dataset (or Nothing)
   '     obj            The object (gene box) for which backpage being created. For GeneFinder,
   '                    this will be Nothing
   '     path           A return variable
   '     purpose        Reason for the Backpage
   '?Typically the gene label with object ID appended to it. Eg. MyGene[234]
   '  Return:     The path to the created HTML file or blank if it could not be created
   '     path           Path for backpage. Defaults to appPath\Backpages\. For exports the path
   '                    would be appropriate for that export. If path does not end in \ then the
   '                    filename and path is used (as for GeneFinder).
   '     purpose        -1 if the gene is not in the Gene DB
   'Call
   '  path = CreateBackpage(idIn, systemIn, head, dbGene, dbExpression, label & "[" & objID & "]", path)
   
   Dim geneRemarks As String, sql As String, rsInfo As Recordset, htmFile As String
   Dim rsObjects As Recordset, rs As Recordset
   Dim row As Integer, col As Integer
   Dim geneTitle As String                               'Each column title in HTML with links, etc
   Dim columnHeads As String                                    'GeneID heading row for all columns
   Dim annotations As String                                                       'Annotation data
   Dim expTable As String                                                 'Expression table in HTML
   Dim currentMousePointer As Integer                         'MousePointer on entry, reset on exit
   Dim centerColor As String               '#hex rgb color for HTML output for center and rim genes
   Dim rimColor As String
   Dim colColor As String                                            'Color for a particular column
   Dim legendLink As String                                           'Relative link to Legend page
   'For AllRelatedGenes()
      Dim genes As Integer
      Dim geneIDs(MAX_GENES, 2) As String
      Dim geneFound As Boolean
   'For AllExpressionData()
      Dim rows As Integer
      Dim rowIDs(MAX_GENES, 1) As String
      Dim columns As Integer
      Dim colorSetTitles(MAX_COLORSETS) As String
      Dim colorSets As Integer
      Dim titleColors(MAX_COLORSETS, 1) As Long
      Dim geneColors(MAX_GENES, MAX_COLORSETS) As Long
      Dim orderNos(MAX_GENES) As Long
      Dim legendPage As String
      Dim legendPageTitle As String
   'For AnnotationData()
      Dim jumps As String
   Dim topOfPage As String          'HTML to jump to top of page. Must be same as in AnnotationData
   
   topOfPage = "&nbsp;&nbsp;<font size=2><a href=""#Top"">Top</a></font>"

   currentMousePointer = Screen.ActiveForm.MousePointer
   Screen.ActiveForm.MousePointer = vbHourglass
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine HTML File Name
   If path = "" Then path = appPath & "Backpages\"
   If Right(path, 1) <> "\" Then                              'Literal Path Plus File Name Received
      htmFile = path
   Else                                                                'Just Path, Append File Name
      If Not obj Is Nothing Then
         If obj.id = "" Then                                      'Unidentified object, no Backpage
            CreateBackpage = ""
            GoTo ExitFunction                              'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         Else
            htmFile = TextToFileName(obj.title & "_" & obj.id) & ".htm"
         End If
      Else
         htmFile = TextToFileName(idIn) & ".htm"
      End If
      htmFile = ValidHTMLName(htmFile, False)
      htmFile = path & htmFile
   End If
   
   
   If htmlSuffix <> "" And Dir(htmFile) <> "" Then 'htmlSuffix <> "_1"
      '  There is a htmlSuffix, which means that HTML pages are being produced with MAPPs
      '  for each criterion in a Color Set. The Backpages are the same for all Color Sets,
      '  therefore they are produced only once for all MAPPs in the group.
      CreateBackpage = htmFile
      GoTo ExitFunction                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
'   If InStr(cfgColoring, "R") Then
      AllRelatedGenes idIn, systemIn, dbGene, genes, geneIDs, geneFound
'   Else
'      genes = 1
'      geneIDs(0, 0) = idIn
'      geneIDs(0, 1) = systemIn
'   End If
   
   If Not geneFound And purpose = PURPOSE_FINDER Then                                                          'Gene not found
'  Even if gene not found in Gene DB, it should still produce a backpage if in the ED.
'  AllRelatedGenes always returns the gene sent to it.
      CreateBackpage = ""
      GoTo ExitFunction                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up HTML Backpage File
'   pageTitle = obj.head
'   If pageTitle = "" Then pageTitle = obj.title
   Open htmFile For Output As #31
   Print #31, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
   Print #31, "<html>"
   Print #31, "<head>"
   Print #31, "   <title>" & head & " Backpage</title>"
   Print #31, "   <meta name=""generator"" content=""" & PROGRAM_TITLE & """>"
   Print #31, "</head>"
   Print #31, ""
   Print #31, "<body>"
   Print #31, "<h1 align=center><a name=""Top"">" & head & "</a></h1>"             'Eg: G alpha o 1
   If geneFound Then '===================================System And ID Subhead (Eg: UniProt P18872)
      Set rsInfo = dbGene.OpenRecordset( _
            "SELECT System FROM Systems WHERE SystemCode = '" & systemIn & "'")
      If systemIn = "O" Then
         Set rs = dbGene.OpenRecordset( _
               "SELECT SystemCode FROM Other WHERE ID = '" & idIn & "'")
         s = rs!systemCode
         Set rs = dbGene.OpenRecordset( _
               "SELECT System FROM Systems WHERE SystemCode = '" & s & "'")
         If rs.EOF Then                                            'SystemCode not in Systems table
            Print #31, "<p align=center>" & rsInfo!system & "  " & idIn _
                       & " (Unidentified Gene System """ & s & """)</a></h1>"
         Else
            Print #31, "<p align=center>" & rsInfo!system & "  " & idIn _
                       & " (" & rs!system & ")</a></h1>"
         End If
      Else
         Print #31, "<p align=center>" & rsInfo!system & "  " & idIn & "</a></h1>"
      End If
      annotations = AnnotationData(genes, geneIDs, dbGene, jumps, purpose)
   Else
      Print #31, "<p align=center>" & idIn & " not in Gene Database</a></h1>"
   End If
'GoTo here

'   ExpressionData obj, titles, ids, primaryTypes, values, remarks, rows, columns, 0
   If Not dbExpression Is Nothing Then '++++++++++++++++++++++++++++++++++++++++ Expression Section
      ReDim columnTitles(dbExpression.TableDefs("Expression").Fields.count - 4) As String
      ReDim expData(MAX_GENES, dbExpression.TableDefs!expression.Fields.count - 4) As Variant
      '=============================================================================Expression Data
      AllExpressionData genes, geneIDs, dbExpression, rows, rowIDs, columns, columnTitles, _
                        expData, colorSetTitles, colorSets, titleColors, geneColors, orderNos, _
                        legendPage, legendPageTitle
      legendLink = ValidHTMLName(legendPageTitle, False) & "_Legend.htm"
      If rows Then                                                          'Expression data exists
         expTable = "<h2><i><a name=""ExpressionProfile"">Expression Profile</a></i>" _
                         & topOfPage & "</h2>" & vbCrLf
         expTable = expTable & "<table border=1>" & vbCrLf _
                  & "   <tr>" & vbCrLf _
                  & "      <td align=left><b>Gene&nbsp;ID</b></td>" & vbCrLf
                  
         For col = 0 To rows - 1 '=====================================================Column Heads
            '  The Backpage expression data table reverses rows and columns
            '  from the Expression Dataset. Rows are columns and vice versa.
            geneTitle = "<b>" & GeneIDwithLink(rowIDs(col, 0), rowIDs(col, 1), dbGene) & "</b>"
            columnHeads = columnHeads & "      <td align=right>" _
                     & geneTitle & "</td>" & vbCrLf
            If expData(row, columns - 1) <> "" Then '===============================Pick up Remarks
               '  This is always the last column (zero-based)
'               geneRemarks = geneRemarks & "&nbsp;&nbsp;&nbsp;<b>" & rowIDs(col, 0) & ":</b> " _
'                           & EmbedLinks(expData(col, columns - 1)) & "<br>" & vbCrLf
               geneRemarks = geneRemarks & "  &nbsp;&nbsp;&nbsp;<b>" & rowIDs(col, 0) & ":</b> " _
                           & expData(col, columns - 1) & "<br>" & vbCrLf
            End If
         Next col
         columnHeads = columnHeads & "   </tr>" & vbCrLf
'         expTable = expTable & "   </tr>" & vbCrLf
         expTable = expTable & columnHeads
         
         '========================================================================Color Set Section
         Dim colorSet As Integer, colorSetData As String, noOfColors As Integer, color As Integer
         Dim colorInstances(MAX_CRITERIA, 2) As Long
            '  colorInstances(x, 0)    Color
            '  colorInstances(x, 1)    Instances of that color
            '  colorInstances(x, 2)    Highest order number in that color
         
         For colorSet = 0 To colorSets - 1 '-----------------------------------------Process Colors
            expTable = expTable & "   <tr>" & vbCrLf
            For i = 0 To noOfColors '_____________________________Set Last-Used Array Back To Zeros
               colorInstances(i, 0) = 0                                                      'Color
               colorInstances(i, 1) = 0                                    'Instances of that color
               colorInstances(i, 2) = 0                         'Highest order number in that color
            Next i
            colorSetData = ""
            noOfColors = 0
            For col = 0 To rows - 1
               colorSetData = colorSetData & "      <td align=right bgcolor=""" _
                        & HtmlHexColor(geneColors(col, colorSet)) & """>" & "</td>" & vbCrLf
               For color = 0 To noOfColors - 1
                  If geneColors(col, colorSet) = colorInstances(color, 0) Then
                     colorInstances(color, 1) = colorInstances(color, 1) + 1
                     If orderNos(col) < colorInstances(color, 2) Then
                        colorInstances(color, 2) = orderNos(col)
                     End If
                     Exit For                              '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                  End If
               Next color
               If color >= noOfColors Then                                               'New color
                  colorInstances(color, 0) = geneColors(col, colorSet)
                  colorInstances(color, 1) = colorInstances(color, 1) + 1
                  colorInstances(color, 2) = orderNos(col)
                  noOfColors = noOfColors + 1
               End If
            Next col
            
            '--------------------------------------------------------------------------Order Colors
            '  Order gene sets by the number of occurrences descending, then order number
            '  ascending in each set.
            '  At the end of this, gene colors are ordered by number of occurrences in the set
            '  (number of genes in a particular color -- that satisfy a particular criterion).
            '  If occurrences in 2 or more sets are tied, the set with the lowest order number
            '  comes first. This is the set with the lowest order number among its genes.
            Dim bottom As Integer
            Dim position As Integer
            Dim temp As Long
            
            For bottom = noOfColors - 1 To 1 Step -1 '_________________Few sets, simple bubble sort
               For position = 0 To bottom - 1
                  If colorInstances(position + 1, 1) > colorInstances(position, 1) _
                     Or (colorInstances(position + 1, 1) = colorInstances(position, 1) _
                         And colorInstances(position + 1, 2) < colorInstances(position, 2)) Then
                           '  Order by occurrences descending, then orderNo ascending
                     temp = colorInstances(position, 0)
                     colorInstances(position, 0) = colorInstances(position + 1, 0)
                     colorInstances(position + 1, 0) = temp
                     temp = colorInstances(position, 1)
                     colorInstances(position, 1) = colorInstances(position + 1, 1)
                     colorInstances(position + 1, 1) = temp
                     temp = colorInstances(position, 2)
                     colorInstances(position, 2) = colorInstances(position + 1, 2)
                     colorInstances(position + 1, 2) = temp
                  End If
               Next position
            Next bottom
            
            '--------------------------------------------------------------------------Assemble Row
            centerColor = HtmlHexColor(colorInstances(0, 0))
            If noOfColors > 1 Then
               rimColor = HtmlHexColor(colorInstances(1, 0))
            Else
               rimColor = centerColor
            End If
            expTable = expTable _
                     & "      <td bgcolor=""" & rimColor & """>" & vbCrLf _
                     & "         <table cellspacing=3 cellpadding=1>" & vbCrLf _
                     & "            <tr><td bgcolor=""" & centerColor & """>" _
                     & "<a href=""" & legendLink & "#Set" & colorSet & """ target=""_blank"">" _
                     & colorSetTitles(colorSet) & "</a></td></tr>" & vbCrLf _
                     & "         </table>" & vbCrLf _
                     & "      </td>" & vbCrLf & colorSetData & "   </tr>" & vbCrLf
         Next colorSet
            
         '--------------------------------------------------------------------------Expression Data
'         If Not obj Is Nothing Then '==========================================Assign Column Colors
'            centerColor = HtmlHexColor(obj.color)                       'Is vbWhite if not assigned
'            rimColor = HtmlHexColor(obj.rim)
'         Else
'            centerColor = "#FFFFFF"                                                          'White
'            rimColor = "#FFFFFF"
'         End If
         For row = 0 To columns - 2 '============================================ Expression Values
            '  columns - 2 does not include Remarks
            If row <= dbExpression.TableDefs!expression.Fields.count - 4 Then
               expTable = expTable & "   <tr>" & vbCrLf _
                        & "      <td align=left>" & columnTitles(row) & "</td>" & vbCrLf
               For col = 0 To rows - 1
'                  Select Case col                                                    'Assign column color
'                  Case 0
'                     colColor = centerColor
'                  Case 1
'                     colColor = rimColor
'                  Case Else
'                     colColor = "#FFFFFF"                                                          'White
'                  End Select
                  If expData(col, row) = "" Then
                     expTable = expTable & "      <td align=right>&nbsp;</td>" & vbCrLf
                  Else
                     expTable = expTable & "      <td align=right>" & expData(col, row) _
                              & "</td>" & vbCrLf
                  End If
'                  expTable = expTable & "      <td align=right bgcolor=""" & colColor & """>" _
                           & expData(col, row) & "</td>" & vbCrLf
      '            expTable = expTable & "      <td align=right>" & Format(values(j, i), "###0.00") _
                           & "</td>" & vbCrLf
               Next col
               expTable = expTable & "   </tr>" & vbCrLf
            End If
         Next row
         expTable = expTable & "</table>" & vbCrLf
      End If
   
      '======================================================================Write HTML Legend File
      Open path & legendLink For Output As #33
      Print #33, legendPage
      Print #33, "</body>"
      Print #33, "</html>"
      Close #33
   End If
            
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Remarks Section
   '  Moving upward from gene specific back to Remarks title
   '  If nothing in geneRemarks after all that, don't print title
   If geneRemarks <> "" Then                                                 'Gene specific remarks
'      geneRemarks = "<b>Expression Dataset (gene specific): </b><br>" & EmbedLinks(geneRemarks)
      geneRemarks = "<b>Expression Dataset (gene specific): </b><br>" & vbCrLf & geneRemarks
   End If
   If Not dbExpression Is Nothing Then
      Set rsInfo = dbExpression.OpenRecordset("SELECT * FROM Info")
      If rsInfo!remarks <> "" Then                                 'From Overall Expression Dataset
'         geneRemarks = "<b>Expression Dataset: </b>" & EmbedLinks(rsInfo!remarks) & "<br>" _
'                     & vbCrLf & geneRemarks
         geneRemarks = "<b>Expression Dataset: </b>" & rsInfo!remarks & "<br>" _
                     & vbCrLf & geneRemarks
      End If
   End If
   If Not obj Is Nothing Then
      If obj.remarks <> "" Then                                           'Remarks From Gene Finder
         geneRemarks = "<b>GenMAPP MAPP: </b>" & obj.remarks & "<br>" & vbCrLf _
                     & geneRemarks
      End If
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Print Jumps At Top Of Page
   If expTable <> "" Then                           'Expression table exists, print jump after head
      Print #31, "<p align=center><a href=""#ExpressionProfile"">Expression Profile</a></p>"
   End If
   If jumps <> "" Then
      Print #31, jumps
   End If
   
   Print #31, annotations
   
   If geneRemarks <> "" Then                                     'Remarks exist, put title in front
      Print #31, "<h2><i>Remarks</i>" & topOfPage & "</h2>"
      Print #31, geneRemarks
   End If
   
   If expTable <> "" Then                                        'Expression table exists, print it
      Print #31, expTable
      Print #31, topOfPage & "</p>"
   End If
   
here:
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ End HTML file
   Print #31, "</body>"
   Print #31, "</html>"
   Close #31
   CreateBackpage = htmFile
   
ExitFunction:
   Screen.ActiveForm.MousePointer = currentMousePointer
End Function
'**************************************************************** Displays Backpage For Single Gene
Sub ShowBackpage(obj As Object, dbGene As Database, Optional dbExpression As Database = Nothing)
   '  Entry:
   '     obj            Gene object
   '     dbGene         The Gene Database for the particular drafter window
   '     dbExpression   An open Expression Dataset (or Nothing)
   Dim head As String, path As String
   
   head = obj.head
   If head = "" Then head = obj.title
'   windowTitle = pageTitle ' & " Backpage"
   
   path = CreateBackpage(obj.id, obj.systemCode, head, dbGene, dbExpression, obj, path)
   
   If path = "" Then
      MsgBox "No backpage data available.", vbExclamation + vbOKOnly, "Generating Backpage"
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If

   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Show Browser With Backpage
   Dim IE As Object
On Error GoTo NewBrowser           'If the backpage does not already exist, a new IE object created
   AppActivate head & " Backpage"
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

NewBrowser:
   Set IE = CreateObject("InternetExplorer.Application")
   IE.visible = True
   IE.Navigate path
   IE.StatusText = obj.head
End Sub
'**************************************************************** All Annotation For A Set Of Genes
Function AnnotationData(genes As Integer, geneIDs() As String, dbGene As Database, _
                        jumps As String, Optional purpose As Integer = PURPOSE_BACKPAGE)
   '  Entry:
   '     genes                   Number of genes in set
   '     geneIDs(MAX_GENES, 1)   Gene ID for each related gene. Primary ones listed first
   '                             geneIDs(x, 0) ID
   '                             geneIDs(x, 1) SystemCode
   '                             geneIDs(x, 2) "P" for primary ID (ID column)
   '                                           "S" for secondary ID (eg: Accession in SwissProt)
   '     dbGene         Gene Database for this query (the Gene Database for the particular
   '                    drafter window)
   '     purpose        Return data for PURPOSE_BACKPAGE, PURPOSE_FINDER (Gene Finder), or
   '                    PURPOSE_GENE (Gene Object)
   '  Return:           Annotation data in HTML form
   '     jumps          HTML links to anchors for cataloging systems below
   '                       Eg: <p align=center>  <a href="#MGI>MGI</a>  <a href="#SwissProt>SwissProt</a>  </p>
   Dim rsSystems As Recordset                                                    'The Systems table
   Dim rsSystem As Recordset                                   'Table for one system. Eg: SwissProt
   Dim rsInfo As Recordset
   Dim rsRelations As Recordset
   Dim pipe As Integer, nextPipe As Integer, slash As Integer
   Dim index As Integer
'   Dim keyColumn As String    'Column in a cataloging-system table containing the unique ID. Eg: ID
                              'Now this is always "ID" so we should do away with this variable
                              'or change it to be the cataloging-systems name for the ID column
   Dim codes As String                    'Display codes between \ and | Eg: SBF in |Accession\SBF|
   Dim displayColumns(MAX_2ND_COLS) As String        'Names of columns in gene table that appear in
   Dim displayColumn As Integer                      'the Backpage or Gene Finder. Zero based
   Dim lastDisplayColumn As Integer
   Dim info As String                                                    'The HTML text being built
   Dim systemTitlePrinted As Boolean                          'If data exists in cataloging system,
                                                              'print title but only once
   Dim primaryCode As String       'System code of the primary gene. Compare to the system codes of
                                   'other genes using the Relations table to see if the relationhip
                                   'is inferred or not
   Dim inferred As String
   Dim topOfPage As String          'HTML to jump to top of page. Must be same as in CreateBackpage
   Dim noOfJumps As Integer                             'Do not jump to the first cataloging system
   Dim s As String, dot As Integer, i As Integer
   Dim orderBy As String, orderBys(50) As String, noOfOrderBys As Integer
   
   jumps = "<p align=center>&nbsp;&nbsp;"
   topOfPage = "&nbsp;&nbsp;<font size=2><a href=""#Top"">Top</a></font>"
   
   If ColumnExists(dbGene, "Info", "DisplayOrder") Then '+++++++++++++++ Determine Order Of Systems
      Set rsInfo = dbGene.OpenRecordset("SELECT DisplayOrder FROM Info")
      noOfOrderBys = SeparateValues(orderBys, rsInfo!DisplayOrder, "|")
      If noOfOrderBys >= 1 Then
         For i = 0 To noOfOrderBys - 1
            orderBy = orderBy & "SystemCode = '" & orderBys(i) & "', "
         Next i
         orderBy = Left(orderBy, Len(orderBy) - 2)                                'Dump comma space
         Set rsSystems = dbGene.OpenRecordset( _
               "SELECT * FROM Systems WHERE [Date] IS NOT NULL" & _
               "   ORDER BY " & orderBy, dbOpenForwardOnly)
      Else '=========================================================Nothing In DisplayOrder Column
         Set rsSystems = dbGene.OpenRecordset( _
               "SELECT * FROM Systems WHERE [Date] IS NOT NULL", dbOpenForwardOnly)
      End If
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ DisplayOrder Column Not Found
      Set rsSystems = dbGene.OpenRecordset( _
            "SELECT * FROM Systems WHERE [Date] IS NOT NULL", dbOpenForwardOnly)
   End If
   
   Do Until rsSystems.EOF '=================================================== Data For Each System
      systemTitlePrinted = False
         '  The rest of this stuff could actually be put in the section below under
         '  If Not systemTitlePrinted Then. ??????????????????
      pipe = InStr(rsSystems!columns, "|") '------------------------Get Name Of Each Display Column
      '  Eg: ID|Accession\SBF|Nicknames\SF|Protein|Functions\B|
      lastDisplayColumn = -1
      slash = InStr(pipe + 1, rsSystems!columns, "\")                                   'Next slash
      Do While slash
         pipe = InStrRev(rsSystems!columns, "|", slash)               'Next pipe with slash in unit
         '  If the Columns column had
         '     ID|Whatever|Accession\SBF|Nicknames\sF|Protein|Functions\B|
         '  then the pipe would be the one beginning the |Accession\SMBF| not the
         '  |Whatever| unit.
         nextPipe = InStr(pipe + 1, rsSystems!columns, "|")
         codes = Mid(rsSystems!columns, slash + 1, nextPipe - slash - 1)
         If purpose = PURPOSE_BACKPAGE Then
            If InStr(codes, "B") Then                                              'Backpage column
               lastDisplayColumn = lastDisplayColumn + 1
               displayColumns(lastDisplayColumn) _
                     = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)              'Column name
            End If
         Else
            If InStr(codes, "F") Then                                           'Gene Finder column
               lastDisplayColumn = lastDisplayColumn + 1
               displayColumns(lastDisplayColumn) _
                     = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)              'Column name
            End If
         End If
         slash = InStr(slash + 1, rsSystems!columns, "\")                               'Next slash
      Loop
      lastDisplayColumn = lastDisplayColumn + 1                               'Add "Remarks" column
      displayColumns(lastDisplayColumn) = "Remarks"
      
      primaryCode = geneIDs(0, 1)
      For index = 0 To genes - 1 '-----------------------------------------------Data For Each Gene
         If geneIDs(index, 1) = rsSystems!systemCode And geneIDs(index, 2) = "P" Then
                                                           'Gene in this system and is a primary ID
            If Not systemTitlePrinted Then '___________________________________________Set Up Title
               inferred = ""
'               If geneIDs(index, 1) <> primaryCode Then  '.............See If Relationship Inferred
'                  Set rsRelations = dbGene.OpenRecordset( _
'                        "SELECT [Type] FROM Relations" & _
'                        "   WHERE SystemCode = '" & primaryCode & _
'                        "'        AND RelatedCode = '" & geneIDs(index, 1) & _
'                        "'     OR SystemCode = '" & geneIDs(index, 1) & _
'                        "'        AND RelatedCode = '" & primaryCode & "'")
'                     '  Check both System and Related codes
'                  If rsRelations![Type] = "Inferred" Then
'                     inferred = "<font size=""-1""></b> (inferred relationship)</font>"
'                  End If
'               End If
               info = info & "<h2><i><a name=""" & rsSystems!system & """</a>" _
                           & rsSystems!system & "</i>" & inferred & topOfPage & "</h2>" & vbCrLf
               systemTitlePrinted = True
'               If noOfJumps >= 1 Then
                  jumps = jumps & "<a href=""#" & rsSystems!system & """>" & rsSystems!system _
                                & "</a>&nbsp;&nbsp;"
'               End If
               noOfJumps = noOfJumps + 1
            End If
            If Right(info, 4) <> ", " & vbCrLf Then
               '  GenBanks listed horizontally with commas so don't repeat ID:
               info = info & "   <b>ID: </b>"
            End If
            If Dat(rsSystems!link) <> "" Then '______________________Generate HTML Link To Web Site
               '  Eg: http://rgd.mcw.edu/query/query.cgi?id=~
               '  Where ~ is the gene id. Also allowing for characters after the ~
               '  UniGene is an exception
               '     Eg: http://www.ncbi.nlm.nih.gov/UniGene/clust.cgi?ORG=~&CID=~
               If rsSystems!system = "UniGene" Then
                  i = InStr(rsSystems!link, "~")
                  dot = InStr(geneIDs(index, 0), ".")
                  s = Left(rsSystems!link, i - 1) & Left(geneIDs(index, 0), dot - 1) _
                    & Mid(rsSystems!link, i + 1)
                  i = InStr(s, "~")
                  s = Left(s, i - 1) & Mid(geneIDs(index, 0), dot + 1) _
                    & Mid(s, i + 1)
               Else
                  i = InStr(rsSystems!link, "~")
                  s = Left(rsSystems!link, i - 1) & geneIDs(index, 0) _
                    & Mid(rsSystems!link, i + 1)
               End If
               info = info & GeneIDwithLink(geneIDs(index, 0), geneIDs(index, 1), dbGene) & "</a>"
'               info = info & "<a href=""" & s & """>" & geneIDs(index, 0) & "</a>"
            Else
               info = info & geneIDs(index, 0)
            End If
            If rsSystems!system = "GenBank" Then '________________________________End That ID Entry
               '  List GenBanks horizontally
               info = info & ", " & vbCrLf
            Else
               '  Others, put break between IDs
               info = info & "<br>" & vbCrLf
            End If
On Error GoTo NoTable
            Set rsSystem = dbGene.OpenRecordset( _
                  "SELECT * FROM " & rsSystems!system & _
                  "   WHERE ID = '" & geneIDs(index, 0) & "'", _
                  dbOpenForwardOnly)           'Eg: SELECT * FROM SwissProt WHERE ID = 'CALM_HUMAN'
            If Not rsSystem.EOF Then '______________________________Gene In Cataloging-System Table
               For displayColumn = 0 To lastDisplayColumn
                  s = SeparatePipes(Dat(rsSystem.Fields(displayColumns(displayColumn))))
                  If s <> "" Then                                 'Value actually in display column
                     info = info & "   <b>" & displayColumns(displayColumn) & ":</b> "
                     info = info & s & "<br>" & vbCrLf
                  End If
               Next displayColumn
               rsSystem.MoveNext
            End If
NoTableReturn:
         End If
      Next index
      rsSystems.MoveNext
      If Right(info, 4) = ", " & vbCrLf Then
         info = Left(info, Len(info) - 4) & "<br>" & vbCrLf
      End If
   Loop
   
'   If noOfJumps > 1 Then
      jumps = jumps & "</p>"
'   Else
'      jumps = ""
'   End If
   AnnotationData = info
   Exit Function                                              '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
NoTable: '++++++++++++++++++++++++++++++++++++++++++++++++ Table For Annotation Data Does Not Exist
   If Err.number = 3078 Then
      Resume NoTableReturn
   End If
   FatalError "GenMAPP:AnnotationData:", Err.Description
End Function
Function GeneIDwithLink(geneId As String, systemCode As String, dbGene As Database)
   Dim rsSystems As Recordset, i As Integer, s As String, dot As Integer
   
   Set rsSystems = dbGene.OpenRecordset( _
         "SELECT * FROM Systems " & _
         "   WHERE SystemCode = '" & systemCode & "'", dbOpenForwardOnly)
   If Dat(rsSystems!link) <> "" Then
      '  Eg: http://rgd.mcw.edu/query/query.cgi?id=~
      '  Where ~ is the gene id. Also allowing for characters after the ~
      '  UniGene is an exception
      '     Eg: http://www.ncbi.nlm.nih.gov/UniGene/clust.cgi?ORG=~&CID=~
      If rsSystems!system = "UniGene" Then
      
         i = InStr(rsSystems!link, "~")
         dot = InStr(geneId, ".")
         s = Left(rsSystems!link, i - 1) & Left(geneId, dot - 1) _
           & Mid(rsSystems!link, i + 1)
         i = InStr(s, "~")
         s = Left(s, i - 1) & Mid(geneId, dot + 1) _
           & Mid(s, i + 1)
      
'         i = InStr(rsSystems!link, "~")
'         s = Left(rsSystems!link, i - 1) & Left(geneId, 2) _
'           & Mid(rsSystems!link, i + 1)
'         i = InStr(s, "~")
'         s = Left(s, i - 1) & Mid(geneId, 4) _
'           & Mid(s, i + 1)
      Else
         i = InStr(rsSystems!link, "~")
         s = Left(rsSystems!link, i - 1) & geneId _
           & Mid(rsSystems!link, i + 1)
      End If
      
'      i = InStr(rsSystems!link, "~")
'      s = Left(rsSystems!link, i - 1) & geneId _
'        & Mid(rsSystems!link, i + 1)
      GeneIDwithLink = "<a href=""" & s & """>" & geneId & "</a>"
   Else
      GeneIDwithLink = geneId
   End If
End Function
Sub GeneData(idIn As String, systemIn As String, purpose As Integer, dbGene As Database, _
             Optional info As String, Optional dbExpression As Database = Nothing, _
             Optional value As Single, Optional centerColor As Long, Optional rimColor As Long)
   'Returns all data for a given gene
   '  Entry:
   '     idIn           Gene identification received (may have to search to find primary)
   '     systemIn       Cataloging system for passed idIn
   '     purpose        Return data for PURPOSE_BACKPAGE, PURPOSE_FINDER (Gene Finder), or
   '                    PURPOSE_GENE (Gene Object)
   '     dbGene         Gene Database for this query (the Gene Database for the particular
   '                    drafter window)
   '     dbExpression   Expression Dataset
   '                    Backpages can still be produced with no Expression Datasets,
   '                    so for purpose "Backpage", this might be Nothing.
   '  Return:
   '     info        Information about gene in text (Gene Finder) or HTML (Backpage) form.
   '     value       Expression value for gene object
   '     centerColor Color for center of gene object
   '     rimColor    Color of rim for gene object
   
   Dim ids(MAX_GENES) As String, systems(MAX_GENES) As String, values(MAX_GENES) As Single
   '  All IDs. Primary ones listed first
   Dim index As Integer, lastIndex As Integer
   Dim primaryIndex As Integer
   Dim firstPrimaryIndex As Integer              'Index for genes in the systemIn cataloging system
   '  Typically 0 because the idIn is usually the ID in the key column
   Dim lastPrimaryIndex As Integer                                                'Begins with zero
   '  Last index from the primary cataloging system. May be 0 if idIn not in primary system.
   '  This might be more than zero only if secondary IDs led back to more than one primary ID.
   '  In other words, Secondary ID X1234 occurred in two rows. Not sure this is possible.
   Dim system As String                            'Cataloging system code currently being examined
'   Dim primaryId As String
'   Dim ids(MAX_GENES) As String   'Matches in received system for primary or secondary idIns
'   Dim keyColumn As String    'Column in a cataloging-system table containing the unique ID. Eg: ID
   Dim column As String                                                       'Any secondary column
   Dim rsSystems As Recordset                                                    'The Systems table
   Dim rsSystem As Recordset                              'A cataloging-system table, eg: SwissProt
   Dim rsRelations As Recordset                                                'The Relations table
   Dim rsRelational As Recordset                         'A relational table, eg: SwissProt-GenBank
   Dim pipe As Integer, slash As Integer
   Dim sql As String
   Dim displayColumns(20) As String     'Names of columns in cataloging-system table that appear in
   Dim displayColumn As Integer         'the Backpage or Gene Finder. Zero based
   Dim lastDisplayColumn As Integer
   Dim systemTitlePrinted As Boolean                          'If data exists in cataloging system,
                                                              'print title but only once
   'For getting Expression Data
      Dim rows As Integer
      Dim geneIDs(MAX_GENES, 2) As String                  'Gene ID for each row of expression data
                                                           '  geneIDs(x, 0) ID
                                                           '  geneIDs(x, 1) SystemCode
      Dim columns As Integer
'      ReDim columnTitles(dbExpression.TableDefs!expression.Fields.Count - 4) As String
'      ReDim expData(MAX_GENES, dbExpression.TableDefs!expression.Fields.Count - 4) As Variant                       'All expression data
   
   If dbGene Is Nothing Then Exit Sub                      'No database >>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   Screen.MousePointer = vbHourglass
   
'   '++++++++++++++++++++++++++++++++++++++++++++ Find All IDs Including And Related To The Input ID
'   index = 0
'   ids(0) = idIn
'   systems(0) = systemIn
'      '  The original input is always the first element of the array. It may not exist in the
'      '  tables, in which case there will be no data reported for it.
'   firstPrimaryIndex = 0                                           'First index for a Key column ID
'
'   '===================================================== Check To See If Received ID In Key Column
'   Set rsSystems = dbGene.OpenRecordset( _
'                   "SELECT * FROM Systems WHERE SystemCode = '" & systemIn & "'", _
'                   dbOpenForwardOnly)                                              'Find system row
'                                                 'Eg: SELECT * FROM Systems WHERE SystemCode = 'Rm'
'   pipe = InStr(rsSystems!columns, "|") '------------------------------------- Check Primary Column
'   keyColumn = Left(rsSystems!columns, pipe - 1)          'First column listed in Systems in proper
'                                                          'system table. Eg: ID
'   Set rsSystem = dbGene.OpenRecordset( _
'                  "SELECT * FROM " & rsSystems!system & _
'                  "   WHERE " & keyColumn & " = '" & idIn & "'", _
'                  dbOpenForwardOnly)           'Eg: SELECT * FROM SwissProt WHERE ID = 'CALM_HUMAN'
'
'   If rsSystem.EOF Then '================================================== Check Secondary Columns
'      '  If the received ID is not in the Key column of the cataloging system table (systemIn)
'      '  go to the secondary columns and search
'      '  Typical column listing in Systems!Columns
'      '     ID|Accession\SMBF|Nicknames\sF|Protein|Functions\B|
'      firstPrimaryIndex = 1                                        'First index for a Key column ID
'         '  To be in this section of the program, the idIn (ids(0)) must not have been found in
'         '  the key column so the idIn is not an ID to be used in relational tables
'      slash = InStr(pipe + 1, rsSystems!columns, "\")                                   'Next slash
'      Do While slash '--------------------------------------------------------Each Secondary Column
'         pipe = InStrRev(rsSystems!columns, "|", slash)               'Next pipe with slash in unit
'         '  If the Columns column had ID|Whatever|Accession\SMBF|Nicknames\sF|Protein|Functions\B|
'         '  then the pipe would be the one beginning the |Accession\SMBF| not the |Whatever| unit.
'         If UCase(Mid(rsSystems!columns, slash + 1, 1)) = "S" Then         'This is a search column
'            column = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)                'Column name
'            If Mid(rsSystems!columns, slash + 1, 1) = "S" Then        'Multiple IDs surrounded by |
'               sql = "SELECT [" & keyColumn & "] AS keyColumn" & _
'                     "   FROM " & rsSystems!system & _
'                     "   WHERE [" & column & "] LIKE '*|" & idIn & "|*'"                  'Use LIKE
'               'Eg: SELECT ID AS keyColumn FROM SwissProt WHERE Accession LIKE '*|A1234|*'
'            Else                                                           'Single ID without pipes
'               sql = "SELECT [" & keyColumn & "] AS keyColumn" & _
'                     "   FROM " & rsSystems!system & _
'                     "   WHERE [" & column & "] = '" & idIn & "'"                            'Use =
'                    'Eg: SELECT ID AS keyColumn FROM SwissProt WHERE Accession = 'A1234'
'                    'Doesn't happen with SwissProt!
'            End If
'            Set rsSystem = dbGene.OpenRecordset(sql, dbOpenForwardOnly)            'Get all records
'            Do Until rsSystem.EOF                               'All rows where idIn in this column
'               index = index + 1
'               ids(index) = rsSystem!keyColumn
'               systems(index) = systemIn
'               rsSystem.MoveNext
'            Loop
'         End If
'         slash = InStr(slash + 1, rsSystems!columns, "\")                               'Next slash
'      Loop
'   End If
'   lastPrimaryIndex = index
'
'   '============================================= Follow Each Primary Gene To All Its Related Genes
'   Set rsRelations = dbGene.OpenRecordset( _
'                     "SELECT * FROM Relations" & _
'                     "   WHERE SystemCode = '" & systemIn & "'" & _
'                     "      OR RelatedCode = '" & systemIn & "'")
'                     'Eg: SELECT * FROM Relations WHERE SystemCode = 'S' OR RelatedCode = 'S'
'   For primaryIndex = firstPrimaryIndex To lastPrimaryIndex '-----------------------Each Primary ID
'      rsRelations.MoveFirst
'      Do Until rsRelations.EOF
'         If rsRelations!systemCode = systemIn Then '_____________________SystemIn In Primary Column
'            Set rsRelational = dbGene.OpenRecordset( _
'                              "SELECT * FROM [" & rsRelations!Relation & "]" & _
'                              "   WHERE Primary = '" & ids(primaryIndex) & "'")
'                              'Eg: SELECT * FROM [SwissProt-GenBank] WHERE Primary = 'CALM_HUMAN'
'            Do Until rsRelational.EOF
'               index = index + 1
'               ids(index) = rsRelational!Related
'               systems(index) = rsRelations!RelatedCode
'               rsRelational.MoveNext
'            Loop
'         Else '__________________________________________________________SystemIn In Related Column
'            Set rsRelational = dbGene.OpenRecordset( _
'                              "SELECT * FROM [" & rsRelations!Relation & "]" & _
'                              "   WHERE Related = '" & ids(primaryIndex) & "'")
'                              'Eg: SELECT * FROM [SwissProt-GenBank] WHERE Related = 'X1234'
'            Do Until rsRelational.EOF
'               index = index + 1
'               ids(index) = rsRelational!Primary
'               systems(index) = rsRelations!systemCode
'               rsRelational.MoveNext
'            Loop
'         End If
'         rsRelations.MoveNext
'      Loop
'   Next primaryIndex
'   lastIndex = index
'
'   For i = 0 To lastIndex
'      Debug.Print systems(i); " "; ids(i)
'   Next i
'   info = ""
'
'   If purpose = PURPOSE_BACKPAGE And Not dbExpression Is Nothing Then
'      AllExpressionData lastIndex, ids, systems, dbExpression, rows, geneIDs, columns, _
'                        columnTitles, expData
'
'
'
'   Debug.Print "           ";
'   For j = 0 To columns - 1
'      Debug.Print columnTitles(j); " ";
'   Next j
'   Debug.Print
'   For i = 0 To rows
'      Debug.Print geneIDs(i, 1); " "; geneIDs(i, 0); " ";
'      For j = 0 To columns - 1
'         Debug.Print expData(i, j); " ";
'      Next j
'      Debug.Print
'   Next i
'
'
'
''      i = 0 '==================================================================== Get Column Titles
''      For i = 3 To dbExpression.TableDefs!expression.Fields.Count - 3
''         '  After SystemCode field to before Remark field
''         For Each fld In dbExpression.TableDefs!expression.Fields
'
'
'
'   End If
   
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Produce Information
   If purpose = PURPOSE_BACKPAGE Or purpose = PURPOSE_FINDER Then
      Set rsSystems = dbGene.OpenRecordset("SELECT * FROM Systems", dbOpenForwardOnly)
      Do Until rsSystems.EOF '================================================ Data For Each System
         systemTitlePrinted = False
         pipe = InStr(rsSystems!columns, "|") '-------------------------------------Find Key Column
         '  Eg: ID|Accession\SBF|Nicknames\SF|Protein|Functions\B|
'         keyColumn = Left(rsSystems!columns, pipe - 1)              'First column listed in Systems
         lastDisplayColumn = -1 '----------------------------------Get Names Of Each Display Column
         slash = InStr(pipe + 1, rsSystems!columns, "\")                                'Next slash
         Do While slash
            pipe = InStrRev(rsSystems!columns, "|", slash)            'Next pipe with slash in unit
            '  If the Columns column had
            '     ID|Whatever|Accession\SMBF|Nicknames\sF|Protein|Functions\B|
            '  then the pipe would be the one beginning the |Accession\SMBF| not the
            '  |Whatever| unit.
            If purpose = PURPOSE_BACKPAGE Then
               If UCase(Mid(rsSystems!columns, slash + 1, 1)) = "B" Then           'Backpage column
                  lastDisplayColumn = lastDisplayColumn + 1
                  displayColumns(lastDisplayColumn) _
                        = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)           'Column name
               End If
            Else
               If UCase(Mid(rsSystems!columns, slash + 1, 1)) = "F" Then        'Gene Finder column
                  lastDisplayColumn = lastDisplayColumn + 1
                  displayColumns(lastDisplayColumn) _
                        = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)           'Column name
               End If
            End If
            slash = InStr(slash + 1, rsSystems!columns, "\")                            'Next slash
         Loop
         For index = 0 To lastIndex '--------------------------------------------Data For Each Gene
            If systems(index) = rsSystems!systemCode Then                      'Gene in this system
               Set rsSystem = dbGene.OpenRecordset( _
                     "SELECT * FROM " & rsSystems!system & _
                     "   WHERE ID = '" & ids(index) & "'", _
                     dbOpenForwardOnly)        'Eg: SELECT * FROM SwissProt WHERE ID = 'CALM_HUMAN'
               Do Until rsSystem.EOF                'Should be no more than one but just to be safe
                  If Not systemTitlePrinted Then
                     info = info & "<h2><i>" & rsSystems!system & "</i></h2>" & vbCrLf
                     systemTitlePrinted = True
                  End If
                  info = info & "   <b>ID: </b>"
                     '  Eg: <b>Name: </b>
                  If Dat(rsSystems!link) <> "" Then
                     '  Eg: http://rgd.mcw.edu/query/query.cgi?id=~
                     '  Where ~ is the gene id. Also allowing for characters after the ~
                     '  UniGene is an exception
                     '     Eg: http://www.ncbi.nlm.nih.gov/UniGene/clust.cgi?ORG=~&CID=~
                     i = InStr(rsSystems!link, "~")
                     s = Left(rsSystems!link, i - 1) & ids(index) & Mid(rsSystems!link, i + 1)
                     info = info & "<a href=""" & s & """>" & ids(index) & "</a><br>" & vbCrLf
                  Else
                     info = info & ids(index) & "<br>" & vbCrLf
                  End If
                  For displayColumn = 0 To lastDisplayColumn
                     info = info & "   <b>" & displayColumns(displayColumn) & ":</b> "
                     info = info & rsSystem.Fields(displayColumns(displayColumn)) & vbCrLf
'                     info = info & dbGene.TableDefs(rsSystems!system).Fields(displayColumns(displayColumn)) & vbCrLf
                  Next displayColumn
                  rsSystem.MoveNext
               Loop
            End If
         Next index
         rsSystems.MoveNext
      Loop
   End If
      
ExitSub:
'Debug.Print info
   Screen.MousePointer = vbDefault
End Sub
Function UpdateMAPP(dbMapp As Database) As Boolean '**************** Update MAPP To Current Version
   Dim rsInfo As Recordset, rsObjects As Recordset
   Dim tdfTable As TableDef
   Dim column As Field
   Dim idxIndex As index
   Dim objKey As Long
   Dim mappPath As String
   Dim ok As Boolean, confirm As Boolean, slash As Integer, oldSource As String
   Dim tiles As Boolean
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++ Checks For Notes Field In Objects Table
   '  This is a little funky. We dropped the Notes column as unneeded in MAPPs; now we are using
   '  it again to store unreconciled, extra gene IDs. Therefore, we are checking even relatively
   '  recent MAPPs to see if the Notes field exists.
   Set tdfTable = dbMapp.TableDefs("Objects")
   With tdfTable
      For Each column In tdfTable.Fields '======================================Change column names
         If column.name = "Notes" Then
            ok = True
            Exit For
         End If
      Next column
   End With
   If Not ok Then
      If GetAttr(dbMapp.name) And vbReadOnly Then
         MsgBox "MAPP is an old version and requires updating. It has been set to read-only " _
                & "through Windows and cannot be manipulated by GenMAPP. To update and open " _
                & "the MAPP, you must first unset its read-only attribute in Windows and then " _
                & "open it in GenMAPP.", vbCritical + vbOKOnly, "Read-Only MAPP"
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If

      mappPath = dbMapp.name
      Set dbMapp = OpenDatabase(mappPath)                                 'Reopen without read-only
      Set tdfTable = dbMapp.TableDefs("Objects")
      Set column = tdfTable.CreateField("Notes", dbMemo)
      tdfTable.Fields.Append column
      tdfTable.Fields("Notes").AllowZeroLength = True
      Set dbMapp = OpenDatabase(mappPath, False, True)                        'Go back to read-only
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For Current Database Version
   '  Old versions were either stated as a m/d/y date or were before date below
   Set rsInfo = dbMapp.OpenRecordset("Info", dbOpenTable)
   If InStr(rsInfo!version, "/") = 0 And rsInfo!version >= "20020725" Then                 'Current
      UpdateMAPP = True
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If GetAttr(mappPath) And vbReadOnly Then
      MsgBox "MAPP is an old version and requires updating. It has been set to read-only " _
             & "through Windows and cannot be manipulated by GenMAPP. To update and open " _
             & "the MAPP, you must first unset its read-only attribute in Windows and then " _
             & "open it in GenMAPP.", vbCritical + vbOKOnly, "Read-Only MAPP"
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If creatingMappSet Then
      MsgBox "MAPP Sets may be created only from a collection of MAPPs in which all the " _
             & "MAPPs are the current version. GenMAPP has encountered" & vbCrLf & vbCrLf _
             & dbMapp.name & vbCrLf & vbCrLf & _
             "that is not current. There may also be other MAPPs in the folder tree that " _
             & "are not current. GenMAPP will shut down. To correct the problem, run GenMAPP, " _
             & "click the ""Tools"" menu, ""Converter"" option, and convert all the MAPPs " _
             & "in the folder tree to the current version. If you use the same folder " _
             & "tree, be sure to delete copies of the old-version MAPPs. Then recreate " _
             & "your MAPP Set.", _
             vbCritical + vbOKOnly, "Old MAPP in Set"
      End                                                  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   Else
      MsgBox dbMapp.name & vbCrLf & vbCrLf & "was created in a previous version of GenMAPP " _
             & "and must be converted to the current version before it may be opened. " _
             & "In the Drafter window, click the ""Tools"" menu, ""Converter"" " _
             & "option to convert your MAPP.", _
             vbCritical + vbOKOnly, "Old MAPP"
   End If
   dbMapp.Close
   mappWindow.mappName = ""
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   
   'Old code below, may want to resurrect some later.
   
   confirm = True '++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check To Confirm Update
   If creatingMappSet Then
         '  This must be true for frmMAPPSet to be valid
      If frmMAPPSet.chkConvert = vbChecked Then
         confirm = False
      End If
   End If
   
   If confirm Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Confirm Update
      If MsgBox("Your MAPP was created in a previous version of GenMAPP. If you open it, " _
                & "it will be changed to a form compatible with the current version but " _
                & "not compatible with the previous version. Your original MAPP will be " _
                & "saved in its original location with a ""V1_"" prefix.", _
                vbExclamation + vbOKCancel, "GenMAPP Version Change") _
            = vbCancel Then
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   End If
   
   If GetAttr(dbMapp.name) And vbReadOnly Then '++++++++++++++++++++++++++++++++++++ Read-only MAPP
      MsgBox "Your MAPP was created in a previous version of GenMAPP. GenMAPP must update " _
             & "it to be compatible with the current version but it has been set to " _
             & "read-only through windows. To open this MAPP, you must first unset the " _
             & "read-only attribute in Windows.", vbOKOnly, "GenMAPP Version Change"
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
      
   rsInfo.Close
   mappPath = dbMapp.name
   dbMapp.Close
   
   slash = InStrRev(mappPath, "\")
   oldSource = Left(mappPath, slash) & "V1_" & Mid(mappPath, slash + 1)
   FileCopy mappPath, oldSource
   
   UpdateSingleMAPP mappPath
   
   If Not creatingMappSet Then
      '  Assume that if creating a MAPP set that the user doesn't want MOD conversions or
      '  tiled MAPPs on it
      Set dbMapp = OpenDatabase(mappPath)
      Set rsObjects = dbMapp.OpenRecordset("SELECT * FROM Objects WHERE SystemCode IN ('G', 'S')")
      If Not rsObjects.EOF Then
         dbMapp.Close
         If MsgBox("Your MAPP contains GenBank and/or Swiss-Prot gene IDs. Do you want " _
                   & "these converted to Model Organism Database IDs?", vbQuestion + vbYesNo, _
                   "MAPP Update") = vbYes Then
            ConvertGBorSPinFile "GS", mappWindow.dbGene, mappPath, , tiles
            If tiles Then TileWarning
         End If
      Else
         dbMapp.Close
      End If
   End If
   
'   Set dbMapp = OpenDatabase(mappPath)                                    'Reopen without read-only
'
'   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Adjust Objects Table
'   Set tdfTable = dbMapp.TableDefs("Objects")
'   With tdfTable
'      For Each column In tdfTable.Fields '======================================Change column names
'         If column.name = "Primary" Then column.name = "ID"
'         If column.name = "PrimaryType" Then column.name = "SystemCode"
'      Next column
'      ok = False '=================================================================Add ObjKey Field
'      For Each column In tdfTable.Fields
'         If column.name = "ObjKey" Then
'            ok = True
'            Exit For
'         End If
'      Next column
'      If Not ok Then
'         Set column = .CreateField("ObjKey", dbLong)
'            '  Can't make this a primary key because Jet orders the table that way. Open
'            '  creates the graphic in the order it encounters the objects in the file. Lines,
'            '  for example, must come before genes to appear behind them.
'         .Fields.Append column
'         .Fields("ObjKey").OrdinalPosition = 0
'         Set idxIndex = .CreateIndex("ixObjKey")
'         idxIndex.Fields.Append .CreateField("ObjKey")                         'Why does this work?
'         .Indexes.Append idxIndex
'      End If
'      Set rsObjects = dbMapp.OpenRecordset("SELECT MAX(ObjKey) AS objectKey FROM Objects")
'         '  Previous MAPPs may have field but not object keys, so the return here will be NULL
'      If VarType(rsObjects!objectKey) = vbNull Then
'         objKey = 0 '===============================================================Add Object Keys
'         Set rsObjects = dbMapp.OpenRecordset("Select * FROM Objects")
'         Do Until rsObjects.EOF
'            rsObjects.edit
'            If rsObjects!Type = "Curve" Then
'               '  Only old MAPPs will have the Curve object instead of Arc
'               rsObjects!Type = "Arc"
'               rsObjects!Width = (rsObjects!centerX - rsObjects!SecondX) / 2
'               rsObjects!centerX = (rsObjects!centerX + rsObjects!SecondX) / 2
'               rsObjects!SecondX = 0
'               rsObjects!Height = -rsObjects!Width / 2           'Old Curve fixed at aspect ratio 2
'               rsObjects!centerY = (rsObjects!centerY + rsObjects!SecondY) / 2
'               rsObjects!SecondY = 0
'            End If
'            objKey = objKey + 1
'            rsObjects!objKey = objKey
'            rsObjects.Update
'            rsObjects.MoveNext
'         Loop
'         rsObjects.Close
'      End If
'      For Each column In tdfTable.Fields '====================================Delete GenMAPP Column
'         '  Do this last or column positions and names screwed up
'         If column.name = "GenMAPP" Then .Fields.Delete "GenMAPP"
'      Next column
'   End With
'
'   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Adjust Info Table
'   Set tdfTable = dbMapp.TableDefs("Info")
'   With tdfTable
'      ok = False
'      For Each column In .Fields                                                 'Add GeneDB Column
'         If column.name = "GeneDB" Then
'            ok = True
'            Exit For
'         End If
'      Next column
'      If Not ok Then
'         Set column = .CreateField("GeneDB", dbMemo)
'         .Fields.Append column
'         .Fields("GeneDB").OrdinalPosition = 14
'      End If
'   End With
'   dbMapp.Execute "UPDATE Info SET Version = '" & BUILD & "'"
'   dbMapp.Close
   
   Set dbMapp = OpenDatabase(mappPath, False, True)                           'Go back to read-only
   UpdateMAPP = True
End Function
