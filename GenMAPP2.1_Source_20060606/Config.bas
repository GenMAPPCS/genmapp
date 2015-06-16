Attribute VB_Name = "Config"
Option Explicit
Rem ********************************************************************************** Config Items
'  The config file, GenMAPP.cfg, resides in the App folder and contains the following:
'     Don't change any of this stuff; you could really screw something up!!
'     ProgramVersion: 20020112                                                              [BUILD]
'     mruGeneDB: C:\GenMAPP\Gene Databases\MyGenes.gdb
'     mruMAPPPath: C:\GenMAPP\MAPPs\
'     mruDataSet: C:\GenMAPP\Expression Datasets\GoodStuff.gex
'     mruColorSet: 8 wk average
'     mruCatalog: C:\GenMAPP\Catalogs\MyCatalog.gct
'     mruExportPath: C:\GenMAPP\Exports\
'     mruExportSourcePath: C:\GenMAPP\MAPPs\
'     mruImportPath: C:\GenMAPP\Imports\
'     mruMappConvertSource: C:\GenMAPP\MAPPs\
'     mruEDConvertSource: C:\GenMAPP\Expression Datasets\
'     Coloring: S
'     Legend: DFGEF8|
'     Options: C
'     JPEGQuality: 90

'  mru is "Most Recently Used"

Public Const CFG_FILE = "GenMAPP.cfg"

Public cfgBaseFolder As String                              'Original base folder for GenMAPP Stuff
Public mruGeneDB As String                                 'Startup Gene DB for all Drafter windows
   '  Most Recently Used Gene DB. It changes each time a Gene DB is opened in any Drafter
   '  window. It is saved in GenMAPP.cfg when the program is exited.
   '  If no Gene DB has yet been specified, this is set to the default path for Gene DBs. The
   '  opening procedure for Gene DBs will detect this and ask for a specific .gdb file.
Public mruMappPath As String                                                         'MRU MAPP Path
Public mruDataSet As String                                                 'MRU Expression Dataset
   '  This is always the MRU Expression Dataset. It changes each time an Expression Dataset
   '  is opened in any Drafter window. It is saved in GenMAPP.cfg when the program is exited.
   '  If no Expression Dataset has yet been specified, this is set to the default path for
   '  Expression Datasets. The opening procedure for Expression Datasets will detect this and
   '  ask for a specific .gex file.
Public mruColorSet As String                                                         'MRU Color Set
   '  See above
Public mruCatalog As String                                                       'MRU MAPP Catalog
   '  Same as for mruGeneDB
Public mruExportPath As String     'MRU Root path for exported files such as HTMLs, BMP, JPEGs, etc
Public mruExportSourcePath As String              'MRU Root path for MAPP sets, etc, to be exported
Public mruImportPath As String          'MRU Root path for imported files such as csv's, txt's, etc
Public mruMappConvertSource As String, mruEDConvertSource As String
   '  Source folders for conversions to later GenMAPP versions. Destination folders are
   '  mruMappPath and mruDataSet
Public cfgColoring As String                                  'Basis for colors and values of genes
   '  R  All genes related to the object's ID
   '  S  Only genes with the object's specific ID
Public cfgLegend As String
   '  Existence of characters determines what to display
   '     D     Display legend
   '     G     Show Gene Database
   '     E     Show Expression Dataset Name
   '     C     Show Color Set name
   '     V     Show name of expression-value column
   '     R     Show Remarks
   '     L     Show Legend itself (colors and criteria)
   '     Fn|   Legend font size (n)
   '     I     Display Information Area
Public cfgOptions As String
   '  Existence of characters determines options
   '     C     Open program with mruDataSet Expression Dataset opened
Public cfgJPEGQuality As String
Public cfgInitialRun As String
   '  If this item exists, it will be "InitialRun: False" indicating the the program is not
   '  running for the first time.
Public cfgCheckForUpdatesOnStart As String
   '  If this entry is "True", GenMAPP checks for updates at the start of each run. "False" or
   '  empty does nothing.
Public cfgLegendAllColorSets As Boolean
   '  If true, When multiple Color Sets are displayed, shows all the colors in the Legend.
   '  If false, shows only the first color set in the Legend.
   

Sub ReadConfig() '************************************************* Get Basic Info From Config File
   Dim cfgValue As String, cfgItem As String, colon As Integer
   
'MsgBox "Opening config file"
   If Dir(App.path & "\" & CFG_FILE) = "" Then                           'Config File Doesn't Exist
      Open App.path & "\" & CFG_FILE For Output As #31
      Close #31
      UpdateConfig mruMappConvertSource, "C:\"
      UpdateConfig mruEDConvertSource, "C:\"
   End If
      
   Open App.path & "\" & CFG_FILE For Input As #31
   Do Until EOF(31)
      Line Input #31, cfgValue
      colon = InStr(cfgValue, ": ")
      If colon Then
         cfgItem = Left(cfgValue, colon - 1)
         cfgValue = Mid(cfgValue, colon + 2)
         Select Case cfgItem
         Case "ProgramVersion"
         Case "baseFolder"
            cfgBaseFolder = cfgValue
         Case "mruGeneDB"
            mruGeneDB = cfgValue                                      'Start with MRU Gene Database
         Case "mruMAPPPath"
            mruMappPath = cfgValue                                  'In case MAPP name gets in here
         Case "mruDataSet"
            mruDataSet = cfgValue
         Case "mruColorSet"
            mruColorSet = cfgValue
            If mruColorSet <> "" And InStr(mruColorSet, "\") = 0 Then   'Old style, single colorset
               mruColorSet = mruColorSet & "\" & mruColorSet                 'DisplayValue\ColorSet
            End If
         Case "mruCatalog"
            mruCatalog = cfgValue
         Case "mruExportPath"
            mruExportPath = cfgValue
         Case "mruExportSourcePath"
            mruExportSourcePath = cfgValue
         Case "mruImportPath"
            mruImportPath = cfgValue
         Case "mruMappConvertSource"
            mruMappConvertSource = cfgValue
         Case "mruEDConvertSource"
            mruEDConvertSource = cfgValue
'         Case "mruLocalSource"
'            mruLocalSource = cfgValue
         Case "Coloring"
            cfgColoring = cfgValue
         Case "Legend"
            cfgLegend = cfgValue
         Case "Options"
            cfgOptions = cfgValue
         Case "JPEGQuality"
            cfgJPEGQuality = cfgValue
         Case "CheckForUpdatesOnStart"
            cfgCheckForUpdatesOnStart = cfgValue
         Case "LegendColorSets"
         If cfgValue = "All" Then
            cfgLegendAllColorSets = True
         Else
            cfgLegendAllColorSets = False
         End If
         Case "InitialRun"
            cfgInitialRun = cfgValue
         End Select
      End If
   Loop
   Close #31
   If cfgJPEGQuality = "" Then cfgJPEGQuality = "90"
End Sub
   
Sub WriteConfig() '************************************************* Rewrite The Configuration File
   '  Done whenever the program closes
   '  Also done if MAPPFinder called up to make sure it has the latest config info.
   '  All global configuration variables must be active
   
   Open App.path & "\" & CFG_FILE For Output As #31
   Print #31, "Do not change this file, serious GenMAPP errors may occur."
   Print #31, "ProgramVersion: " & BUILD
   Print #31, "baseFolder: " & cfgBaseFolder
   Print #31, "mruGeneDB: " & mruGeneDB
   mruMappPath = GetFolder(mruMappPath)                                   'Always write just folder
   Print #31, "mruMAPPPath: " & mruMappPath
   Print #31, "mruDataSet: " & mruDataSet
   Print #31, "mruColorSet: " & mruColorSet
   Print #31, "mruCatalog: " & mruCatalog
   Print #31, "mruExportPath: " & mruExportPath
   Print #31, "mruExportSourcePath: " & mruExportSourcePath
   Print #31, "mruImportPath: " & mruImportPath
   Print #31, "mruMappConvertSource: " & mruMappConvertSource
   Print #31, "mruEDConvertSource: " & mruEDConvertSource
'   Print #31, "mruLocalSource: " & mruLocalSource
   Print #31, "Coloring: " & cfgColoring
   Print #31, "Legend: " & cfgLegend
   Print #31, "Options: " & cfgOptions
   Print #31, "JPEGQuality: " & cfgJPEGQuality
   Print #31, "CheckForUpdatesOnStart: " & cfgCheckForUpdatesOnStart
   If cfgLegendAllColorSets Then
      Print #31, "LegendColorSets: All"
   Else
      Print #31, "LegendColorSets: First"
   End If
   Print #31, "InitialRun: False"
   Close #31
End Sub

Sub UpdateConfig(item As String, value As String) '************** Update Single Item In Config File
   '  All global configuration variables must be active
   Dim valueWritten As Boolean, Lin As String, cfgItem As String, colon As Integer
   
   If Dir(App.path & "\config.$tm") <> "" Then
      Kill App.path & "\config.$tm"
   End If
   If item = "mruMAPPPath" Then
      value = GetFolder(value)
   End If
   Open App.path & "\" & CFG_FILE For Input As #31
   Open App.path & "\config.$tm" For Output As #33
   Do Until EOF(31)
      Line Input #31, Lin
      colon = InStr(Lin, ": ")
      If colon Then
         cfgItem = Left(Lin, colon - 1)
         If cfgItem = item Then
            Lin = cfgItem & ": " & value
            valueWritten = True
         End If
      End If
      Print #33, Lin
   Loop
   If Not valueWritten Then
      Print #33, item; ": "; value
   End If
   Close #31, #33
   Kill App.path & "\" & CFG_FILE
   Name App.path & "\config.$tm" As App.path & "\" & CFG_FILE
End Sub


