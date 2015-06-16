Attribute VB_Name = "GenMAPP_Convert"
Option Explicit

Rem ***************************************************************************** DataSet Constants
   Public Const MAX_EXP_COLUMNS = 300                                  'Max expression data columns
   Public Const MAX_CRITERIA = 30
   Public Const MAX_COLORSETS = 100

Rem ************************************************************************* File Number Constants
   Public Const FILE_RAW_DATA = 50
   Public Const FILE_EXCEPTIONS = 51
   Public Const FILE_CONVERT_LOG = 52
   Public Const FILE_TREE = 60
   Public Const FILE_TEMP = 99                          'File must be opened and closed immediately
   Public Const F_TEMP = 99                          'File must be opened and closed immediately

Rem ************************************************************************* Misc Global Constants
   Public Const MAX_GENES = 2000           'Maximum number of genes to be considered for one object
      '  Eg: a conversion from SwissProt to GenBank will return a maximum of MAX_GENES results.
   Public Const MAX_SYSTEMS = 100                         'Max number of cataloging systems allowed
   Public Const TITLE_CHAR_LIMIT = 50        'Limit on character count for expression column titles
   Public Const SYSTEM_TITLE_CHAR_LIMIT = 30      'Limit on character count for system table titles
   Public Const CHAR_DATA_LIMIT = 50        'Limit on character count for searchable character data
                                            'This includes unique IDs in Gene Tables
   Public Const MAX_2ND_COLS = 100           'Limit on searchable secondary columns in a Gene Table

Rem ******************************************************************* Conversion Tables Available
   Public modSP As String, modGB As String                'These are actual tables, eg: MGD-GenBank
   Public GBLocusLink As Boolean, GBUniGene As Boolean         'True if the table exists in Gene DB
   Public SPLocusLink As Boolean, SPUniGene As Boolean, GBSwissProt As Boolean
   
Rem ****************************************************************** Other Module-Level Variables
   Public modSys As String, modCode As String

'******************************************************* Find Which Conversion Tables Are Available
Function DetermineMODConversionTables(dbGene As Database) As Boolean
   '  Entry    dbGene      Current open Gene Database.
   '  Return   True if conversion possible.
   '           Sets module-level variables to use in other parts of the conversion.
   '           These show MOD parameters and the various conversion tables available.
   '              ModGB, ModSP   The MOD-GenBank and MOD-SwissProt tables
   '              GBSwissProt, GBLocusLink, GBUniGene, SPLocusLink, SPUniGene
   '                 These are true if these tables exist in the Gene Database
   '                 If modGB is SwissProt-GenBank (as for Human) then GBSwissProt is False
   '              modSys      MOD System. Eg: "SwissProt" (Human)
   '              modCode     System code for MOD, Eg: "S" for SwissProt (Human)

   Dim rsSystems As Recordset, tablesExist As Boolean, rsInfo As Recordset
   Dim tdf As TableDef
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Start As If Conversion Not Possible
   modGB = ""
   modSP = ""
   GBSwissProt = False
   GBLocusLink = False
   GBUniGene = False
   SPLocusLink = False
   SPUniGene = False
   modCode = ""
   modSys = ""
   
   If dbGene Is Nothing Then
      MsgBox "No Gene Database. Choose Gene Database before attempting conversions.", _
             vbOKOnly + vbExclamation, "Gene ID Conversion"
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   
   
   Set rsInfo = dbGene.OpenRecordset("SELECT MODSystem FROM Info")
   If Dat(rsInfo!MODSystem) = "" Then                                       'There is no MOD system
      MsgBox "Gene Database has no designated Model Organism table. Cannot perform " _
             & "conversions.", vbOKOnly + vbExclamation, "Gene ID Conversion"
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   modSys = rsInfo!MODSystem
   Set rsSystems = dbGene.OpenRecordset( _
         "SELECT SystemCode FROM Systems WHERE System = '" & modSys & "'")
   modCode = rsSystems!systemCode
   
   For Each tdf In dbGene.TableDefs
      Select Case tdf.name
      Case modSys & "-GenBank"
         modGB = tdf.name
         tablesExist = True
      Case "SwissProt-GenBank"
         If modGB <> "SwissProt-GenBank" Then
            '  If Human, SwissProt-GenBank is the MOD table, don't list it again
            GBSwissProt = True
            tablesExist = True
         End If
      Case "LocusLink-GenBank"
         GBLocusLink = True
         tablesExist = True
      Case "UniGene-GenBank"
         GBUniGene = True
         tablesExist = True
      Case modSys & "-SwissProt"
         modSP = tdf.name
         tablesExist = True
      Case "SwissProt-LocusLink"
         SPLocusLink = True
         tablesExist = True
      Case "SwissProt-UniGene"
         SPUniGene = True
         tablesExist = True
      End Select
   Next tdf
   If tablesExist Then
      DetermineMODConversionTables = True
   Else
      MsgBox "This Gene Database has no conversion tables for GenBank or SwissProt. Choose " _
             & "another Gene Database.", vbExclamation + vbOKOnly, "Gene ID Conversion"
   End If
End Function

Sub AllRelatedGenes(ByVal idIn As String, systemIn As String, dbGene As Database, _
                    genes As Integer, geneIDs() As String, geneFound As Boolean, _
                    Optional supportedSystem As Boolean = False, _
                    Optional systemsList As Variant)
   '  Entry:
   '     idIn           Gene identification received (may have to search to find primary)
   '     systemIn       Cataloging system code for passed idIn
   '     dbGene         Gene Database for this query (the Gene Database for the particular
   '                    drafter window)
   '     geneFound      If True and Specific Gene option checked, don't bother searching for the
   '                    gene. The gene is just being matched to an Expression Dataset.
   '     systemsList    List of system codes that appear in this dataset. Only passed if coloring
   '                    genes to eliminate looking for IDs in systems that don't appear in the
   '                    dataset. If not passed, defaults to "ALL".
   '                    If "EXISTS" then routine being used to find only the existence of the
   '                    specific gene, not all related genes. As soon as a gene is found,
   '                    we can exit the routine. The only return needed is genes(x, 2), where
   '                    x is the last gene found, to test to see whether the gene was found in
   '                    a [P]rimary or [S]econdary column. Use last gene because gene passed
   '                    (genes(0, x)) may not be found but leads to relational tables.
   '                    Used in converting Expression Datasets.
   '  Return:
   '     genes                   Number of related genes (counting from 1)
   '                             This is also used as the index for the geneIDs array. It is zero
   '                             based, so at exit, genes is increased by one.
   '     geneIDs(MAX_GENES, 2)   Gene ID for each related gene. Primary ones listed first
   '                             geneIDs(x, 0) ID
   '                             geneIDs(x, 1) SystemCode
   '                             geneIDs(x, 2) "P" for primary ID (ID column)
   '                                           "S" or "s" for secondary ID
   '                                           (eg: Accession in SwissProt)
   '                             The first gene is always the one passed to the procedure whether
   '                             it is found in relational tables or not.
   '     geneFound               True if gene found in any primary or secondary column or
   '                             relational tables.
   '     supportedSystem         systemIn is a system supported in the Gene Database. If false and
   '                             genes > 0 then search has looked in Related column of
   '                             Relational tables and found gene.
   '  The calling function can compare idIn with geneIDs(0, 0) to see if the passed gene ID exists
   '  in the database. This function will find related genes even if the passed gene does not exist
   '  in a supported system but only in a relational table.
   'For AllRelatedGenes()
   '   Dim genes as integer
   '   Dim geneIDs(MAX_GENES, 2) As String
   '   Dim geneFound as boolean
   '   'Dim supportedSystem as Boolean                'System supported in Gene Database [optional]
   '   'Dim systemsList As Variant                                    'Systems to search [optional]
   '  Call:
   '     AllRelatedGenes idIn, systemIn, dbGene, genes, geneIDs, geneFound, _
                         [supportedSystem], [systemsList]
   
   '  At this point related genes are found only if the systemIn is a supported system. We can
   '  find genes if the systemIn is represented in the Related column of a Relational table but
   '  handling that is not been defined yet. This sub should set supportedSystem = False,
   '  no matter what we do with related genes found.

   Dim index As Integer, lastIndex As Integer
   Dim primaryIDs(MAX_GENES) As String                          'Gene IDs to search relationals for
      '  Primary IDs are those IDs in the systemIn that are used to search relational tables.
      '  A Primary ID is added to Genes() returned if it is found in the systemIn. (And if the
      '  systemIn is in the SystemsList, i.e. represented in the Expression Dataset.)
      '  The search for Primary IDs is in both the ID and Secondary ID columns of the systemIn.
      '  It is possible that more than one Primary ID may be found if the secondary ID shows
      '  up in more than one row.
   Dim primaryIndex As Integer                                                          'Zero based
   Dim firstPrimaryIndex As Integer
   Dim lastPrimaryIndex As Integer
   '  Last index for Primary IDs. May be 0 if idIn not in primary system.
   '  This might be more than zero only if secondary IDs led back to more than one primary ID.
   '  In other words, Secondary ID X1234 occurred in two rows. Not sure this is possible.
   Dim primaryIDReplaced As Integer                                                 'Either 0 or -1
   Dim system As String                            'Cataloging system code currently being examined
   Dim column As String                                                       'Any secondary column
   Dim rsSystems As Recordset                                                    'The Systems table
   Dim rsSystem As Recordset                              'A cataloging-system table, eg: SwissProt
   Dim rsRelations As Recordset                                                'The Relations table
   Dim rsRelational As Recordset                         'A relational table, eg: SwissProt-GenBank
   Dim pipe As Integer, slash As Integer
   Dim sql As String
   Dim searchSeconds As Boolean   'True if doing a GeneFinder or Backpage search or if coloring
                                  'and Expression Dataset has secondary IDs in it. If ED has
                                  'them, the Info table SystemCodes column will contain "|~|".
   Dim secondaryCols(MAX_2ND_COLS, 1) As String
   '           secondaryCols(x, 0)     Names of secondary columns
   '           secondaryCols(x, 1)     "S" if multiple, pipe-surrounded IDs allowed
   '                                   "s" if single, non-pipe-surrounded IDs allowed
   Dim lastSecondCol As Integer
   Dim singleGene As Boolean                   'True if systemList comes in as "EXISTS". See above.
   Dim currentMousePointer As Integer                  'Form's MousePointer on entry. Reset on exit
 
   If dbGene Is Nothing Then Exit Sub                      'No database >>>>>>>>>>>>>>>>>>>>>>>>>>>
   If idIn = "" Then Exit Sub                              'Gene not identified >>>>>>>>>>>>>>>>>>>
   
   currentMousePointer = Screen.ActiveForm.MousePointer
   Screen.ActiveForm.MousePointer = vbHourglass
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Assign Gene Passed As Gene #1
   genes = 0                                                              'Zero based at this point
   geneIDs(genes, 0) = idIn
   geneIDs(genes, 1) = systemIn
   geneIDs(genes, 2) = "P"                                                   'Default to primary ID
   
   If geneFound And InStr(cfgColoring, "S") Then '++++++++++++++++++++++ Just Return Gene Passed In
      '  User has chosen the Specific Gene Option and is just matching to an Expression Dataset
      '  rather than trying to find a gene in GeneFinder or creating a Backpage. In this case,
      '  who cares whether the gene is in a supported system, found in a relational system or
      '  exists anywhere.
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   If IsMissing(systemsList) Then systemsList = "ALL"
   If VarType(systemsList) = vbNull Then systemsList = "ALL"
   If systemsList = "" Then systemsList = "ALL"
   If systemsList = "EXISTS" Then
      systemsList = "ALL"
      singleGene = True
   End If
   
   geneFound = False
   supportedSystem = False
   searchSeconds = True
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ See If Supported System
'      '  Gene in (idIn and systemIn) is classified as:
'      '     Supported   Its system exists as one of the system tables in the Gene DB
'      '                 supportedSystem = True
'      '     Relational  Its system not supported but exists in a relational table
'      '                 supportedSystem = False
'      '     Neither     Exits sub with only the passed gene returned
   Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems WHERE SystemCode = '" & systemIn & "'", _
                   dbOpenForwardOnly)                                    'Get the system table name
                                                 'Eg: SELECT * FROM Systems WHERE SystemCode = 'Rm'
   If Not rsSystems.EOF Then '========================================================System Exists
      If Dat(rsSystems![Date]) <> "" Or systemIn = "O" Then        'Date or Other, supported system
         '  Other always supported. See comments at beginning of sub
         supportedSystem = True
      End If
   End If
   
   If supportedSystem Then '++++++++++++++++++++++++++++++++++++++++++++++++ Look For Specific Gene
      '  Specific gene can only be in a supported system
      Set rsSystem = dbGene.OpenRecordset( _
                     "SELECT * FROM " & rsSystems!system & _
                     "   WHERE ID = '" & idIn & "'", _
                     dbOpenForwardOnly)        'Eg: SELECT * FROM SwissProt WHERE ID = 'CALM_HUMAN'
      If Not rsSystem.EOF Then                                             'Found idIn in ID column
         geneFound = True
         If singleGene Then GoTo ExitSub                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
   End If

'  Per Kam, the only time Specific Gene option has effect is in coloring. GeneFinder and
'  Backpages will always search for related genes.
'   If InStr(cfgColoring, "S") Then '++++++++++++++++++++++++++++++++++++++++ "Specific Gene" Option
'      '  If user Options specify "Specific Gene", no relations are searched for nor are
'      '  secondary columns
'      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'   End If
   
   lastSecondCol = SecondCols(rsSystems, secondaryCols)                     'Find Secondary Columns
   If supportedSystem And Not geneFound Then '++++++++++++++++++ Look For Gene In Secondary Columns
      '  If the received idIn is not in the ID column of the gene system table (systemIn)
      '  go to the secondary columns and search
      '  Typical column listing in Systems!Columns
      '     ID|Accession\SBF|Nicknames\sF|Protein|Functions\B|
      For i = 0 To lastSecondCol - 1 '========================================Each Secondary Column
         If secondaryCols(i, 1) = "S" Then                           'Multiple-ID secondary columns
            sql = "SELECT ID" & _
                  "   FROM " & rsSystems!system & " " & _
                  "   WHERE [" & secondaryCols(i, 0) & "] LIKE '*|" & idIn & "|*' ORDER BY ID"
               'Eg: SELECT ID FROM SwissProt WHERE Accession LIKE '*|A1234|*'
         Else                                                              'Single ID without pipes
            sql = "SELECT ID" & _
                  "   FROM " & rsSystems!system & " " & _
                  "   WHERE [" & secondaryCols(i, 0) & "] = '" & idIn & "' ORDER BY ID"
               'Eg: SELECT ID FROM SGD WHERE Gene = 'TFC3'
         End If
         Set rsSystem = dbGene.OpenRecordset(sql, dbOpenForwardOnly)
         If Not rsSystem.EOF Then '----------------------------------Gene Found In Secondary Column
            geneFound = True
            primaryIndex = 1                                                    'First primary gene
            geneIDs(0, 2) = "S"             'First gene actually in a secondary column. Change code
            If singleGene Then GoTo ExitSub                'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
               '  For singleGene, must exit here because converting Expression Datasets checks
               '  last gene found to determine if [P]rimary or [S]econdary.
            Do Until rsSystem.EOF                               'All rows where idIn in this column
               '  Add primary ID for each instance of the secondary ID in the system.
               '  This should not add another instance of idIn because that was never found
               '  in the first place
'               If systemsList = "ALL" Or InStr(systemsList, "|" & systemIn & "|") <> 0 Then
                  '  Either looking at all systems or gene system is in the
                  '  Expression Dataset (for coloring), add to returned related genes.
                  '  Add the primary ID to the list.
                  If genes < MAX_GENES Then
                     genes = genes + 1
                     geneIDs(genes, 0) = rsSystem!id
                     geneIDs(genes, 1) = systemIn
                     geneIDs(genes, 2) = "P"
                  End If
'               End If
               rsSystem.MoveNext
            Loop
         End If
      Next i
   End If
   firstPrimaryIndex = primaryIndex
   lastPrimaryIndex = genes
   '  At this point we have created a list of primary IDs in the systemIn.
   '  geneIDs() up to lastPrimaryIndex has all the IDs we want to search for.
      
   If supportedSystem And Not singleGene Then '+++++++++++++++++++++ Find Secondary IDs In systemIn
      '  If singleGene, don't do this because this routine only finds secondary columns for genes
      '  that have been found above and secondary IDs are not listed in the relational tables.
      For primaryIndex = firstPrimaryIndex To lastPrimaryIndex
         '  GeneIDs(0, 0) might not be a primary ID
         '  Uses the SecondCols return from above because only dealing with systemIn at this point
         For i = 0 To lastSecondCol - 1 '=====================================Each Secondary Column
            sql = "SELECT [" & secondaryCols(i, 0) & "] AS Secondary" & _
                  "   FROM " & rsSystems!system & _
                  "   WHERE ID = '" & geneIDs(primaryIndex, 0) & "'" & _
                  "   ORDER BY [" & secondaryCols(i, 0) & "]"
               'Eg: SELECT Accession FROM SwissProt WHERE ID = 'CALM_HUMAN'
            Set rsSystem = dbGene.OpenRecordset(sql, dbOpenForwardOnly)
            If Not rsSystem.EOF Then '------------------------------------Found In Secondary Column
               If secondaryCols(i, 1) = "S" Then '_________________Multiple IDs In Secondary Column
                  '  Might return something like "|A1234|B5678|"
                  Dim nextPipe As Integer, strOut As String
                  
                  pipe = 1
                  Do While pipe < Len(rsSystem!Secondary)
                     nextPipe = InStr(pipe + 1, rsSystem!Secondary, "|")
                     If nextPipe = 0 Then nextPipe = Len(rsSystem!Secondary) + 1
                     If Mid(rsSystem!Secondary, pipe + 1, nextPipe - pipe - 1) _
                           <> geneIDs(0, 0) Then 'Don't repeat first gene
                        genes = genes + 1
                        geneIDs(genes, 0) = Mid(rsSystem!Secondary, pipe + 1, nextPipe - pipe - 1)
                        geneIDs(genes, 1) = systemIn
                        geneIDs(genes, 2) = "S"
                     End If
                     pipe = nextPipe
                  Loop
               Else '_________________________________________________Single ID In Secondary Column
                  If rsSystem!Secondary <> geneIDs(0, 0) Then              'Don't repeat first gene
                     genes = genes + 1
                     geneIDs(genes, 0) = rsSystem!Secondary
                     geneIDs(genes, 1) = systemIn
                     geneIDs(genes, 2) = "S"
                  End If
               End If
            End If
         Next i
      Next primaryIndex
   End If
   '  We now have all the secondary IDs for the primary ones
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Search "Other" Table
   '  Finds any ID in the Other table whether or not it is found in any other Systems table.
   '  Will not find any relations (unless the gene was coded as Other to begin with).
   Set rsSystem = dbGene.OpenRecordset( _
         "SELECT * FROM Other WHERE ID = '" & idIn & "' AND SystemCode = '" & systemIn & "'", _
         dbOpenForwardOnly)        'Eg: SELECT * FROM Other WHERE ID = '12345' AND SystemCode = 'G'
   Do Until rsSystem.EOF                                        'All rows where idIn in this column
      If geneFound Then                                      'Found in SystemIn, add to found genes
         genes = genes + 1
         geneIDs(genes, 0) = rsSystem!id
         geneIDs(genes, 1) = "O"
         geneIDs(genes, 2) = "P"
      Else                                    'Not found in SystemIn, change systemIn to "O" system
         geneFound = True
         systemIn = "O"
         genes = 0
         geneIDs(genes, 0) = rsSystem!id
         geneIDs(genes, 1) = systemIn
         geneIDs(genes, 2) = "P"
      End If
      rsSystem.MoveNext
   Loop

   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find All Applicable Relational Tables
   If systemsList <> "ALL" Then '===========================Only System Codes In Expression Dataset
      '  Look at only those relational tables where the SystemCode is represented in the
      '  Expression Dataset
      '  Matches Expression data from supported and nonsupported systems
      Set rsRelations = dbGene.OpenRecordset( _
                        "SELECT * FROM Relations" & _
                        "   WHERE SystemCode = '" & systemIn & "' " & _
                        "         AND instr('" & systemsList & "', relatedCode )" & _
                        "      OR RelatedCode = '" & systemIn & "' " & _
                        "         AND instr('" & systemsList & "', systemCode)" & _
                        "   ORDER BY Relation NOT LIKE '*GenBank*'")
                        'This query selects all relational tables, supported or not. GenBanks
                        'always listed last.
         '  Eg: SELECT * FROM Relations
         '         WHERE SystemCode = 'G'
         '               AND instr('|G|O|', 'O')                             'In Expression Dataset
         '            OR RelatedCode = 'G')
         '               AND instr('|G|O|', 'G')
   Else '==========================================================================Any System Codes
      Set rsRelations = dbGene.OpenRecordset( _
                        "SELECT * FROM Relations" & _
                        "   WHERE (SystemCode = '" & systemIn & "'" & _
                        "          OR RelatedCode = '" & systemIn & "')" & _
                        "   ORDER BY Relation NOT LIKE '*GenBank*'")
                        'This query selects all relational tables, supported or not. GenBanks
                        'always listed last.
   End If
   '  At this point, rsRelations has all the relational tables we want to look at
            
   Do Until rsRelations.EOF '+++++++++++++++++++++++++++++++++++++++++ Search All Relational Tables
'Debug.Print rsRelations!Relation
      '======================================================================Find Secondary Columns
      If rsRelations!systemCode = systemIn Then '------------------------SystemIn In Primary Column
         '  Look for the system referenced in the related (the other) column
         Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems WHERE SystemCode = '" & rsRelations!relatedCode & "' AND [Date] IS NOT NULL", _
                   dbOpenForwardOnly)
      Else '-------------------------------------------------------------SystemIn In Related Column
         '  Look for the system referenced in the system (the other) column
         Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems WHERE SystemCode = '" & rsRelations!systemCode & "' AND [Date] IS NOT NULL", _
                   dbOpenForwardOnly)
      End If
      lastSecondCol = SecondCols(rsSystems, secondaryCols)
      
      For primaryIndex = firstPrimaryIndex To lastPrimaryIndex '====================Each Primary ID
         '  For singleGene, there should be only one Primary ID, the gene passed in. We did
         '  not search for secondaries for the idIn, and to reach here idIn must not have
         '  been found.
         If rsRelations!systemCode = systemIn Then '---------------------SystemIn In Primary Column
            sql = "SELECT * FROM [" & rsRelations!Relation & "]" & _
                 "   WHERE [Primary] = '" & geneIDs(primaryIndex, 0) & "' ORDER BY Related"
                        'Eg: SELECT * FROM [SwissProt-GenBank] WHERE Primary = 'CALM_HUMAN'
            Set rsRelational = dbGene.OpenRecordset(sql)
               '  For some stupid reason Windows does not like SwissProt-Affy Primary 'Q61018'.
               '  It occurred when the SwissProt-Affy table was blank. It does not occur in the new
               '  database with UniProt-Affy, with data in the table.
            Do Until rsRelational.EOF
               If genes < MAX_GENES - 1 Then                          'Anything more we just forget
                  geneFound = True
                  genes = genes + 1
                  geneIDs(genes, 0) = rsRelational!related
                  geneIDs(genes, 1) = rsRelations!relatedCode
                  geneIDs(genes, 2) = "P"
                  If singleGene Then GoTo ExitSub          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                     '  For singleGene, must exit here because converting Expression Datasets
                     '  checks last gene found to determine if [P]rimary or [S]econdary.
                  If searchSeconds Then
                     AddSecondIDs secondaryCols, lastSecondCol, geneIDs, genes, rsSystems, dbGene
                  End If
               End If
               rsRelational.MoveNext
            Loop
         Else '----------------------------------------------------------SystemIn In Related Column
            Set rsRelational = dbGene.OpenRecordset( _
                              "SELECT * FROM [" & rsRelations!Relation & "]" & _
                              "   WHERE Related = '" & geneIDs(primaryIndex, 0) & "'" & _
                              "   ORDER BY [Primary]")
                              'Eg: SELECT * FROM [SwissProt-GenBank] WHERE Related = 'X1234'
            Do Until rsRelational.EOF
               If genes < MAX_GENES - 1 Then                          'Anything more we just forget
                  geneFound = True
                  genes = genes + 1
                  geneIDs(genes, 0) = rsRelational!primary
                  geneIDs(genes, 1) = rsRelations!systemCode
                  geneIDs(genes, 2) = "P"
                  If singleGene Then GoTo ExitSub          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                     '  For singleGene, must exit here because converting Expression Datasets
                     '  checks last gene found to determine if [P]rimary or [S]econdary.
                  If searchSeconds Then
                     AddSecondIDs secondaryCols, lastSecondCol, geneIDs, genes, rsSystems, dbGene
                  End If
               End If
               rsRelational.MoveNext
            Loop
         End If
      Next primaryIndex
      rsRelations.MoveNext
   Loop
   
ExitSub:
   genes = genes + 1        'genes was zero based, not it should be count of actual number of genes
   Screen.ActiveForm.MousePointer = currentMousePointer

'   For i = 0 To genes - 1
'      Debug.Print geneIDs(i, 1); " "; geneIDs(i, 0)
'   Next i
End Sub
Function UpdateSingleMAPP(mappPath As String) As Boolean '********** Update MAPP To Current Version
   Dim rsInfo As Recordset, rsObjects As Recordset, dbMapp As Database
   Dim tdfTable As TableDef
   Dim column As Field
   Dim idxIndex As index
   Dim objKey As Long
   Dim ok As Boolean
   
   Screen.ActiveForm.MousePointer = vbHourglass
   Set dbMapp = OpenDatabase(mappPath)                                    'Reopen without read-only
    
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Adjust Objects Table
   Set tdfTable = dbMapp.TableDefs("Objects")
   With tdfTable
      For Each column In tdfTable.Fields '======================================Change column names
         If column.name = "Primary" Then column.name = "ID"
         If column.name = "SystemCode" Then                 'Consistently change column to old name
            column.name = "PrimaryType"                     'To allow for various versions
         End If
      Next column
      Set column = .CreateField("SystemCode", dbText, 2)
      .Fields.Append column
      dbMapp.Execute "Update Objects SET SystemCode = PrimaryType"
      dbMapp.Execute "ALTER TABLE Objects DROP COLUMN PrimaryType"
      .Fields("SystemCode").AllowZeroLength = True
      .Fields("SystemCode").OrdinalPosition = 2
      ok = False '==================================================================Add Notes Field
      For Each column In tdfTable.Fields
         If column.name = "Notes" Then
            ok = True
            Exit For
         End If
      Next column
      If Not ok Then
         Set column = .CreateField("Notes", dbMemo)
         .Fields.Append column
      End If
      .Fields("Notes").AllowZeroLength = True
      ok = False '=================================================================Add ObjKey Field
      For Each column In tdfTable.Fields
         If column.name = "ObjKey" Then
            ok = True
            Exit For
         End If
      Next column
      If Not ok Then
         Set column = .CreateField("ObjKey", dbLong)
            '  Can't make this a primary key because Jet orders the table that way. Open
            '  creates the graphic in the order it encounters the objects in the file. Lines,
            '  for example, must come before genes to appear behind them.
'         .Fields("ObjKey").OrdinalPosition = 0
         column.OrdinalPosition = 0
         .Fields.Append column
         Set idxIndex = .CreateIndex("ixObjKey")
         idxIndex.Fields.Append .CreateField("ObjKey")                         'Why does this work?
         .Indexes.Append idxIndex
      End If
      Set rsObjects = dbMapp.OpenRecordset("SELECT MAX(ObjKey) AS objectKey FROM Objects")
         '  Previous MAPPs may have field but not object keys, so the return here will be NULL
      If VarType(rsObjects!objectKey) = vbNull Then
         objKey = 0 '===============================================================Add Object Keys
         Set rsObjects = dbMapp.OpenRecordset("Select * FROM Objects")
         If UCase(App.EXEName) = "CONVERT" Then
            History "Updating " & FileAbbrev(mappPath, 50)
            rsObjects.MoveLast
            SetProgressBase rsObjects.recordCount, "records"
            rsObjects.MoveFirst
         End If
         Do Until rsObjects.EOF
            If UCase(App.EXEName) = "CONVERT" Then History , rsObjects.AbsolutePosition
            rsObjects.edit
            If rsObjects!Type = "Curve" Then
               '  Only old MAPPs will have the Curve object instead of Arc
               rsObjects!Type = "Arc"
               If rsObjects!SecondX = rsObjects!centerX Then  '------------------Calculate Rotation
                  rsObjects!Width = Abs(rsObjects!centerY - rsObjects!SecondY) / 2
                  If rsObjects!SecondY >= rsObjects!centerY Then
                     rsObjects!rotation = 0.5 * PI                                      '90 degrees
                  Else
                     rsObjects!rotation = 1.5 * PI                                     '270 degrees
                  End If
               Else
                  rsObjects!Width = Abs(rsObjects!centerX - rsObjects!SecondX) / 2
                  If rsObjects!SecondX >= rsObjects!centerX Then
                     rsObjects!rotation = 0                                              '0 degrees
'                     rsObjects!Height = rsObjects!Width / 2      'Old Curve fixed at aspect ratio 2
                  Else
                     rsObjects!rotation = PI                                           '180 degrees
'                     rsObjects!Height = rsObjects!Width / 2      'Old Curve fixed at aspect ratio 2
                  End If
               End If
               rsObjects!Height = rsObjects!Width / 2            'Old Curve fixed at aspect ratio 2
'               rsObjects!Width = (rsObjects!centerX - rsObjects!SecondX) / 2
               rsObjects!centerX = (rsObjects!centerX + rsObjects!SecondX) / 2
               rsObjects!SecondX = 0
'               rsObjects!Height = -rsObjects!Width / 2           'Old Curve fixed at aspect ratio 2
               rsObjects!centerY = (rsObjects!centerY + rsObjects!SecondY) / 2
               rsObjects!SecondY = 0
            End If
            objKey = objKey + 1
            rsObjects!objKey = objKey
            rsObjects.Update
            rsObjects.MoveNext
         Loop
         rsObjects.Close
         If UCase(App.EXEName) = "CONVERT" Then
            SetProgressBase
            History
         End If
      End If
      For Each column In tdfTable.Fields '====================================Delete GenMAPP Column
         '  Do this last or column positions and names screwed up
         If column.name = "GenMAPP" Then .Fields.Delete "GenMAPP"
      Next column
   End With
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Adjust Info Table
   Set tdfTable = dbMapp.TableDefs("Info")
   With tdfTable
      ok = False
      For Each column In .Fields '================================================Add GeneDB Column
         If column.name = "GeneDB" Then
            ok = True
            Exit For
         End If
      Next column
      If Not ok Then
         Set column = .CreateField("GeneDB", dbMemo)
         .Fields.Append column
         .Fields("GeneDB").OrdinalPosition = 3
      End If
      ok = False
      For Each column In .Fields '=========================================Add GeneDBVersion Column
         If column.name = "GeneDBVersion" Then
            ok = True
            Exit For
         End If
      Next column
      If Not ok Then
         Set column = .CreateField("GeneDBVersion", dbText, 10)
         .Fields.Append column
         .Fields("GeneDBVersion").OrdinalPosition = 4
      End If
      For Each column In .Fields '=========================================Delete Expression Column
         '  Do this last or column positions and names screwed up
         If column.name = "Expression" Then .Fields.Delete "Expression"
      Next column
      For Each column In .Fields '===========================================Delete ColorSet Column
         '  Must do these separately or the Delete unsets the Field object
         If column.name = "ColorSet" Then .Fields.Delete "ColorSet"
      Next column
   End With
   dbMapp.Execute "UPDATE Info SET Version = '" & BUILD & "'"
   dbMapp.Close
   
   UpdateConfig "mruMAPPPath", GetFolder(mappPath)
   
   Screen.ActiveForm.MousePointer = vbDefault
   UpdateSingleMAPP = True
End Function
'Function SingleMAPPtoMOD(mappPath As String, dbGene As Database, tiles As Boolean) As Boolean
''Will probably do away with this procedure, using ConvertGenBanksInFile instead.
'   '  Entry    mappPath    Full path to MAPP being converted
'   '           dbGene      Current Gene DB. Always from frmConvert at this point
'   '  Return   True if successful
'   '           Tiles       True if gene objects tiled, ie, more than one substitution found
'   '                       Must be initialized to false elsewhere because if working with a set
'   '                       of MAPPs, a previous MAPP may be tiled and we want the warning to
'   '                       be displayed.
'
'   Dim dbMapp As Database, rsObjects As Recordset, rsNew As Recordset, rsSwissProt As Recordset
'   Dim rsSystems As Recordset, rsInfo As Recordset
'   Dim rs As Recordset
'   Dim changes As Integer                                               'Number of GenBanks changed
'   Dim changeLog As String                                                'File with changes logged
'   Dim modSys As String, modCode As String
'   Dim tdf As TableDef
'   Dim otherPossibilities As String, conversionFound As Boolean
'   Dim sql As String, newID As String
'   Dim id As String, systemCode As String
'   Dim objectKey As Long, geneRecords As Long
'   Dim newIDs(MAX_GENES, 2) As String, noOfNewIDs As Integer     'Substitute IDs for passed GenBank
'   '  newIDs(x, 0)   ID code
'   '  newIDs(x, 1)   System code. Eg: "L"
'   '  newIDs(x, 2)   System Name. Eg: "LocusLink"
'
'   Screen.ActiveForm.MousePointer = vbHourglass
'
'   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find MOD System
'   Set rsInfo = dbGene.OpenRecordset("SELECT MODSystem FROM Info")
'   If Dat(rsInfo!MODSystem) = "" Then                                       'There is no MOD system
'      SingleMAPPtoMOD = False
'      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'   End If
'   modSys = rsInfo!MODSystem
'   Set rsSystems = dbGene.OpenRecordset( _
'         "SELECT System, SystemCode, Species FROM Systems WHERE System = '" & modSys & "'")
'   modCode = rsSystems!systemCode
'
'   Set dbMapp = OpenDatabase(mappPath) '+++++++++++++++++++++++++++++++++++++++++++++++++ Open MAPP
'   If dbMapp Is Nothing Then                                                       'Can't open MAPP
'      SingleMAPPtoMOD = False
'      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'   End If
'
'   changeLog = Left(mappPath, InStrRev(mappPath, ".") - 1) & ".log" '++++++++++++ Set Up Change Log
'   Open changeLog For Output As #FILE_CONVERT_LOG
'   Print #FILE_CONVERT_LOG, "Gene Label"; vbTab; "Old ID"; vbTab; "Old System"; vbTab; _
'         "New ID"; vbTab; " New System"; vbTab; "Other Possibilities: ID[System]"
'
'   DetermineMODConversionTables dbGene '++++++++++++++++++++++++++++++++++ Find Relationship Tables
'
'   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Convert MAPP Records
'   Set rs = dbMapp.OpenRecordset("SELECT MAX(ObjKey) AS MaxKey FROM Objects")
'   objectKey = rs!maxKey      'Last object key. Determined to allow adding of tiled genes for dupes
'   Set rsObjects = dbMapp.OpenRecordset( _
'         "SELECT * FROM Objects WHERE Type = 'Gene' AND SystemCode IN ('G', 'S')")
'   If UCase(App.EXEName) = "CONVERT" Then
'      History "Converting to MOD " & FileAbbrev(mappPath, 50)
'      If rsObjects.EOF Then
'         MsgBox "Nothing to convert.", vbInformation + vbOKOnly, "Converting to MOD"
'      Else
'         rsObjects.MoveLast
'         geneRecords = rsObjects.RecordCount 'Go only this far through the database rather than to
'                                             'EOF because tiled objects may be added
'         SetProgressBase geneRecords, "records"
'         rsObjects.MoveFirst
'      End If
'   End If
'   With rsObjects
''      Do While .AbsolutePosition <= geneRecords - 1 '++++++++++++++++++++++++++++++ Each Gene Object
'      Do Until .EOF '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Each MAPP Object
''If !Label = "SOD2" Then Stop
'         If UCase(App.EXEName) = "CONVERT" Then History , rsObjects.AbsolutePosition
'         conversionFound = False
'         otherPossibilities = ""
'         id = !id
'         systemCode = !systemCode
'         Print #FILE_CONVERT_LOG, !Label; vbTab; !id; vbTab;
'         If systemCode = "G" Then '=================================================Convert GenBank
'            noOfNewIDs = ConvertGenBank(dbGene, id, newIDs())
'            Print #FILE_CONVERT_LOG, "GenBank"; vbTab;
'            If modGB <> "" Then '-----------------------------------Look For MOD-GenBank Conversion
'               Set rsNew = dbGene.OpenRecordset( _
'                     "SELECT [Primary] FROM [" & modGB & "] WHERE Related = '" & id & "'")
'               If Not rsNew.EOF Then '_________________________________________Found In MOD-GenBank
'                  '  Take the first MOD entry for the given GenBank.
'                  .Edit
'                  !id = rsNew![Primary]
'                  !systemCode = modCode
'                  .Update
'                  Print #FILE_CONVERT_LOG, rsNew![Primary]; vbTab; modSys; vbTab;
'                  conversionFound = True
'                  changes = changes + 1
'                  rsNew.MoveNext
'                  Do Until rsNew.EOF '...........................................Pick Up Any Extras
'                     '  If there are any other MOD entries for the given GenBank, put them in
'                     '  the conversion log as Other Possibilities and show them on the MAPP
'                     '  tiled at 50 pixel offsets from the first conversion.
'                     otherPossibilities = otherPossibilities & rsNew![Primary] _
'                                        & "[" & modSys & "] "
'                     objectKey = objectKey + 1
'                     Dim newCenterX As Single, newCenterY As Single
'                     Dim newWidth As Single, newHeight As Single
'                     Dim newLabel As Variant, newHead As Variant, newRemarks As Variant
'                     Dim newLinks As Variant, newNotes As Variant
'                     newCenterX = rsObjects!centerX + 50
'                     newCenterY = rsObjects!centerY + 50
'                     newWidth = rsObjects!Width
'                     newHeight = rsObjects!Height
'                     newLabel = rsObjects!Label
'                     newHead = rsObjects!head
'                     newRemarks = rsObjects!remarks
'                     newLinks = rsObjects!links
'                     newNotes = rsObjects!notes
'                     .AddNew         'Do this instead of INSERT so that the rsNew loop doesn't lose
'                                     'its place
'                     !objKey = objectKey
'                     !id = rsNew![Primary]
'                     !systemCode = modCode
'                     !Type = "Gene"
'                     !centerX = newCenterX + 50
'                     !centerY = newCenterY + 50
'                     !Width = newWidth
'                     !Height = newHeight
'                     !Label = newLabel
'                     !head = newHead
'                     !remarks = newRemarks
'                     !links = newLinks
'                     !notes = newNotes
'                     .Update
'                     tiles = True
''                     sql = "INSERT INTO Objects (ObjKey, ID, SystemCode, Type, CenterX, CenterY, Width, Height, Label, Head, Remarks, Links, Notes)"
''                     sql = sql & " VALUES (" & objectKey & ", '" & rsNew![Primary] & "', '" & modCode & "', 'Gene', " & rsObjects!CenterX + 50 & ", " & rsObjects!CenterY + 50 & ", " & rsObjects!Width & ", " & rsObjects!Height & ", '" & rsObjects!Label & "', '" & rsObjects!Head & "', '" & rsObjects!Remarks & "', '" & rsObjects!Links & "', '" & rsObjects!Notes & "')"
''                     dbMapp.Execute sql
'                     rsNew.MoveNext
'                  Loop
'               End If
'            End If
'            If GBSwissProt Then '-----------------------------Look For GenBank-SwissProt Conversion
'               Set rsNew = dbGene.OpenRecordset( _
'                     "SELECT [Primary] FROM [SwissProt-GenBank]" & _
'                     "   WHERE Related = '" & id & "' ORDER BY INSTR([Primary], '_') <> 0 DESC")
'                  '  SwissProt IDs as opposed to Accession numbers have underscores. Eg: CALM_HUMAN
'               If Not rsNew.EOF Then '___________________________________Found In SwissProt-GenBank
'                  Set rs = dbGene.OpenRecordset( _
'                        "SELECT ID FROM SwissProt" & _
'                        "   WHERE ID = '" & rsNew![Primary] & "'" & _
'                        "      OR Accession LIKE '*|" & rsNew![Primary] & "|*'" & _
'                        "   ORDER BY ID = '" & rsNew![Primary] & "'")
'                  '  Look for IDs for accession numbers in SwissProt. May be more than one.
'                  '  Order should put SwissProt-GenBank Primaries that are SwissProt IDs first.
'                  If Not rs.EOF Then                                      'Found as ID in SwissProt
'                     Do Until rs.EOF
'                        If Not conversionFound Then
'                           .Edit
'                           !id = rs!id
'                           !systemCode = "S"
'                           .Update
'                           Print #FILE_CONVERT_LOG, rs!id; vbTab; "SwissProt"; vbTab;
'                           conversionFound = True
'                           changes = changes + 1
'                        Else
'                           otherPossibilities = otherPossibilities & rs!id & "[SwissProt] "
'                        End If
'                        rs.MoveNext
'                     Loop
'                  Else                                                'Not found as ID in SwissProt
'                     If Not conversionFound Then
'                        .Edit
'                        !id = rsNew![Primary]
'                        !systemCode = "S"
'                        .Update
'                        Print #FILE_CONVERT_LOG, rsNew![Primary]; vbTab; "SwissProt"; vbTab;
'                        conversionFound = True
'                        changes = changes + 1
'                     Else
'                        otherPossibilities = otherPossibilities & rsNew![Primary] & "[SwissProt] "
'                     End If
'                  End If
'                  rsNew.MoveNext
'                  Do Until rsNew.EOF
'                     otherPossibilities = otherPossibilities & rsNew![Primary] & "[SwissProt] "
'                     rsNew.MoveNext
'                  Loop
'               End If
'            End If
'            If GBLocusLink Then '-----------------------------Look For GenBank-LocusLink Conversion
'               Set rsNew = dbGene.OpenRecordset( _
'                     "SELECT [Primary] FROM [LocusLink-GenBank] WHERE Related = '" & id & "'")
'               If Not rsNew.EOF Then '___________________________________Found In LocusLink-GenBank
'                  If Not conversionFound Then
'                     .Edit
'                     !id = rsNew![Primary]
'                     !systemCode = "L"
'                     .Update
'                     Print #FILE_CONVERT_LOG, rsNew![Primary]; vbTab; "LocusLink"; vbTab;
'                     conversionFound = True
'                     changes = changes + 1
'                  Else
'                     otherPossibilities = otherPossibilities & rsNew![Primary] & "[LocusLink] "
'                  End If
'                  rsNew.MoveNext
'                  Do Until rsNew.EOF
'                     otherPossibilities = otherPossibilities & rsNew![Primary] & "[LocusLink] "
'                     rsNew.MoveNext
'                  Loop
'               End If
'            End If
'            If GBUniGene Then '---------------------------------Look For GenBank-UniGene Conversion
'               Set rsNew = dbGene.OpenRecordset( _
'                     "SELECT [Primary] FROM [UniGene-GenBank] WHERE Related = '" & id & "'")
'               If Not rsNew.EOF Then '_____________________________________Found In UniGene-GenBank
'                  If Not conversionFound Then
'                     .Edit
'                     !id = rsNew![Primary]
'                     !systemCode = "U"
'                     .Update
'                     Print #FILE_CONVERT_LOG, rsNew![Primary]; vbTab; "UniGene"; vbTab;
'                     conversionFound = True
'                     changes = changes + 1
'                  Else
'                     otherPossibilities = otherPossibilities & rsNew![Primary] & "[UniGene] "
'                  End If
'                  rsNew.MoveNext
'                  Do Until rsNew.EOF
'                     otherPossibilities = otherPossibilities & rsNew![Primary] & "[UniGene] "
'                     rsNew.MoveNext
'                  Loop
'               End If
'            End If
'         ElseIf !systemCode = "S" And modSys = "SwissProt" Then '==================SwissProt Is MOD
'            '  No conversion made
'            Print #FILE_CONVERT_LOG, "SwissProt"; vbTab; vbTab; vbTab; vbTab;
'         ElseIf !systemCode = "S" Then '=========================================Old SwissProt Code
'            Print #FILE_CONVERT_LOG, "SwissProt"; vbTab;
'            '-------------------------------------------------Pick Up IDs For All Accession Numbers
'            Set rsSwissProt = dbGene.OpenRecordset( _
'                  "SELECT DISTINCT ID FROM SwissProt WHERE Accession LIKE '*|" & id & "|*'")
'            sql = "'" & id & "'"                                       'Begin List With Original ID
'            Do Until rsSwissProt.EOF
'               sql = sql & ", '" & rsSwissProt!id & "'"        'Eg: 'P12345', 'CALM_HUMAN', 'A9687'
'               rsSwissProt.MoveNext
'            Loop
'            If modSP <> "" Then '---------------------------------Look For SwissProt-MOD Conversion
'               Set rsNew = dbGene.OpenRecordset( _
'                     "SELECT [Primary] FROM [" & modSP & "] WHERE Related IN(" & sql & ")")
'               If Not rsNew.EOF Then '_______________________________________Found In MOD-SwissProt
'                  .Edit
'                  !id = rsNew![Primary]
'                  !systemCode = modCode
'                  .Update
'                  Print #FILE_CONVERT_LOG, rsNew![Primary]; vbTab; modSys; vbTab;
'                  conversionFound = True
'                  changes = changes + 1
'                  rsNew.MoveNext
'                  Do Until rsNew.EOF '...........................................Pick Up Any Extras
'                     otherPossibilities = otherPossibilities & rsNew![Primary] & "[" & modSys & "] "
'                     rsNew.MoveNext
'                  Loop
'               End If
'            End If
'            If SPLocusLink Then '---------------------------Look For SwissProt-LocusLink Conversion
'               Set rsNew = dbGene.OpenRecordset( _
'                     "SELECT Related FROM [SwissProt-LocusLink] WHERE [Primary] IN(" & sql & ")")
'               If Not rsNew.EOF Then '_________________________________Found In SwissProt-LocusLink
'                  If Not conversionFound Then
'                     .Edit
'                     !id = rsNew!related
'                     !systemCode = "L"
'                     .Update
'                     Print #FILE_CONVERT_LOG, rsNew!related; vbTab; "LocusLink"; vbTab;
'                     conversionFound = True
'                     changes = changes + 1
'                  Else
'                     otherPossibilities = otherPossibilities & rsNew!related & "[LocusLink] "
'                  End If
'                  rsNew.MoveNext
'                  Do Until rsNew.EOF
'                     otherPossibilities = otherPossibilities & rsNew!related & "[LocusLink] "
'                     rsNew.MoveNext
'                  Loop
'               End If
'            End If
'            If SPUniGene Then '-------------------------------Look For SwissProt-UniGene Conversion
'               Set rsNew = dbGene.OpenRecordset( _
'                     "SELECT Related FROM [SwissProt-UniGene] WHERE [Primary] IN(" & sql & ")")
'               If Not rsNew.EOF Then '___________________________________Found In UniGene-SwissProt
'                  If Not conversionFound Then
'                     .Edit
'                     !id = rsNew!related
'                     !systemCode = "U"
'                     .Update
'                     Print #FILE_CONVERT_LOG, rsNew!related; vbTab; "UniGene"; vbTab;
'                     conversionFound = True
'                     changes = changes + 1
'                  Else
'                     otherPossibilities = otherPossibilities & rsNew!related & "[UniGene] "
'                  End If
'                  rsNew.MoveNext
'                  Do Until rsNew.EOF
'                     otherPossibilities = otherPossibilities & rsNew!related & "[UniGene] "
'                     rsNew.MoveNext
'                  Loop
'               End If
'            End If
'         Else '===============================================================Old Code Unidentified
'            Print #FILE_CONVERT_LOG, "No conversion made"; vbTab; vbTab; vbTab
''            Print #FILE_CONVERT_LOG, "Unidentified System '" & SystemCode & "'"; _
''                  vbTab; vbTab; vbTab
'         End If
'         Print #FILE_CONVERT_LOG, otherPossibilities
'         .MoveNext
'      Loop
'   End With
'   Close #FILE_CONVERT_LOG
'   If UCase(App.EXEName) = "CONVERT" Then
'      SetProgressBase
'      History
'   End If
'
'ExitSub:
'   Screen.ActiveForm.MousePointer = vbDefault
'   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'End Function

Function ConvertGBorSP(dbGene As Database, id As String, systemCode As String, _
                       newIDs() As String) As Integer
   '  Entry    dbGene      Open Gene DB being used for conversion
   '           id          ID to be converted
   '           systemCode  "G" for GenBank, "S" for SwissProt
   '           Module-level variables. Must be set before entry. These show MOD parameters
   '                 and the various conversion tables available.
   '              ModGB, ModSP   The MOD-GenBank and MOD-SwissProt tables
   '              GBSwissProt, GBLocusLink, GBUniGene, SPLocusLink, SPUniGene
   '                 These are true if these tables exist in the Gene Database
   '                 If modGB is SwissProt-GenBank (as for Human) then GBSwissProt is false
   '              modSys      MOD System. Eg: "SwissProt" (Human)
   '              modCode     System code for MOD, Eg: "S" for SwissProt (Human)
   '
   '  Return   Number of new IDs returned. Max is UBound(newIDs) (usually MAX_GENES).
   '           NewIDs   Substitute IDs for passed ID. Call must pass an array in this form:
   '                       newIDs(x, 0)   ID
   '                       newIDs(x, 1)   System code. Eg: "L"
   '                       newIDs(x, 2)   System Name. Eg: "LocusLink"
   Dim lastNewID As Integer
   Dim rsNew As Recordset, rs As Recordset
   Dim uniqueIDs(MAX_GENES) As String, noOfUniqueIDs As Integer
   Dim sql As String
      
   lastNewID = -1                                                 'Index of zero-based newIDs array
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find MOD IDs
   Set rsNew = Nothing
   If systemCode = "G" And modGB <> "" Then                                         'GenBank to MOD
      Set rsNew = dbGene.OpenRecordset( _
            "SELECT [Primary] FROM [" & modGB & "] WHERE Related = '" & id & "'")
   ElseIf systemCode = "S" And modSP <> "" Then                                   'SwissProt to MOD
      Set rsNew = dbGene.OpenRecordset( _
            "SELECT [Primary] FROM [" & modSP & "] WHERE Related = '" & id & "'")
      If rsNew.EOF Then                                                    'Check secondary columns
         noOfUniqueIDs = FindUniqueIDs(id, "S", uniqueIDs(), dbGene)
         sql = ""
         For i = 1 To noOfUniqueIDs
            sql = sql & "'" & uniqueIDs(i - 1) & "', "
         Next i
         If sql <> "" Then
            sql = Left(sql, Len(sql) - 2)                      'Drop off comma space
            Set rsNew = dbGene.OpenRecordset( _
                  "SELECT [Primary] FROM [" & modSP & "] WHERE Related IN(" & sql & ")")
         End If
      End If
   End If
   If Not rsNew Is Nothing Then '===============================================Found In MOD-System
      Do Until rsNew.EOF
         If lastNewID < UBound(newIDs) Then
            lastNewID = lastNewID + 1
            newIDs(lastNewID, 0) = rsNew![primary]
            newIDs(lastNewID, 1) = modCode
            newIDs(lastNewID, 2) = modSys
         End If
         rsNew.MoveNext
      Loop
   End If
   
   If systemCode = "G" And GBSwissProt Then '+++++++++++++++++++++++++++++++++++ Find SwissProt IDs
      '  Once we are this far, we only want to convert GenBanks
      Set rsNew = dbGene.OpenRecordset( _
            "SELECT [Primary] FROM [SwissProt-GenBank]" & _
            "   WHERE Related = '" & id & "' ORDER BY INSTR([Primary], '_') <> 0 DESC")
         '  SwissProt IDs as opposed to TrEMBL Accession numbers have underscores. Eg: CALM_HUMAN
      Do Until rsNew.EOF '===============================================Found In SwissProt-GenBank
         Set rs = dbGene.OpenRecordset( _
               "SELECT ID FROM SwissProt" & _
               "   WHERE ID = '" & rsNew![primary] & "'" & _
               "      OR Accession LIKE '*|" & rsNew![primary] & "|*'" & _
               "   ORDER BY INSTR(ID, '_') <> 0 DESC")
         '  Look for IDs for accession numbers in SwissProt. May be more than one.
         '  Order should put SwissProt-GenBank Primaries that are SwissProt IDs before
         '  TrEMBL IDs or accession numbers.
         Do Until rs.EOF '-------------------------------------------Found as ID in SwissProt Table
            If lastNewID < UBound(newIDs) Then
               lastNewID = lastNewID + 1
               newIDs(lastNewID, 0) = rs!id
               newIDs(lastNewID, 1) = "S"
               newIDs(lastNewID, 2) = "SwissProt"
            End If
            rs.MoveNext
         Loop
         rsNew.MoveNext
      Loop
   End If
   
   If systemCode = "G" And GBLocusLink Then '+++++++++++++++++++++++++++++++++++ Find LocusLink IDs
      Set rsNew = dbGene.OpenRecordset( _
            "SELECT [Primary] FROM [LocusLink-GenBank] WHERE Related = '" & id & "'")
      Do Until rsNew.EOF  '==============================================Found In LocusLink-GenBank
         If lastNewID < UBound(newIDs) Then
            lastNewID = lastNewID + 1
            newIDs(lastNewID, 0) = rsNew![primary]
            newIDs(lastNewID, 1) = "L"
            newIDs(lastNewID, 2) = "LocusLink"
         End If
         rsNew.MoveNext
      Loop
   End If
   If systemCode = "G" And GBUniGene Then  '++++++++++++++++++++++++++++++++++++++ Find UniGene IDs
      Set rsNew = dbGene.OpenRecordset( _
            "SELECT [Primary] FROM [UniGene-GenBank] WHERE Related = '" & id & "'")
      Do Until rsNew.EOF  '================================================Found In UniGene-GenBank
         If lastNewID < UBound(newIDs) Then
            lastNewID = lastNewID + 1
            newIDs(lastNewID, 0) = rsNew![primary]
            newIDs(lastNewID, 1) = "U"
            newIDs(lastNewID, 2) = "UniGene"
         End If
         rsNew.MoveNext
      Loop
   End If
   ConvertGBorSP = lastNewID + 1                                           'Number of new IDs found
End Function
'************************************************************** All Unique IDs For Any ID In System
Function FindUniqueIDs(id As String, systemCode As String, uniqueIDs() As String, _
                       dbGene As Database) As Integer
   '  Entry    id          ID to find. Eg: "P123456"
   '           systemCode  For ID. Eg: "S"
   '                       This can also be a system name, eg: "SwissProt". Checks for code first.
   '           uniqueIDs() Passed array for unique IDs
   '           dbGene      Open Gene Database
   '  Return   Number of unique IDs found.
   '           uniqueIDs() UniqueIDs found, zero based.
   Dim rsSystems As Recordset, rsID As Recordset
   Dim slash As Integer, pipe As Integer, column As String
   Dim lastUniqueID As Integer
   
   Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT System, Columns FROM Systems WHERE SystemCode = '" & systemCode & "'")
   If rsSystems.EOF Then                                    'Check if systemCode really System name
      Set rsSystems = dbGene.OpenRecordset( _
                      "SELECT System, Columns FROM Systems WHERE System = '" & systemCode & "'")
   End If
   If rsSystems.EOF Then                                                      'System doesn't exist
      FindUniqueIDs = 0
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ See If ID Passed Is Unique
   Set rsID = dbGene.OpenRecordset( _
              "SELECT ID FROM " & rsSystems!system & " WHERE ID = '" & id & "'")
   If Not rsID.EOF Then                                                               'It is unique
      uniqueIDs(0) = id
      FindUniqueIDs = 1
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Look For IDs In Piped Columns
   slash = InStr(rsSystems!columns, "\S")
   Do Until slash = 0
      pipe = InStrRev(rsSystems!columns, "|", slash)
      column = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)
      Set rsID = dbGene.OpenRecordset( _
                 "SELECT ID FROM " & rsSystems!system & _
                 "   WHERE INSTR(" & column & ", '|" & id & "|') <> 0")
      Do Until rsID.EOF
         If lastUniqueID < UBound(uniqueIDs) Then
            uniqueIDs(lastUniqueID) = rsID!id
            lastUniqueID = lastUniqueID + 1
         End If
         rsID.MoveNext
      Loop
      slash = InStr(slash + 1, rsSystems!columns, "\S")
   Loop
      
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Look For IDs In Nonpiped Columns
   slash = InStr(rsSystems!columns, "\s")
   Do Until slash = 0
      pipe = InStrRev(rsSystems!columns, "|", slash)
      column = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)
      Set rsID = dbGene.OpenRecordset( _
                 "SELECT ID FROM " & rsSystems!system & _
                 "   WHERE " & column & " = '" & id & "'")
      Do Until rsID.EOF
         If lastUniqueID < UBound(uniqueIDs) Then
            uniqueIDs(lastUniqueID) = rsID!id
            lastUniqueID = lastUniqueID + 1
         End If
         rsID.MoveNext
      Loop
      slash = InStr(slash + 1, rsSystems!columns, "\s")
   Loop
      
   FindUniqueIDs = lastUniqueID
End Function
'*********************************************************** Switch GenBank And/Or SwissProt To MOD
Sub ConvertGBorSPinFile(GBorSP As String, dbGene As Database, file As String, _
                        Optional changeLog As String = "", Optional tiles As Boolean = False)
   '  Entry    GBorSP   "G" for GenBank, "S" for SwissProt, "GS" for both
   '           dbGene   Open Gene DB being used for conversion
   '           file     Full path to a database, either a MAPP or an Expression Dataset
   '           tiles    True if any tiles (multiple substitutions for genes on a MAPP) exist
   '  Return   tiles    True if any tiles (multiple substitutions for genes on a MAPP) made
   '  Controls required on active form:   prgProgress, lblOperation, lblDetail
   '  DetermineMODConversionTables() must be called before entry
   Dim dbFile As Database
   Dim dbType As String                               '"MAPP" for MAPP, "ED" for Expression Dataset
   Dim tdf As TableDef
   Dim sql As String
   Dim rs As Recordset
   Dim newIDs(MAX_GENES, 2) As String                            'Substitute IDs for passed GenBank
   '  newIDs(x, 0)   ID
   '  newIDs(x, 1)   System code. Eg: "L"
   '  newIDs(x, 2)   System Name. Eg: "LocusLink"
   Dim noOfNewIDs As Integer, i As Integer, notes As String
   Dim objKey As Integer, maxKey As Integer
   
   Select Case GBorSP
   Case "G"
      Screen.ActiveForm.lblOperation = "Converting GenBank"
   Case "S"
      Screen.ActiveForm.lblOperation = "Converting SwissProt"
   Case Else
      Screen.ActiveForm.lblOperation = "Converting GenBank and SwissProt"
   End Select
   Screen.ActiveForm.lblDetail = ""
   Screen.ActiveForm.lblOperation.visible = True
   Screen.ActiveForm.lblDetail.visible = True
   
   Set dbFile = OpenDatabase(file)
   
   dbType = "ED" '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine DB Type
   For Each tdf In dbFile.TableDefs
      If tdf.name = "Objects" Then                                         'Only exists in MAPP DBs
         dbType = "MAPP"
         Exit For                                          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
   Next tdf
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up Change Log File
   If changeLog = "" Then
      changeLog = Left(dbFile.name, InStrRev(dbFile.name, ".")) & "log"
   End If
   Open changeLog For Output As #FILE_CONVERT_LOG
   
   Select Case dbType '++++++++++++++++++++++++++++++++++++++++++++++ Set Up For Different DB Types
   Case "MAPP"
      Print #FILE_CONVERT_LOG, "Gene Label"; vbTab; "Old ID"; vbTab; "Old System"; vbTab; _
            "New ID"; vbTab; " New System"; vbTab; "Other Possibilities: ID[System]"
      Set rs = dbFile.OpenRecordset("SELECT MAX(ObjKey) AS MaxKey FROM Objects")
      maxKey = rs!maxKey                                             'For adding tiled gene objects
      objKey = maxKey
      Set rs = dbFile.OpenRecordset( _
               "SELECT Count(0) AS Records FROM Objects" & _
               "   WHERE INSTR('" & GBorSP & "', SystemCode) <> 0 AND Type = 'Gene'")
      sql = "SELECT * FROM Objects WHERE INSTR('" & GBorSP & "', SystemCode) <> 0 AND Type = 'Gene' ORDER BY ObjKey"
   Case "ED"
      Print #FILE_CONVERT_LOG, "Old ID"; vbTab; "Old System"; vbTab; _
            "New ID"; vbTab; " New System"; vbTab; "Other Possibilities: ID[System]"
      Set rs = dbFile.OpenRecordset( _
               "SELECT Count(0) AS Records FROM Expression" & _
               "   WHERE INSTR('" & GBorSP & "', SystemCode) <> 0")
      sql = "SELECT * FROM Expression WHERE INSTR('" & GBorSP & "', SystemCode) <> 0"
   End Select
   Screen.ActiveForm.prgProgress.value = 0
   Screen.ActiveForm.prgProgress.Max = Max(rs!records, 1)      'Quick fix. If rs returns no records
                                                  'then Max is zero -- invalid, so we limit it to 1
   Screen.ActiveForm.prgProgress.visible = True
   
   Set rs = dbFile.OpenRecordset(sql) '++++++++++++++++++++++++++++++++++++++++++++++ Go Through DB
   Do Until rs.EOF
      If dbType = "MAPP" Then
         If rs!objKey > maxKey Then Exit Do                'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            '  maxKey is the highest objKey value in original MAPP. Beyond this point we are
            '  looking at only tiled, added, objects
         Print #FILE_CONVERT_LOG, rs!Label; vbTab;
      End If
      Select Case rs!systemCode
      Case "G"
         Print #FILE_CONVERT_LOG, rs!id; vbTab; "GenBank"; vbTab;
      Case "S"
         Print #FILE_CONVERT_LOG, rs!id; vbTab; "SwissProt"; vbTab;
      Case Else
         Print #FILE_CONVERT_LOG, rs!id; vbTab; rs!systemCode; vbTab;
      End Select
      Screen.ActiveForm.prgProgress.value = Min(rs.AbsolutePosition, _
                                                Screen.ActiveForm.prgProgress.Max)
      Screen.ActiveForm.lblDetail = rs!id
      DoEvents
      noOfNewIDs = ConvertGBorSP(dbGene, rs!id, rs!systemCode, newIDs())
      If noOfNewIDs > 0 Then '====================================================Found Substitutes
         notes = Dat(rs!notes)                                        'May add unchosen IDs to this
         If noOfNewIDs > 1 Then '--------------------------------------Add Extra IDs To Notes Field
            notes = notes & " «"                                                          'ANSI 171
            For i = 1 To noOfNewIDs - 1
               notes = notes & newIDs(i, 0) & "[" & newIDs(i, 1) & "] "
            Next i
            notes = notes & "»"                                                           'ANSI 187
            '  Notes field ends up looking like this:
            '     Whatever «A12345[G] SNOUT_AARDVARK[S] »
         End If
         If dbType = "MAPP" Then
            dbFile.Execute "UPDATE Objects" & _
                           "   SET ID = '" & newIDs(0, 0) & "'," & _
                           "       SystemCode = '" & newIDs(0, 1) & "'," & _
                           "       Notes = '" & notes & "'" & _
                           "   WHERE ObjKey = " & rs!objKey
               '  Without tiling, this update can be done just like the ED update below
            Dim newCenterX As Single, newCenterY As Single
            newCenterX = rs!centerX
            newCenterY = rs!centerY
         Else
            rs.edit '-----------------------------------------Assign First Return To Current Record
            rs!id = newIDs(0, 0)
            rs!systemCode = newIDs(0, 1)
            rs!notes = notes
            rs.Update
         End If
         Print #FILE_CONVERT_LOG, newIDs(0, 0); vbTab; newIDs(0, 2); vbTab;
         
         For i = 1 To noOfNewIDs - 1 '------------------List Subsequent Ones In Other Possibilities
            Print #FILE_CONVERT_LOG, newIDs(i, 0) & "[" & newIDs(i, 2) & "]   ";
            If dbType = "MAPP" Then '__________________________________Tile Subsequent Ones On MAPP
               '  If there are any other MOD entries for the given GenBank, put them in
               '  the conversion log as Other Possibilities and show them on the MAPP
               '  tiled at 50 pixel offsets from the first conversion.
               Dim newWidth As Single, newHeight As Single
               Dim newLabel As Variant, newHead As Variant, newRemarks As Variant
               Dim newLinks As Variant, newNotes As Variant
               objKey = objKey + 1
               newCenterX = newCenterX + 50
               newCenterY = newCenterY + 50
               With rs
                  newWidth = !Width
                  newHeight = !Height
                  newLabel = !Label
                  newHead = !head
                  newRemarks = !remarks
                  newLinks = !links
                  newNotes = !notes
                  
'                  sql = "INSERT INTO Objects (ObjKey, ID, SystemCode, Type, CenterX, CenterY, Width, Height, Label, Head, Remarks, Links, Notes)"
'                  sql = sql & " VALUES (" & objKey & ", '" & newIDs(i, 0) & "', '" & newIDs(i, 1) & "', 'Gene', " & newCenterX & ", " & newCenterY & ", " & rs!Width & ", " & rs!Height & ", '" & rs!Label & "', '" & rs!Head & "', '" & rs!remarks & "', '" & rs!Links & "', '" & rs!Notes & "')"
'                  dbFile.Execute sql
                  .AddNew  'Do this instead of INSERT so that the rsNew loop doesn't lose its place
                  !objKey = objKey
                  !id = newIDs(i, 0)
                  !systemCode = newIDs(i, 1)
                  !Type = "Gene"
                  !centerX = newCenterX
                  !centerY = newCenterY
                  !Width = newWidth
                  !Height = newHeight
                  !Label = newLabel
                  !head = newHead
                  !remarks = newRemarks
                  !links = newLinks
                  !notes = newNotes
                  .Update
               End With
               tiles = True
            End If
         Next i
         Print #FILE_CONVERT_LOG, " "
      Else
         Print #FILE_CONVERT_LOG, "No conversion available"; vbTab; vbTab; " "
      End If
      rs.MoveNext
   Loop
   Close #FILE_CONVERT_LOG
   Screen.ActiveForm.prgProgress.visible = False
   Screen.ActiveForm.lblDetail = ""
   Screen.ActiveForm.lblOperation = ""
   Screen.ActiveForm.lblOperation.visible = False
   Screen.ActiveForm.lblDetail.visible = False
   DoEvents
End Sub
'**************************************************** Brings Expression Dataset Up To Current Specs
Function UpdateSingleDataset(expression As String) As Boolean
   Dim dbExpression As Database
   Dim rsInfo As Recordset, fld As Field, ok As Boolean
   Dim tbl As TableDef

   Screen.ActiveForm.MousePointer = vbHourglass
   Set dbExpression = OpenDatabase(expression)
   DeleteIndexes dbExpression, "Expression"
   
   ok = True '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Change Expression Table
   For Each fld In dbExpression.TableDefs!expression.Fields
      '  Must do these separately or the Delete unsets the Field object
      If fld.name = "GenMAPP" Then
         dbExpression.TableDefs!expression.Fields.Delete "GenMAPP"
      End If
   Next fld
   For Each fld In dbExpression.TableDefs!expression.Fields
      If fld.name = "Experiment" Then
         dbExpression.TableDefs!expression.Fields.Delete "Experiment"
      End If
   Next fld
   For Each fld In dbExpression.TableDefs!expression.Fields
      If fld.name = "DisplayName" Then
         dbExpression.TableDefs!expression.Fields.Delete "DisplayName"
      End If
   Next fld
   For Each fld In dbExpression.TableDefs!expression.Fields
      If fld.name = "DisplayColorSet" Then
         dbExpression.TableDefs!expression.Fields.Delete "DisplayColorSet"
      End If
   Next fld
   For Each fld In dbExpression.TableDefs!expression.Fields
      If fld.name = "DisplayRemarks" Then
         dbExpression.TableDefs!expression.Fields.Delete "DisplayRemarks"
      End If
   Next fld
   For Each fld In dbExpression.TableDefs!expression.Fields
      If fld.name = "DisplayValue" Then
         dbExpression.TableDefs!expression.Fields.Delete "DisplayValue"
      End If
   Next fld
   For Each fld In dbExpression.TableDefs!expression.Fields
      If fld.name = "Primary" Then fld.name = "ID"
      If fld.name = "PrimaryType" Then fld.name = "SystemCode"
      If fld.name = "Notes" Then fld.AllowZeroLength = True                   'Old EDs not this way
'   dbExpression.TableDefs!expression.Fields!notes.AllowZeroLength = True      'Old EDs not this way
   Next fld
   dbExpression.Execute "UPDATE Expression SET SystemCode = 'S' WHERE SystemCode = 'N'"
      '  Old SWISS-PROT accession numbers were given Primary Type N
   
   ok = False '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Add Info!GeneDB Field
   For Each fld In dbExpression.TableDefs!info.Fields
      If fld.name = "GeneDB" Then
         ok = True
         Exit For
      End If
   Next fld
   If Not ok Then
      dbExpression.TableDefs!info.Fields.Append _
                   dbExpression.TableDefs!info.CreateField("GeneDB", dbMemo)
   End If
   dbExpression.TableDefs!info.Fields("GeneDB").OrdinalPosition = 3       'Set Position in any case
   
   ok = False '+++++++++++++++++++++++++++++++++++++++++++++++++++++++ Add Info!GeneDBVersion Field
   For Each fld In dbExpression.TableDefs!info.Fields
      If fld.name = "GeneDBVersion" Then
         ok = True
         Exit For
      End If
   Next fld
   If Not ok Then
      dbExpression.TableDefs!info.Fields.Append _
                   dbExpression.TableDefs!info.CreateField("GeneDBVersion", dbText, 10)
   End If
   dbExpression.TableDefs!info.Fields("GeneDBVersion").OrdinalPosition = 4            'Set Position
   
   ok = False '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Add Info!SystemCodes Field
      '  This field is filled when the Display table is created
   For Each fld In dbExpression.TableDefs!info.Fields
      If fld.name = "SystemCodes" Then
         ok = True
         Exit For
      End If
   Next fld
   If Not ok Then
      dbExpression.TableDefs!info.Fields.Append _
                   dbExpression.TableDefs!info.CreateField("SystemCodes", dbMemo)
   End If
   dbExpression.TableDefs!info.Fields!systemCodes.OrdinalPosition = 7     'Set Position in any case
   
   ok = False '+++++++++++++++++++++++++++++++++++++++++++++++++++++ Add Info!MOD and MODCode Field
      '  Not currently used
   For Each fld In dbExpression.TableDefs!info.Fields
      If fld.name = "MOD" Then
         ok = True
         Exit For
      End If
   Next fld
   If Not ok Then
      dbExpression.TableDefs!info.Fields.Append _
                   dbExpression.TableDefs!info.CreateField("MOD", dbMemo)
      dbExpression.TableDefs!info.Fields.Append _
                   dbExpression.TableDefs!info.CreateField("MODCode", dbText, 2)
   End If
   dbExpression.TableDefs!info.Fields!MOD.OrdinalPosition = 8             'Set Position in any case
   dbExpression.TableDefs!info.Fields!modCode.OrdinalPosition = 9         'Set Position in any case
   
   ok = False '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Add ColorSet!SetNo Field
   For Each fld In dbExpression.TableDefs!colorSet.Fields
      If fld.name = "Graphic" Then fld.name = "Remarks"
      If fld.name = "SetNo" Then
         ok = True
      End If
   Next fld
   If Not ok Then
      dbExpression.TableDefs!colorSet.Fields.Append _
                   dbExpression.TableDefs!colorSet.CreateField("SetNo", dbInteger)
      dbExpression.TableDefs!colorSet.Fields("SetNo").OrdinalPosition = 1
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Reindex Expression Table
   dbExpression.Execute "CREATE INDEX idxID ON Expression (ID, SystemCode)"
   dbExpression.Execute "CREATE INDEX idxOrder ON Expression (OrderNo)"
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Save Updated Dataset
   dbExpression.Execute "UPDATE Info SET Version = '" & BUILD & "'"
   dbExpression.Close
   
   UpdateConfig "mruDataSet", expression
   
   Screen.ActiveForm.MousePointer = vbDefault
   UpdateSingleDataset = True
End Function
'*********************************************************** Imports Raw Data To Expression Dataset
Function ConvertExpressionData(rawDataFile As String, dbExpression As Database, _
                               dbGene As Database, Optional returnExceptions As String = "") _
                              As Boolean
   '  Entry    rawDataFile    Path to raw data file (unopened). Must be valid file,
   '                          not checked here. It also provides the name of the ED. It must
   '                          agree with the frmExpression.expressionName if called from that form.
   '           dbExpression   Expression Dataset created and open. Must have all the
   '                          columns etc to fit the raw data file. This could be a temporary ED
   '                          depending on where it was called
   '           dbGene         An open Gene Database
   '           returnExceptions  If anything but "", do not print exception messages from here,
   '                             just return them in this variable.
   '  Return   True if successful. At this point, nothing would make this false.
   '           expressionDB is filled in, errors or not, including Display table
   '           If errors exist then an exception file is generated
   '  Required controls on calling form:
   '     lblProgresssTitle
   '     lblErrors
   '     lblPrgMax
   '     prgProgress
   Dim tdfExpression As TableDef, fld As Field
   Dim tempGex As Database
   Dim delimiter As String * 1                                'Delimiter (tab or comma) in raw data
   Dim inLine As String                                                   'Input line from raw data
   Dim i As Integer, j As Integer
   Dim orderNo As Long
   Dim invalidChrs As String, s As String
   Dim system As String, systemCode As String, rsSystems As Recordset, rsInfo As Recordset
   Dim rs As Recordset
   Dim systems(MAX_SYSTEMS, 2), lastSystem As Integer
      '  systems(x, 0)     Name of cataloging system. Eg: GenBank
      '  systems(x, 1)     System code. Eg: G
      '  systems(x, 2)     Additional search columns. Eg: |Gene\sBF|Orf\SBF|
   Dim slash As Integer, pipe As Integer, dot As Integer, idColumn As String
   Dim errors As String, errorFile As String, exceptionFile As String, geneId As String
   Dim errorsExists As Integer                           'One if ~Errors~ column exists in raw data
   Dim sql As String
   Dim columns As Integer 'Number of columns of raw data including Notes and Remarks if they exist.
                          'One based
   Dim notesIndex As Integer                     'Position of Notes column in raw data (zero-based)
   Dim notes As String
      '  Notess can be in any column of the raw dataset. If it exists it is shifted to the
      '  second-to-last column of the ED. Otherwise, an empty column is added to the ED.
   Dim remarksIndex As Integer                 'Position of Remarks column in raw data (zero-based)
   Dim remarks As String
      '  Remarks can be in any column of the raw dataset. If it exists it is
      '  shifted to the last column of the ED. Otherwise, an empty column is added to the ED.
   Dim AddToOther As Boolean                    'If true, add all unidentified genes to Other table
   Dim addToDB As Boolean                       'Gene valid enough to add to the Expression Dataset
   Dim warningColumns As String
   Dim otherGene As Boolean                                          'This gene gets added to Other
   'For AllRelatedGenes()
      Dim genes As Integer
      Dim geneIDs(MAX_GENES, 2) As String
      Dim geneFound As Boolean
      'Dim supportedSystem as Boolean                 'System supported in Gene Database [optional]
      Dim systemsList As Variant                                      'Systems to search [optional]
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++ Open Raw Data And Determine Characteristics
   History "Converting raw expression data"
   Screen.ActiveForm.lblDetail.visible = True
   Screen.ActiveForm.lblDetail = FileAbbrev(rawDataFile)
   Open rawDataFile For Binary As #FILE_RAW_DATA
   Screen.ActiveForm.prgProgress.visible = True
   Screen.ActiveForm.prgProgress.Max = LOF(FILE_RAW_DATA)
   Screen.ActiveForm.prgProgress.value = 0
   inLine = RemoveQuotes(InputUnixLine(FILE_RAW_DATA))             'First line (titles) of raw data
   If InStr(inLine, vbTab) Then
      delimiter = vbTab
   Else
      delimiter = ","
   End If
   notesIndex = -1                                                             'Default to no Notes
   i = InStr(1, inLine, delimiter & "NOTES", vbTextCompare)             'Look for Notes in raw Data
   If i Then  '=================================================================Notes Column Exists
      If Mid(inLine, i + 6, 1) = delimiter Or Len(inLine) = i + 5 Then
         '  "Notes" either followed by a delimiter or at end of inLine
         notesIndex = 0
         j = 0                                                                      'Next delimiter
         Do
            notesIndex = notesIndex + 1
            j = InStr(j + 1, inLine, delimiter)
         Loop Until j = i   'Delimiter before column name same location as delimiter before Remarks
         notesIndex = notesIndex - 1                   'Index doesn't include GeneID and SystemCode
      End If
   End If
   remarksIndex = -1                                                         'Default to no Remarks
   i = InStr(1, inLine, delimiter & "REMARKS", vbTextCompare)         'Look for Remarks in raw Data
   If i Then  '===============================================================Remarks Column Exists
      If Mid(inLine, i + 8, 1) = delimiter Or Len(inLine) = i + 7 Then
         '  "Remarks" either followed by a delimiter or at end of inLine
         remarksIndex = 0
         j = 0                                                                      'Next delimiter
         Do
            remarksIndex = remarksIndex + 1
            j = InStr(j + 1, inLine, delimiter)
         Loop Until j = i   'Delimiter before column name same location as delimiter before Remarks
         remarksIndex = remarksIndex - 1               'Index doesn't include GeneID and SystemCode
      End If
   End If
   If Right(inLine, 9) = delimiter & "~Errors~" Then                          'Is an exception file
      errorsExists = 1
      '  Only ask this after the user has been through the raw data file once
      If MsgBox("Automatically add all unidentified genes to the ""Other"" table?", _
                vbQuestion + vbYesNo, "Converting Raw Data") = vbYes Then
         AddToOther = True
      End If
   End If

   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Open Error File And Print Headings
   errorFile = Left(rawDataFile, InStrRev(rawDataFile, ".") - 1) & ".$tm"
   Open errorFile For Output As #FILE_EXCEPTIONS
   Set tdfExpression = dbExpression.TableDefs("Expression")
      '  The Expression DB will always have a Remarks column
   For Each fld In tdfExpression.Fields                              'Use existing Expression table
      Select Case fld.name
      Case "OrderNo", "Notes", "Remarks"                            'OrderNo not an exception file,
                                                                     'Notes & Remarks always at end
      Case "ID"
         '  Second field of the ED is "ID" but Excel has an acknowledged bug that will not open
         '  a csv file that begins with "ID" so we make the column "Gene ID" for the
         '  exception file.
         Print #FILE_EXCEPTIONS, "Gene ID"; delimiter;
      Case Else                                                           '
         Print #FILE_EXCEPTIONS, fld.name; delimiter;
      End Select
   Next fld
   Print #FILE_EXCEPTIONS, "Notes"; delimiter; "Remarks"; delimiter; "~Errors~"; vbLf;
      
   '++++++++++++++++++++++++++++++++++++++++++++++++ Determine Number Of Columns Of Expression Data
   columns = tdfExpression.Fields.count - 3                                              'One-based
                                                              'Subtract OrderNo, GeneID, SystemCode
   If notesIndex = -1 Then                                                   'Notes not in raw data
      columns = columns - 1                                         'Reduce no. of expValue columns
   End If
   If remarksIndex = -1 Then                                               'Remarks not in raw data
      columns = columns - 1                                         'Reduce no. of expValue columns
   End If
   ReDim expValues(columns) As String                       'Row of values from raw expression file
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Get Valid Systems
   Set rsSystems = dbGene.OpenRecordset( _
                   "SELECT * FROM Systems ORDER BY System", dbOpenForwardOnly)
   lastSystem = -1
   Do Until rsSystems.EOF
      If VarType(rsSystems!Date) <> vbNull Or rsSystems!system = "Other" Then     'Supported system
         '  Other system is always supported, date or not
         lastSystem = lastSystem + 1
         systems(lastSystem, 0) = rsSystems!system
         systems(lastSystem, 1) = rsSystems!systemCode
         slash = InStr(1, rsSystems!columns, "\S", vbTextCompare)
         Do While slash                                              'Get additional search columns
            pipe = InStrRev(rsSystems!columns, "|", slash)
            systems(lastSystem, 2) = systems(lastSystem, 2) _
                                   & Mid(rsSystems!columns, pipe, slash - pipe + 2) & "|"
            slash = InStr(slash + 1, rsSystems!columns, "\S", vbTextCompare)
         Loop
      End If
      rsSystems.MoveNext
   Loop
   
   orderNo = 0                                          'Skip first line because it is the headings
   dbExpression.Execute "DELETE FROM Expression"                 'Start with empty Expression table
   systemsList = "|"
   Screen.ActiveForm.lblPrgMax.visible = True
   Screen.ActiveForm.lblPrgMax.Caption = "Errors:"
   Screen.ActiveForm.lblErrors.visible = True
   Screen.ActiveForm.lblErrors = 0
   DoEvents
   errors = GetExpressionRow(errorsExists, geneId, systemCode, expValues, delimiter)    'Second row
   Do '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Each Raw Data Row
      '  Each line, whether from a raw data file or exception file goes through same error checks
      '  First line of data already parsed to get datatypes above. If first line has too many
      '  or too few rows, routine stops before this.
      orderNo = orderNo + 1
      otherGene = False
      addToDB = True

      '=======================================================================Validate Gene ID Name
      '  If Gene ID invalid, don't try to find it in any System tables
      If Len(geneId) > CHAR_DATA_LIMIT Then                                                            'Too long
         errors = errors & "Gene ID more than " & CHAR_DATA_LIMIT & " characters. "
         Screen.ActiveForm.lblErrors = Screen.ActiveForm.lblErrors + 1
      End If
      invalidChrs = "gene ID"       'Send to InvalidChr function and get return of any invalid chrs
         '  Does not allow apostrophe because it must match other outside data and apostrophes
         '  must be changed in DBs
      If InvalidChr(geneId, invalidChrs) Then
         errors = errors & "Invalid character(s) " & invalidChrs & "found in gene ID. "
         addToDB = False
      End If
      
      i = 0 '==================================================================See If System Exists
      Do Until systemCode = systems(i, 1) Or i > lastSystem
         i = i + 1
      Loop
      If i <= lastSystem And addToDB Then '===========================================Identify Gene
         '  If addToDB false then name must be invalid. Don't search.
         geneFound = False
'         systemsList = "EXISTS"
         AllRelatedGenes geneId, systemCode, dbGene, genes, geneIDs, geneFound, , "EXISTS"
         If geneFound Then
            If UCase(geneIDs(genes - 1, 2)) = "S" Then                       'Secondary code exists
               If InStr(systemsList, "|~|") = 0 Then
                  systemsList = systemsList & "~|"
               End If
            End If
            If InStr(systemsList, "|" & systemCode & "|") = 0 Then                  'New SystemCode
               systemsList = systemsList & systemCode & "|"
            End If
         Else
            If AddToOther Then
               otherGene = True
            ElseIf errorsExists = 0 Then
'               lblErrors = lblErrors + 1
               errors = errors & "Gene not found in " & systems(i, 0) & " or any related system. "
            End If
         End If
      Else  '-----------------------------------------------------------System Not In Gene Database
         If AddToOther Then
            otherGene = True
         ElseIf errorsExists = 0 Then
            '  This is only an error on the first time through the data. Subsequent times,
            '  the gene is just accepted
'            lblErrors = lblErrors + 1
            errors = errors & "Gene ID system " & systemCode & " not in Gene Database. "
         End If
      End If
      
      '=======================================================================Check For Data Errors
      sql = ""       'Assemble SQL here even though addToDB may be false because we are testing for
                     'datatypes, etc. anyway and it would take more time to repeat the tests later
                     'than throw away an sql variable here.
      notes = ""
      remarks = ""
      j = 0                'Columns in ED with Remarks shifted, not including GeneID and SystemCode
      For i = 1 To columns                                                    'Each raw data column
         If i = notesIndex Then                                           'Move Notes to end of row
            notes = expValues(i)
         ElseIf i = remarksIndex Then                                   'Move Remarks to end of row
            remarks = expValues(i)
         ElseIf VarType(expValues(i)) = vbNull Then                'NULL accepted for any data type
            sql = sql & ", NULL"
            j = j + 1
         ElseIf tdfExpression.Fields(j + 3).Type = dbSingle Then
            If Not IsNumeric(expValues(i)) Then                            'Turn empties into NULLs
               If Trim(expValues(i)) <> "" Then
'                  lblErrors = lblErrors + 1
                  errors = errors & "Nonnumeric value """ & expValues(i) & """ in numeric column '" _
                        & tdfExpression.Fields(j + 3).name & "'. "
                  addToDB = False
               End If
               sql = sql & ", NULL"                        'Empty columns will be turned into NULLs
            Else
               sql = sql & ", " & expValues(i)
            End If
            j = j + 1
         ElseIf tdfExpression.Fields(j + 3).Type = dbText Then
            If Len(expValues(i)) > CHAR_DATA_LIMIT Then
'               lblErrors = lblErrors + 1
               errors = errors & "Too many characters """ & expValues(i) & """ in column '" _
                     & tdfExpression.Fields(j + 3).name & "'. "
'               addToDB = False
               If InStr(warningColumns, tdfExpression.Fields(j + 3).name & vbCrLf) = 0 Then
                  warningColumns = warningColumns & "   " & tdfExpression.Fields(j + 3).name _
                                 & vbCrLf
               End If
            End If
            s = TextToSql(Left(expValues(i), CHAR_DATA_LIMIT))                    'Shorten to limit
            invalidChrs = "char data"   'Send to InvalidChr function and get return of invalid chrs
            If InvalidChr(s, invalidChrs) Then
               errors = errors & "Invalid character(s) " & invalidChrs & "found in " _
                      & tdfExpression.Fields(j + 3).name & ". "
               addToDB = False
            End If
            sql = sql & ", '" & s & "'"
            j = j + 1
         Else                                                     'Other datatypes (should be none)
            sql = sql & ", '" & expValues(i) & "'"
            j = j + 1
         End If
      Next i
      sql = sql & ", '" & notes & "', '" & remarks & "')"                        'Notes and Remarks

      If addToDB Then
         If otherGene Then
            dbGene.Execute "INSERT INTO Other (ID, SystemCode, [Date])" & _
                           "   VALUES ('" & geneId & "', '" & systemCode & "', '" _
                                       & Format(Now, "dd-mmm-yyyy") & "')"
         End If
         sql = "INSERT INTO Expression VALUES (" & orderNo & ", '" & geneId & "', '" _
             & systemCode & "'" & sql
'On Error GoTo ConvertError
         dbExpression.Execute sql       'Errors here fall to ConvertError and come back at QuitLine
On Error GoTo 0
      End If
      
QuitLine:
      '=====================================================================Write To Exception File
      '  Remarks moved to end and then errors added
      Print #FILE_EXCEPTIONS, geneId; delimiter; systemCode; delimiter;
      For i = 1 To columns
         If i <> notesIndex And i <> remarksIndex Then                           'Notes and Remarks
            Print #FILE_EXCEPTIONS, expValues(i); delimiter;
         End If
      Next i
      Print #FILE_EXCEPTIONS, ""; delimiter; remarks; delimiter; errors & " " & vbLf;
      '  Blank column for Notes before Remarks
      '  Space always added to errors because some spreadsheet, etc, software drops the delimiter
      '  if the last column empty
      '  vbLf for Unix
      If errors <> "" Then
         Screen.ActiveForm.lblErrors = Screen.ActiveForm.lblErrors + 1
         DoEvents
      End If
      errors = GetExpressionRow(errorsExists, geneId, systemCode, expValues, delimiter)
Detail geneId
   Loop While errors <> "**eof**"
   Screen.ActiveForm.lblErrors.visible = False
                                                                           'Row Processing Finished
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Write Info Table Data
   Set rs = dbGene.OpenRecordset("SELECT Version FROM Info")
   Set rsInfo = dbExpression.OpenRecordset("SELECT * FROM Info")
   If Not rsInfo.EOF Then               'Info already exists probably because processing exceptions
      dbExpression.Execute _
         "UPDATE Info SET Version = '" & BUILD & "'," & _
         "                GeneDB = '" & GetFile(dbGene.name) & "', " & _
         "                GeneDBVersion = '" & rs!version & "', " & _
         "                Modify = '" & Format(Now, "dd-mmm-yyyy") & "'," & _
         "                SystemCodes = '" & systemsList & "'"
   Else
      s = GetFile(rawDataFile)                                                'Determine name of ED
      s = Left(s, InStrRev(s, ".") - 1)                                             'Drop extension
      If Right(s, 3) = ".EX" Then
         s = Left(s, Len(s) - 3)
      End If
      dbExpression.Execute _
         "INSERT INTO Info (Title, Version, GeneDB, GeneDBVersion, Modify, SystemCodes)" & _
         "   VALUES ('" & s & "', '" & BUILD & "', '" & GetFile(dbGene.name) & "', '" & _
                     rs!version & "', '" & Format(Now, "dd-mmm-yyyy") & "', '" & systemsList & "')"
   End If
   
   CreateDisplayTable dbExpression

   Close #FILE_RAW_DATA, #FILE_EXCEPTIONS '++++++++++++++++++++++++++++++++++++++++++++++ Finish Up
   If Screen.ActiveForm.lblErrors <> "0" Or warningColumns <> "" Then '+++++++++++++ Problems exist
      If InStr(rawDataFile, ".EX.") <> 0 Then                  'Raw data file was an exception file
         Kill rawDataFile
         exceptionFile = rawDataFile                                'Replace the old exception file
      Else                                                                'Create an exception file
         dot = InStrRev(rawDataFile, ".")
         If dot = 0 Then dot = Len(rawDataFile) + 1
         exceptionFile = Left(rawDataFile, dot - 1) & ".EX" & Mid(rawDataFile, dot)
      End If
      If Dir(exceptionFile) <> "" Then Kill exceptionFile
      Name errorFile As exceptionFile
      
      '======================================================Set Up Warning And Exceptions Messages
      Dim file As String
      Dim warningMsg As String, exceptionMsg As String
      If warningColumns <> "" Then '-------------------------------------------------Warnings Exist
         warningMsg = "Data that exceeds the " & CHAR_DATA_LIMIT & "-character limit " _
                & "was detected in one or more rows in these column(s):" & vbCrLf & _
                warningColumns _
                & "These data have been automatically truncated. If this is not acceptable, " _
                & "correct in your exception file" & vbCrLf & exceptionFile & vbCrLf _
                & "and process the exception file."
      Else
         warningMsg = ""
      End If
      If Screen.ActiveForm.lblErrors <> "0" Then '---------------------------------Exceptions Exist
         exceptionMsg = Screen.ActiveForm.lblErrors & " errors were detected in your " _
                & "raw data. Check the exception file:" _
                & vbCrLf & "   " & exceptionFile _
                & vbCrLf & "Your Expression Dataset has been created. If the errors in " _
                & "the above file are critical, you may correct them and process the " _
                & "exceptions to recreate the Gene Table."
      Else
         exceptionMsg = ""
      End If
      
      If returnExceptions = "" Then '===============================Display Exceptions And Warnings
         If warningMsg <> "" Then '-------------------------------------------------Display Warning
            MsgBox warningMsg, vbInformation + vbOKOnly, "Data Conversion Warning"
         End If
         If exceptionMsg <> "" Then '--------------------------------------------Display Exceptions
            MsgBox exceptionMsg, vbExclamation + vbOKOnly, "Raw Data File Conversion"
         End If
      ElseIf Left(returnExceptions, 9) = "To file: " Then '====================Put Messages In File
         file = Mid(returnExceptions, 10)
         If warningMsg <> "" Then '-------------------------------------------------Display Warning
            Open file For Append As #99
            Print #99, warningMsg
            Print #99, ""
            Close #99
         End If
         If exceptionMsg <> "" Then '--------------------------------------------Display Exceptions
            Open file For Append As #99
            Print #99, exceptionMsg
            Print #99, ""
            Close #99
         End If
      Else '===================================================Return Exceptions To Calling Routine
         returnExceptions = exceptionFile
      End If
'      txtName = ""
'      Kill rawDataFile
'      expressionDirty = False
'      makeDisplayTable = False
'      Kill tempGex
'      tempGex = ""
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ No Problems Exist
      Kill errorFile
   End If
'   FileCopy tempGex, expressionName                                                 'Make permanent
'   Kill tempGex
'   tempGex = ""
   
'   CreateDisplayTable dbExpression
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For GenBanks Still In ED
   Set rs = dbExpression.OpenRecordset("SELECT COUNT(0) as Records FROM Expression WHERE SystemCode = 'G'")
   If rs!records Then
      exceptionMsg = rawDataFile & vbCrLf & "Expression Dataset includes GenBank IDs. " _
             & "For maximum efficiency, these IDs should be changed to IDs from the model " _
             & "organism database or another Gene ID system. Contact your chip maker for " _
             & "alternative gene IDs."
      If Left(returnExceptions, 9) = "To file: " Then '========================Put Messages In File
         file = Mid(returnExceptions, 10)
         Open file For Append As #99
         Print #99, exceptionMsg
         Print #99, ""
         Close #99
      Else
         MsgBox exceptionMsg, vbInformation + vbOKOnly, "GenBanks in Expression Data"
      End If
   End If
   ConvertExpressionData = True
   dbExpression.Close                                         'Opened again in FillExpressionValues
   Set dbExpression = Nothing
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
ConvertError: '==============================================================Trap Conversion Errors
   errors = errors & "Trapped error '" & Err.Description & "' adding to Expression Dataset. "
   Screen.ActiveForm.lblErrors = Screen.ActiveForm.lblErrors + 1
'   Select Case Err '==========================================================Errors In Input Lines
'   Case 3346
'      frmException!lblMessage = "Number of values don't match number of column titles."
'      exception = "number of values"
'   Case 3075
'      frmException!lblMessage = "Invalid character in numeric value."
'      exception = "invalid number"
'   Case Else '------------------------------------------------------------------Unidentified Errors
'      frmException!lblMessage = Err.Description
'      exception = "unknown"
'   End Select
   Resume QuitLine
End Function
Function GetExpressionRow(errorsExists As Integer, geneId As String, systemCode As String, _
                          expValues() As String, Optional delimiter As String = vbTab) As String
   'Entry:  The next line in file #fILE_RAW_DATA, which must be open and set correctly
   '        ErrorsExists   1 if ~Errors~ column exists in raw data. Last column ignored
   '        delimiter      Tab or comma character
   'Return: Blank if successful. Error message or **eof** if not.
   '        geneID      First column, which must be the gene ID.
   '        systemCode  Second column, which must be the cataloging system code.
   '        expValues() All subsequent columns. one-based. If the upper bound of this array
   '                    is zero, there are no expression values to look for.
   '  Controls required on active form:   prgProgress
   
   Dim lin As String                                                                'Line from file
   Dim prevMark As Integer, mark As Integer, lastExpValue As Integer
   
   Do '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Get Line
      lin = RemoveQuotes(InputUnixLine(FILE_RAW_DATA), delimiter)
   Loop While lin = ""                                                          'Ignore blank lines
   
   If lin = "**eof**" Then
      GetExpressionRow = "**eof**"
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   Screen.ActiveForm.prgProgress.value = _
         Min(Seek(FILE_RAW_DATA), Screen.ActiveForm.prgProgress.Max)     'To avoid going beyond Max
         
   expValues(UBound(expValues)) = ""                    'In case of empty column, hopefully Remarks
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Parse The Line
   lastExpValue = 0                                                                'One-based index
   mark = InStr(lin, delimiter)
   If mark = 0 Then mark = Len(lin) + 1
   geneId = Mid(lin, prevMark + 1, mark - prevMark - 1)
   If geneId = "" Then
      GetExpressionRow = "No Gene ID. "
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   prevMark = mark
   mark = InStr(prevMark + 1, lin, delimiter)
   If mark = 0 Then mark = Len(lin) + 1
   systemCode = Mid(lin, prevMark + 1, mark - prevMark - 1)
   If systemCode = "" Then
      GetExpressionRow = "No gene ID type. "
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
'   '===============================================================Make SystemCode Case Insensitive
'      '  First character always uppercase, second if it exists always lowercase
'      Mid(SystemCode, 1, 1) = UCase(Mid(SystemCode, 1, 1))
'      If Len(SystemCode) = 2 Then
'         Mid(SystemCode, 2, 1) = LCase(Mid(SystemCode, 2, 1))
'      End If
   prevMark = mark
   Do Until prevMark > Len(lin)
      mark = InStr(prevMark + 1, lin, delimiter)
      If mark = 0 Then mark = Len(lin) + 1
      lastExpValue = lastExpValue + 1
      If lastExpValue > UBound(expValues) Then
         '  We are beyond the bounds of the array and must exit the function.
         If errorsExists Then                      'There is an error column that should be ignored
            mark = InStr(mark + 1, lin, delimiter)
            If mark Then                                                  'There is a column beyond
               GetExpressionRow = "Too many columns. "
            End If
         Else                                                    'No error column, too many columns
            GetExpressionRow = "Too many columns. "
         End If
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
      expValues(lastExpValue) = Mid(lin, prevMark + 1, mark - prevMark - 1)
'      i = InStr(expValues(lastExpValue), """") '----------------------------------Get Rid Of Quotes
'      Do While i
'         expValues(lastExpValue) = Left(expValues(lastExpValue), i - 1) _
'                                 & Mid(expValues(lastExpValue), i + 1)
'         i = InStr(expValues(lastExpValue), """")
'      Loop
      expValues(lastExpValue) = Trim(TextToSql(expValues(lastExpValue)))
      prevMark = mark
   Loop
   If lastExpValue - errorsExists < UBound(expValues) - 1 Then             'Allow for blank column,
      GetExpressionRow = "Too few columns. "                               'hopefully Remarks
   End If
End Function
Sub CreateDisplayTable(dbExpression As Database) '*************************************************
   '  Entry:   dbExpression   Open Expression database
   '  Controls required on active form:
   '     lblPrgMax
   '     lblPrgValue
   '     prgProgress.value
   '
   Dim rsColorSetLocal As Recordset, rsExpression As Recordset, rsDisplay As Recordset
      '  ColorSet may be open so use a local recordset
      'Use frmExpression.ColorSet instead????????????????????????????????
   Dim tdfDisplay As TableDef, tdf As TableDef, fld As Field, displayExists As Boolean
   Dim setNo As Integer, lastColorSet As Integer, lastCriterion As Integer, expression As String
   Dim expressionColumns(MAX_COLORSETS) As String
   Dim criterions(MAX_COLORSETS, MAX_CRITERIA + 1) As String, criterion As Integer
   Dim criter As String                                              'To assemble criterion to test
   Dim rgbs(MAX_COLORSETS, MAX_CRITERIA + 1) As String, color As Integer
      '  These are ragged arrays for the second subscript. The end of each subarray is indicated
      '  by nothing in criterions(i, j)
   'For GetColorSet()
      Dim labels(MAX_CRITERIA) As String, criteria(MAX_CRITERIA) As String, _
          colors(MAX_CRITERIA) As Long
      Dim notFoundIndex As Integer                       'Index of 'Not found' criterion (last one)
   
   If dbExpression Is Nothing Then Exit Sub                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   If Not dbExpression.Updatable Then
      expression = dbExpression.name
      dbExpression.Close
      Set dbExpression = OpenDatabase(expression)
   End If
      
   Set rsExpression = dbExpression.OpenRecordset("SELECT * FROM Expression")
   If rsExpression.EOF Then Exit Sub                       'No Expression Data >>>>>>>>>>>>>>>>>>>>
   
'   Screen.ActiveForm.lblPrgMax.Visible = True
'   Screen.ActiveForm.lblPrgMax = "Total records: "
'   Screen.ActiveForm.lblPrgValue.Visible = True
'   Screen.ActiveForm.prgProgress.value = 0
   History "Creating Display table"
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Clean Up Anything Open
'   If colorSet <> "" Then
'      colorSet = rsColorSet!colorSet
'      rsColorSet.Close
'   End If
   
   For Each tdf In dbExpression.TableDefs '++++++++++++++++++++++++++++++++++++ Empty Display Table
      If tdf.name = "Display" Then
         displayExists = True
         Exit For
      End If
   Next tdf
   If displayExists Then dbExpression.TableDefs.Delete "Display"
'   dbExpression.Execute "DROP TABLE Display"                         'Doesn't work for some reason
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Criteria And Colors
   Set rsColorSetLocal = dbExpression.OpenRecordset("SELECT * FROM ColorSet")
   setNo = 0
   With rsColorSetLocal                                                             'Reassign SetNo
      Do Until .EOF
         .edit
         !setNo = setNo
         .Update
         GetColorSet dbExpression, rsColorSetLocal, labels, criteria, colors, notFoundIndex
         expressionColumns(setNo) = rsColorSetLocal!column
         For criterion = 0 To notFoundIndex
            criterions(setNo, criterion) = criteria(criterion)
            rgbs(setNo, criterion) = colors(criterion)
         Next criterion
         rgbs(setNo, criterion) = -1
         .MoveNext
         setNo = setNo + 1
      Loop
   End With
   lastColorSet = setNo - 1
   
'For i = 0 To lastColorSet
'   Debug.Print i; "  "; expressionColumns(i)
'   j = 0
'   Do While rgbs(i, j) >= 0
'      Debug.Print , criterions(i, j); ": "; rgbs(i, j)
'      j = j + 1
'   Loop
'Next i
         
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Create New Display Table
   Set tdfDisplay = dbExpression.CreateTableDef("Display")
   With tdfDisplay
      .Fields.Append .CreateField("OrderNo", dbLong)
      .Fields.Append .CreateField("ID", dbText, CHAR_DATA_LIMIT)
      .Fields.Append .CreateField("SystemCode", dbText, 2)
      For setNo = 0 To lastColorSet
         .Fields.Append .CreateField("Value" & setNo, dbMemo)
            '  This field is Memo so that blanks can be stored as well as zeros, which are
            '  valid values
         tdfDisplay.Fields("Value" & setNo).AllowZeroLength = True
         .Fields.Append .CreateField("Color" & setNo, dbLong)
      Next setNo
   End With
   dbExpression.TableDefs.Append tdfDisplay
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Add Values To Display Table
   Dim prevID As String, prevSystemCode As String, prevColor As Long
   Dim systemCodes As String                    'SystemCodes represented in this Expression Dataset
      '  Eg: "|G|Rm|S|"
   Dim values(MAX_COLORSETS) As Single, hues(MAX_COLORSETS) As Long, hue As Long, nextHue As Long
   '  How many different names for "color" can you think of!
   Dim rs As Recordset, sql As String
   Dim idxID As index
   Dim openBracket As Integer, closeBracket As Integer
   Dim sql1 As String
   
   systemCodes = "|"                                                                'No systems yet
   Set rsExpression = dbExpression.OpenRecordset("SELECT * FROM Expression")
   With rsExpression
      .MoveLast
'      Screen.ActiveForm.prgProgress.Max = .RecordCount
'      Screen.ActiveForm.lblPrgMax = "Total records: " & .RecordCount
      SetProgressBase .recordCount, "records"
      .MoveFirst
      Do Until .EOF '=============================================Each Record In Expression Dataset
'         Screen.ActiveForm.prgProgress.value = .AbsolutePosition
         History , .AbsolutePosition
         DoEvents
         If InStr(systemCodes, "|" & !systemCode & "|") = 0 Then                    'New SystemCode
            systemCodes = systemCodes & !systemCode & "|"
         End If
'         If !ID <> prevID And !systemCode <> prevSystemCode Then
         '---------------------------------------------------------------------------Different Gene
            '  Will begin with prevID = "" so force processing of the first gene.
               '  Criterion example:
               '     [Fold change] > 1.2 AND [p value] < 0.05
            sql = "INSERT INTO Display" & _
                  "   VALUES (" & !orderNo & ", '" & !id & "', '" & !systemCode & "'"
            For setNo = 0 To lastColorSet                                           'Each Color Set
               criterion = 0
               If expressionColumns(setNo) = "[None]" Then
                  sql = sql & ", ''"
               Else
                  sql = sql & ", '" & .Fields(expressionColumns(setNo)).value & "'"
               End If
               hue = -1                                          'Default to "No criteria met" flag
               Do Until criterions(setNo, criterion) = ""                           'Each criterion
                  '  The first empty criterion is "No criteria met"
                  If TestCriterion(criterions(setNo, criterion), rsExpression, dbExpression) Then
                     hue = rgbs(setNo, criterion)
                     Exit Do                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                  End If
                  criterion = criterion + 1
               Loop
               If hue = -1 Then                                           'No criterion match found
                  hue = rgbs(setNo, criterion)                         'The "No criteria met" color
               End If
               sql = sql & ", " & hue
            Next setNo
            dbExpression.Execute sql & ")"
               '  Eg: INSERT INTO Display
               '         VALUES ('AA000004', 'G', '', 13421772, '1.27', 65535,
               '                 '1.27', 65535, '1.27', 65535)
         .MoveNext
      Loop
   End With
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Index Display Table
   Set idxID = dbExpression.TableDefs("Display").CreateIndex("ixID")
   With idxID                     'Probably faster to index after filling table?????
      .Fields.Append .CreateField("ID")
      .Fields.Append .CreateField("SystemCode")
   End With
   dbExpression.TableDefs("Display").Indexes.Append idxID
   
   dbExpression.Execute "UPDATE Info SET SystemCodes = '" & systemCodes & "'"
   
'   Screen.ActiveForm.prgProgress.Visible = False
'   Screen.ActiveForm.lblPrgMax.Visible = False
'   Screen.ActiveForm.lblPrgValue.Visible = False

   SetProgressBase
   History
   '  dbExpression is left as writable by this function. I hope this causes no problems
   '  in frmDrafter
End Sub
'********************************************************************************* Decode Color Set
Sub GetColorSet(dbExpression As Database, rsColorSet As Recordset, labels() As String, _
                criteria() As String, colors() As Long, notFoundIndex As Integer)
   '  Entry:
   '     dbExpression   Expression dataset for this window
   '     rsColorSet     In use for this window
   '  Return:
   '     labels()       For each criterion in colorSet, zero based
   '     criteria()     Each of the criteria for the Color Set
   '     colors()       Color for each criterion
   '     notFoundIndex  Index of 'Not found' criterion
   '                    It must exist and always last
   'For GetColorSet()
   '  Dim labels(MAX_CRITERIA) As String, criteria(MAX_CRITERIA) As String, _
          colors(MAX_CRITERIA) As Long
   '  Dim notFoundIndex As Integer                       'Index of 'Not found' criterion (last one)
   '  Call:
   '     GetColorSet dbExpression, rsColorSet, labels, criteria, colors, notFoundIndex
   Dim index As Integer
   Dim pipe1 As Integer, pipe2 As Integer, CrLf As Integer, nextCrLf As Integer
   Dim allCriteria As String                                      'The value of the Criteria column
   '  Typical Criteria value:
   '     Big 8 wk|[8Wk] > 35|123
   '     8 wk|[8Wk] > 20|456
   '     No criteria met||13421772    'Always here
   '     Not found||16777215          'Always here
   '  Each criterion is label|SQL criterion|color and newline
      
   If dbExpression Is Nothing Then                                           'No Expression Dataset
      notFoundIndex = -1
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   If rsColorSet Is Nothing Then                                                      'No Color Set
      notFoundIndex = -1
      GoTo ExitSub                                         'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   allCriteria = rsColorSet!criteria
   CrLf = -1                                                                      'vbCrLf is 2 chrs
   pipe1 = InStr(CrLf + 2, allCriteria, "|")
   Do While pipe1
      nextCrLf = InStr(pipe1 + 1, allCriteria, vbCrLf)
      If nextCrLf = 0 Then nextCrLf = Len(allCriteria) + 1
      pipe2 = InStr(pipe1 + 1, allCriteria, "|")
      labels(index) = Mid(allCriteria, CrLf + 2, pipe1 - CrLf - 2)
      criteria(index) = Mid(allCriteria, pipe1 + 1, pipe2 - pipe1 - 1)
      colors(index) = Val(Mid(allCriteria, pipe2 + 1, nextCrLf - pipe2 - 1))
      index = index + 1
      CrLf = nextCrLf
      pipe1 = InStr(CrLf + 2, allCriteria, "|")
   Loop
   notFoundIndex = index - 1             'Always the last in allCriteria. Indicates end of criteria
ExitSub:
End Sub
'*********************************************************** See If Expression Data Meets Criterion
Function TestCriterion(criterion As String, rsExpression As Recordset, dbExpression As Database) _
         As Boolean
   '  Entry:   Criterion      From ColorSet Table. Eg: [Fold change] > 1.2
   '           rsExpression   Expression table at record being tested
   '           dbExpression   Expression dataset so we can use the Info table
   '  Return:  True if criterion met
   '  All column names must be enclosed in [ ] (Help reinforces this).
   '  Program assembles an SQL query from the criterion with all values, then
   '  tests it by SELECTing from the Info table (because it is a one-record table).
   '  For example, the criterion
   '     [Fold change] > 1.2 AND [p value] < 0.05
   '  is assembled into
   '     rsExpression![Fold change] > 1.2 AND rsExpression![p value] < 0.05
   '  rsExpression is at the current record being tested and so has the
   '  values for that row. If the SELECT results in other than EOF, the
   '  criterion must have evaluated to True.
   Dim sql As String, openBracket As Integer, closeBracket As Integer, s As String
   Dim rs As Recordset
   
   If criterion = "" Then                                'This should always be "No criteria found"
      TestCriterion = True
      Exit Function                                         '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   openBracket = InStr(criterion, "[")                                    'Beginning of column name
   sql = Left(criterion, openBracket - 1)
   Do Until openBracket > Len(criterion)
      closeBracket = InStr(openBracket + 1, criterion, "]")
      s = Mid(criterion, openBracket, closeBracket - openBracket + 1)
      openBracket = InStr(closeBracket + 1, criterion, "[")
      If openBracket = 0 Then openBracket = Len(criterion) + 1
      If openBracket > 1 Then
         If Mid(criterion, openBracket - 1, 1) <> " " Then              'For ill-formed expressions
            sql = sql & " "
         End If
      End If
      If VarType(rsExpression(s).value) = vbNull Then
         sql = sql & "NULL" & Mid(criterion, closeBracket + 1, openBracket - closeBracket - 1)
      Else
         Select Case rsExpression(s).Type
         Case dbSingle, dbDouble
            sql = sql & rsExpression(s).value _
                & Mid(criterion, closeBracket + 1, openBracket - closeBracket - 1)
         Case Else
            sql = sql & "'" & rsExpression(s).value & "'" _
                & Mid(criterion, closeBracket + 1, openBracket - closeBracket - 1)
         End Select
      End If
   Loop
   Set rs = dbExpression.OpenRecordset("SELECT Title FROM Info WHERE " & sql)
   '  Eg: SELECT Title FROM Info
   '          WHERE rsExpression![Fold change] > 1.2       WHERE 2.5 > 1.2
   '            AND rsExpression![p value] < 0.05             AND 0.02 < 0.05
   '  Use the SQL engine to evaluate the criterion.
   TestCriterion = Not rs.EOF
End Function
'*********************************************************** Secondary Column Names for Gene System
Function SecondCols(rsSystems As Recordset, secondaryCols() As String) As Integer
   '  Entry    rsSystem    Record from the systems table for the particular system
   '  Return   The number of secondary columns found, counting from 1
   '           secondaryCols(x, 0)     Names of secondary columns
   '           secondaryCols(x, 1)     "S" if multiple, pipe-surrounded IDs allowed
   '                                   "s" if single, non-pipe-surrounded IDs allowed
   '                                   Anything else if not a secondary column (actually, doesn't
   '                                      get returned)
   Dim pipe As Integer, slash As Integer, column As String, columns As Integer
   
   If rsSystems.EOF Then
      SecondCols = 0
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   pipe = 3                                                                           'End of "ID|"
   slash = InStr(pipe + 1, rsSystems!columns, "\")                                      'Next slash
   Do While slash '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Each Secondary Column
      pipe = InStrRev(rsSystems!columns, "|", slash)                  'Next pipe with slash in unit
      '  If the Columns column had
      '     ID|Whatever|Accession\SBF|Nicknames\sF|Protein|Functions\B|
      '  then the pipe would be the one beginning the |Accession\SBF|
      '  not the |Whatever| unit.
      If UCase(Mid(rsSystems!columns, slash + 1, 1)) = "S" Then '===========This is a search column
         columns = columns + 1                          'One based, used for zero-based index below
         secondaryCols(columns - 1, 0) = Mid(rsSystems!columns, pipe + 1, slash - pipe - 1)
         secondaryCols(columns - 1, 1) = Mid(rsSystems!columns, slash + 1, 1)
            '  Value will be either "S" for multiple IDs surrounded by pipes or "s" for a single
            '  value not surrounded by pipes
      End If
      slash = InStr(slash + 1, rsSystems!columns, "\")                                  'Next slash
   Loop
   SecondCols = columns
End Function
Sub AddSecondIDs(secondaryCols() As String, lastSecondCol As Integer, geneIDs() As String, _
                 genes As Integer, rsSystems As Recordset, dbGene As Database)
   Dim rsSystem As Recordset, pipe As Integer, nextPipe As Integer, sql As String
   
   For i = 0 To lastSecondCol - 1 '===========================================Each Secondary Column
      sql = "SELECT [" & secondaryCols(i, 0) & "] AS Secondary" & _
            "   FROM " & rsSystems!system & _
            "   WHERE ID = '" & geneIDs(genes, 0) & "' ORDER BY [" & secondaryCols(i, 0) & "]"
         'Eg: SELECT Accession FROM SwissProt WHERE ID = 'CALM_HUMAN'
      Set rsSystem = dbGene.OpenRecordset(sql, dbOpenForwardOnly)
      If Not rsSystem.EOF Then '------------------------------------------Found In Secondary Column
         If secondaryCols(i, 1) = "S" Then '_______________________Multiple IDs In Secondary Column
            '  Might return something like "|A1234|B5678|"
            pipe = 1
            Do While pipe < Len(rsSystem!Secondary)
               nextPipe = InStr(pipe + 1, rsSystem!Secondary, "|")
               If nextPipe = 0 Then nextPipe = Len(rsSystem!Secondary) + 1
               genes = genes + 1
               geneIDs(genes, 0) = Mid(rsSystem!Secondary, pipe + 1, nextPipe - pipe - 1)
               geneIDs(genes, 1) = rsSystems!systemCode
               geneIDs(genes, 2) = "S"
               pipe = nextPipe
            Loop
         Else '_______________________________________________________Single ID In Secondary Column
            genes = genes + 1
            geneIDs(genes, 0) = rsSystem!Secondary
            geneIDs(genes, 1) = rsSystems!systemCode
            geneIDs(genes, 2) = "s"
         End If
      End If
   Next i
End Sub
Sub TileWarning()
   MsgBox "In the conversion, some gene IDs had " _
          & "multiple matches. Your MAPP(s) will display these multiple matches " _
          & "as a set of tiled gene boxes. Delete the inappropriate ones and " _
          & "reposition the correct one.", vbInformation + vbOKOnly, _
          "Multiple Match Warning"
End Sub

Function EDToRawData(dbExpression As Database) As String '************ Converts ED To Raw Data File
   '  Entry    An open Expression Dataset
   '  Return   Path to raw data file if successful, empty if not
   '  Result   A tab-delimited file of expression data with the same path as the ED but .txt
   Dim rsExpression As Recordset, tdf As TableDef, fld As Field
   Dim file As String, i As Integer, dataRow As Variant, columns As Integer
   
   If dbExpression Is Nothing Then Exit Function           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   History "Converting to raw data"
   
   file = dbExpression.name
   file = Left(file, InStrRev(file, ".")) & "txt"
   
   Open file For Output As #FILE_RAW_DATA
   
   columns = dbExpression.TableDefs("Expression").Fields.count
   Print #FILE_RAW_DATA, "Gene ID"; vbTab; "SystemCode"; '+++++++++++++++++++++++++++ Column Titles
         '  Excel's bug errors if first column named "ID"
   For i = 3 To columns - 1
      '  Start past OrderNo, ID, and SystemCode
      Print #FILE_RAW_DATA, vbTab; dbExpression.TableDefs("Expression").Fields(i).name;
   Next i
   Print #FILE_RAW_DATA, vbLf;
   
   Set rsExpression = dbExpression.OpenRecordset("SELECT * FROM Expression ORDER BY OrderNo")
   rsExpression.MoveLast
   SetProgressBase rsExpression.recordCount, "records"
   rsExpression.MoveFirst
   Do Until rsExpression.EOF
      dataRow = rsExpression.GetRows(1)
      History , rsExpression.AbsolutePosition
      Print #FILE_RAW_DATA, dataRow(1, 0);                                                 'Gene ID
      For i = 2 To columns - 1
         Print #FILE_RAW_DATA, vbTab; dataRow(i, 0);
      Next i
      Print #FILE_RAW_DATA, vbLf;
   Loop
   Close #FILE_RAW_DATA
   SetProgressBase
   EDToRawData = file
End Function

Sub SwitchIDsInFile(fromSystems() As String, toSystem As String, dbGene As Database, _
                   tempDB As String, Optional changeLog As String = "", _
                   Optional tiles As Boolean = False)
   '  Entry    fromSystems()     Array of systems for systems to be switched.
   '                             UBound is last system
   '           toSystem          System to change source systems to.
   '           dbGene            Open Gene Database.
   '           tempDB            Path of Temporary MAPP to be switched.
   '           changeLog         Path of change log file.
   '           tiles    True if any tiles (multiple substitutions for genes on a MAPP) exist
   '  Return   tiles    True if any tiles (multiple substitutions for genes on a MAPP) made
   '  Calling routine must have dimensioned an array to pass to fromSystems.
   '  Both from and to systems must exist. This checking is done by the calling routine
   '     by choosing from lists of existing systems.
   '  Controls required on active form:   prgProgress, lblOperation, lblDetail
   Dim dbFile As Database
   Dim tdf As TableDef
   Dim sql As String
   Dim rs As Recordset, rsSystems As Recordset
   Dim newIDs(MAX_GENES) As String                                   'Substitute IDs for passed one
   Dim noOfNewIDs As Integer, i As Integer, notes As String
   Dim objKey As Integer, maxKey As Integer
   ReDim fromCodes(MAX_SYSTEMS) As String
   Dim fromCode As String, fromSystem As String, toCode As String
   
   If dbGene Is Nothing Then Exit Sub                      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   ReDim fromCodes(UBound(fromSystems)) As String '+++++++++++++++++++++++++ Find Codes For Systems
   For i = 0 To UBound(fromSystems)
      Set rs = dbGene.OpenRecordset("SELECT SystemCode FROM Systems" & _
                                    "   WHERE System = '" & fromSystems(i) & "'")
      fromCodes(i) = rs!systemCode
   Next i
   Set rs = dbGene.OpenRecordset("SELECT SystemCode FROM Systems" & _
                                 "   WHERE System = '" & toSystem & "'")
   toCode = rs!systemCode
   
   Screen.ActiveForm.lblOperation = "Converting to " & toSystem
   Screen.ActiveForm.lblDetail = ""
   Screen.ActiveForm.lblOperation.visible = True
   Screen.ActiveForm.lblDetail.visible = True
   
   Set dbFile = OpenDatabase(tempDB)
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up Change Log File
   If changeLog = "" Then
      changeLog = Left(dbFile.name, InStrRev(dbFile.name, ".")) & "log"
   End If
   Open changeLog For Output As #FILE_CONVERT_LOG
   
   '==================================================================== Set Up for Converting MAPP
   Print #FILE_CONVERT_LOG, "Gene Label"; vbTab; "Old ID"; vbTab; "Old System"; vbTab; _
         "New ID"; vbTab; " New System"; vbTab; "Other Possibilities: ID[System]"
   Set rs = dbFile.OpenRecordset("SELECT MAX(ObjKey) AS MaxKey FROM Objects")
   maxKey = rs!maxKey                                             'For adding tiled gene objects
   objKey = maxKey
   Set rs = dbFile.OpenRecordset( _
            "SELECT Count(0) AS Records FROM Objects WHERE Type = 'Gene'")
   sql = "SELECT * FROM Objects WHERE Type = 'Gene' ORDER BY ObjKey"
      '  Go through all Gene records to make progress indicator more even
   Screen.ActiveForm.prgProgress.value = 0
   Screen.ActiveForm.prgProgress.Max = Max(rs!records, 1)      'Quick fix. If rs returns no records
                                                  'then Max is zero -- invalid, so we limit it to 1
   Screen.ActiveForm.prgProgress.visible = True
   
   Set rs = dbFile.OpenRecordset(sql) '++++++++++++++++++++++++++++++++++++++++++++++ Go Through DB
   Do Until rs.EOF
      If rs!objKey > maxKey Then Exit Do                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         '  maxKey is the highest objKey value in original MAPP. Beyond this point we are
         '  looking at only tiled, added, objects
      
      '===========================================================Find System Name For Code In MAPP
      Set rsSystems = dbGene.OpenRecordset("SELECT System FROM Systems " & _
                                        "   WHERE SystemCode = '" & rs!systemCode & "'")
      
      If rsSystems.EOF Then                                                'Unidentified In Gene DB
         fromSystem = "Code """ & rs!systemCode & """"                          'Leave it "Code Xx"
         fromCode = ""
      Else
         fromSystem = rsSystems!system
         fromCode = rs!systemCode
      End If
      Print #FILE_CONVERT_LOG, rs!Label; vbTab; rs!id; vbTab; fromSystem; vbTab;
      Screen.ActiveForm.prgProgress.value = rs.AbsolutePosition
      Screen.ActiveForm.lblDetail = rs!id
      DoEvents
      
      '================================================================See If ID Should Be Switched
      noOfNewIDs = 0                                                                     'One based
         For i = 0 To UBound(fromCodes)
            If Trim(fromCode) = fromCodes(i) Then '---------------------------------ID In From List
               '  The Trim() is only there because MAPPBuilder adds a space to the end of the
               '  SystemCode. When MAPPBuilder is fixed, we should remove the Trim().
               noOfNewIDs = SwitchIDs(dbGene, rs!id, fromSystem, toSystem, newIDs())
               Exit For                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            End If
         Next i
      
      If noOfNewIDs > 0 Then '====================================================Found Substitutes
         notes = Dat(rs!notes)                                        'May add unchosen IDs to this
         If noOfNewIDs > 1 Then '--------------------------------------Add Extra IDs To Notes Field
            notes = notes & " «"                                                          'ANSI 171
            For i = 1 To noOfNewIDs - 1
               If toCode = "Ls" Then                       'Special exception for LocusLink Symbols
                  notes = notes & newIDs(i) & "[L] "
               Else
                  notes = notes & newIDs(i) & "[" & toCode & "] "
               End If
            Next i
            notes = notes & "»"                                                           'ANSI 187
            '  Notes field ends up looking like this:
            '     Whatever «A12345[G] SNOUT_AARDVARK[S] »
         End If
         With rs
            .edit
            !id = newIDs(0)
            If toCode = "Ls" Then                          'Special exception for LocusLink Symbols
               !systemCode = "L"
            Else
               !systemCode = toCode
            End If
            !notes = notes
            .Update
         End With
'         dbFile.Execute "UPDATE Objects" & _
'                        "   SET ID = '" & newIDs(0) & "'," & _
'                        "       SystemCode = '" & toCode & "'," & _
'                        "       Notes = '" & notes & "'" & _
'                        "   WHERE ObjKey = " & rs!objKey
            '  Without tiling, this update can be done just like the ED update below
         Dim newCenterX As Single, newCenterY As Single
         newCenterX = rs!centerX
         newCenterY = rs!centerY
         Print #FILE_CONVERT_LOG, newIDs(0); vbTab; toSystem; vbTab;
         
         For i = 1 To noOfNewIDs - 1 '------------------List Subsequent Ones In Other Possibilities
            Print #FILE_CONVERT_LOG, newIDs(i) & "[" & toSystem & "]   ";
            '__________________________________________________________Tile Subsequent Ones On MAPP
            '  If there are any other MOD entries for the given ID, put them in
            '  the conversion log as Other Possibilities and show them on the MAPP
            '  tiled at 50 pixel offsets from the first conversion.
            Dim newWidth As Single, newHeight As Single
            Dim newLabel As Variant, newHead As Variant, newRemarks As Variant
            Dim newLinks As Variant, newNotes As Variant
            objKey = objKey + 1
            newCenterX = newCenterX + 50
            newCenterY = newCenterY + 50
            With rs
               newWidth = !Width
               newHeight = !Height
               newLabel = !Label
               newHead = !head
               newRemarks = !remarks
               newLinks = !links
               newNotes = !notes
               
'                  sql = "INSERT INTO Objects (ObjKey, ID, SystemCode, Type, CenterX, CenterY, Width, Height, Label, Head, Remarks, Links, Notes)"
'                  sql = sql & " VALUES (" & objKey & ", '" & newIDs(i, 0) & "', '" & newIDs(i, 1) & "', 'Gene', " & newCenterX & ", " & newCenterY & ", " & rs!Width & ", " & rs!Height & ", '" & rs!Label & "', '" & rs!Head & "', '" & rs!remarks & "', '" & rs!Links & "', '" & rs!Notes & "')"
'                  dbFile.Execute sql
               .AddNew  'Do this instead of INSERT so that the rsNew loop doesn't lose its place
               !objKey = objKey
               !id = newIDs(i)
               If toCode = "Ls" Then                       'Special exception for LocusLink Symbols
                  !systemCode = "L"
               Else
                  !systemCode = toCode
               End If
               !Type = "Gene"
               !centerX = newCenterX
               !centerY = newCenterY
               !Width = newWidth
               !Height = newHeight
               !Label = newLabel
               !head = newHead
               !remarks = newRemarks
               !links = newLinks
               !notes = newNotes
               .Update
            End With
            tiles = True
         Next i
         Print #FILE_CONVERT_LOG, " "
      Else
         Print #FILE_CONVERT_LOG, "No conversion made"; vbTab; vbTab; " "
      End If
      rs.MoveNext
   Loop
   Close #FILE_CONVERT_LOG
   Screen.ActiveForm.prgProgress.visible = False
   Screen.ActiveForm.lblDetail = ""
   Screen.ActiveForm.lblOperation = ""
   Screen.ActiveForm.lblOperation.visible = False
   Screen.ActiveForm.lblDetail.visible = False
   DoEvents
End Sub

'************************************************* Switch Single Gene ID From One System To Another
Function SwitchIDs(dbGene As Database, id As String, fromSystem As String, toSystem As String, _
                   newIDs() As String) As Integer
   '  Entry    dbGene      Open Gene DB being used for conversion
   '           id          ID to be converted
   '           fromSystem  System to be converted from, ie. for the above ID.
   '           toSystem    System to be converted to
   '  Return   Number of new IDs returned. One based. Max is MAX_GENES.
   '           NewIDs()    Substitute IDs for passed ID. Call must pass an array
   
   Dim noOfNewIDs As Integer                                                             'One based
   Dim tdf As TableDef, rs As Recordset
   Dim uniqueIDs(MAX_GENES) As String, noOfUniqueIDs As Integer, i As Integer, sql As String
   
   If id = "" Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ No ID, Can't Switch
      SwitchIDs = 0
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   Set rs = Nothing '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Substitute IDs
   For Each tdf In dbGene.TableDefs '=============================================Find Proper Table
      If tdf.name = fromSystem & "-" & toSystem Then
         Set rs = dbGene.OpenRecordset("SELECT Related AS ID FROM [" & tdf.name & "]" & _
                                       "   WHERE [Primary] = '" & id & "'")
         If rs.EOF Then                                                    'Check secondary columns
            noOfUniqueIDs = FindUniqueIDs(id, fromSystem, uniqueIDs(), dbGene)
            sql = ""
            For i = 1 To noOfUniqueIDs
               sql = sql & "'" & uniqueIDs(i - 1) & "', "
            Next i
            If sql <> "" Then
               sql = Left(sql, Len(sql) - 2)                                  'Drop off comma space
               Set rs = dbGene.OpenRecordset("SELECT Related AS ID FROM [" & tdf.name & "]" & _
                                             "   WHERE [Primary] IN(" & sql & ")")
            End If
         End If
         Exit For                                          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      ElseIf tdf.name = toSystem & "-" & fromSystem Then
         Set rs = dbGene.OpenRecordset("SELECT [Primary] AS ID FROM [" & tdf.name & "]" & _
                                       "   WHERE Related = '" & id & "'")
         If rs.EOF Then                                                    'Check secondary columns
            noOfUniqueIDs = FindUniqueIDs(id, fromSystem, uniqueIDs(), dbGene)
            sql = ""
            For i = 1 To noOfUniqueIDs
               sql = sql & "'" & uniqueIDs(i - 1) & "', "
            Next i
            If sql <> "" Then
               sql = Left(sql, Len(sql) - 2)                                  'Drop off comma space
               Set rs = dbGene.OpenRecordset("SELECT [Primary] AS ID FROM [" & tdf.name & "]" & _
                                             "   WHERE Related IN(" & sql & ")")
            End If
         End If
         Exit For                                          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
   Next tdf
      
   If Not rs Is Nothing Then '+++++++++++++++++++++++++++++++++++++++++++ Build Return Array Of IDs
      Do Until rs.EOF
         noOfNewIDs = noOfNewIDs + 1
         newIDs(noOfNewIDs - 1) = rs!id                                                 'Zero based
         rs.MoveNext
      Loop
   End If
   SwitchIDs = noOfNewIDs
End Function

