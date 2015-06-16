Attribute VB_Name = "Expressions"
Option Explicit

Type GeneRec
   systemCode As String
   color(MAX_COLORSETS) As Long                                           'Works for expression Set
      '  color(0) is the number of Color Sets
   value As String
   orderNo As Long
   occurrences As Integer
End Type

'********************************************************** Sets Display Properties For Gene Object
Sub ExpressionDisplay(obj As Object, dbGene As Database, dbExpression As Database, _
                      colorIndexes() As Integer, valueIndex As Integer, _
                      Optional systemCodes As String = "")
   '  Entry:
   '     obj            Gene object
   '     dbGene         The Gene Database for the particular drafter window
   '     dbExpression   An open Expression Dataset (or Nothing)
   '     colorIndexes() The numbers of the value columns in the Display table.
   '                    colorIndexes(0) is the number of ColorSets
   '                    Eg: if colorIndexes(2) is 4 then data found in Value4 and Color4 columns
   '     valueIndex     Value column to display
   '     systemCodes    List of SystemCodes from the ED Info table. These are the systems
   '                    to search for related genes. "|~|" means to also search secondary IDs in
   '                    AllRelatedGenes()
   '     cfgColoring    The global variable
   '                       R  Related gene IDs
   '                             All related genes are assembled. Center color is the mode color;
   '                             rim is the second mode. If tie, highest order number in mode
   '                             sets center, second highest in differing color sets rim.
   '                       S  Specific gene ID
   '                             Essentially the same but only primary gene IDs considered.
   '  Return:
   '     obj properties:
   '        value       Value to display (or NULL)
   '        center()    colors of gene object center
   '        rim()       colors of gene object rim
   '        lineStyle   LINE_STYLE_SOLID     Only one gene found in Expression Dataset
   '                    LINE_STYLE_BROKEN    More than one gene exists in the Expression Dataset
   '  Call:
   '     ExpressionDisplay obj, dbGene, dbExpression, colorindexes(), valueIndex, [systemCodes]
   
   '  This should be combined with the Backpage coloring routine. ???????????????????
   
   Dim rsDisplay As Recordset                                                        'Display table
   Dim rsCriterion As Recordset                   'Gene currently in order to test against criteria
   Dim rsInfo As Recordset
   Dim sql As String
   Dim index As Integer, row As Integer
   Static lastColorSet As String, lastExpression As String
   Dim metCriterion As Integer      'Criterion number met by center color. Don't even check for rim
'  Don't call the following from here
'   'For GetColorSet()
'     Static labels(MAX_CRITERIA) As String, criteria(MAX_CRITERIA) As String, _
'          colors(MAX_CRITERIA) As Long
'     Static notFoundIndex As Integer                     'Index of 'Not found' criterion (last one)
   'For AllRelatedGenes()
      Dim genes As Integer
      Dim geneIDs(MAX_GENES, 2) As String
      Dim geneFound As Boolean
      Dim systemsList As String
   'For AllExpressionData()
'      Dim rows As Integer
'      Dim rowIDs(MAX_GENES, 1) As String
'      Dim columns As Integer
'      ReDim columnTitles(dbExpression.TableDefs!Expression.Fields.Count - 4) As String
'      ReDim expData(MAX_GENES, dbExpression.TableDefs!Expression.Fields.Count - 4) As Variant
   Dim AllGenes(MAX_GENES) As GeneRec, gene As Integer, lastGene As Integer             'Zero based
      '  All genes within an IDCode and system
   Dim GeneSets(MAX_GENES) As GeneRec, geneSet As Integer, lastGeneSet As Integer       'Zero based
      '  All related genes
      '  value from lowest orderNo
      '  orderno is lowest orderNo
   Dim i As Integer, j As Integer, s As String, colorSet As Integer, rs As Recordset
   
   If colorIndexes(0) = 0 Then '+++++++++++++++++++++++++++++++++++++++++++ No Expression Data Desired
      obj.centerOrderNo = -1
      obj.centerSystemCode = ""
      obj.rimOrderNo = -1
      obj.rimSystemCode = ""
      obj.value = ""
      obj.color(0) = 1
      obj.color(1) = vbWhite
      obj.rim(0) = 1
      obj.rim(1) = vbWhite
      obj.lineStyle = LINE_STYLE_SOLID                                                  'Solid line
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If valueIndex = -1 Then obj.value = ""                                       'Not showing values
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find All Relevant Genes
   '  Find all relevant genes and apply colors (criteria satisfied) to them
   If cfgColoring = "S" Then '=====================================================Specific Gene ID
      lastGene = -1                                                                     'Zero based
      s = "SELECT OrderNo, ID, SystemCode, Value" & valueIndex & " AS Value" & _
          "   FROM Display" & _
          "   WHERE ID = '" & obj.id & "' AND SystemCode = '" & obj.systemCode & "'" & _
          "   ORDER BY OrderNo"
      Set rsDisplay = dbExpression.OpenRecordset(s)
      Do Until rsDisplay.EOF '-----------------------------------------Each Display Record For Gene
         lastGene = lastGene + 1                                           'One more display record
         AllGenes(lastGene).systemCode = rsDisplay!systemCode
         AllGenes(lastGene).value = rsDisplay!value                         'This is first in order
         AllGenes(lastGene).orderNo = rsDisplay!orderNo
         For i = 1 To colorIndexes(0) '....................................Each ColorSet For That Gene
            s = "SELECT Color" & colorIndexes(i) & " AS Color" & _
                "   FROM Display" & _
                "   WHERE OrderNo = " & rsDisplay!orderNo
            Set rs = dbExpression.OpenRecordset(s)                             'Single Color Column
            AllGenes(lastGene).color(i) = rs!color
            '  color() now contains the colors for each chosen ColorSet
            '  SQL and database structures have no good way of handling arrays of varying
            '  length, so we use this clumsy workaround. Arrays were not necessary in the
            '  first incarnation of the database and we are stuck with that structure.
         Next i
         rsDisplay.MoveNext
      Loop
   Else '=========================================================================All Related Genes
      geneFound = True            'Flag to AllRelatedGenes that just matching to Expression Dataset
      AllRelatedGenes obj.id, obj.systemCode, dbGene, genes, geneIDs, geneFound, , systemCodes
'AllRelatedGenes idIn, systemIn, dbGene, genes, geneIDs, geneFound
      lastGene = -1                                                                     'Zero based
'If obj.objKey = "114" Then Stop
'If obj.title = "RhoA" Then Stop
      For gene = 0 To genes - 1 '--------------------------------------For Each Gene In All Related
         '  This is each distinct gene ID. Display table may have many instances of that ID.
         If valueIndex = -1 Then                                               'Don't display value
            s = "SELECT OrderNo, ID, SystemCode, '' AS [Value]" & _
                "   FROM Display" & _
                "   WHERE ID = '" & geneIDs(gene, 0) & "'" & _
                "      AND SystemCode = '" & geneIDs(gene, 1) & "'" & _
                "   ORDER BY OrderNo"
         Else                                                                        'Display value
            s = "SELECT OrderNo, ID, SystemCode, Value" & valueIndex & " AS [Value]" & _
                "   FROM Display" & _
                "   WHERE ID = '" & geneIDs(gene, 0) & "'" & _
                "      AND SystemCode = '" & geneIDs(gene, 1) & "'" & _
                "   ORDER BY OrderNo"
         End If
         Set rsDisplay = dbExpression.OpenRecordset(s)           'All Display Records For That Gene
         Do Until rsDisplay.EOF '_________________________________Each Display Record For That Gene
            lastGene = lastGene + 1                                        'One more display record
            AllGenes(lastGene).systemCode = rsDisplay!systemCode
'            s = "SELECT Color" & colorIndexes(i) & " AS Color" & _
'                "   FROM Display" & _
'                "   WHERE OrderNo = " & rsDisplay!orderNo
'            Set rs = dbExpression.OpenRecordset(s)                          'Single Color Column
            AllGenes(lastGene).orderNo = rsDisplay!orderNo
            For i = 1 To colorIndexes(0) '..............................Each ColorSet For That Gene
               s = "SELECT Color" & colorIndexes(i) & " AS Color" & _
                   "   FROM Display" & _
                   "   WHERE OrderNo = " & rsDisplay!orderNo
               Set rs = dbExpression.OpenRecordset(s)                          'Single Color Column
               AllGenes(lastGene).color(i) = rs!color
               '  color() now contains the colors for each chosen ColorSet
               '  SQL and database structures have no good way of handling arrays of varying
               '  length, so we use this clumsy workaround. Arrays were not necessary in the
               '  first incarnation of the database and we are stuck with that structure.
            Next i
            If valueIndex <> -1 Then '..................................Display Value For That Gene
               s = "SELECT value" & valueIndex & " AS [Value]" & _
                   "   FROM Display" & _
                   "   WHERE OrderNo = " & rsDisplay!orderNo
               Set rs = dbExpression.OpenRecordset(s)                          'Single Value Column
               AllGenes(lastGene).value = rsDisplay!value
            Else
               AllGenes(lastGene).value = ""
            End If
            rsDisplay.MoveNext
         Loop
      Next gene
   End If
      '  allGenes() contains display colors and value for all instances of all related genes
   
   If lastGene >= 0 Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Genes Found
'If obj.title = "G gamma 4" Then Stop
      For colorSet = 1 To colorIndexes(0) '=======================================For Each ColorSet
         '-----------------------------------------------------------------------Assemble Gene Sets
         '  Put genes in sets within a Color Set based on color.
         '  Keep track of occurrences in each set.
         lastGeneSet = 0
         GeneSets(0).systemCode = AllGenes(0).systemCode '--------------------First Gene In geneSet
         GeneSets(0).value = AllGenes(0).value
         GeneSets(0).color(0) = AllGenes(0).color(colorSet)
            '  Use color(0) of GeneRec to keep track of color in each Color Set
         GeneSets(0).orderNo = AllGenes(0).orderNo
            '  allGenes() in orderNo order so this is the lowest
         GeneSets(0).occurrences = 1
         For gene = 1 To lastGene '___________________Count Occurrences Of That Color In AllGenes()
            '  allGenes(0) is in geneSets(0), so can start search at gene = 1
            For i = 0 To lastGeneSet '.......................................Find GeneSet For Color
               '  Go through geneSet to add occurrence to existing color
               If AllGenes(gene).color(colorSet) = GeneSets(i).color(0) Then
                  GeneSets(i).occurrences = GeneSets(i).occurrences + 1
                  Exit For                                 'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
               End If
            Next i
            If i > lastGeneSet Then '...................Existing Color Not Found, Start New GeneSet
               lastGeneSet = lastGeneSet + 1
               GeneSets(lastGeneSet).systemCode = AllGenes(gene).systemCode
               GeneSets(lastGeneSet).color(0) = AllGenes(gene).color(colorSet)
               GeneSets(lastGeneSet).value = AllGenes(gene).value      'Lowest OrderNo always first
               GeneSets(lastGeneSet).orderNo = AllGenes(gene).orderNo
               GeneSets(lastGeneSet).occurrences = 1
            End If
         Next gene
         
         '==========================================================================Order Gene Sets
         '  Order gene sets by the number of occurrences descending, then order number ascending
         '  in each set.
         '  At the end of this, gene sets are ordered by number of occurrences in the set (number
         '  of genes in a particular color -- that satisfy a particular criterion). If occurrences
         '  in 2 or more sets are tied, the set with the lowest order number comes first. This is
         '  the set with the lowest order number among its genes.
         Dim bottom As Integer
         Dim position As Integer
         Dim temp As GeneRec
         
         For bottom = lastGeneSet To 1 Step -1                        'Few sets, simple bubble sort
            For position = 0 To bottom - 1
               If GeneSets(position + 1).occurrences > GeneSets(position).occurrences _
                  Or (GeneSets(position + 1).occurrences = GeneSets(position).occurrences _
                      And GeneSets(position + 1).orderNo < GeneSets(position).orderNo) Then
                        '  Order by occurrences descending, then orderNo ascending
                  temp = GeneSets(position)
                  GeneSets(position) = GeneSets(position + 1)
                  GeneSets(position + 1) = temp
               End If
            Next position
         Next bottom
         
         '=============================================================================Apply Colors
         obj.color(colorSet) = GeneSets(0).color(0)
         If lastGeneSet > 0 Then '-------------------------------------More Than One Gene Set In ED
            obj.rim(colorSet) = GeneSets(1).color(0)
         Else
            obj.rim(colorSet) = obj.color(colorSet)
         End If
         
         If colorIndexes(colorSet) = valueIndex Then '==================================Apply Value
            '  If the Color Set index is the Value index (eg, Color4 and Value4) then that value is
            '  applied. Like the Color Set, it is the value with the most occurrences.
            If IsNumeric(GeneSets(0).value) Then
               obj.value = CSng(GeneSets(0).value)
            Else
               obj.value = GeneSets(0).value
            End If
         End If
      Next colorSet
      
      If lastGene > 0 Then '-----------------------------------------------More Than One Gene In ED
         obj.lineStyle = LINE_STYLE_BROKEN                                                                    'Solid line
      Else
         obj.lineStyle = LINE_STYLE_SOLID                                                                    'Solid line
      End If
      obj.color(0) = colorIndexes(0)                                 'Number of Color Sets for gene
   Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ No Genes Found
      obj.centerOrderNo = -1
      obj.centerSystemCode = ""
      obj.rimOrderNo = -1
      obj.rimSystemCode = ""
      obj.value = ""
      obj.color(0) = 1
      obj.color(1) = vbWhite
      obj.rim(0) = 1
      obj.rim(1) = vbWhite
      obj.lineStyle = LINE_STYLE_SOLID                                                                    'Solid line
   End If
End Sub
'*************************************************** Find All Expression Data For A Single Gene Set
Sub AllExpressionData(genes As Integer, geneIDs() As String, _
                      dbExpression As Database, rows As Integer, rowIDs() As String, _
                      columns As Integer, columnTitles() As String, expData() As Variant, _
                      colorSetTitles() As String, colorSets As Integer, titleColors() As Long, _
                      geneColors() As Long, orderNos() As Long, _
                      legendPage As String, legendPageTitle As String)
                      
'                      Optional obj As Object = Nothing, Optional centerColor As Long = -1, _
'                      Optional rimColor As Long = -1)
   '  At this point, only called from CreateBackpage().
   '  Entry:   genes                   Number of genes to find expression data for
   '           geneIDs(MAX_GENES, 1)   Gene ID for each gene for which to find expression data
   '              geneIDs(x, 0)        ID. Eg: AA073456
   '              geneIDs(x, 1)        SystemCode. Eg: G
   '                                   This is all the IDs for which to find data. No search for
   '                                   related IDs takes place. If needed that is done first and
   '                                   the results send here.
   '           dbExpression            An open Expression Dataset (or Nothing)
'   '           [obj]                   The gene object
   '  Return:  rows                    Number of rows of Expression Data
   '           rowIDs(MAX_GENES, 1)    Gene ID for each row of expression data
   '              rowIDs(x, 0)         ID
   '              rowIDs(x, 1)         SystemCode
   '           columns                 Number of data columns in Expression Dataset
   '                                   This includes Remarks, which is always the last column
   '           columnTitles(Expression Columns)       Column titles from Expression Dataset
   '           expData(MAX_GENES, Expression Columns) All expression data
   '           colorSetTitles(MAX_COLORSETS)          Titles for each ColorSet
   '           colorSets                              Number of Color Sets
   '           titleColors(MAX_COLORSETS, 1)          Colors for titles (and gene object if
   '                                                  this Color Set was used)
   '              titleColors(x, 0)                   Center color
   '              titleColors(x, 1)                   Rim color
   '           geneColors(MAX_GENES, MAX_COLORSETS)   Gene color within Color Set
   '           orderNos(MAX_GENES)                    Order number for each gene
   '           legendPage              HTML Legend page
   '           legendPageTitle         Name of Legend page to use for file name and URL
'   '           centerColor             Color of the first column
'   '           rimColor                Color of second column
'   '                                      If the obj exists, returned genes are always arranged
'   '                                      with the gene that provides the center color first
'   '                                      and the rim color next.
   'For AllExpressionData()
   '   Dim rows as Integer
   '   Dim rowIDs(MAX_GENES, 1) As String
   '   Dim columns as Integer
   '   Dim colorSetTitles(MAX_COLORSETS) as String
   '   Dim colorSets as Integer
   '   Dim titleColors(MAX_COLORSETS, 1) As Long
   '   Dim geneColors(MAX_GENES, MAX_COLORSETS) As Long
   '   Dim orderNos(MAX_GENES) As Long
   '   Dim legendPage As String
   '   Dim legendPageTitle As String
   '   ReDim columnTitles(dbExpression.TableDefs!Expression.Fields.Count - 4) as String
   '   ReDim expData(MAX_GENES, dbExpression.TableDefs!Expression.Fields.Count - 4) As Variant
   'Call:
   '  AllExpressionData genes, geneIDs, dbExpression, rows, rowIDs, columns, _
                        columnTitles, expData, colorSetTitles, colorSets, _
                        titleColors, geneColors, legendPage, legendPageTitle
                                                           
   '  Rows: individual genes. Columns: data columns from the Expression Dataset including Remarks.
   '  The Backpage requires all expression data. Coloring a Gene object requires only getting
   '  expression data until, or if, a criterion does not match the primary (center) criterion
   '  and only for the specific criterion columns.
   Dim lastColumnTitle As Integer                                                       'Zero based
   Dim index As Integer, col As Integer                                          'Expression values
   Dim row As Integer, colorSet As Integer, criter As Integer
   Dim centerRow As Integer, rimRow As Integer             'Row (gene) that determines these colors
      '  These rows, if set, are moved to the beginning of the returned rowIDs() and
      '  expData() arrays.
   Dim rsExpression As Recordset, rsColorSets As Recordset, rsInfo As Recordset
   Dim links As String
   Dim criterions(MAX_COLORSETS, MAX_CRITERIA) As String
   Dim criterion As Integer
   Dim criterionColors(MAX_COLORSETS, MAX_CRITERIA) As Long
   Dim position As Integer, bottom As Integer
   'For GetColorSet()
     Dim labels(MAX_CRITERIA) As String, criteria(MAX_CRITERIA) As String, _
         colors(MAX_CRITERIA) As Long
     Dim notFoundIndex As Integer                       'Index of 'Not found' criterion (last one)
   '  Call:
   '     GetColorSet dbExpression, rsColorSet, labels, criteria, colors, notFoundIndex

   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Initializations
   rows = 0
'   If Not obj Is Nothing Then
'      If obj.centerOrderNo <> -1 Then                                      'Save room for first row
'         rows = rows + 1
'      End If
'      If obj.rimOrderNo <> -1 Then                                        'Save room for second row
'         rows = rows + 1
'      End If
'   End If
'   centerRow = -1
'   rimRow = -1
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Color Set Data And Legend
   Set rsColorSets = dbExpression.OpenRecordset("ColorSet", dbOpenTable)
   colorSets = 0
   Do Until rsColorSets.EOF
      legendPage = legendPage _
                 & "<p>&nbsp;</p>" & vbCrLf & "<p><a name=""Set" & colorSets & """><b>" _
                 & rsColorSets!colorSet & "</b></a>&nbsp;&nbsp;" _
                 & "<font size=2><a href=""#Top""><b>Top</b></a></font></p>" _
                 & vbCrLf _
                 & "  <table border=1>" & vbCrLf _
                 & "   <tr><td><b>Color&nbsp;&nbsp;&nbsp;</b></td><td><b>Label</b></td>" _
                 & "<td><b>Criterion</b></td></tr>" & vbCrLf
      links = links & "&nbsp;<a href=""#Set" & colorSets & """>" & rsColorSets!colorSet _
            & "</a>&nbsp;"
      colorSetTitles(colorSets) = rsColorSets!colorSet
      GetColorSet dbExpression, rsColorSets, labels, criteria, colors, notFoundIndex
      For criterion = 0 To notFoundIndex - 1                            'Includes "No criteria met"
         legendPage = legendPage _
                    & "   <tr>" & vbCrLf & "      <td align=right bgcolor=""" _
                    & HtmlHexColor(colors(criterion)) & """>&nbsp;</td>" & vbCrLf & "      <td>" _
                    & labels(criterion) & "</td>" & vbCrLf & "      <td>" & criteria(criterion) _
                    & "</td>" & vbCrLf & "   </tr>" & vbCrLf
         criterions(colorSets, criterion) = criteria(criterion)
         criterionColors(colorSets, criterion) = colors(criterion)
      Next criterion
      legendPage = legendPage & "</table>" & vbCrLf
      If rsColorSets!remarks <> "" Then
         legendPage = legendPage & "<p><b>Remarks:</b> " & EmbedLinks(rsColorSets!remarks) _
                    & "</p>" & vbCrLf
      End If
      rsColorSets.MoveNext
      colorSets = colorSets + 1
   Loop
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Finish Legend
   legendPageTitle = GetFile(dbExpression.name)
   legendPageTitle = Left(legendPageTitle, InStrRev(legendPageTitle, ".") - 1)
   '  Back up to beginning
   Set rsInfo = dbExpression.OpenRecordset("Info", dbOpenTable)
   If rsInfo!remarks <> "" Then
      legendPage = "<p><b>Remarks:</b> " & EmbedLinks(rsInfo!remarks) & "</p>" & vbCrLf _
                 & legendPage
   End If
   legendPage = "<p align=center>" & links & "</p>" & vbCrLf & legendPage
   
   legendPage = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">" & vbCrLf _
              & "<html>" & vbCrLf _
              & "<head>" & vbCrLf _
              & "   <title>" & legendPageTitle & " Legend</title>" & vbCrLf _
              & "   <meta name=""generator"" content=""GenMAPP 2.1"">" & vbCrLf _
              & "</head>" & vbCrLf & vbCrLf _
              & "<body>" & vbCrLf _
              & "<h1 align=center><a name=""Top"">" & legendPageTitle & " Legend</a></h1>" _
              & vbCrLf & legendPage
              
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Column Titles
   columns = dbExpression.TableDefs!expression.Fields.count - 3      'Only data columns and Remarks
      '  Remarks are always in the last (columns - 1) column
      '  Fields.Count less OrderNo, ID, and SystemCode
   For j = 0 To columns - 1
      columnTitles(j) = dbExpression.TableDefs!expression.Fields(j + 3).name
   Next j
   
   For index = 0 To genes - 1 '++++++++++++++++++++++++++++++++++++++++++++++++++++ Each Gene (Row)
      Set rsExpression = dbExpression.OpenRecordset( _
            "SELECT * FROM Expression" & _
            "   WHERE ID = '" & geneIDs(index, 0) & "'" & _
            "      AND SystemCode = '" & geneIDs(index, 1) & "'" & _
            "   ORDER BY OrderNo", _
            dbOpenForwardOnly)
            '  Eg: SELECT * FROM Expression where ID = 'M12345' AND SystemCode = 'G'
            '         ORDER BY OrderNo
      Do Until rsExpression.EOF '==================================================IDs For Each Row
         row = rows                                                      'Put in next row of arrays
         rows = rows + 1     'rows starts at zero or beyond colors, set to the next to be processed
                             'At the end of processing, rows is the number of rows
                             'processed (zero-based index + 1)
         rowIDs(row, 0) = geneIDs(index, 0)                                       'ID for this gene
         rowIDs(row, 1) = geneIDs(index, 1)
         orderNos(row) = rsExpression!orderNo
         
         For colorSet = 0 To colorSets - 1 '=============================Colors For Each Row (Gene)
         '  Determines colors for each gene in each Color Set for the top of the display
            criter = 0
            Do
               If TestCriterion(criterions(colorSet, criter), rsExpression, dbExpression) Then
                  geneColors(row, colorSet) = criterionColors(colorSet, criter)
                  Exit Do                                  'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
               End If
               criter = criter + 1
            Loop
         Next colorSet
            
         For col = 0 To columns - 1 '==============================Data For Each Column In Each Row
            expData(row, col) = Dat(rsExpression.Fields(col + 3).value)
         Next col
         rsExpression.MoveNext
      Loop
   Next index
   
'Exit Sub
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Sort Return By Order No
   '           rowIDs(MAX_GENES, 1)    Gene ID for each row of expression data
   '              rowIDs(x, 0)         ID
   '              rowIDs(x, 1)         SystemCode
   '           expData(MAX_GENES, Expression Columns) All expression data
   '           geneColors(MAX_GENES, MAX_COLORSETS)   Gene color within Color Set
   '           orderNos(MAX_GENES)                    Order number for each gene
   Dim swap As Boolean, lTemp As Long, sTemp As String, vTemp As Variant, i As Integer
   
   swap = True
   bottom = rows - 1
   Do While bottom > 0 And swap
      swap = False
      For position = 0 To bottom - 1
         If orderNos(position + 1) < orderNos(position) Then
            lTemp = orderNos(position)
            orderNos(position) = orderNos(position + 1)
            orderNos(position + 1) = lTemp
            sTemp = rowIDs(position, 0)
            rowIDs(position, 0) = rowIDs(position + 1, 0)
            rowIDs(position + 1, 0) = sTemp
            sTemp = rowIDs(position, 1)
            rowIDs(position, 1) = rowIDs(position + 1, 1)
            rowIDs(position + 1, 1) = sTemp
            For i = 0 To columns - 1
               vTemp = expData(position, i)
               expData(position, i) = expData(position + 1, i)
               expData(position + 1, i) = vTemp
            Next i
            For i = 0 To colorSets - 1
               lTemp = geneColors(position, i)
               geneColors(position, i) = geneColors(position + 1, i)
               geneColors(position + 1, i) = lTemp
            Next i
            swap = True
         End If
      Next position
      bottom = bottom - 1
   Loop
            
End Sub

'**************************************************************** Update Earlier Expression Dataset
Function UpdateDataset(dbExpression As Database, Optional dirty As Boolean) As Boolean
   '  This should all be handled by DatasetCurrent() instead of calling it..
   Dim rsInfo As Recordset, fld As Field, ok As Boolean, expression As String
   Dim tbl As TableDef, slash As Integer, oldSource As String, rawFile As String

   If DatasetCurrent(dbExpression) Then
      UpdateDataset = True
      GoTo ExitFunction                                    'Version current vvvvvvvvvvvvvvvvvvvvvvv
   End If
   
'   MsgBox dbExpression.name & vbCrLf & vbCrLf & "was created in a previous version of " _
'          & "GenMAPP and must be converted to the current version before it may be " _
'          & "opened. In the Drafter window, click the ""Tools"" menu, ""Converter"" " _
'          & "option to convert your Expression Dataset.", _
'          vbCritical + vbOKOnly, "Old Expression Dataset"
   GoTo ExitFunction                                       'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   
'Old code. May want to resurrect some of it later.
   '  Version 2 Expression Info table has a GeneDB field for the name of the Gene Database
   '  used to create the Expression Dataset. This adds this field to old expression datasets
   '  if necessary.
   expression = dbExpression.name
   If MsgBox("This Expression Dataset was created with an earlier version of GenMAPP and must " _
             & "be converted to the GenMAPP Version 2 format, after which it will not work " _
             & "with the earlier GenMAPP Version. To convert, click OK, otherwise " _
             & "click Cancel. If you convert, your original Expression Dataset will be " _
             & "saved in its original location with a ""V1_"" prefix.", _
             vbExclamation + vbOKCancel, "Old Expression Dataset Version") = vbCancel Then
      GoTo ExitFunction                                    'Canceled update vvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   frmExpression.MousePointer = vbHourglass
   
   dbExpression.Close                                                'Be sure all tables are closed
   Set dbExpression = Nothing
      
   slash = InStrRev(expression, "\")
   oldSource = Left(expression, slash) & "V1_" & Mid(expression, slash + 1)
   FileCopy expression, oldSource
   
   With frmExpression
      .grdCriteria.visible = False
      .lblOperation.visible = True
      .lblDetail.visible = True
      .prgProgress.visible = True
      .lblPrgMax.visible = True
      .lblPrgValue.visible = True
   End With
   UpdateSingleDataset expression                                      'Will open the database here
   Set dbExpression = OpenDatabase(expression)
   rawFile = EDToRawData(dbExpression)
   ConvertExpressionData rawFile, dbExpression, mappWindow.dbGene
      '  This will leave a .EX file if exceptions exist
   Kill rawFile
   With frmExpression
      .grdCriteria.visible = True
      .lblOperation.visible = False
      .lblDetail.visible = False
      .prgProgress.visible = False
      .lblPrgMax.visible = False
      .lblPrgValue.visible = False
   End With
   
'   frmExpression.expressionDirty = True
   frmExpression.FillExpressionValues expression
'   frmExpression.makeDisplayTable = True
'   frmExpression.colorSetDirty = True                                'Force remake of Display table
'   frmExpression.mnuSaveExp_Click
   UpdateDataset = True
   
ExitFunction:
   frmExpression.MousePointer = vbDefault
End Function

'**************************************************************** Update Earlier Expression Dataset
Function DatasetCurrent(dbExpression As Database) As Boolean
   Dim rsInfo As Recordset
   
   If dbExpression Is Nothing Then
      DatasetCurrent = True
      Exit Function
   End If
   
   Set rsInfo = dbExpression.OpenRecordset("SELECT * FROM Info")
   If rsInfo!version <> "" And InStr(rsInfo!version, "/") = 0 _
         And rsInfo!version >= "20020717" Then
      If Not HasDisplayTable(dbExpression) Then         'Conversion does not create a Display table
         MsgBox "This Expression Dataset requires updating, which could take several minutes. " _
                & "The mouse pointer over your window will be changed to an hourglass. " _
                & "When it becomes an arrow again, the process is finished and you may proceed.", _
                vbExclamation + vbOKOnly, "Creating Display Table"
         frmSplash.MousePointer = vbHourglass    'If starting at command line, Splash screen active
         mappWindow.MousePointer = vbHourglass
         CreateDisplayTable dbExpression
         frmSplash.MousePointer = vbDefault
         mappWindow.MousePointer = vbDefault
      End If
      DatasetCurrent = True
   Else
      MsgBox dbExpression.name & vbCrLf & vbCrLf & "was created in a previous version of " _
             & "GenMAPP and must be converted to the current version before it may be " _
             & "opened. In the Drafter window, click the ""Tools"" menu, ""Converter"" " _
             & "option to convert your Expression Dataset.", _
             vbCritical + vbOKOnly, "Old Expression Dataset"
   End If
'   rsInfo.Close
End Function
Function HasDisplayTable(dbExpression As Database) As Boolean
   Dim tbl As TableDef
   
   If dbExpression Is Nothing Then
      HasDisplayTable = True
      Exit Function
   End If
   
   For Each tbl In dbExpression.TableDefs                          'See if there is a Display table
      If tbl.name = "Display" Then
         HasDisplayTable = True
         Exit For                                          'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
   Next tbl
End Function
