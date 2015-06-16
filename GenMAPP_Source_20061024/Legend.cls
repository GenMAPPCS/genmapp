VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objLegend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarSelectMode As Boolean 'local copy
Public objType As String
Public objKey As String
Public editMode As Boolean
Public centerX As Single                                                           'Left for Legend
Public centerY As Single                                                            'Top for Legend
Public Display As String                                        'Display characteristics for Legend
   '  Stored in ID in MAPP database. (Not currently used)
   '  Defaults to global cfgLegend (see GenMAPP.bas)
Public dataObj As Object                                 'The frmDrafter from which data is derived
   '  This object must always know its Gene Database and Expression Dataset. For the printer, for
   '  example, the outputObj is Printer but the dataObj is frmDrafter.
Public canvas As Object                                'Drawing surface on which originally created

Public minX As Single, minY As Single          'Left edge. Left is reserved for the string function
Public maxX As Single, maxY As Single                     'Farthest X and Y Board points for object
                           '  Used for figuring minimum board size (autosize)
                           
'Private valueColumn As String
Private remarks As String

'*********************************************************************** Create And Draw Legend Box
Public Sub Create(Optional X As Single = -1, Optional Y As Single = -1, _
                  Optional disp As String = "", _
                  Optional dataObject As Object = Nothing)
   '  Entry:   X, Y        Upper-left corner of Legend in twips
   '                       if -1 then don't change current centerX or centerY
   '                       if Y is -2, set to initial Y position that does not hide the Info Area
   '           disp        display parameters. If empty, uses global cfgLegend
   '                       if exists, changes cfgLegend
   '           dataObject  Passed into dataObj property.
   '  Creating a Legend simply passes parameters to DrawObj()
   Dim rsColorSet As Recordset, rsInfo As Recordset                        'From Expression Dataset
   
   If objType = "" Or Y = -2 Then                                         'Legend Doesn't Exist Yet
      X = 0                'Offset its starting position so that new Info Area will be visible, too
'      drawingBoard.fontSize = 9
'      drawingBoard.Font.name = "Arial"
'      Y = 7 * drawingBoard.TextHeight("")
      Y = 3 * TWIPS_CM                                                          'To clear Info Area
   End If
   objType = "Legend"                       'Type doesn't exist until some change made from default
   If X >= 0 Then centerX = X
   If Y >= 0 Then centerY = Y
   If disp <> "" Then '++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Legend From Options
      cfgLegend = disp
   End If
   If dataObject Is Nothing Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++ Set dataObj
      Set dataObj = mappWindow                              'At this point, it is always frmDrafter
    Else
      Set dataObj = dataObject
   End If
   
   If Not dataObj.dbExpression Is Nothing Then '+++++++++++++++++ Set Expression Dataset Parameters
'      If Not dataObj.rsColorSet Is Nothing Then                           'Expression and Color Set
'         Set rsColorSet = dataObj.dbExpression.OpenRecordset( _
'                         "SELECT Column FROM ColorSet" & _
'                         "   WHERE ColorSet = '" & dataObj.rsColorSet!colorSet & "'")
'         If Not rsColorSet.EOF Then
'            valueColumn = rsColorSet!column
'         End If
'      End If
      Set rsInfo = dataObj.dbExpression.OpenRecordset("SELECT * FROM Info")
      remarks = Dat(rsInfo!remarks)
   End If
   DrawObj
End Sub
Public Sub DrawObj(Optional draw As Boolean = True, Optional outputObj As Object = Nothing)
   '  Entry    OutputObj   Object the Legend is drawn in; for frmDrafter it is picDrafter, others
   '                       it is the form or context.
   Const COLOR_SIZE = 24       'Width and height of legend color square when multiplied by fontSize
   Const UNDERLINE_WIDTH = 270               'Width of underscore lines when multiplied by fontSize
   Const LINE_SPACE = 3             'Extra vertical space between lines when multiplied by fontSize
   Const GENEBOX_WIDTH = GENE_MIN_HEIGHT * 4
   Const GENEBOX_LEFT = 200
   Dim i As Integer, s As String, xWidth As Single
   Dim rs As Recordset
   Dim contName As String
   Dim valueColumn As String
   Dim zoom As Single
   Dim cXzoom As Single, cYzoom As Single               'Actual center coordinates on output object
      '  This is the actual center coordinates on the zoomed output object.
      '  cXzoom = centerX * zoom
   Dim drawWidthAdjust As Single
   Dim fontSize As Single                                                 'Font size from cfgLegend
   Dim stripeWidth As Single
   'For GetColorSet()
     Dim labels(MAX_CRITERIA) As String, criteria(MAX_CRITERIA) As String, _
          colors(MAX_CRITERIA) As Long
     Dim notFoundIndex As Integer                        'Index of 'Not found' criterion (last one)
   
   If objType = "" Then Exit Sub                           'Legend not created >>>>>>>>>>>>>>>>>>>>
   
   If outputObj Is Nothing Then '++++++++++++++++++++++++++++++++++++++++++++++++++++ Set outputObj
      Set outputObj = drawingBoard
   End If
   zoom = dataObj.zoom
   If outputObj Is Printer Then '-----------------------------------------------------------Printer
      contName = "Printer"
      drawWidthAdjust = 2.5 * zoom                              'Must adjust for printer resolution
   Else '------------------------------------------------------------------------------------Screen
      contName = outputObj.name                                  'For frmDrafter, name is picLegend
      drawWidthAdjust = 1 * zoom      'Standard DrawWidth for screen, can't go smaller with preview
   End If
   
RelocateObject:
   If callingFunction <> "MnuRedraw" Then '+++++++++++++++++++++++++++++++ Clear Out Any Old Legend
      outputObj.Line (minX * zoom, minY * zoom)-(maxX * zoom, maxY * zoom), vbWhite, BF
                                                           'Draw an opaque box where the Legend was
   End If
   If Not draw Or dataObj.dbExpression Is Nothing Or colorIndexes(0) = 0 _
         Or InStr(cfgLegend, "D") = 0 Then '===================================Don't Display Legend
      minX = 0
      minY = 0
      maxX = 0
      maxY = 0
      GoTo ExitSub                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Apply Zoom
   cXzoom = centerX * zoom: cYzoom = centerY * zoom
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Draw Legend
   Dim XTemp As Single, YTemp As Single
   Dim maxWidth As Single
   
   fontSize = Val(Mid(cfgLegend, InStr(cfgLegend, "F") + 1, _
                      InStr(cfgLegend, "|") - InStr(cfgLegend, "F") - 1))
                      
   With outputObj                                          '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~With
   .Font.Underline = False
   .FontStrikethru = False
   .CurrentY = cYzoom
   i = 0
   maxWidth = 0
   If InStr(cfgLegend, "G") And Not dataObj.dbGene Is Nothing Then '==========Display Gene Database
      .FontName = "Arial"
      FontSizeFloor fontSize * 1.25 * zoom, outputObj
      .FontItalic = True
      .FontBold = True
      .CurrentX = cXzoom                                                               'Left margin
      outputObj.Print "Gene Database"
      .CurrentX = cXzoom
      .DrawWidth = Max(2 * drawWidthAdjust, 0.51)
      outputObj.Line (.CurrentX, .CurrentY)-Step(UNDERLINE_WIDTH * fontSize * zoom, 0), .foreColor
      maxWidth = Max(maxWidth, UNDERLINE_WIDTH * fontSize * zoom)
      .CurrentX = cXzoom
      .CurrentY = .CurrentY + LINE_SPACE * fontSize * zoom
      FontSizeFloor fontSize * zoom, outputObj
      .FontItalic = False
      s = dataObj.dbGene.name
      s = Mid(s, InStrRev(s, "\") + 1)
      outputObj.Print s
      maxWidth = Max(maxWidth, .TextWidth(s))                                  'If this line longer
      .CurrentY = .CurrentY + LINE_SPACE * fontSize * zoom
   End If
   If InChrs(cfgLegend, "ECVR") Then '===========================Display "Expression Dataset" Title
      .FontName = "Arial"
      FontSizeFloor fontSize * 1.25 * zoom, outputObj
      .FontItalic = True
      .FontBold = True
      .CurrentX = cXzoom
      outputObj.Print "Expression Dataset"                        'Don't print heading without name
      .DrawWidth = Max(2 * drawWidthAdjust, 0.51)
      .CurrentX = cXzoom
      outputObj.Line (.CurrentX, .CurrentY)-Step(UNDERLINE_WIDTH * fontSize * zoom, 0), _
                      outputObj.foreColor
      maxWidth = Max(maxWidth, UNDERLINE_WIDTH * fontSize * zoom)
      .CurrentY = .CurrentY + LINE_SPACE * fontSize * zoom
   End If
   If InStr(cfgLegend, "E") Then '==========================================Display Name Of Dataset
      FontSizeFloor fontSize * zoom, outputObj
      .FontItalic = False
      s = dataObj.dbExpression.name
      s = Mid(s, InStrRev(s, "\") + 1)
      s = "Name: " & Left(s, InStr(s, ".") - 1)
      .CurrentX = cXzoom
      outputObj.Print s
      maxWidth = Max(maxWidth, .TextWidth(s))                                  'If this line longer
   End If
   If InStr(cfgLegend, "C") Then '========================================Display Name Of Color Set
      '  Name of color set should not display unless name of dataset displayed
      '  Could dataObj.rsColorSet ever be nothing?
      FontSizeFloor fontSize * zoom, outputObj
      .FontItalic = False
      .CurrentX = cXzoom
      SetRsColorSet
      If colorIndexes(0) = 1 Then '------------------------------------------------Single Color Set
         s = "Color Set: " & dataObj.rsColorSet!colorSet
         outputObj.Print s
         maxWidth = Max(maxWidth, .TextWidth(s))
      ElseIf colorIndexes(0) > 1 Then '-----------------------------------------Multiple Color Sets
         Dim xColor As Single              'Zoomed X position of Color Set title relative to cXzoom
         Dim boxTop As Single                                               'Top of Legend gene box
         Dim firstTitle As Single             'Y position of first title so we can draw lines later
         
         dataObj.rsColorSet.MoveFirst
         outputObj.Print "Color Sets: "  ' & dataObj.rsColorSet!colorSet
         xColor = GENEBOX_LEFT * zoom
         stripeWidth = GENEBOX_WIDTH * zoom / colorIndexes(0)   'Zoomed width of stripe in gene box
         .CurrentX = cXzoom + xColor
         firstTitle = .CurrentY
         Do Until dataObj.rsColorSet.EOF '................................Print Names Of Color Sets
            s = dataObj.rsColorSet!colorSet
            outputObj.Print s
            maxWidth = Max(maxWidth, xColor + .TextWidth(s))
            dataObj.rsColorSet.MoveNext
            xColor = xColor + stripeWidth
            .CurrentX = cXzoom + xColor
         Loop
         boxTop = .CurrentY
         .DrawWidth = Max(0.51, 1 * drawWidthAdjust)                       'DrawWidth must be > 0.5
         outputObj.Line (cXzoom + GENEBOX_LEFT * zoom, boxTop) _
                        -Step(GENEBOX_WIDTH * zoom, GENE_MIN_HEIGHT * zoom), vbGray, BF
                                                                                    'Draws gray box
         For i = 1 To colorIndexes(0) - 1 '..................................Divides Box With Lines
            xColor = cXzoom + GENEBOX_LEFT * zoom + stripeWidth * i
            outputObj.Line (xColor, boxTop)-Step(0, GENE_MIN_HEIGHT * zoom), vbBlack
            outputObj.Line (xColor - stripeWidth / 2, firstTitle + .TextHeight(" ") * i) _
                           -(xColor - stripeWidth / 2, boxTop), vbBlack
               'Line from title down to box
         Next i
         outputObj.Line (cXzoom + GENEBOX_LEFT * zoom, boxTop + .TextHeight("Gene") / 2) _
                        -Step(GENEBOX_WIDTH * zoom, .TextHeight("Gene") * 0.67), vbGray, BF
                                                    'Draws gray box to split lines for "Gene" title
                                                 '.67 adjustment because of no descenders in "Gene"
         .CurrentX = cXzoom + (GENEBOX_LEFT + GENEBOX_WIDTH / 2) * zoom - .TextWidth("Gene") / 2
         .CurrentY = boxTop + GENE_MIN_HEIGHT * zoom / 2 - .TextHeight("Gene") / 2
         outputObj.foreColor = vbBlack
         outputObj.Print "Gene";
         outputObj.Line (cXzoom + GENEBOX_LEFT * zoom, boxTop) _
                        -Step(GENEBOX_WIDTH * zoom, GENE_MIN_HEIGHT * zoom), vbBlack, B
                                                                                'Draws black border
         dataObj.rsColorSet.MoveFirst
      End If
      Set rs = dataObj.dbExpression.OpenRecordset( _
                   "SELECT ColorSet, Column FROM ColorSet WHERE SetNo = " & valueIndex)
      If rs.EOF Then '-------------------------------------------------Determine Value Column Title
         valueColumn = ""
      Else
         valueColumn = rs!colorSet & ": " & rs!column                      'Color Set: column title
      End If
   End If
   If InStr(cfgLegend, "V") And valueColumn <> "" Then '==========Display Name Of Gene Value Column
      '  Name of gene value column should not display unless name of dataset displayed
      FontSizeFloor fontSize * zoom, outputObj
      .FontItalic = False
      s = "Gene Value: " & valueColumn
      .CurrentX = cXzoom
      outputObj.Print s
      maxWidth = Max(maxWidth, .TextWidth(s))
   End If
   If InStr(cfgLegend, "R") And remarks <> "" Then '================================Display Remarks
      '  Remarks will not display unless name of dataset displayed
      FontSizeFloor fontSize * zoom, outputObj
      .FontItalic = False
      .CurrentX = cXzoom
      outputObj.Print remarks
      maxWidth = Max(maxWidth, .TextWidth(remarks))
   End If
   If InStr(cfgLegend, "L") And colorIndexes(0) >= 1 Then '==========================Display Legend
      '  colorIndexes(0) is the number of colors sets being displayed -- 1 thru colorIndexes(0)
      SetRsColorSet
      Do Until dataObj.rsColorSet.EOF '----------------------------------------------Each Color Set
         s = dataObj.rsColorSet!colorSet
         .CurrentY = .CurrentY + LINE_SPACE * fontSize * zoom
         .FontName = "Arial"
         FontSizeFloor fontSize * 1.25 * zoom, outputObj
         .FontItalic = True
         .FontBold = True
         .CurrentX = cXzoom
         s = "Legend: " & s
         outputObj.Print s
         maxWidth = Max(maxWidth, .TextWidth(s))
         .DrawWidth = Max(0.51, 2 * drawWidthAdjust)                       'DrawWidth must be > 0.5
         .CurrentX = cXzoom
         outputObj.Line (.CurrentX, .CurrentY)-Step(UNDERLINE_WIDTH * fontSize * zoom, 0), _
                        .foreColor
         .CurrentY = .CurrentY + LINE_SPACE * fontSize * zoom
         FontSizeFloor fontSize * zoom, outputObj
         .FontItalic = False
         GetColorSet dataObj.dbExpression, dataObj.rsColorSet, labels, criteria, colors, _
                     notFoundIndex
             '  Possibly store this with Legend object if can declare a Private array member of
             '  object??????????????????????
         .CurrentX = cXzoom
         .DrawWidth = Max(0.51, 1 * drawWidthAdjust)
         For i = 0 To notFoundIndex '......................Draw Boxes And Titles For This Color Set
            XTemp = .CurrentX: YTemp = .CurrentY                           'Drawing box moves these
            outputObj.Line (XTemp, YTemp) _
                           -Step(COLOR_SIZE * fontSize * zoom, COLOR_SIZE * fontSize * zoom), _
                           colors(i), BF
            outputObj.Line (XTemp, YTemp) _
                           -Step(COLOR_SIZE * fontSize * zoom, COLOR_SIZE * fontSize * zoom), 0, B
            .CurrentY = YTemp
            .CurrentX = cXzoom + (COLOR_SIZE * fontSize + 40) * zoom
            outputObj.Print labels(i)
            .CurrentX = cXzoom
            maxWidth = Max(maxWidth, (COLOR_SIZE * fontSize + 40) * zoom + .TextWidth(labels(i)))
            .CurrentY = .CurrentY + LINE_SPACE * fontSize * zoom
         Next i
         If Not cfgLegendAllColorSets Then Exit Do
         dataObj.rsColorSet.MoveNext
      Loop
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Boundary Dimensions
   '  These are all nonzoomed to maintain consistency with other objects. Because of text,
   '  etc., maxWidth and .CurrentY are calculated zoomed, so zoom must be removed.
   minX = centerX
   minY = centerY
   maxX = centerX + maxWidth / zoom + 20                                     '20 is extra clearance
   maxY = .CurrentY / zoom + 20                                          'to clear any bottom lines
   If OutsideBoard(Me) Then GoTo RelocateObject                      'Only if minY or minX negative
   End With                                                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End With
ExitSub:
End Sub

Public Sub Move(X As Single, Y As Single, ByVal movePoint As Integer) '*************** Moves Object
   '  X and Y in board not screen coordinates
   '  movePoint   0: Move entire object    1: Center, move
   '  Negative movePoint indicates a shift move
   '  movePoint 0 is used for multiple selections. It is a relative move whereas others move to
   '     an absolute position. Relative moves are always in multiples of the GRID_SIZE.
   '  movePoint 1 is absolute, not relative.
   '  movePoint 1 comes in on grid coordinates.
   Dim wide As Single, high As Single, Shift As Boolean
   
   Screen.ActiveForm.dirty = True
   
   If movePoint < 0 Then                                                                'Shift move
      Shift = True
      movePoint = -movePoint
   End If
   
'   Set oldObj = Copy                               'Make copy of object before moving to allow undo
   
'   DrawObj False                                                                  'Erase old object
   Select Case movePoint
   Case 0 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ For Multiple Selections
      centerX = centerX + X - POINT_SIZE / 2                               'Move is always relative
      centerY = centerY + Y - POINT_SIZE / 2
   Case 1 '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Move
      centerX = X - POINT_SIZE / 2                             'Placing the point snaps to the grid
      centerY = Y - POINT_SIZE / 2
   End Select
   DrawObj
   mappWindow.mnuRedraw_Click
End Sub

'*********************************************************************************** See If Clicked
Public Function CheckClick(ByVal X As Single, ByVal Y As Single) As Boolean
   X = X / mappWindow.zoom: Y = Y / mappWindow.zoom
   If X >= minX Then                             'Check each separately to take less time on nonhit
      If X <= maxX Then
         If Y >= minY Then
            If Y <= maxY Then
               CheckClick = True
            End If
         End If
      End If
   End If
End Function
Public Sub SetEdit(Optional edit As Boolean = True)
   '  activeObject variable always set somewhere else (SetActiveObject)
   If edit Then
      EditPoint Me                                                                          'Moving
      editMode = True
   Else
      EditPoint , , False
      editMode = False
   End If
End Sub
Sub MaxOnBoard()
   '  Maintain compatibility with other objects
End Sub

Public Property Let SelectMode(ByVal vData As Boolean)
    mvarSelectMode = vData
End Property
Public Property Get SelectMode() As Boolean
    SelectMode = mvarSelectMode
End Property


