VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objLump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Rem ///////////////////////////////////////////////////////////////////////////////// Drawn Objects
'  Lumps can be (value of objType):
'     Rectangle   Sizable and rotatable rectangle
'     Oval        Sizable and rotatable oval
'     Brace       Sizable and rotatable in 90 degree increments
'     Label       Text on white background with backpage
'     Gene        Rectangle with value to right and backpage
'     Vesicle     Single circle
'     ProteinA    Five ovals
'     ProteinB    Four filled circles
'     Ribosome    Two filled ovals
'     OrganA      Big oval
'     OrganB      Kidney bean
'     OrganC      Double circle
'     CellA       Oval with oval dot off axis
'     CellB       Rotatable oval for cell nucleus
'     Poly        Regular polygon of any number of sides. Sides stored in DB in SecondX
'  Order of drawing with Open and Redraw (determines order on screen -- last on top):
'     Other objects
'     Genes
'     Labels

Option Explicit

'Const GENE_MIN_HEIGHT = GRID_SIZE * 6                'Defined globally because Legend also uses it
Const GENE_MIN_WIDTH = GRID_SIZE * 10
Const GENE_BORDER_CLEARANCE = 16  'Extra space for border line. Approx half the width of the border
Const GENE_TEXT_CLEARANCE = 40                                 'Extra space between text and border
Const GENE_DESCENDER_ADJUST = 20
Const GENE_VALUE_WIDTH = 432 + GENE_TEXT_CLEARANCE  'Width of 6-character expression value. Must be
                                                'recalculated if font changes. Use statement below:
                                                '         frmDrafter.Font.name = "Arial"
                                                '         frmDrafter.Font.Size = 10
                                                '         frmDrafter.Font.Bold = True
                                                '         frmDrafter.Font.Italic = False
                                                '         debug.Print frmDrafter.TextWidth(1.2345)
                                                '         Stop
Const GENE_RIM = 50                                       'Size of rim color for farthest out value

Const BRACE_THICK = 30                                                  'Thickness of straight part
Const BRACE_RADIUS = 50                                                      'Radius of curved ends
Const TOP_BRACE = 0                                               'Rotation of braces     0 degrees
Const RIGHT_BRACE = 1                                             '                      90 degrees
Const BOTTOM_BRACE = 2                                            '                     180 degrees
Const LEFT_BRACE = 3                                              '                     270 degrees
Const VESICLE_CENTER = 4 * GRID_SIZE                    'Initial center to horizontal edge (radius)
                                                               '  Vesicle.wide is radius of vesicle
Const PROTEINA_RADIUS = 300                                 'Initial radius of an individual circle
                                                                         '  Stored in ProteinA.wide
Const PROTEINA_ASPECT = 1.2                  'Adjust center-to-edge distance by half the difference
                                             'between this And 1
Const PROTEINB_RADIUS = 300                                 'Initial radius of an individual circle
                                                                         '  Stored in ProteinB.wide
Const RIBOSOME_CENTER = 230                                              'Center to horizontal edge
Const RIBOSOME_ASPECT = 2.2
Const ORGANA_CENTER = 9 * GRID_SIZE                       'Center to horizontal edge (major radius)
Const ORGANA_ASPECT = 5
Const ORGANB_CENTER = 9 * GRID_SIZE                                      'Center to horizontal edge
Const ORGANB_HEIGHT = ORGANB_CENTER * 5                                         'From center to top
Const ORGANB_RADIUS = ORGANB_HEIGHT * 2
Const ORGANC_CENTER = 7 * GRID_SIZE                       'Center to horizontal edge (outer radius)
Const CELLA_RADIUS = 30 * GRID_SIZE                                              'Major axis radius
Const CELLA_ASPECT = 4                                  'Divide major axis to get minor axis radius
Const CELLA_DOT_ASPECT = 2.5                    'Divide major axis to get minor axis radius for dot
Const CELLA_ANGLE = -75 * PI / 180
                        '  CellB only a nucleus because angles make lining up rectangles difficult.
                        '  Make rectangles with lines
Const CELLB_RADIUS = 4 * GRID_SIZE             'Major-axis radius. Stored in object's wide property
Const CELLB_ASPECT = 2

'************************************************************************ Control and Process Items
Public editMode As Boolean
Private mvarSelectMode As Boolean
Public canvas As Object                                'Drawing surface on which originally created

'************************************************************* Identification And Descriptive Items
Public objType As String                                                         'Gene, Brace, etc.
Public objKey As String                                            'Unique identifier for an object
   '  String because it is also the identifier in the objLumps collection. Goes in the MAPP
   '  database as long.
Public title As String                                                       'Label for the graphic
Public id As String                                                           'Primary ID for genes
                                                                              'Font name for labels
Public systemCode As String                         'Code for cataloging-system. See database specs
                                                               'Font style for labels chr()
                                                                     '  0  Regular
                                                                     '  1  Bold
                                                                     '  2  Italic
                                                                     '  4  Underline
                                                                     '  8  Strikethru
Public head As String                                                      'Backpage head for genes
Public remarks As String                                                'Backpage remarks for genes
                                                                          'Straight text for Labels
Public notes As String                                   'Notes field, contains other IDs for genes
Public links As String                                                   'Backpage links for Labels
Public mouseOver As String                                               'Mouseover text for object
   '  Currently used only in the Tools and Object toolboxes
   '  Saved for each instance of the object to allow for different mouseovers for each instance
   '  in future

'************************************************************************* Coloring And Value Items
'  All these items are temporary and never stored with the object in the MAPP. They store
'  coloring and value data for genes as related to the application of a specific data and
'  color set. They are stored here to save lookups each time an object is redrawn.
Public centerOrderNo As Long, centerSystemCode As String       'Gene that provides colors to object
Public rimOrderNo As Long, rimSystemCode As String              'OrderNo is unique ID within system
Public lastColorSet As Integer                       'Single gene with multiple colorsets displayed
   '  Last index of color and rim.
Public displayColorSet As Integer                          'Index of Color Set whose value displays
   '  Both indexes above are zero except for genes with multiple Color Sets displayed.
Public value As Variant                                                 'Expression value for genes
Public prevValue As Variant                                     'Previous expression value for gene
   '  This is saved so that when the previous value is wiped out, only those characters are
   '  overwritten instead of leaving a big blank box
Public lineStyle As Integer 'Used for box line style for genes. Uses line style constants for lines
                              '  LINE_STYLE_SOLID = 0          Solid
                              '  LINE_STYLE_BROKEN = 1         Broken
Private oColor(MAX_COLORSETS) As Long                                              'Color of object
   '  Must be accessed by Lets and Gets because it is an array.
   '  color(0) is the number of Color Sets to display in stripes. 1 on are the colors.
   '  -1 indicates no color (transparent)
   '  For genes, this is set to white when the gene is created and changed to the appropriate color
   '  when expression data is applied. color(1) is saved in the Objects table only to maintain
   '  compatibility and common routines with other objects.
   '  For other objects, it is the first color of the object.
Private oRim(MAX_COLORSETS) As Long                                         'Color of rim of object
   '  Color of the rim of the object if different from color, otherwise -1.

'******************************************************************************** Dimensional Items
Public centerX As Single, centerY As Single                    'Coords in twips of center of object
   '  centerX is the number of sides for a polygon (Poly)
Public Size As Single                                      'Font size for labels. Stored in SecondX
Public wide As Single, high As Single                             'Oval:   Major and minor radii
                                                                  'Brace:  wide  Span, tip to tip
                                                                  '        high  picPoint offset
Public rotation As Single                                                'Rotation angle in radians
   '  Zero degrees is right along the X axis.
   '  Rotation moves first in a positive Y direction, i.e. down
   '  formulas for rotating a point
   '        X1 = X * Cos(rotation) - Y * Sin(rotation)
   '        Y1 = X * Sin(rotation) + Y * Cos(rotation)
Public sides As Integer                                        'Sides of polygon. Stored in SecondY
Public minY As Single, minX As Single                        'Least X and Y board points for object
Public maxX As Single, maxY As Single                     'Farthest X and Y board points for object
   '  Used for figuring minimum board size (autosize)
   '  These points are always unzoomed (zoom = 1)
Private zoom As Single

Public Function copy() As Object '******************************************** Exact Copy Of Object
   '  Use only for undo ?????????????????????????????
   Dim newObj As New objLump
   Dim i As Integer
   
   newObj.objKey = objKey
   Set newObj.canvas = canvas
   newObj.editMode = False
   newObj.SelectMode = False   'mvarSelectMode

   newObj.objType = objType
   newObj.centerX = centerX
   newObj.centerY = centerY
   newObj.minY = minY
   newObj.minX = minX
   newObj.maxX = maxX
   newObj.maxY = maxY
   newObj.wide = wide
   newObj.high = high
   newObj.rotation = rotation
   newObj.title = title
   newObj.id = id
   newObj.systemCode = systemCode
   For i = 0 To oColor(0)
      newObj.color(i) = oColor(i)
      newObj.rim(i) = oRim(i)
   Next i
   newObj.Size = Size
   newObj.sides = sides
   newObj.value = value
   newObj.head = head
   newObj.remarks = remarks
   newObj.links = links
   newObj.lineStyle = lineStyle
   newObj.mouseOver = mouseOver
   Set copy = newObj
End Function
Public Sub Restore() '********************************** Copy Object Into Current Object's Location
   '  Uses global oldObj to perform undos
   Dim i As Integer
   
'   mvarID = oldObj.objKey
   Set canvas = oldObj.canvas
   editMode = oldObj.editMode
   mvarSelectMode = oldObj.SelectMode

   objType = oldObj.objType
   centerX = oldObj.centerX
   centerY = oldObj.centerY
   minY = oldObj.minY
   minX = oldObj.minX
   maxX = oldObj.maxX
   maxY = oldObj.maxY
   wide = oldObj.wide
   high = oldObj.high
   rotation = oldObj.rotation
   title = oldObj.title
'   GenMAPP = oldObj.GenMAPP
   id = oldObj.id
   systemCode = oldObj.systemCode
   For i = 0 To oldObj.color(0)
      oColor(i) = oldObj.color(i)
      oRim(i) = oldObj.rim(i)
   Next i
   Size = oldObj.Size
   sides = oldObj.sides
   value = oldObj.value
   head = oldObj.head
   remarks = oldObj.remarks
   links = oldObj.links
   lineStyle = oldObj.lineStyle
   mouseOver = oldObj.mouseOver
End Sub
'************************************************************** Creates and Draws a New Lump Object
Public Sub Create(X As Single, Y As Single, lumpType As String, Optional w As Single = 0, _
                  Optional h As Single = 0, Optional r As Single = 0, Optional c As Long = -1, _
                  Optional options As String = "", _
                  Optional objectKey As Variant, Optional outputObj As Object = Nothing)
   '  Entry:   X, Y     Unzoomed coordinates (may come from picDrafter or MAPP file)
   '                    Snap to grid coordinates
   
   '           w        wide
   '           h        high
   '           r        rotation
   '           c        color       -1 indicates transparent
   '           options  sides for polygon (Poly)
   '  Legacy:  wide, high, rotation, color can be set in code after Create with a drawObj
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Optional Parameters
   If VarType(objectKey) = vbNull Then objectKey = -1
   If IsMissing(objectKey) Then objectKey = -1                  'Previous MAPPs have no object keys
   If outputObj Is Nothing Then
      Set outputObj = drawingBoard
   End If
   
   Set canvas = outputObj                    'Objects are never created for the printer, only drawn
   objType = lumpType
   wide = w
   high = h
   rotation = r
   oColor(0) = 1
   oColor(1) = c
   Select Case objType '----------------------------Assign Status Bar Text For New Object Placement
   Case "Rectangle", "Poly"
      If wide = 0 Then wide = VESICLE_CENTER * 2
      If high = 0 Then high = VESICLE_CENTER * 2
      If objType = "Poly" Then
'         oColor(1) = -1                                                      'Indicates transparent
         If options <> "" Then
            '  Original opening sets sides before creation, others use options
            sides = Val(options)
         End If
      Else
'         oColor(1) = vbBlack
         oColor(1) = -1                                                      'Indicates transparent
      End If
   Case "Oval", "CellB"
      If objType = "CellB" Then
         wide = CELLB_RADIUS                                    'Store here to use in other modules
         high = CELLB_RADIUS / CELLB_ASPECT
         oColor(1) = vbBlack
         If rotation = 0 Then rotation = -75 * PI / 180                           'Initial rotation
         objType = "Oval"                                            'From now on, treat as an oval
      Else
         If wide = 0 Then wide = VESICLE_CENTER
         If high = 0 Then high = VESICLE_CENTER
      End If
'      oColor(1) = -1                                                        'Indicates transparent
   Case "Arc"
      If wide = 0 Then wide = VESICLE_CENTER
      If high = 0 Then high = VESICLE_CENTER
      If oColor(1) = -1 Then oColor(1) = vbBlack
   Case "Brace"
      high = BRACE_RADIUS * 2 + BRACE_THICK              'Offset between center and other picPoints
                     'This is stored in database for compatibility with other lumps but it is
                     'always the above formula
      If wide = 0 Then wide = 800                                          'Default span for braces
      If oColor(1) = -1 Then oColor(1) = vbBlack                                  'Default to black
   Case "Label"
      If title = "" Then title = "Label"
      If id = "" Then id = "Arial"
      If Size = 0 Then Size = 12
      If systemCode = "" Then systemCode = Chr(1)                                  'Default to bold
      If oColor(1) = -1 Then oColor(1) = vbBlack                                  'Default to black
   Case "MAPPTitle"
      If id = "" Then id = "Arial"
      If Size = 0 Then Size = 16
      If systemCode = "" Then systemCode = Chr(1)                                  'Default to bold
      If oColor(1) = -1 Then oColor(1) = vbBlack                                          'Default to black
      title = ""                                                   'MAPP title always created blank
      '  In DrawObj, forces other parameters to 0
   Case "Gene"
      If title = "" Then title = "Gene"
      mouseOver = "Represents genes or gene products.\Put it where you want it on the\" _
                & "Drafting Board"
      oColor(0) = 1                                                      'Transparent to begin with
      oColor(1) = -1                                                     'Transparent to begin with
      oRim(0) = 1                                                                           'No rim
      oRim(1) = -1                                                                          'No rim
   Case "Vesicle"
      If wide = 0 Then wide = VESICLE_CENTER
   Case "ProteinA"
      If wide = 0 Then wide = PROTEINA_RADIUS
   Case "ProteinB"
      If wide = 0 Then wide = PROTEINB_RADIUS
   Case "Ribosome"
   Case "OrganA", "OrganB", "OrganC"
   Case "CellA"
      wide = CELLA_RADIUS                                       'Store here to use in other modules
      high = CELLA_RADIUS / CELLA_ASPECT
      oColor(1) = vbBlack
      If rotation = 0 Then rotation = CELLA_ANGLE                                 'Initial rotation
   Case "CellB"
      wide = CELLB_RADIUS                                       'Store here to use in other modules
      high = CELLB_RADIUS / CELLB_ASPECT
      oColor(1) = vbBlack
      If rotation = 0 Then rotation = -75 * PI / 180                              'Initial rotation
   End Select
   If objectKey = -1 And outputObj.name = "picDrafter" Then                             'New object
      outputObj.container.objKey = outputObj.container.objKey + 1    'Increment picDrafter's objKey
      objKey = outputObj.container.objKey              'Numeric to string conversion automatic here
   Else                                                                  'Object from MAPP database
      objKey = objectKey
   End If
   centerX = GridCoord(X)                              'Set center point already to grid coordinate
   centerY = GridCoord(Y)
   DrawObj True, canvas
   If outputObj.name = "picDrafter" Then outputObj.container.dirty = True
End Sub
Public Sub Duplicate(obj As Object) '******************************************** Duplicates object
   title = obj.title
   id = obj.id
   systemCode = obj.systemCode
   value = obj.value
   lineStyle = obj.lineStyle
   head = obj.head
   remarks = obj.remarks
   links = obj.links
   Size = obj.Size
   sides = obj.sides
   Create obj.centerX + 100, obj.centerY + 100, obj.objType, obj.wide, obj.high, _
          obj.rotation, obj.color(1)
End Sub
'************************************************************************************* Draws Object
Sub DrawObj(Optional draw As Boolean = True, Optional outputObj As Object = Nothing)
   '  draw       True - new object. False - erase object
   '  Each object must figure maxX and maxY properties if drawing on frmDrafter
   '     OutsideBoard called at end of procedure if drawing on frmDrafter
   
   Dim drawRim(MAX_COLORSETS) As Long                                         'rim in this function
   Dim drawColor(MAX_COLORSETS) As Long                                     'color in this function
      '  (0) element is the number of colors
   Dim centerSize As Single                     'Size of objects that display smaller on frmObjects
   Dim centerHeight As Single
   Dim cXzoom As Single, cYzoom As Single               'Actual center coordinates on output object
      '  This is the actual center coordinates on the zoomed output object.
      '  cXzoom = centerX * zoom
   Dim rSin As Single, rCos As Single                            'Sine and cosine of rotation angle
   Dim newWidth As Single
   Dim angle As Single                                              'Loop counter for drawing ovals
   Dim major As Single                                                        'Major radius of oval
'   Dim cont As Object                    'Where object drawn. On object's canvas unless printing
   Dim contName As String                    'Name of canvas. Printer doesn't support Name property
   Dim zoomWidthAdjust As Single                              'Adjustment for zoomWidth for Printer
   Dim zoomWidth As Single, zoomHeight As Single
   Dim stripeWidth As Single
   Dim X As Single, Y As Single                                                            'A point
   Dim Xr As Single, Yr As Single                                                  'A point rotated
   Dim i As Integer, l As Long
   
   If objKey = "" Then Exit Sub                            'Object doesn't exist >>>>>>>>>>>>>>>>>>
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Output Object
   If outputObj Is Nothing Then Set outputObj = drawingBoard
   If outputObj Is frmObjects Then
      zoom = 1
   Else
      zoom = mappWindow.zoom
   End If
   If outputObj Is Printer Then '-----------------------------------------------------------Printer
      contName = "Printer"
      zoomWidthAdjust = 2.5 * zoom                              'Must adjust for printer resolution
   Else '------------------------------------------------------------------------------------Screen
      contName = outputObj.name
      zoomWidthAdjust = 1 * zoom      'Standard zoomWidth for screen, can't go smaller with preview
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Set Minimum Size
   wide = Max(Abs(wide), GRID_SIZE) * Sgn(wide)                    'For negative widths and heights
   high = Max(Abs(high), GRID_SIZE) * Sgn(high)
   outputObj.DrawWidth = Max(0.51, zoomWidthAdjust)                      '0.51 is minimum line size
   If draw Then '-----------------------------------------------------------------------Draw Object
      For i = 0 To oColor(0)
            If mvarSelectMode Then
               drawRim(i) = vbRed
            Else
               drawRim(i) = oRim(i)
            End If
            drawColor(i) = oColor(i)
      Next i
   Else '------------------------------------------------------------------------------Erase Object
      drawRim(1) = vbWhite
      drawColor(1) = vbWhite
      drawColor(0) = 1                                                           'Only single color
   End If
   
RelocateObject:
   '  These parameters are always recalculated with Min and Max functions. Set to values
   '  that will always be replaced
   minX = 1E+38: minY = 1E+38                                             'Make huge to be replaced
   maxX = 0: maxY = 0                                                    'Make small to be replaced
   
   cXzoom = centerX * zoom
   cYzoom = centerY * zoom
   Select Case objType '++++++++++++++++++++++++++++++++++++++++++++++++++++++ Draw Specific Object
      '  Always lay out the object unzoomed first, figure min and max values, then zoom
   Case "Poly" '=================================================================Adjustable Polygon
      '  Polys always have black outlines
      Dim Xpoints(9) As Single, Ypoints(9) As Single
      
      
'      rotation = 45 / 360 * 2 * PI
      'rSin = Sin(rotation): rCos = Cos(rotation)
      '  Base points on 0,0, rotate, and add starting point
      
      wide = Min(wide, MAX_BOARD_WIDTH - 50) 'What is this for ??????????????????
      high = Min(high, MAX_BOARD_HEIGHT - 50)
      rSin = Sin(rotation): rCos = Cos(rotation)
      '  Base points on 0,0, rotate, and add starting point
      
      '-----------------------------------------------------------------------Locate Rotated Points
         '  Also figure X and Y limits
         '  Then apply zoom
      maxX = centerX
      minX = centerX
      maxY = centerY
      minY = centerY
      For i = 0 To sides - 1
         Xpoints(i) = (wide / 2) * Cos(2 * PI * i / sides + rotation) + centerX
         If Xpoints(i) > maxX Then maxX = Xpoints(i)
         If Xpoints(i) < minX Then minX = Xpoints(i)
         Xpoints(i) = Xpoints(i) * zoom
         Ypoints(i) = (wide / 2) * Sin(2 * PI * i / sides + rotation) + centerY
         If Ypoints(i) > maxY Then maxY = Ypoints(i)
         If Ypoints(i) < minY Then minY = Ypoints(i)
         Ypoints(i) = Ypoints(i) * zoom
'outputObj.DrawWidth = 5
'outputObj.PSet (Xpoints(i), Ypoints(i)), vbRed
'outputObj.DrawWidth = 1
      Next i
      If OutsideBoard(Me) Then GoTo RelocateObject                   'Only if minY or minX negative
      
      If oColor(1) >= 0 Then '------------------------------------------------------------Draw Fill
         '  Draw fill first so clean outline can overlay fill
         outputObj.DrawWidth = 5
      
         For X = wide / 2 - outputObj.DrawWidth * 8 / zoom To 1 Step -outputObj.DrawWidth * 8 / zoom
            '  Draw smaller and smaller polygons toward center
            '  Must refigure center, rotation, and zoom, in other words, each point
            outputObj.CurrentX = (X * Cos(2 * PI * (sides - 1) / sides + rotation) + centerX) _
                     * zoom
            outputObj.CurrentY = (X * Sin(2 * PI * (sides - 1) / sides + rotation) + centerY) _
                     * zoom
               '  Start line at last point
            For i = 0 To sides - 1
               If draw Then
                  l = drawColor(1)
               Else
                  l = vbWhite
               End If
               outputObj.Line -((X * Cos(2 * PI * i / sides + rotation) + centerX) * zoom, _
                                (X * Sin(2 * PI * i / sides + rotation) + centerY) * zoom), l
            Next i
         Next X
         outputObj.DrawWidth = 1
      End If

      '------------------------------------------------------------------------------------Draw Rim
      outputObj.CurrentX = Xpoints(sides - 1): outputObj.CurrentY = Ypoints(sides - 1)
      For i = 0 To sides - 1
         outputObj.Line -(Xpoints(i), Ypoints(i)), drawRim(1)
      Next i
   Case "Rectangle" '==========================================================Adjustable Rectangle
      '  Rectangles always have black outlines
      Dim XULr As Single, YULr As Single                          'Upper-left point on rotated axis
      Dim XURr As Single, YURr As Single                         'Upper-right point on rotated axis
      Dim XLLr As Single, YLLr As Single                          'Lower-left point on rotated axis
      Dim XLRr As Single, YLRr As Single                         'Lower-right point on rotated axis
      Dim XBeg As Single, YBeg As Single                                    'Beginning of fill line
      Dim xEnd As Single, YEnd As Single                                          'End of fill line
   
      wide = Min(wide, MAX_BOARD_WIDTH - 50) 'What is this for ??????????????????
      high = Min(high, MAX_BOARD_HEIGHT - 50)
      rSin = Sin(rotation): rCos = Cos(rotation)
      '  Base points on 0,0, rotate, and add starting point
      '-----------------------------------------------------------------------Locate Rotated Points
         X = -wide / 2                                                            'Upper-left point
         Y = -high / 2
         XULr = (X * rCos - Y * rSin)                                             'Rotate the point
         YULr = (X * rSin + Y * rCos)
         XLLr = -XULr + centerX                    'Lower-right equidistant from center, add center
         YLLr = -YULr + centerY
         XULr = XULr + centerX                                            'Add center to upper-left
         YULr = YULr + centerY
         
         X = -X                 'Upper-right point, X equidistant from center, Y the same as before
         XURr = (X * rCos - Y * rSin)                                             'Rotate the point
         YURr = (X * rSin + Y * rCos)
         XLRr = -XURr + centerX                    'Lower-right equidistant from center, add center
         YLRr = -YURr + centerY
         XURr = XURr + centerX                                           'Add center to upper-right
         YURr = YURr + centerY
      If contName = "picDrafter" And draw Then '------------------------------Figure X and Y Limits
         maxX = XULr
         If XURr > maxX Then maxX = XURr
         If XLLr > maxX Then maxX = XLLr
         If XLRr > maxX Then maxX = XLRr
         minX = XULr
         If XURr < minX Then minX = XURr
         If XLLr < minX Then minX = XLLr
         If XLRr < minX Then minX = XLRr
         maxY = YULr
         If YURr > maxY Then maxY = YURr
         If YLLr > maxY Then maxY = YLLr
         If YLRr > maxY Then maxY = YLRr
         minY = YULr
         If YURr < minY Then minY = YURr
         If YLLr < minY Then minY = YLLr
         If YLRr < minY Then minY = YLRr
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
      '----------------------------------------------------------------------------------Apply zoom
         XULr = XULr * zoom: YULr = YULr * zoom
         XURr = XURr * zoom: YURr = YURr * zoom
         XLLr = XLLr * zoom: YLLr = YLLr * zoom
         XLRr = XLRr * zoom: YLRr = YLRr * zoom
      If oColor(1) >= 0 Then '------------------------------------------------------------Draw Fill
         '  Draw fill first so clean outline can overlay fill
         X = -wide * zoom / 2 + 5         '5 for clearance so fill doesn't overlap surrounding line
         For Y = -high * zoom / 2 + 5 To high * zoom / 2 - 5
            XBeg = (X * rCos - Y * rSin) + cXzoom                 'Rotate the points and add center
            YBeg = (X * rSin + Y * rCos) + cYzoom
            xEnd = (-X * rCos - Y * rSin) + cXzoom                'Rotate the points and add center
            YEnd = (-X * rSin + Y * rCos) + cYzoom
            If draw Then
               outputObj.Line (XBeg, YBeg)-(xEnd, YEnd), drawColor(1)
            Else
               outputObj.Line (XBeg, YBeg)-(xEnd, YEnd), vbWhite
            End If
         Next Y
      End If
      '--------------------------------------------------------------------------------Draw Outline
         outputObj.Line (XULr, YULr)-(XURr, YURr), drawRim(1)   'vbBlack
'         outputObj.Line (XURr, YURr)-(XLLr, YLLr), vbBlack
         outputObj.Line -(XLLr, YLLr), drawRim(1)   'vbBlack
         outputObj.Line -(XLRr, YLRr), drawRim(1)   'vbBlack
         outputObj.Line -(XULr, YULr), drawRim(1)   'vbBlack
   Case "Arc" '===============================++++++++=======================Adjustable Oval Or Arc
      If mvarSelectMode Then
         If Not DrawEllipse(drawRim(1), drawRim(1), outputObj) Then GoTo RelocateObject
      Else
         If Not DrawEllipse(drawColor(1), drawColor(1), outputObj) Then GoTo RelocateObject
      End If
   Case "Oval" '=============================================================Adjustable Oval Or Arc
      If Not DrawEllipse(drawRim(1), drawColor(1), outputObj) Then GoTo RelocateObject
   Case "CellB" '====================================================================== Filled Oval
      If Not DrawEllipse(drawColor(1), drawRim(1), outputObj) Then GoTo RelocateObject
   Case "Brace" '========================================================================All Braces
      Const CURVE_STEP = 8                                      'Step for loop drawing curved parts
      Dim extension As Single                                   'Size of each of the straight parts
      Dim dotX As Single, dotY As Single                          'Center of radius for curved ends
      Dim side As Single
      Dim thickness As Single, curvature As Single 'BRACE_THICK and BRACE_RADIUS after zoom applied
      Dim zoomSpan As Single                                               'Span after zoom applied
      
      Dim fudge As Single
      
      fudge = 3 * zoom                   'Add to extensions to meet curves and to curves to thicken
      thickness = BRACE_THICK * zoom
      curvature = BRACE_RADIUS * zoom
      extension = ((wide - BRACE_RADIUS * 4) / 2 - 20) * zoom
      zoomSpan = wide * zoom                                                            'Tip to tip
      
      
      outputObj.DrawWidth = 1                                             'Never prints single line
      
      If draw And SelectMode Then
         drawColor(1) = drawRim(1)                                           'Whole brace in select color
      End If
      Select Case rotation
      Case LEFT_BRACE, RIGHT_BRACE
         If rotation = LEFT_BRACE Then
            side = 1
            If contName = "picDrafter" And draw Then
               maxX = centerX + BRACE_THICK + BRACE_RADIUS * 2          'Center point to end points
                     '  Why in the Hell would anyone put a left brace at the right side of the
                     '  board anyway
               maxY = centerY + wide / 2
               minX = centerX                                                         'Center point
               minY = centerY - wide / 2
               If OutsideBoard(Me) Then GoTo RelocateObject          'Only if minY or minX negative
            End If
         Else
            side = -1
            If contName = "picDrafter" And draw Then
               maxX = centerX                                                         'Center point
               maxY = centerY + wide / 2
               minX = centerX - BRACE_THICK - BRACE_RADIUS * 2          'Center point to end points
               minY = centerY - wide / 2
               If OutsideBoard(Me) Then GoTo RelocateObject          'Only if minY or minX negative
            End If
         End If
         dotX = cXzoom + side * (curvature * 2 + thickness)                                'Top Arc
         dotY = cYzoom - (zoomSpan / 2 - curvature)
         For Y = dotY To dotY - curvature Step -CURVE_STEP / zoom
            X = dotX - side * (Abs(curvature ^ 2 - (dotY - Y) ^ 2) ^ 0.5)
            outputObj.Line (X, Y)-Step(-side * (thickness + fudge), 0), drawColor(1)
'            outputObj.Line (X, Y)-Step(33, 0), drawColor(1)
         Next Y
         outputObj.Line (cXzoom + side * curvature, cYzoom - (curvature + fudge)) _
                        -Step(side * thickness, -(extension + 5 * fudge)), drawColor(1), BF
         dotX = cXzoom                                                              'Top Center Arc
         dotY = cYzoom - curvature
         For Y = dotY To dotY + curvature Step CURVE_STEP / zoom
            X = dotX + side * Abs(curvature ^ 2 - (Y - dotY) ^ 2) ^ 0.5
            outputObj.Line (X, Y)-Step(side * (thickness + fudge), 0), drawColor(1)
         Next Y
         dotX = cXzoom                                                           'Bottom Center Arc
         dotY = cYzoom + curvature
         For Y = dotY To dotY - curvature Step -CURVE_STEP / zoom
            X = dotX + side * Abs(curvature ^ 2 - (dotY - Y) ^ 2) ^ 0.5
            outputObj.Line (X, Y)-Step(side * (thickness + fudge), 0), drawColor(1)
         Next Y
         outputObj.Line (cXzoom + side * curvature, cYzoom + (curvature + fudge)) _
                        -Step(side * thickness, extension + 5 * fudge), drawColor(1), BF
         dotX = cXzoom + side * (curvature * 2 + thickness)                             'Bottom Arc
         dotY = cYzoom + (zoomSpan / 2 - curvature)
         For Y = dotY To dotY + curvature Step CURVE_STEP / zoom
            X = dotX - side * Abs(curvature ^ 2 - (Y - dotY) ^ 2) ^ 0.5
            outputObj.Line (X, Y)-Step(-side * (thickness + fudge), 0), drawColor(1)
         Next Y
      Case TOP_BRACE, BOTTOM_BRACE
         If rotation = TOP_BRACE Then
            side = 1
            If contName = "picDrafter" And draw Then
               maxX = centerX + wide / 2
               maxY = centerY + BRACE_THICK + BRACE_RADIUS * 2          'Center point to end points
                     '  Why in the Hell would anyone put a top brace at the bottom side of the
                     '  board anyway
               minX = centerX - wide / 2
               minY = centerY                                                         'Center point
               If OutsideBoard(Me) Then GoTo RelocateObject          'Only if minY or minX negative
            End If
         Else
            side = -1
            If contName = "picDrafter" And draw Then
               maxX = centerX + wide / 2
               maxY = centerY                                                         'Center point
               minX = centerX - wide / 2
               minY = centerY - BRACE_THICK - BRACE_RADIUS * 2          'Center point to end points
               If OutsideBoard(Me) Then GoTo RelocateObject          'Only if minY or minX negative
            End If
         End If
         dotY = cYzoom + side * (curvature * 2 + thickness)                               'Left Arc
         dotX = cXzoom - (zoomSpan / 2 - curvature)
         For X = dotX To dotX - curvature Step -CURVE_STEP / zoom                       'Center out
            Y = dotY - side * (Abs(curvature ^ 2 - (dotX - X) ^ 2) ^ 0.5)
            outputObj.Line (X, Y)-Step(0, -side * (thickness + fudge)), drawColor(1)
         Next X
         '  Produces straight line from center arc out to left arc. Fudge must make it longer to
         '  reach left arc. Evidently, starting at the center arc without fudge is OK
         outputObj.Line (cXzoom - curvature, cYzoom + side * curvature) _
                        -Step(-(extension + 5 * fudge), side * thickness), drawColor(1), BF
         dotY = cYzoom                                                             'Left Center Arc
         dotX = cXzoom - curvature
         For X = dotX To dotX + curvature Step CURVE_STEP / zoom
            Y = dotY + side * Abs(curvature ^ 2 - (X - dotX) ^ 2) ^ 0.5
            outputObj.Line (X, Y)-Step(0, side * (thickness + fudge)), drawColor(1)
         Next X
         dotY = cYzoom                                                            'Right Center Arc
         dotX = cXzoom + curvature
         For X = dotX To dotX - curvature Step -CURVE_STEP / zoom
            Y = dotY + side * Abs(curvature ^ 2 - (dotX - X) ^ 2) ^ 0.5
            outputObj.Line (X, Y)-Step(0, side * (thickness + fudge)), drawColor(1)
         Next X
         outputObj.Line (cXzoom + curvature, cYzoom + side * curvature) _
                        -Step(extension + 5 * fudge, side * thickness), drawColor(1), BF
         dotY = cYzoom + side * (curvature * 2 + thickness)                              'Right Arc
         dotX = cXzoom + (zoomSpan / 2 - curvature)
         For X = dotX To dotX + curvature Step CURVE_STEP / zoom
            Y = dotY - side * Abs(curvature ^ 2 - (X - dotX) ^ 2) ^ 0.5
            outputObj.Line (X, Y)-Step(0, -side * (thickness + fudge)), drawColor(1)
         Next X
      End Select
   Case "Label", "MAPPTitle" '===============================================================Labels
      With outputObj
         .Font.name = id
         FontSizeFloor Size * zoom, outputObj
         .Font.Bold = Asc(systemCode) And 1
         .Font.Italic = Asc(systemCode) And 2
         .Font.Underline = Asc(systemCode) And 4
         .FontStrikethru = Asc(systemCode) And 8
         wide = .TextWidth(title) / zoom + 2 * GENE_TEXT_CLEARANCE
            '  TextWidth is zoomed, wide is not
         If wide > MAX_BOARD_WIDTH And contName = "picDrafter" Then                 'Label too wide
            FontSizeFloor .Font.Size / (wide / MAX_BOARD_WIDTH) - 0.2, outputObj   'Adjust size
            Size = .Font.Size
            wide = .TextWidth(title) / zoom + 2 * GENE_TEXT_CLEARANCE
            centerX = MAX_BOARD_WIDTH / 2                                                 'Recenter
         End If
         high = .TextHeight(title) / zoom
      End With
      If objType = "MAPPTitle" And draw And contName = "picDrafter" Then           'Always centered
         '  MAPPTitles are always centered.
         '  Should be done only on drawing because erasing must use the old CenterX.
         '  Only frmDrafter needs this routine because other forms never put a title in a
         '  different place or change form sizes.
         If title = "" Then                                                    'Title doesn't exist
            centerX = 0             'Don't allow it to affect any of the board parameters like maxX
            centerY = 0
            wide = 0
            high = 0
         Else
            centerX = Max(wide / 2, outputObj.Width / zoom / 2)
            '  Do not allow title to start before board. This way, if the board is too narrow,
            '  it will be widened by OutsideBoard() without the title being relocated.
            centerY = outputObj.TextHeight(title) / zoom / 2 + GENE_TEXT_CLEARANCE
         End If
         cXzoom = centerX * zoom
         cYzoom = centerY * zoom
      End If
      If contName <> "Printer" Then                         'Printer never needs colored background
         outputObj.Line (cXzoom - wide * zoom / 2, cYzoom - high * zoom / 2) _
                        -Step(wide * zoom, high * zoom), outputObj.BackColor, BF
      End If
      outputObj.CurrentX = cXzoom - outputObj.TextWidth(title) / 2
      outputObj.CurrentY = cYzoom - outputObj.TextHeight(title) / 2
      If Not mvarSelectMode Then outputObj.foreColor = drawColor(1)
      If draw Then
         outputObj.Print title
         If contName = "picDrafter" Then
            minX = centerX - wide / 2
            minY = centerY - high / 2
            maxX = centerX + wide / 2
            maxY = centerY + high / 2
            If OutsideBoard(Me) Then GoTo RelocateObject             'Only if minY or minX negative
         End If
      End If
   Case "Gene" '==============================================================================Genes
'If title = "RhoA" Then Stop
      With outputObj
         .Font.name = "Arial"
         FontSizeFloor 10, outputObj
         .Font.Bold = True
         .Font.Italic = False
         .Font.Underline = False
         .FontStrikethru = False
         If contName = "picDrafter" And draw Then
                                            'New height, width if drawing on frmDrafter or frmTools
            newWidth = GridCoord((.TextWidth(title) + 2 * GENE_TEXT_CLEARANCE _
                                  + 2 * GENE_BORDER_CLEARANCE) / 2) * 2
                       'Size must be double the grid size to align each side to grid
            If newWidth + GENE_VALUE_WIDTH > MAX_BOARD_WIDTH Then                'Too big for board
               FontSizeFloor .Font.Size * MAX_BOARD_WIDTH / (newWidth + GENE_VALUE_WIDTH) - 0.3, _
                             outputObj
               newWidth = GridCoord((.TextWidth(title) + 2 * GENE_TEXT_CLEARANCE _
                                     + 2 * GENE_BORDER_CLEARANCE) / 2) * 2
            End If
            If newWidth + 5 > wide Then       'Fudge factor because genes were resizing differently
               wide = newWidth                'at different times
            End If
            high = GENE_MIN_HEIGHT
                                                                                     'Autosize here
            minX = centerX - wide / 2
            minY = centerY - high / 2
            maxX = centerX + wide / 2 + GENE_VALUE_WIDTH              'Always clear for gene values
            maxY = centerY + high / 2
            If OutsideBoard(Me) Then GoTo RelocateObject             'Only if minY or minX negative
         End If
         .FillStyle = vbFSSolid

         '-------------------------------------------------------------------------------Apply Zoom
         FontSizeFloor 10 * zoom, outputObj
         zoomWidth = wide * zoom                      'Don't refigure width set by frmDrafter ????
         zoomHeight = high * zoom
         
         '-----------------------------------------------------------------------------Dump Old Box
         .DrawStyle = vbSolid
         outputObj.Line (cXzoom - zoomWidth / 2, cYzoom - zoomHeight / 2) _
                   -Step(zoomWidth, zoomHeight), vbWhite, BF
         DrawGeneValue outputObj, False
         If Not draw Then GoTo EndSub                      'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         
         '-----------------------------------------------------------------------------Draw Stripes
         If drawColor(0) = 0 Then
            stripeWidth = zoomWidth
         Else
            stripeWidth = zoomWidth / drawColor(0)
         End If
'If title = "G gamma 4" Then Stop
         For i = 1 To drawColor(0)
            '__________________________________________________________________Determine Rim Colors
            '  The complete box is drawn using the rim color, then overlayed with the smaller
            '  center color, then overlayed by the border, then text is applied, then value
            '  tacked on the end.
            drawRim(i) = oRim(i)
               '  mvarSelectMode affects the border, not the rim on genes
            If drawRim(i) = -1 Then drawRim(i) = drawColor(i)
               '  No rim, make it the same as the center. If no center (-1), neither is drawn.
               
            '_______________________________________________________________________Display Stripes
            If drawRim(i) >= 0 Then '...........................................................Rim
               outputObj.Line (cXzoom - zoomWidth / 2 + stripeWidth * (i - 1), _
                               cYzoom - zoomHeight / 2) _
                         -Step(zoomWidth - stripeWidth * (i - 1), zoomHeight), drawRim(i), BF
                  '  start position: cXzoom - zoomWidth / 2 + stripeWidth * (i - 1)
                  '                  center of gene - half the width + stripeWidth * stripe no - 1
                  '  width: zoomWidth - stripeWidth * (i - 1)
                  '         width of gene minus stripes already drawn.
                  '  Stripe 1 displays in entire gene box. Stripe 2 starts one stripe width from
                  '  the left to the end of the box leaving stripe 1 showing, and so forth.
            End If
            outputObj.Line (cXzoom - zoomWidth / 2 + stripeWidth * (i - 1), _
                            cYzoom - zoomHeight / 2) _
                      -Step(0, zoomHeight), vbBlack
            If drawColor(i) >= 0 Then '......................................................Center
               outputObj.Line (cXzoom - zoomWidth / 2 + stripeWidth * (i - 1), _
                               cYzoom - zoomHeight / 2 + GENE_RIM * zoom) _
                         -Step(zoomWidth - stripeWidth * (i - 1), _
                               zoomHeight - GENE_RIM * 2 * zoom), drawColor(i), BF
            End If
         Next i
         
         '------------------------------------------------------------------------------Draw Border
         If lineStyle = LINE_STYLE_BROKEN Then                                              'Broken
            .DrawStyle = vbDot
         Else
            .DrawStyle = vbSolid
         End If
         .FillStyle = vbFSTransparent
         If mvarSelectMode Then
            outputObj.Line (cXzoom - zoomWidth / 2, cYzoom - zoomHeight / 2) _
                      -Step(zoomWidth, zoomHeight), vbRed, B
         Else
            outputObj.Line (cXzoom - zoomWidth / 2, cYzoom - zoomHeight / 2) _
                      -Step(zoomWidth, zoomHeight), vbBlack, B
         End If

         '------------------------------------------------------------------------------------Title
         If InStr(notes, "«") <> 0 Then .Font.Italic = True
            '  Don't resize because of italic but do center
         .CurrentX = cXzoom - .TextWidth(title) / 2
         .CurrentY = cYzoom - .TextHeight(title) / 2
         outputObj.foreColor = vbBlack
         outputObj.Print title;
         .Font.Italic = False
      End With
      DrawGeneValue outputObj, False
      If oColor(1) <> vbWhite And value <> "" Then '-------------------------------------Gene Value
         DrawGeneValue outputObj
      End If
'      If Not draw Then 'Erase remainder of old gene. Printer & preview will always have draw = true
'         outputObj.Line (cXzoom - wide * zoom / 2, cYzoom - high * zoom / 2) _
'                        -Step((wide + 400) * zoom, high * zoom), &HFFFFFF, BF
'            '  Very interim step. The wide + 400 erases the gene value also  ???????????????????
'      End If
      outputObj.FillStyle = vbFSTransparent
      outputObj.DrawStyle = vbSolid
   Case "Vesicle"
      outputObj.Circle (cXzoom, cYzoom), wide * zoom, outputObj.foreColor
      If contName = "picDrafter" And draw Then
         minX = centerX - wide
         minY = centerY - wide
         maxX = centerX + wide
         maxY = centerY + wide
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
   Case "ProteinA" '=================================================================5 open circles
      '  Don't need these per Kam
      zoomWidth = wide * zoom
      outputObj.DrawWidth = Max(0.51, 5 * zoomWidthAdjust)
      outputObj.Circle (cXzoom - zoomWidth * 0.7, cYzoom - zoomWidth * 0.5), zoomWidth, _
                       outputObj.foreColor, , , PROTEINA_ASPECT
      outputObj.Circle (cXzoom, cYzoom - zoomWidth * 0.4), zoomWidth, outputObj.foreColor, , , _
                       PROTEINA_ASPECT
      outputObj.Circle (cXzoom + zoomWidth * 0.7, cYzoom - zoomWidth * 0.5), zoomWidth, _
                       outputObj.foreColor, , , PROTEINA_ASPECT
      outputObj.Circle (cXzoom - zoomWidth * 0.7, cYzoom + zoomWidth * 0.5), zoomWidth, _
                       outputObj.foreColor, , , PROTEINA_ASPECT
      outputObj.Circle (cXzoom + zoomWidth * 0.7, cYzoom + zoomWidth * 0.5), zoomWidth, _
                       outputObj.foreColor, , , PROTEINA_ASPECT
      outputObj.DrawWidth = Max(0.51, zoomWidthAdjust)
   Case "ProteinB" '===============================================================4 Shaded Circles
      If draw Then '-----------------------------------------------------------------Set Fill Color
         outputObj.FillColor = &H888888
         If contName = "picDrafter" Then
            minX = centerX - wide * 1.79
            maxX = centerX + wide * 1.79
            minY = centerY - wide * 1.59
            maxY = centerY + wide * 1.59
            If OutsideBoard(Me) Then GoTo RelocateObject             'Only if minY or minX negative
         End If
      Else '---------------------------------------------------------------------------Erase Object
         outputObj.FillColor = vbWhite
      End If
      zoomWidth = wide * zoom                                                           'Apply Zoom
      outputObj.DrawWidth = Max(5 * zoomWidthAdjust, 1)
      outputObj.FillStyle = vbFSSolid
      outputObj.Circle (cXzoom - zoomWidth * 0.7, cYzoom), zoomWidth, drawRim(1)
      outputObj.Circle (cXzoom + zoomWidth * 0.7, cYzoom), zoomWidth, drawRim(1)
      outputObj.Circle (cXzoom, cYzoom - zoomWidth * 0.5), zoomWidth, drawRim(1)
      outputObj.Circle (cXzoom, cYzoom + zoomWidth * 0.5), zoomWidth, drawRim(1)
      outputObj.FillColor = vbWhite
      outputObj.FillStyle = vbFSTransparent
      outputObj.DrawWidth = Max(0.51, zoomWidthAdjust)
   Case "Ribosome"
      If draw Then '-----------------------------------------------------------------Set Fill Color
         outputObj.FillColor = vbBlack
         If contName = "picDrafter" Then
            minX = centerX - RIBOSOME_CENTER * 1.1
            maxX = centerX + RIBOSOME_CENTER * 1.1
            minY = centerY - RIBOSOME_CENTER * 1.49
            maxY = centerY + RIBOSOME_CENTER * 1.49
            If OutsideBoard(Me) Then GoTo RelocateObject             'Only if minY or minX negative
         End If
      Else '---------------------------------------------------------------------------Erase Object
         outputObj.FillColor = vbWhite
      End If
      outputObj.DrawWidth = Max(0.51, 4 * zoomWidthAdjust)                              'Apply Zoom
      outputObj.FillStyle = vbFSSolid
      outputObj.Circle _
         (cXzoom - (RIBOSOME_CENTER - RIBOSOME_CENTER * 0.8 / RIBOSOME_ASPECT) * zoom, cYzoom), _
         RIBOSOME_CENTER * 0.8 * zoom, drawRim(1), , , RIBOSOME_ASPECT
      outputObj.Circle _
         (cXzoom + (RIBOSOME_CENTER - RIBOSOME_CENTER * 1.4 / RIBOSOME_ASPECT) * zoom, cYzoom), _
         RIBOSOME_CENTER * 1.4 * zoom, drawRim(1), , , RIBOSOME_ASPECT
      outputObj.FillColor = vbWhite
      outputObj.FillStyle = vbFSTransparent
      outputObj.DrawWidth = Max(0.51, zoomWidthAdjust)
   Case "OrganA"
      Select Case contName
      Case "frmObjects"
         outputObj.Circle (cXzoom, cYzoom), ORGANA_CENTER * ORGANA_ASPECT / 3, _
                           drawRim(1), , , ORGANA_ASPECT
      Case Else
         If draw And contName = "picDrafter" Then
            minX = centerX - ORGANA_CENTER
            maxX = centerX + ORGANA_CENTER
            minY = centerY - ORGANA_CENTER * ORGANA_ASPECT
            maxY = centerY + ORGANA_CENTER * ORGANA_ASPECT
            If OutsideBoard(Me) Then GoTo RelocateObject             'Only if minY or minX negative
         End If
         outputObj.Circle (cXzoom, cYzoom), ORGANA_CENTER * ORGANA_ASPECT * zoom, drawRim(1), , , _
                          ORGANA_ASPECT
      End Select
   Case "OrganB" '===================================================================== Kidney Bean
      Dim outer As Single, inner As Single, littleRadius As Single
      Dim lStart As Single, lEnd As Single
      Dim littleX As Single, littleY As Single, radiusMult As Single
      
      Select Case contName
      Case "frmObjects"
         centerSize = ORGANB_CENTER / 3
         centerHeight = ORGANB_HEIGHT / 3
      Case Else
         centerSize = ORGANB_CENTER '* zoom
         centerHeight = ORGANB_HEIGHT ' * zoom
      End Select
      outer = 12.6: inner = 10.3
      radiusMult = 3.9
      littleRadius = centerSize * 0.75
      littleX = centerSize - littleRadius        'Center of end circles (offset from object center)
      littleY = centerHeight - littleRadius
      lStart = 260: lEnd = 165                                'Start and end angles for end circles
      If draw And contName = "picDrafter" Then
         minX = centerX - centerSize
         maxX = centerX + centerSize
         minY = centerY - littleY - littleRadius
         maxY = centerY + littleY + littleRadius
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
      
      '----------------------------------------------------------------------------------Apply Zoom
      centerSize = centerSize * zoom
      centerHeight = centerHeight * zoom
      littleRadius = littleRadius * zoom
      littleX = littleX * zoom                   'Center of end circles (offset from object center)
      littleY = littleY * zoom
      
      '-------------------------------------------------------------------------------Draw Big Arcs
      outputObj.Circle (cXzoom + radiusMult * centerHeight, cYzoom), _
                       radiusMult * centerHeight + centerSize / 4, drawRim(1), _
                       (180 - inner) * PI / 180, (180 + inner) * PI / 180
      outputObj.Circle (cXzoom + radiusMult * centerHeight, cYzoom), _
                       radiusMult * centerHeight + centerSize, drawRim(1), _
                       (180 - outer) * PI / 180, (180 + outer) * PI / 180
      
      '----------------------------------------------------------------------------Draw End Circles
      outputObj.Circle (cXzoom + littleX, cYzoom - littleY), littleRadius, drawRim(1), _
                       lStart * PI / 180, lEnd * PI / 180
      outputObj.Circle (cXzoom + littleX, cYzoom + littleY), littleRadius, drawRim(1), _
                       (360 - lEnd) * PI / 180, (360 - lStart) * PI / 180
   Case "OrganC" '====================================================================Double Circle
      If draw And contName = "picDrafter" Then
         minX = centerX - ORGANC_CENTER
         maxX = centerX + ORGANC_CENTER
         minY = centerY - ORGANC_CENTER
         maxY = centerY + ORGANC_CENTER
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
      outputObj.Circle (cXzoom, cYzoom), ORGANC_CENTER * zoom, drawRim(1)
      outputObj.Circle (cXzoom, cYzoom), ORGANC_CENTER * 0.9 * zoom, drawRim(1)
   Case "CellA" '===================================================Oval With Filled Oval At Center
      '  Cell with nucleus
      Dim minXTemp As Single, minYTemp As Single, maxXTemp As Single, maxYTemp As Single
      
      '  X1, Y1                                                           'Points on unrotated axis
      '  X2, Y2                                                             'Points on rotated axis
      '  X3, Y3                                                             'Points on rotated axis
      '  rSin, rCos                                              'Sine and cosine or rotation angle
'      Dim dotX As Single, dotY As Single                           'Center of filled oval in middle
   
      Select Case contName
      Case "frmObjects"
         major = CELLA_RADIUS / 3
      Case Else
         major = CELLA_RADIUS
      End Select
'      If draw And contName = "picDrafter" Then
'         '  Multipliers are to adjust for the angle of the oval. They are approximations, not
'         '  calculated figures. If the angle changes, these multipliers will have to reestimated.
'         minX = CenterX - major / CELLA_ASPECT * 1.4
'         maxX = CenterX + major / CELLA_ASPECT * 1.4
'         minY = CenterY - major * 0.97
'         maxY = CenterY + major * 0.97
'         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
'      End If
      
      '-----------------------------------------------------------------------------Draw Outer Oval
      If Not DrawEllipse(drawRim(1), , outputObj, , , major, major / CELLA_ASPECT) Then
         GoTo RelocateObject
      End If
      minXTemp = minX                       'Save these because drawing inner oval will change them
      minYTemp = minY
      maxXTemp = maxX
      maxYTemp = maxY
      
      '-----------------------------------------------------------------------------Draw Inner Oval
      If Not DrawEllipse(drawRim(1), oColor(1), outputObj, centerX - (major / 3) * Cos(CELLA_ANGLE), _
                         centerY - (major / 3) * Sin(CELLA_ANGLE), _
                         major / 3, major / CELLA_DOT_ASPECT / 3) Then
         GoTo RelocateObject             'This should never happen since inner is inside outer oval
      End If
      minX = minXTemp                                    'Restore the parameters for the outer oval
      minY = minYTemp
      maxX = maxXTemp
      maxY = maxYTemp
   End Select

EndSub:
   outputObj.foreColor = vbBlack                  'Reset in case something follows that doesn't set
   If Not oldObj Is Nothing Then                                             'Keeping track of undo
      Set newObj = Me                               'Address of current object that might be undone
'      frmDrafter.mnuUndo.Enabled = True
   End If
End Sub
Function DrawEllipse(Optional rimColor As Long = -1, Optional insideColor As Long = -1, _
                     Optional outputObj As Object = Nothing, _
                     Optional ByVal Xc As Single = -1, Optional ByVal Yc As Single = -1, _
                     Optional ByVal major As Single = -1, Optional ByVal minor As Single = -1, _
                     Optional rot As Variant, Optional ByVal lineWidth As Single = 1) As Boolean
   '  Works only within a Lump object
   '  Entry:   rimColor       Defaults to black
   '           insideColor    Defaults to unfilled
   '           outputObj      Defaults to drawingBoard
   '           Xc, Yc         Center of oval. Default to centerX and CenterY
   '           major, minor   Axes. Default to wide and high parameters
   '           rot            Rotation in radians. Defaults to rotation parameter
   '                             Must be Variant so that a Null can be passed.
   '           lineWidth      Defaults to One
   '           All values enter unzoomed. Uses the class variable zoom to adjust
   '  Return   True if oval could be drawn within the board.
   '           Sets minX, minY, maxX, maxY
   Dim rotate As Single
   Dim X As Single, Y As Single                                                            'A point
   Dim Xr As Single, Yr As Single                                                  'A point rotated
   Dim rSin As Single, rCos As Single                             'Sine and Cosine of rotated angle
   Dim angle As Single, endAngle As Single
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Optional Parameters
   
   If rimColor = -1 Then rimColor = oColor(1)
   If rimColor = -1 Then rimColor = vbBlack
'   If insideColor = -1 Then insideColor = oColor(1)
   If outputObj Is Nothing Then Set outputObj = drawingBoard
   If Xc = -1 Then Xc = centerX
   If Yc = -1 Then Yc = centerY
   If major = -1 Then major = wide
   If minor = -1 Then minor = high
   If IsMissing(rot) Then
      rotate = rotation
   Else
      rotate = rot
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Basic Parameters
'   If objType = "Arc" And draw And Not SelectMode Then
'      outputObj.foreColor = rimColor
'   End If
   rSin = Sin(-rotate)
   rCos = Cos(-rotate)
'   If draw And objType = "Arc" And contName = "picDrafter" Then '----Make End Point Snap To Grid
'      Xr = GridCoord(wide * rCos)                              'Rotate each point from 0 Degrees
'      Yr = GridCoord(wide * rSin)                          'and set to nearest grid intersection
'      wide = (Xr ^ 2 + Yr ^ 2) ^ 0.5                                                'Reset width
'   End If
'   printWidth = wide * zoom     'Should these be Abs() ???????????????????????
'   printHeight = high * zoom
   If objType = "Arc" Then
      endAngle = PI + 0.1
   Else
      endAngle = 2 * PI + 0.1
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++ Calculate Min And Max To See If Within Board
   X = major * Cos(0)                                                         'Find the first point
   Y = minor * Sin(0)
   Xr = Xc + X * rCos + Y * rSin
   Yr = Yc - X * rSin + Y * rCos
   For angle = 0 To endAngle Step 0.1
      X = major * Cos(angle)                                                                 'Point
      Y = minor * Sin(angle)
      Xr = Xc + X * rCos + Y * rSin                                            'Rotated from center
      Yr = Yc - X * rSin + Y * rCos
      maxX = Max(maxX, Xr)
      maxY = Max(maxY, Yr)
      minX = Min(minX, Xr)
      minY = Min(minY, Yr)
   Next angle
   If OutsideBoard(Me) Then
      '  If minY or minX negative then OutsideBoard resets centerX and centerY. The calling
      '  function must goto RelocateObject
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
      
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Apply Zoom
   major = major * zoom
   minor = minor * zoom
   Xc = Xc * zoom
   Yc = Yc * zoom
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Draw Object
   If insideColor >= 0 And objType <> "Arc" Then '=============================================Fill
      Dim fillWidth As Single, fillHeight As Single                       'Moving toward center (0)
      fillWidth = major
      fillHeight = minor
      
      Do While fillWidth > 0 And fillHeight > 0  '--------------------------Move fill toward center
         X = fillWidth * Cos(0)                                               'Find the first point
         Y = fillHeight * Sin(0)
         Xr = Xc + X * rCos + Y * rSin
         Yr = Yc - X * rSin + Y * rCos
         outputObj.CurrentX = Xr
         outputObj.CurrentY = Yr
         For angle = 0 To endAngle Step 0.05                                        'Rest of points
            X = fillWidth * Cos(angle)                                                       'Point
            Y = fillHeight * Sin(angle)
            Xr = Xc + X * rCos + Y * rSin                                                  'Rotated
            Yr = Yc - X * rSin + Y * rCos
            outputObj.Line -(Xr, Yr), insideColor
         Next angle
         fillWidth = fillWidth - 5
         fillHeight = fillHeight - 5
      Loop
   End If
      
   '=======================================================================================Draw Rim
   X = major * Cos(0)                                                         'Find the first point
   Y = minor * Sin(0)
   Xr = Xc + X * rCos + Y * rSin
   Yr = Yc - X * rSin + Y * rCos
   outputObj.CurrentX = Xr
   outputObj.CurrentY = Yr
   For angle = 0 To endAngle Step 0.1                                               'Rest of points
      X = major * Cos(angle)                                                                 'Point
      Y = minor * Sin(angle)
      Xr = Xc + X * rCos + Y * rSin                                                        'Rotated
      Yr = Yc - X * rSin + Y * rCos
      outputObj.Line -(Xr, Yr), rimColor
   Next angle
   DrawEllipse = True
End Function
Public Sub DrawGeneValue(outputObj As Object, Optional draw As Boolean = True)
   Dim foreColor As Long, X As Single, Y As Single, fontSize As Single
   
   With outputObj
   fontSize = .Font.Size
   foreColor = .foreColor                                                    'Save to restore later
   If draw Then
      .foreColor = vbBlack
   Else
      .foreColor = vbWhite
   End If
   FontSizeFloor 8 * zoom, outputObj
'   X = (CenterX + wide) * zoom + (GENE_TEXT_CLEARANCE + GENE_BORDER_CLEARANCE) * zoom
'   Y = .CurrentY + .TextHeight(title) * 0.15 'Adjust for font change ??????????
   X = (centerX + wide / 2) * zoom + (GENE_TEXT_CLEARANCE + GENE_BORDER_CLEARANCE) * zoom
   Y = (centerY - high * 0.31) * zoom               'Adjust for font change ??????????
'   .CurrentX = (CenterX + wide / 2) * zoom + (GENE_TEXT_CLEARANCE + GENE_BORDER_CLEARANCE) * zoom
'   .CurrentY = (CenterY - high * 0.31) * zoom               'Adjust for font change ??????????
'   .CurrentY = (CenterY - high / 2 - 30) * zoom           'Adjust for font change ??????????
   .CurrentX = X
   .CurrentY = Y
   If draw Then
      outputObj.Line (X, Y)-Step(450 * zoom, .TextHeight(title)), _
                     vbWhite, BF
      .CurrentX = X
      .CurrentY = Y
      outputObj.Print ValueDisplay(value, 6)
      prevValue = value
   Else
      outputObj.Print ValueDisplay(prevValue, 6)
   End If
   .foreColor = foreColor                                                        'Restore foreColor
   .Font.Size = fontSize
   End With
End Sub

Public Sub Move(ByVal X As Single, ByVal Y As Single, ByVal movePoint As Integer) '*** Moves Object
   '  X and Y are nonzoomed coordinates
   '  movePoint   0: Move entire object    1: Center, move    2: Edge, size    3: Edge, rotate
   '  Negative movePoint indicates a shift move
   '     For ovals and rectangles, makes them circles and squares
   '  movePoint 0 is used for multiple selections. It is a relative move whereas others move to
   '     an absolute position. Relative moves are always in multiples of the GRID_SIZE.
   '  movePoints 1, 2, and 3 are absolute, not relative.
   '  movePoints 1 and 2 come in on grid coordinates.
   Dim XOffset As Single, YOffset As Single, Shift As Boolean
   Dim gridX As Single, gridY As Single
   
   mappWindow.dirty = True
   
   If movePoint < 0 Then                                                                'Shift move
      Shift = True
      movePoint = -movePoint
   End If
   
   Set oldObj = copy                               'Make copy of object before moving to allow undo
   
'   X = X / mappWindow.zoom
'   Y = Y / mappWindow.zoom
'
   DrawObj False                                                                  'Erase old object
   Select Case movePoint
   Case 0 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ For Multiple Selections
      centerX = centerX + X                                                'Move is always relative
      centerY = centerY + Y
   Case 1 '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Move
      centerX = X                                              'Placing the point snaps to the grid
      centerY = Y
   Case 2 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ General Size Or Width
      Select Case objType
      Case "Rectangle"
         wide = ((X - centerX) ^ 2 + (Y - centerY) ^ 2) ^ 0.5 * 2
         If Shift Then
            high = wide
         End If
      Case "Oval", "Arc"
         wide = ((X - centerX) ^ 2 + (Y - centerY) ^ 2) ^ 0.5
         If Shift Then
            high = wide
         End If
      Case "Brace"
         Select Case rotation
         Case TOP_BRACE, BOTTOM_BRACE
            wide = Abs(X - centerX) * 2
         Case Else
            wide = Abs(Y - centerY) * 2
         End Select
      Case "Gene", "Poly"
         wide = (X - centerX) * 2
      Case "Vesicle"
         wide = X - centerX
      Case "ProteinA"
         wide = (X - centerX) / 1.7 / 1.1                                'See EditPoint explanation
      Case "ProteinB"
         wide = (X - centerX) / 1.7                                                 'Circle overlap
      End Select
   Case 3 '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Height
      Select Case objType
      Case "Rectangle"
         high = ((X - centerX) ^ 2 + (Y - centerY) ^ 2) ^ 0.5 * 2
         If Shift Then
            wide = high
         End If
      Case "Oval", "Arc"
         high = ((X - centerX) ^ 2 + (Y - centerY) ^ 2) ^ 0.5
         If Shift Then
            wide = high
         End If
      End Select
   Case 4 '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Rotate
      XOffset = centerX - X
      YOffset = centerY - Y
      Select Case objType
      Case "Rectangle"
         Dim pointAngle As Single, cornerAngle As Single
         cornerAngle = Atn(high / wide)
         If XOffset = 0 Then
            If YOffset >= 0 Then
               pointAngle = 0.5 * PI                                                    '90 degrees
            Else
               pointAngle = 1.5 * PI                                                   '270 degrees
            End If
         Else
            pointAngle = Atn(YOffset / XOffset)
            If XOffset <= 0 Then                                             '2nd and 3rd quadrants
               pointAngle = PI + pointAngle                                                'Add 180
            End If
         End If
         rotation = pointAngle - cornerAngle
         If Shift Then                                                        '45 degree increments
            rotation = rotation / PI * 4                        'Should be a number between 0 and 8
            rotation = Round(rotation)
            rotation = rotation * PI / 4
         End If
      Case Else
         If XOffset = 0 Then
            If YOffset >= 0 Then
               rotation = 0.5 * PI                                                      '90 degrees
            Else
               rotation = 1.5 * PI                                                     '270 degrees
            End If
         Else
            rotation = Atn(YOffset / XOffset)
            
            If XOffset <= 0 Then                                             '2nd and 3rd quadrants
               rotation = PI + rotation                                                    'Add 180
            End If
         End If
      End Select
      If Shift Then                                                           '45 degree increments
         rotation = rotation / PI * 4                           'Should be a number between 0 and 8
         rotation = Round(rotation)
         rotation = rotation * PI / 4
      End If
      rotation = InCircle(rotation)
      Select Case objType
      Case "Brace"                                                    'Only in 90 degree increments
         rotation = rotation / PI * 2                           'Should be a number between 0 and 3
         rotation = Round(rotation) Mod 4
'         rotation = rotation * PI / 2
      End Select
   End Select
   DrawObj
   If movePoint <> 0 Then '++++++++++++++++++++++++++++ Redraw Edit Points If Not a Multiple Select
      Select Case objType
      Case "Rectangle", "Oval", "Arc", "Brace", "Gene", "Vesicle", "ProteinA", "ProteinB", "Poly"
               'Always redraw second edit point since it does not necessarily have to stay at edge.
               'It may be above or below the edge, or inside if size is too small.
               'DrawObj before this because gene may change wide if too small
         EditPoint Me, 2
      End Select
      Select Case objType
      Case "Rectangle", "Oval", "Arc"
         EditPoint Me, 3
      End Select
      Select Case objType
      Case "Rectangle", "Oval", "Arc", "Brace", "CellB", "Poly"  'Always redraw rotation edit point
                                               'since it does not necessarily have to stay at edge.
         EditPoint Me, 4
      End Select
   End If
End Sub
Public Sub SetEdit(Optional edit As Boolean = True) '***************************** Put In Edit Mode
   '  activeObject variable always set somewhere else (SetActiveObject)
   If edit Then
      EditPoint Me                                                                          'Moving
      editMode = True
      Select Case objType
      Case "Rectangle", "Oval", "Arc", "Brace", "Gene", "Vesicle", "ProteinA", "ProteinB", "Poly"
         EditPoint Me, 2                                             'Sizing, width or general size
      End Select
      Select Case objType
      Case "Rectangle", "Oval", "Arc"
         EditPoint Me, 3                                                            'Sizing, height
      End Select
      Select Case objType
      Case "Rectangle", "Oval", "Arc", "Brace", "CellB", "Poly"
         EditPoint Me, 4                                                                  'Rotating
      End Select
   Else
      EditPoint , , False
      editMode = False
'      frmDrafter.sbrBar.Panels(1).Text = ""
      Select Case objType
      Case "Rectangle", "Oval", "Arc", "Brace", "Gene", "Vesicle", "ProteinA", "ProteinB", "Poly"
         EditPoint , 2, False
      End Select
      Select Case objType
      Case "Rectangle", "Oval", "Arc"
         EditPoint , 3, False
      End Select
      Select Case objType
      Case "Rectangle", "Oval", "Arc", "Brace", "CellB", "Poly"
         EditPoint , 4, False
      End Select
   End If
End Sub
Public Function CheckClick(ByVal X As Single, ByVal Y As Single) As Boolean '******* See If Clicked
   '  X and Y passed in nonzoomed coordinates
   '  Lump usually can be clicked anywhere within, not just around border
   Const PLAY = 75                                           'Amount on either side of line for hit
   Dim XOffset As Single, YOffset As Single
'   Dim PLAY As Single
   Dim centerSize As Single                     'Size of objects that display smaller on frmObjects
   Dim centerHeight As Single
   Dim zoom As Single
   Dim rSin As Single, rCos As Single                            'Sine and cosine of rotation angle
   
   If canvas.name = "picDrafter" Then
      zoom = canvas.container.zoom
   Else
      zoom = 1
   End If
   
'   X = X / zoom
'   Y = Y / zoom
   
   Select Case objType
   Case "Poly" '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Polygon
      '  Everything is converted to relate to the center of the polygon
      '  Each poly line is checked against the point to see if the point's distance is within
      '  PLAY of the poly lines intersection with a line drawn from the center to the point.
      '  As soon as there is a hit on one poly line, the process stops. If not hits, no click.
      Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single, xL As Single, yL As Single
      Dim pointDistance As Single, lineDistance As Single
      
      '====================================================Distance From Point To Center Of Polygon
'      X = X / zoom - centerX
'      Y = Y / zoom - centerY
      X = X - centerX
      Y = Y - centerY
      pointDistance = (X ^ 2 + Y ^ 2) ^ 0.5
      
      If pointDistance > wide / 2 + PLAY Then               'Totally outside regualr polygon radius
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
      
      x1 = (wide / 2) * Cos(2 * PI * (sides - 1) / sides + rotation)
      y1 = (wide / 2) * Sin(2 * PI * (sides - 1) / sides + rotation)
'canvas.DrawWidth = 5
'canvas.PSet (x1 + centerX, y1 + centerY), vbGreen
'canvas.DrawWidth = 1
      For i = 0 To sides - 1 '======================================================Check Each Line
         x2 = (wide / 2) * Cos(2 * PI * i / (sides) + rotation)
         y2 = (wide / 2) * Sin(2 * PI * i / (sides) + rotation)
'canvas.DrawWidth = 5
'canvas.PSet (x2 + centerX, y2 + centerY), vbRed
'canvas.DrawWidth = 1
         If LineIntersection(X, Y, 0, 0, x1, y1, x2, y2, xL, yL) Then                 'Not Parallel
'canvas.DrawWidth = 3
'canvas.PSet (xL + centerX, yL + centerY), vbBlue
'canvas.DrawWidth = 1
            lineDistance = (xL ^ 2 + yL ^ 2) ^ 0.5
            If pointDistance <= lineDistance + PLAY Then
               If pointDistance >= lineDistance - PLAY Then
                  CheckClick = True
                  Exit Function                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
               End If
            End If
         End If
         x1 = x2
         y1 = y2
'canvas.DrawWidth = 5
'canvas.PSet (x1 + centerX, y1 + centerY), vbGreen
'canvas.DrawWidth = 1
      Next i
   Case "Rectangle" '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Rectangle
      Dim XRotate As Single, YRotate As Single                         'After rotation back to zero
                  '  Click will be checked at zero rotation from center of object, so set click
                  '  X and Y in relation to object center and rotate click point back to zero
   
      rSin = Sin(-rotation): rCos = Cos(-rotation)             'Rotate back to object zero rotation
      XRotate = ((X - centerX) * rCos - (Y - centerY) * rSin)     'Rotate the points and add center
      YRotate = ((X - centerX) * rSin + (Y - centerY) * rCos)
      
      '-------------------------------------------------------------------If Within Outer Rectangle
      '  Both point and object based on object's center
      If XRotate >= -wide / 2 - PLAY And XRotate <= wide / 2 + PLAY Then
         If YRotate >= -high / 2 - PLAY And YRotate <= high / 2 + PLAY Then
            CheckClick = True
            '------------------------------------------------------------Not Within Inner Rectangle
            If XRotate >= -wide / 2 + PLAY And XRotate <= wide / 2 - PLAY Then
               If YRotate >= -high / 2 + PLAY And YRotate <= high / 2 - PLAY Then
                  CheckClick = False
               End If
            End If
         End If
      End If
   Case "Oval", "Arc" '+++++++++++++++++++++++++++++++++++++++++++++++++++++ Adjustable Oval Or Arc
      '  Clickable within 50 twips of oval
      'Dim rSin As Single, rCos As Single                        'Sine and cosine of rotation angle
      'Dim XRotate As Single, YRotate As Single                        'After rotation back to zero
                  '  Click will be checked at zero rotation from center of object, so set click
                  '  X and Y in relation to object center and rotate click point back to zero
   
'      PLAY = 50
      '---------------------------------------------Rotate Click Point Back To Object Zero Rotation
      rSin = Sin(-rotation): rCos = Cos(-rotation)
      XRotate = ((X - centerX) * rCos - (Y - centerY) * rSin)      'Rotate the points based
      YRotate = ((X - centerX) * rSin + (Y - centerY) * rCos)      'on center
      
      '------------------------------------------------------------------------If Within Outer Oval
      '  high / wide is ratio of major and minor radii
      If YRotate < 0 And objType = "Arc" Then                                   'Upper part of oval
         CheckClick = False                                                           'No arc there
      ElseIf Abs(XRotate) <= wide + PLAY Then                                           'In X range
         If Abs(YRotate) <= ((wide + PLAY) ^ 2 - XRotate ^ 2) ^ 0.5 * high / wide Then
            CheckClick = True
            '-----------------------------------------------------------------Not Within Inner Oval
            '  If oval is filled (oColor(1) >= 0) it can be clicked anywhere
            If oColor(1) < 0 And Abs(XRotate) <= wide - PLAY Then                       'In X range
               If Abs(YRotate) <= ((wide - PLAY) ^ 2 - XRotate ^ 2) ^ 0.5 * high / wide Then
                  CheckClick = False
               End If
            End If
         End If
      End If
   Case "Brace"
      Dim thickness As Single, curvature As Single
      
      thickness = BRACE_THICK * zoom
      curvature = BRACE_RADIUS * zoom
      Select Case rotation
      Case TOP_BRACE
         If X >= centerX - wide / 2 And X <= centerX + wide / 2 Then
            If Y >= centerY And Y <= centerY + curvature * 2 + thickness Then
               CheckClick = True
            End If
         End If
      Case BOTTOM_BRACE
         If X >= centerX - wide / 2 And X <= centerX + wide / 2 Then
            If Y <= centerY And Y >= centerY - (curvature * 2 + thickness) Then
               CheckClick = True
            End If
         End If
      Case LEFT_BRACE
         If Y >= centerY - wide / 2 And Y <= centerY + wide / 2 Then
            If X >= centerX And X <= centerX + curvature * 2 + thickness Then
               CheckClick = True
            End If
         End If
      Case RIGHT_BRACE
         If Y >= centerY - wide / 2 And Y <= centerY + wide / 2 Then
            If X <= centerX And X >= centerX - (curvature * 2 + thickness) Then
               CheckClick = True
            End If
         End If
      End Select
   Case "Label"
      If X >= centerX - wide / 2 And X <= centerX + wide / 2 Then
         If Y >= centerY - high / 2 And Y <= centerY + high / 2 Then
               CheckClick = True
         End If
      End If
   Case "MAPPTitle" '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Not Clickable
   Case "Gene"
      Dim rightSide As Single
      rightSide = centerX + wide / 2
      If X >= centerX - wide / 2 And X <= rightSide Then
         If Y >= centerY - high / 2 And Y <= centerY + high / 2 Then
            CheckClick = True
         End If
      End If
   Case "Vesicle" '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Open Circle
   '  Clickable within 50 twips of outer circle
'      PLAY = 50
      XOffset = wide + PLAY
      YOffset = Abs(XOffset ^ 2 - (X - centerX) ^ 2) ^ 0.5
      '----------------------------------------------------------------------If Within Outer Circle
      XOffset = wide + PLAY
      YOffset = Abs(XOffset ^ 2 - (X - centerX) ^ 2) ^ 0.5
      If X >= centerX - XOffset And X <= centerX + XOffset Then                     'Within X range
         If Y >= centerY - YOffset And Y <= centerY + YOffset Then                   'Within Circle
            CheckClick = True
            '----------------------------------------------------------------If Within Inner Circle
            XOffset = wide - PLAY
            YOffset = Abs(XOffset ^ 2 - (X - centerX) ^ 2) ^ 0.5
            If X >= centerX - XOffset And X <= centerX + XOffset Then               'Within X range
               If Y >= centerY - YOffset And Y <= centerY + YOffset Then             'Within Circle
                  CheckClick = False
               End If
            End If
         End If
      End If
   Case "ProteinA"
      If X >= centerX - wide * 1.8 And X <= centerX + wide * 1.8 Then
         If Y >= centerY - wide * 1.6 And Y <= centerY + wide * 1.6 Then
               CheckClick = True
         End If
      End If
   Case "ProteinB"
      If X >= centerX - wide * 1.8 And X <= centerX + wide * 1.8 Then
         If Y >= centerY - wide * 1.5 And Y <= centerY + wide * 1.5 Then
               CheckClick = True
         End If
      End If
   Case "Ribosome"
      If X >= centerX - RIBOSOME_CENTER And X <= centerX + RIBOSOME_CENTER Then
         If Y >= centerY - RIBOSOME_CENTER * RIBOSOME_ASPECT * 0.65 And Y <= centerY + RIBOSOME_CENTER * RIBOSOME_ASPECT * 0.65 Then
               CheckClick = True
         End If
      End If
   Case "OrganA"
      Select Case canvas.name                                     'Only a third the size in toolbox
      Case "frmObjects"
         centerSize = ORGANA_CENTER / 3
      Case "picDrafter"
         centerSize = ORGANA_CENTER
      End Select
      If X >= centerX - centerSize And X <= centerX + centerSize Then
         If Y >= centerY - Abs((centerSize * ORGANA_ASPECT) ^ 2 - (X - centerX) ^ 2) ^ 0.5 _
            And Y <= centerY + Abs((centerSize * ORGANA_ASPECT) ^ 2 - (X - centerX) ^ 2) ^ 0.5 Then
               CheckClick = True
         End If
      End If
   Case "OrganB"
      Select Case canvas.name                                     'Only a third the size in toolbox
      Case "frmObjects"
         centerSize = ORGANB_CENTER / 3
         centerHeight = ORGANB_HEIGHT / 3
      Case "picDrafter"
         centerSize = ORGANB_CENTER
         centerHeight = ORGANB_HEIGHT
      End Select
      If X >= centerX - centerSize And X <= centerX + centerSize Then
         If Y >= centerY - centerHeight And Y <= centerY + centerHeight Then
               CheckClick = True
         End If
      End If
   Case "OrganC"
      If X >= centerX - ORGANC_CENTER And X <= centerX + ORGANC_CENTER Then
         If Y >= centerY - Abs(ORGANC_CENTER ^ 2 - (X - centerX) ^ 2) ^ 0.5 _
            And Y <= centerY + Abs(ORGANC_CENTER ^ 2 - (X - centerX) ^ 2) ^ 0.5 Then
               CheckClick = True
         End If
      End If
   Case "CellA"
      '  Must click near center -- within minor-axis radius
      Dim clickRadius As Single
      
      clickRadius = CELLA_RADIUS / CELLA_ASPECT
      If canvas.name = "frmObjects" Then clickRadius = clickRadius / 3
      If X >= centerX - clickRadius And X <= centerX + clickRadius Then
         If Y >= centerY - Abs(clickRadius ^ 2 - (X - centerX) ^ 2) ^ 0.5 _
            And Y <= centerY + Abs(clickRadius ^ 2 - (X - centerX) ^ 2) ^ 0.5 Then
               CheckClick = True
         End If
      End If
   Case "CellB"
      '  Must click near center -- within minor-axis radius
      clickRadius = CELLB_RADIUS / CELLB_ASPECT
      If X >= centerX - clickRadius And X <= centerX + clickRadius Then
         If Y >= centerY - Abs(clickRadius ^ 2 - (X - centerX) ^ 2) ^ 0.5 _
            And Y <= centerY + Abs(clickRadius ^ 2 - (X - centerX) ^ 2) ^ 0.5 Then
               CheckClick = True
         End If
      End If
   End Select
End Function
Sub MaxOnBoard() '*********************************** Makes Object Max Height Or Width To Fit Board
   If maxX - minX > MAX_BOARD_WIDTH Then '+++++++++++++++++++++++++++++++++++++++++ Object Too Wide
'      MsgBox "Object too wide to fit on drafting board. It is adjusted to the maximum width of " _
'             & "the board and recentered.", vbExclamation + vbOKOnly, "Object Size"
      DrawObj False                                                                          'Erase
      wide = GridMax(MAX_BOARD_WIDTH)
      Select Case objType
      Case "Gene"
         wide = wide - GENE_VALUE_WIDTH - GRID_SIZE                     'Extra grid square as fudge
         centerX = GridCoord(wide / 2)
      Case "Oval"
         wide = wide / 2 - GRID_SIZE                                          'wide is major radius
         '  This does not completely solve the problem because if the oval is rotated high and
         '  wide will not be on these axes
         centerX = GridMin(wide)
      Case Else
         centerX = GridCoord(wide / 2)
      End Select
   End If
   If maxY - minY > MAX_BOARD_HEIGHT Then '++++++++++++++++++++++++++++++++++++++++ Object Too High
'      MsgBox "Object too high to fit on drafting board. It is adjusted to the maximum height of "
'             & "the board and recentered.", vbExclamation + vbOKOnly, "Object Size"
      DrawObj False                                                                          'Erase
      high = GridMax(MAX_BOARD_HEIGHT)
      Select Case objType
      Case "Oval"
         high = (high - GRID_SIZE) / 2                                        'high is minor radius
         centerY = GridMin(high)
      Case Else
         centerY = GridCoord(high / 2)
      End Select
   End If
   DrawObj                                                                                  'Redraw
End Sub

Rem/////////////////////////////////////////////////////////////////////////////// Property Methods
Public Property Let SelectMode(ByVal vData As Boolean)
   If mvarSelectMode <> vData Then
      mvarSelectMode = vData
      If Screen.ActiveForm Is frmObjects Then
         DrawObj True, frmObjects
      Else
         DrawObj True
      End If
   End If
End Property
Public Property Get SelectMode() As Boolean
    SelectMode = mvarSelectMode
End Property

Public Property Let color(index As Integer, value As Long)
   oColor(index) = value
End Property
Public Property Get color(index As Integer) As Long
   color = oColor(index)
End Property

Public Property Let rim(index As Integer, value As Long)
   oRim(index) = value
End Property
Public Property Get rim(index As Integer) As Long
   rim = oRim(index)
End Property

