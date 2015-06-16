VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'  See Objects.bas for line types, constants, and specifics

Private mvarSelectMode As Boolean

Public objKey As String                                            'Unique identifier for an object
   '  String because it is also the identifier in the objLumps collection. Goes in the MAPP
   '  database as long.
Public canvas As Object                                'Drawing surface on which originally created

Public style As String                                                         'Solid, Broken, etc.
Public objType As String
Public color As Long                                                                 'Color of line
Public wide As Single                                      'Width of line (Width column in Objects)
Public remarks As String                                                   'Straight text for lines
Public links As String                                           'Links for lines (not implemented)
Public centerX As Single, centerY As Single
   '  Makes this compatible with other objects. Should probably use this instead
   '  of XStart and YStart
Public XStart As Single, xEnd As Single
Public YStart As Single, YEnd As Single
Public minY As Single, minX As Single
Public maxX As Single, maxY As Single                     'Farthest X and Y Board points for object
   '  These are always in nonzoomed coordinates
Public editMode As Boolean
                           '  Used for figuring minimum board size (autosize)

'************************************************************** Creates and Draws a New Line Object
Public Sub Create(XSt As Single, YSt As Single, XEn As Single, YEn As Single, _
          sty As String, Optional objectKey As Variant, Optional outputObj As Object = Nothing)
   'Entry:  Xs, Ys   Unzoomed
   
   If VarType(objectKey) = vbNull Then objectKey = -1
   If IsMissing(objectKey) Then objectKey = -1                  'Previous MAPPs have no object keys
   If outputObj Is Nothing Then
      Set outputObj = drawingBoard
   End If
   Set canvas = outputObj                 'Objects are never created for the printer, only drawn
   If color < 0 Then color = 0                        'Colors default to -1 (transparent) in Access
   objType = "objLine"
   If objectKey = -1 Then                                                               'New object
      outputObj.container.objKey = outputObj.container.objKey + 1    'Increment picDrafter's objKey
      objKey = outputObj.container.objKey              'Numeric to string conversion automatic here
   Else                                                                  'Object from MAPP database
      objKey = objectKey
   End If
   XStart = GridCoord(XSt)                                          'Set points to grid coordinates
   YStart = GridCoord(YSt)
   xEnd = GridCoord(XEn)
   YEnd = GridCoord(YEn)
   style = sty
   DrawObj True
   outputObj.container.dirty = True
'   frmDrafter.sbrBar.Panels(1).Text = ""
End Sub
Public Sub Duplicate(obj As Object)
   color = obj.color
   Create obj.XStart + 100, obj.YStart + 100, obj.xEnd + 100, obj.YEnd + 100, obj.style
End Sub
'****************************************************************** Draws Or Erases a Line Or Arrow
Sub DrawObj(Optional draw As Boolean = True, Optional outputObj As Object = Nothing)
   '  draw       True - new line. False - erase line
   
   Dim XOffset As Single, YOffset As Single
   Dim index As Single, XInc As Single, YInc As Single
   Dim XOrig As Single, YOrig As Single, XAdj As Single, YAdj As Single
   Dim XStartZoom As Single, XEndZoom As Single                             'End coordinates zoomed
   Dim YStartZoom As Single, YEndZoom As Single
   Dim angle As Double
   Dim leftX As Single, leftY As Single, rightX As Single, rightY As Single    'Receptor end points
                                                                         '  In relation to line end
   Dim arrowWidth As Single, receptorExtension As Single
   Dim outName As String                       'Name of canvas. Print doesn't support Name property
   
   Dim drawWidthAdjust As Single                              'Adjustment for DrawWidth for Printer
   Dim printWidth As Single, printHeight As Single
   Dim zoom As Single
   
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Output Object
   If outputObj Is Nothing Then Set outputObj = canvas             'Where it was originally created
   zoom = mappWindow.zoom
   If outputObj Is Printer Then '-----------------------------------------------------------Printer
      outName = "Printer"
      drawWidthAdjust = 2.5                                     'Must adjust for printer resolution
   Else '------------------------------------------------------------------------------------Screen
      outName = outputObj.name
      drawWidthAdjust = 1             'Standard DrawWidth for screen, won't go smaller with preview
   End If
   
RelocateObject: '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Figure Object Dimensions
   If outName = "picDrafter" And draw Then                 'Default autosize values to ends of line
      minX = Min(XStart, xEnd)
      minY = Min(YStart, YEnd)
      maxX = Max(XStart, xEnd)
      maxY = Max(YStart, YEnd)
   End If
   XOffset = xEnd - XStart
   YOffset = YEnd - YStart
   If YOffset = 0 Then
      If XOffset >= 0 Then
         angle = 0.5 * PI                                                               '90 degrees
      Else
         angle = 1.5 * PI                                                              '270 degrees
      End If
   Else
      angle = Atn(XOffset / -YOffset)
      If YOffset > 0 Then                                                          'Lower quadrants
         angle = PI + angle                                                                'Add 180
      ElseIf XOffset < 0 Then                                                      'Fourth quadrant
         angle = 2 * PI + angle                                                            'Add 360
      End If
   End If
   XInc = Abs(Sin(angle)) * Sgn(XOffset)                      'Distance on the axis to create point
   YInc = Abs(Cos(angle)) * Sgn(YOffset)                      'for a line of length 1
   If draw Then '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine Colors
      Select Case outName
      Case "Printer"                                     'These should never show select mode color
         outputObj.foreColor = color
      Case Else
         If mvarSelectMode Then
            outputObj.foreColor = vbRed
         Else
            outputObj.foreColor = color
         End If
      End Select
   Else '--------------------------------------------------------------------------------Erase Line
      outputObj.foreColor = vbWhite
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Apply Zoom
   XStartZoom = XStart * zoom
   XEndZoom = xEnd * zoom
   YStartZoom = YStart * zoom
   YEndZoom = YEnd * zoom
   receptorExtension = RECEPTOR_EXTENSION * zoom
   outputObj.DrawWidth = Max(1 * drawWidthAdjust * zoom, 1)
   
   Select Case style '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Draw Object Ends
      '  The center line is drawn later
   Case "Arrow", "BrokenArrow" '---------------------------------------------------------Draw Arrow
      '  minX, minY, maxX, and maxY don't take into account the width of the arrow
      arrowWidth = ARROW_WIDTH * zoom
      If OutsideBoard(Me) Then GoTo RelocateObject                   'Only if minY or minX negative
      For index = 0 To arrowWidth 'Step 5
         XOrig = XEndZoom + YInc * index                    '90 degrees to the side of arrow center
         YOrig = YEndZoom - XInc * index
         XAdj = -XInc * index * ARROW_ASPECT                    'How far back of origin line starts
         YAdj = -YInc * index * ARROW_ASPECT
         outputObj.Line (XOrig + XAdj, YOrig + YAdj) _
                     -Step(-XInc * arrowWidth * ARROW_ASPECT - XAdj, _
                           -YInc * arrowWidth * ARROW_ASPECT - YAdj), outputObj.foreColor
         XOrig = XEndZoom + YInc * -index                                'Other side of center line
         YOrig = YEndZoom - XInc * -index
         outputObj.Line (XOrig + XAdj, YOrig + YAdj) _
                     -Step(-XInc * arrowWidth * ARROW_ASPECT - XAdj, _
                           -YInc * arrowWidth * ARROW_ASPECT - YAdj), outputObj.foreColor
      Next index
   Case "Receptor" '----------------------------------------------------------------Draw V Receptor
      angle = angle - 0.25 * PI '..................................................Minus 45 Degrees
      If angle < 0 Then angle = angle + 2 * PI
      leftX = Abs(Sin(angle) * receptorExtension)
      leftY = Abs(Cos(angle) * receptorExtension)
      If outName = "picDrafter" And draw Then        'Figure limits with Abs of receptor extensions
         minX = Min(minX, (XEndZoom - leftX) / zoom)
         minY = Min(minY, (YEndZoom - leftY) / zoom)
         maxX = Max(maxX, (XEndZoom + leftX) / zoom)
         maxY = Max(maxY, (YEndZoom + leftY) / zoom)
      End If
      Select Case angle
      Case Is <= 0.5 * PI                                                           'First quadrant
         leftY = -leftY
      Case Is <= PI                                                                'Second quadrant
      Case Is <= 1.5 * PI                                                           'Third quadrant
         leftX = -leftX
      Case Else                                                                    'Fourth quadrant
         leftX = -leftX
         leftY = -leftY
      End Select
      
      angle = angle + 0.5 * PI '....................................................Plus 90 Degrees
      If angle >= 2 * PI Then angle = angle - 2 * PI
      rightX = Abs(Sin(angle) * receptorExtension)
      rightY = Abs(Cos(angle) * receptorExtension)
      If outName = "picDrafter" And draw Then        'Figure limits with Abs of receptor extensions
         minX = Min(minX, (XEndZoom - rightX) / zoom)
         minY = Min(minY, (YEndZoom - rightY) / zoom)
         maxX = Max(maxX, (XEndZoom + rightX) / zoom)
         maxY = Max(maxY, (YEndZoom + rightY) / zoom)
      End If
      Select Case angle
      Case Is <= 0.5 * PI                                                           'First quadrant
         rightY = -rightY
      Case Is <= PI                                                                'Second quadrant
      Case Is <= 1.5 * PI                                                           'Third quadrant
         rightX = -rightX
      Case Else                                                                    'Fourth quadrant
         rightX = -rightX
         rightY = -rightY
      End Select
      
      If outName = "picDrafter" And draw Then
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
      outputObj.Line (XEndZoom, YEndZoom)-Step(leftX, leftY), outputObj.foreColor
      outputObj.Line (XEndZoom, YEndZoom)-Step(rightX, rightY), outputObj.foreColor
   Case "ReceptorSq" '---------------------------------------------------------Draw Square Receptor
      '  receptorExtension is from center to edge of receptor, also length of receptor
      
      leftX = XEndZoom + YInc * receptorExtension                       '90 degrees to left of line
      leftY = YEndZoom - XInc * receptorExtension
      rightX = XEndZoom - YInc * receptorExtension                     '90 degrees to right of line
      rightY = YEndZoom + XInc * receptorExtension
      If outName = "picDrafter" And draw Then
         maxX = Max(maxX, leftX / zoom)                                                      'T-Bar
         maxX = Max(maxX, rightX / zoom)
         maxY = Max(maxY, leftY / zoom)
         maxY = Max(maxY, rightY / zoom)
         minX = Min(minX, leftX / zoom)
         minX = Min(minX, rightX / zoom)
         minY = Min(minY, leftY / zoom)
         minY = Min(minY, rightY / zoom)
         maxX = Max(maxX, maxX + XInc * receptorExtension / zoom)                            'Prong
            '  maxX is the farthest of leftX or rightX, so add the receptor extension to maxX
            '  and so forth for the rest of the stuff
         maxY = Max(maxY, maxY + YInc * receptorExtension / zoom)
         minX = Min(minX, minX + XInc * receptorExtension / zoom)
         minY = Min(minY, minY + YInc * receptorExtension / zoom)
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
      outputObj.Line (leftX, leftY)-(rightX, rightY), outputObj.foreColor                    'T-bar
                                                                                            'Prongs
      outputObj.Line (leftX, leftY)-Step(XInc * receptorExtension, YInc * receptorExtension), _
                     outputObj.foreColor
      outputObj.Line (rightX, rightY)-Step(XInc * receptorExtension, YInc * receptorExtension), _
                     outputObj.foreColor
   Case "Inhibitor" '--------------------------------------------------------Draw T-Bar (Inhibitor)
      '  receptorExtension is length of T in T-bar
      
      leftX = XEndZoom + YInc * receptorExtension / 2                   '90 degrees to left of line
      leftY = YEndZoom - XInc * receptorExtension / 2
      rightX = XEndZoom - YInc * receptorExtension / 2                 '90 degrees to right of line
      rightY = YEndZoom + XInc * receptorExtension / 2
      If outName = "picDrafter" And draw Then
         maxX = Max(maxX, leftX / zoom)
         maxX = Max(maxX, rightX / zoom)
         maxY = Max(maxY, leftY / zoom)
         maxY = Max(maxY, rightY / zoom)
         minX = Min(minX, leftX / zoom)
         minX = Min(minX, rightX / zoom)
         minY = Min(minY, leftY / zoom)
         minY = Min(minY, rightY / zoom)
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
      outputObj.Line (leftX, leftY)-(rightX, rightY), outputObj.foreColor
   Case "LigandSq" '-------------------------------------------------------------Draw Square Ligand
      Dim boxWidth As Single                                            'Width of filled square box
      Dim receptorClearance As Single                         'Distance between ligand and receptor
      
      receptorClearance = GRID_SIZE * 2 * zoom
      boxWidth = (receptorExtension - receptorClearance) * 2
      XEndZoom = XEndZoom - XInc * receptorClearance          'End point lines up with end point on
      YEndZoom = YEndZoom - YInc * receptorClearance          'receptor so that the two align
      If outName = "picDrafter" And draw Then
         '  Both sides of box -- index = 0 and index = boxWidth * 0.75
         maxX = Max(maxX, (XEndZoom + Abs(YInc) * boxWidth / 2 / zoom))           'index = 0
         maxX = Max(maxX, _
                    (XEndZoom - Abs(XInc) * boxWidth * 0.75 + Abs(YInc) * boxWidth / 2 / zoom))
            'Don't understand the minus in - Abs(XInc)    ????????????????
         maxY = Max(maxY, (YEndZoom + Abs(XInc) * boxWidth / 2 / zoom))           'index = 0
         maxY = Max(maxY, _
                    (YEndZoom - Abs(YInc) * boxWidth * 0.75 + Abs(XInc) * boxWidth / 2 / zoom))
         minX = Min(minX, (XEndZoom - Abs(YInc) * boxWidth / 2) / zoom)
         minX = Min(minX, _
                    (XEndZoom - Abs(XInc) * boxWidth * 0.75 + Abs(YInc) * boxWidth / 2) / zoom)
         minY = Min(minY, (YEndZoom - Abs(XInc) * boxWidth / 2) / zoom)
         minY = Min(minY, _
                    (YEndZoom - Abs(YInc) * boxWidth * 0.75 + Abs(XInc) * boxWidth / 2) / zoom)
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
      For index = 0 To boxWidth * 0.75 Step 1
         XOrig = XEndZoom - XInc * index
         YOrig = YEndZoom - YInc * index
         outputObj.Line (XOrig + YInc * boxWidth / 2, YOrig - XInc * boxWidth / 2) _
                       -(XOrig - YInc * boxWidth / 2, YOrig + XInc * boxWidth / 2), _
                       outputObj.foreColor
      Next index
   Case "ReceptorRd" '----------------------------------------------------------Draw Round Receptor
      angle = InCircle(2 * PI - (angle - 0.5 * PI))           'Set angle along +X axis from -Y axis
      If outName = "picDrafter" And draw Then
         maxX = Max(maxX, (XEndZoom + XInc * receptorExtension + YInc * receptorExtension) / zoom)
         maxX = Max(maxX, (XEndZoom + XInc * receptorExtension - YInc * receptorExtension) / zoom)
         maxY = Max(maxY, (YEndZoom + YInc * receptorExtension + XInc * receptorExtension) / zoom)
         maxY = Max(maxY, (YEndZoom + YInc * receptorExtension - XInc * receptorExtension) / zoom)
         minX = Min(minX, (XEndZoom + XInc * receptorExtension + YInc * receptorExtension) / zoom)
         minX = Min(minX, (XEndZoom + XInc * receptorExtension - YInc * receptorExtension) / zoom)
         minY = Min(minY, (YEndZoom + YInc * receptorExtension + XInc * receptorExtension) / zoom)
         minY = Min(minY, (YEndZoom + YInc * receptorExtension - XInc * receptorExtension) / zoom)
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
      outputObj.Circle (XEndZoom + XInc * receptorExtension, _
                        YEndZoom + YInc * receptorExtension), _
                        receptorExtension, outputObj.foreColor, _
                        InCircle(angle + 0.5 * PI), InCircle(angle - 0.5 * PI)
   Case "LigandRd" '--------------------------------------------------------------Draw Round Ligand
      Dim receptorRadius As Single                                       'Zoomed radius of receptor
      receptorRadius = receptorExtension - GRID_SIZE * 2 * zoom
      If outName = "picDrafter" And draw Then
         '  Full circle including space for receptor
         maxX = Max(maxX, (XEndZoom - XInc * receptorRadius + receptorRadius) / zoom)
         maxX = Max(maxX, (XEndZoom - XInc * receptorRadius - receptorRadius) / zoom)
         maxY = Max(maxY, (YEndZoom - YInc * receptorRadius + receptorRadius) / zoom)
         maxY = Max(maxY, (YEndZoom - YInc * receptorRadius - receptorRadius) / zoom)
         minX = Min(minX, (XEndZoom - XInc * receptorRadius + receptorRadius) / zoom)
         minX = Min(minX, (XEndZoom - XInc * receptorRadius - receptorRadius) / zoom)
         minY = Min(minY, (YEndZoom - YInc * receptorRadius + receptorRadius) / zoom)
         minY = Min(minY, (YEndZoom - YInc * receptorRadius - receptorRadius) / zoom)
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
      outputObj.FillStyle = vbFSSolid
      If draw Then
         outputObj.FillColor = color
      Else
         outputObj.FillColor = vbWhite
      End If
      outputObj.Circle (XEndZoom - XInc * receptorExtension, _
                   YEndZoom - YInc * receptorExtension), _
                   receptorExtension - GRID_SIZE * 2 * zoom, outputObj.foreColor
      outputObj.FillColor = vbWhite
      outputObj.FillStyle = vbFSTransparent
               'Why reassign these below???????????????????????
      XEndZoom = XEndZoom - XInc * GRID_SIZE * 2 * zoom   'End point lines up with end point on
      YEndZoom = YEndZoom - YInc * GRID_SIZE * 2 * zoom   'receptor so that the two align
   Case Else '--------------------------------------------------------------------------Plain Lines
      '  All other line objects check OutsideBoard for their contained lines and various
      '  extensions above
      If draw And outName = "picDrafter" Then
         If OutsideBoard(Me) Then GoTo RelocateObject                'Only if minY or minX negative
      End If
   End Select
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Draw Line
   If style = "Broken" Or style = "BrokenArrow" Then
      BrokenLine outputObj, XStartZoom, YStartZoom, XEndZoom, YEndZoom
   Else
      outputObj.Line (XStartZoom, YStartZoom)-(XEndZoom, YEndZoom), outputObj.foreColor
   End If
End Sub
'******************************************************************* Moves a Point On a Line Object
Public Sub Move(ByVal X As Single, ByVal Y As Single, movePoint As Integer)
   '  X and Y are nonzoomed coordinates
   '  movePoint   0: Move entire line    1: Beginning of line    2: End of line
   '  movePoint 0 is used for multiple selections. It is a relative move whereas others move to
   '     an absolute position. Relative moves are always in multiples of the GRID_SIZE.
   '  movePoints 1 and 2 are absolute, not relative.
   '  movePoints 1 and 2 come in on grid coordinates.
   
'   X = X / mappWindow.zoom                                                'Unzoom these coordinates
'   Y = Y / mappWindow.zoom
   
   DrawObj False                                                                    'Erase old line
   Select Case movePoint
   Case 0      'Move of entire object is always relative to previous location. For multiple selects
      XStart = XStart + X
      YStart = YStart + Y
      xEnd = xEnd + X
      YEnd = YEnd + Y
   Case 1
      If X = xEnd And Y = YEnd Then                                            'Line shrunk to zero
         If XStart > xEnd Then
            XStart = xEnd + GRID_SIZE
         ElseIf XStart < xEnd Then
            XStart = xEnd - GRID_SIZE
         End If
         If YStart > YEnd Then
            YStart = YEnd + GRID_SIZE
         ElseIf YStart < YEnd Then
            YStart = YEnd - GRID_SIZE
         End If
      Else
         XStart = X                                            'Placing the point snaps to the grid
         YStart = Y
      End If
   Case 2
      If X = XStart And Y = YStart Then                                        'Line shrunk to zero
         If xEnd > XStart Then
            xEnd = XStart + GRID_SIZE
         ElseIf xEnd < XStart Then
            xEnd = XStart - GRID_SIZE
         End If
         If YEnd > YStart Then
            YEnd = YStart + GRID_SIZE
         ElseIf YEnd < YStart Then
            YEnd = YStart - GRID_SIZE
         End If
      Else
         xEnd = X
         YEnd = Y
      End If
   End Select
   DrawObj
   mappWindow.dirty = True
End Sub
Public Sub SetEdit(Optional edit As Boolean = True) '********************* Puts Object in Edit Mode
   '  activeObject variable always set somewhere else (SetActiveObject)
   If edit Then
      EditPoint Me, 1
      EditPoint Me, 2
      editMode = True
   Else
      EditPoint , 1, False
      EditPoint , 2, False
      editMode = False
   End If
End Sub
Public Function CheckClick(ByVal X As Single, ByVal Y As Single) As Boolean '******* See If Clicked
   '  Enter:   X, Y  Zoomed coordinates
   Dim XOffset As Single, YOffset As Single
   Dim YPoint As Single, XLow As Single, XHigh As Single, YLow As Single, YHigh As Single
   Dim play As Single
   Dim zoom As Single
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Unzoom
   If canvas.name = "picDrafter" Then
      zoom = mappWindow.zoom
   Else
      zoom = 1
   End If
'   X = X / zoom
'   Y = Y / zoom
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check For Click
   XOffset = xEnd - XStart
   YOffset = YEnd - YStart
   
   If XOffset >= 0 Then
      XLow = XStart
      XHigh = xEnd
   Else
      XLow = xEnd
      XHigh = XStart
   End If
   If YOffset >= 0 Then
      YLow = YStart
      YHigh = YEnd
   Else
      YLow = YEnd
      YHigh = YStart
   End If
   
   If XOffset = 0 Then
      YPoint = YEnd
      play = Abs(YEnd - YStart)                                           'Make play the whole line
   Else
      YPoint = (X - XStart) * YOffset / XOffset + YStart             'Y point on line corresponding
                                                                     'to mouse X
      play = Sqr(1 + (YOffset / XOffset) ^ 2) * SLOP                 'Play to either side of YPoint
   End If
   If X < XLow - SLOP Then                                              'See if within bounding box
      CheckClick = False
   ElseIf X > XHigh + SLOP Then
      CheckClick = False
   ElseIf Y < YLow - SLOP Then
      CheckClick = False
   ElseIf Y > YHigh + SLOP Then
      CheckClick = False
   ElseIf Y > YPoint + play Then
      CheckClick = False
   ElseIf Y < YPoint - play Then
      CheckClick = False
   Else
      CheckClick = True
   End If
End Function
Sub MaxOnBoard()
   '  Maintain compatibility with other objects
   If maxX - minX > MAX_BOARD_WIDTH Then '+++++++++++++++++++++++++++++++++++++++++ Object Too Wide
'      MsgBox "Line object too wide to fit on drafting board. It is adjusted to the maximum " _
'             & "width of the board.", vbExclamation + vbOKOnly, "Line Object Size"
'      DrawObj False                                                                         'Erase
      If XStart < xEnd Then
         XStart = 0
         xEnd = GridMax(MAX_BOARD_WIDTH - (maxX - xEnd))
         '  maxX - Xend is the distance that the object extends beyond the farthest X target
         '  for receptor extensions, etc. Subtract it from the MAX_BOARD_WIDTH to get the actual
         '  X-length of the line
      Else
         XStart = GridMax(MAX_BOARD_WIDTH)
         xEnd = GridMin(xEnd - minX)
      End If
   End If
   If maxY - minY > MAX_BOARD_HEIGHT Then '++++++++++++++++++++++++++++++++++++++++ Object Too High
'      MsgBox "Line object too high to fit on drafting board. It is adjusted to the maximum " _
'             & "height of the board.", vbExclamation + vbOKOnly, "Line Object Size"
'      DrawObj False                                                                          'Erase
      If YStart < YEnd Then
         YStart = 0
         YEnd = GridMax(MAX_BOARD_HEIGHT - (maxY - YEnd))
      Else
         YStart = GridMax(MAX_BOARD_HEIGHT)
         YEnd = GridMin(YEnd - minY)
      End If
   End If
   DrawObj                                                                                  'Redraw
End Sub

Rem/////////////////////////////////////////////////////////////////////////////// Property Methods
Public Property Let SelectMode(ByVal vData As Boolean) '************ Redraws Object In Select Color
   If mvarSelectMode <> vData Then
      mvarSelectMode = vData
      DrawObj True
   End If
End Property
Public Property Get SelectMode() As Boolean
    SelectMode = mvarSelectMode
End Property
