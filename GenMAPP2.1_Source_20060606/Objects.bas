Attribute VB_Name = "Objects"
Option Explicit

Rem ********************************************************************** Common Object Properties
'  Action      Action currently underway with object: Moving, Sizing, etc.
'  EditMode    True if object being edited
'  CenterX     X position of center of object (doesn't exist for lines)
'  CenterY     Y position of center of object (doesn't exist for lines)

Rem ************************************************************************* Common Object Methods
'  SetEdit     Sent true (or nothing) if object going into edit mode, false to turn it off
'  +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ For Only Objects, Not Controls
'  CheckClick  See if object clicked. Includes SLOP in hit area

Rem *************************************************************** Constants For Object Parameters
Public Const POINT_SIZE = 150                                             'Size of edit point (box)
Public Const LABEL_FONT = "Arial"
Public Const LABEL_SIZE = 10
Public Const LABEL_BOLD = True
Public Const ARROW_WIDTH = 40
Public Const ARROW_ASPECT = 4

Public Const CURVE_ASPECT = 2
Public Const RECEPTOR_EXTENSION = 6 * GRID_SIZE                               'Length of receptor <

Public Const LEFT_BRACE = 0                                                'Orientations for braces
Public Const TOP_BRACE = 1
Public Const RIGHT_BRACE = 2
Public Const BOTTOM_BRACE = 3

Public Const DEFAULT_INCREASE_COLOR = &HC0C0FF                            'Default colors for genes
Public Const DEFAULT_DECREASE_COLOR = &HFFC0C0
Public Const DEFAULT_NOCHANGE_COLOR = &HC0FFFF
Public Const DEFAULT_NOTMET_COLOR = &HD9D9D9

Public Const LINE_STYLE_SOLID = 0
Public Const LINE_STYLE_BROKEN = 1

Public Const SLOP = 100                           'Amount of space around object to allow click hit

Rem ///////////////////////////////////////////////////////////////////////// Control Mouse Actions
Rem *********************************************************** Produces Backpage For Single Object
Function CreateObjPage(obj As Object, Optional path As String = "")
   '  Entry:
   '     obj            The object (gene box) for which backpage being created. For GeneFinder,
   '                    this will be Nothing
   '     path           Path for backpage. Defaults to appPath\Backpages\. For exports the path
   '                    would be appropriate for that export. If path does not end in \ then the
   '                    filename and path is used (as for GeneFinder).
   '  Return:     The path to the created HTML file or blank if it could not be created
'idIn As String, systemIn As String, head As String, dbGene As Database, _
'                        Optional dbExpression As Database = Nothing, _
'                        Optional obj As Object = Nothing, _
'                        Optional path As String = "", _
'                        Optional purpose As Integer = PURPOSE_BACKPAGE) As String
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
   
   Dim head As String
   Dim currentMousePointer As Integer                         'MousePointer on entry, reset on exit
   Dim htmFile As String
'
'   Dim geneRemarks As String, sql As String, rsInfo As Recordset, htmFile As String
'   Dim rsObjects As Recordset, rs As Recordset
'   Dim row As Integer, col As Integer
'   Dim geneTitle As String                               'Each column title in HTML with links, etc
'   Dim columnHeads As String                                    'GeneID heading row for all columns
'   Dim annotations As String                                                       'Annotation data
'   Dim expTable As String                                                 'Expression table in HTML
'   Dim currentMousePointer As Integer                         'MousePointer on entry, reset on exit
'   Dim centerColor As String               '#hex rgb color for HTML output for center and rim genes
'   Dim rimColor As String
'   Dim colColor As String                                            'Color for a particular column
'   Dim legendLink As String                                           'Relative link to Legend page
'   'For AllRelatedGenes()
'      Dim genes As Integer
'      Dim geneIDs(MAX_GENES, 2) As String
'      Dim geneFound As Boolean
'   'For AllExpressionData()
'      Dim rows As Integer
'      Dim rowIDs(MAX_GENES, 1) As String
'      Dim columns As Integer
'      Dim colorSetTitles(MAX_COLORSETS) As String
'      Dim colorSets As Integer
'      Dim titleColors(MAX_COLORSETS, 1) As Long
'      Dim geneColors(MAX_GENES, MAX_COLORSETS) As Long
'      Dim orderNos(MAX_GENES) As Long
'      Dim legendPage As String
'      Dim legendPageTitle As String
'   'For AnnotationData()
'      Dim jumps As String
'   Dim topOfPage As String          'HTML to jump to top of page. Must be same as in AnnotationData
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ See If Backpage Needed
   If obj Is Nothing Then                                                   'No object, no backpage
      CreateObjPage = ""
      GoTo ExitFunction                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   If obj.id = "" Then                                            'Unidentified object, no Backpage
      CreateObjPage = ""
      GoTo ExitFunction                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   If (obj.objType = "Label" Or obj.title = "") And obj.head = "" _
         And obj.remarks = "" And obj.links = "" Then                                      'No data
      CreateObjPage = ""
      GoTo ExitFunction                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
'   topOfPage = "&nbsp;&nbsp;<font size=2><a href=""#Top"">Top</a></font>"

   currentMousePointer = Screen.ActiveForm.MousePointer
   Screen.ActiveForm.MousePointer = vbHourglass
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Determine HTML File Name
   head = obj.title
   If head = "" Then
      head = obj.head
   End If
   
   If path = "" Then path = appPath & "Backpages\"
   If Right(path, 1) <> "\" Then                              'Literal Path Plus File Name Received
      htmFile = path
   Else                                                                'Just Path, Append File Name
      htmFile = TextToFileName(head & "_" & obj.objKey) & ".htm"
      htmFile = ValidHTMLName(htmFile, False)
      htmFile = path & htmFile
   End If
   
   If htmlSuffix <> "" And Dir(htmFile) <> "" Then 'htmlSuffix <> "_1"
      '  There is a htmlSuffix, which means that HTML pages are being produced with MAPPs
      '  for each criterion in a Color Set. The Backpages are the same for all Color Sets,
      '  therefore they are produced only once for all MAPPs in the group.
      CreateObjPage = htmFile
      GoTo ExitFunction                                    'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up HTML Backpage File
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
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ HTML Data
   If obj.remarks <> "" Then
      Print #31, "<p>&nbsp;</p>"
      Print #31, "<p>" & obj.remarks & "</p>"
   End If
   
   If obj.links <> "" Then
      Print #31, "<p>&nbsp;</p>"
      Print #31, "<a href=""http://" & obj.links & """>Link</a>"
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ End HTML file
   Print #31, "</body>"
   Print #31, "</html>"
   Close #31
   CreateObjPage = htmFile
   
ExitFunction:
   Screen.ActiveForm.MousePointer = currentMousePointer
End Function
'**************************************************************** Displays Backpage For Single Gene
Sub ShowObjPage(obj As Object)
   '  Entry:
   '     obj            Gene object
   Dim head As String, path As String
   
   If obj.objType <> "Gene" Then Exit Sub                  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   head = obj.head
   If head = "" Then head = obj.title
'   windowTitle = pageTitle ' & " Backpage"
   
   path = CreateObjPage(obj)
   
   If path = "" Then
      MsgBox "No backpage data available.", vbExclamation + vbOKOnly, "Generating Object Page"
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If

   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Show Browser With Object Page
   Dim IE As Object
On Error GoTo NewBrowser           'If the backpage does not already exist, a new IE object created
   AppActivate head & " Backpage"
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

NewBrowser:
   Set IE = CreateObject("InternetExplorer.Application")
   IE.visible = True
   IE.Navigate path
   IE.StatusText = head
End Sub

'********************************************* Create a Target With Which To Drag A Point On a Line
Sub EditPoint(Optional element As Variant, Optional number As Integer = 1, _
              Optional draw As Boolean = True)
   '  If draw = false, delete the target
   Dim index As Integer, X As Single, Y As Single
   Dim centerX As Single, centerY As Single
   Dim zoom As Single
   
   With mappWindow
      !picPoint(number).visible = False                                   'Start with point deleted
      If draw Then
         zoom = mappWindow.zoom
         centerX = element.centerX * zoom
         centerY = element.centerY * zoom
         Select Case number
         Case 1 '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Center, Moving picPoint
            Select Case element.objType
            Case "objLine"
               !picPoint(1).Left = element.XStart * zoom - POINT_SIZE / 2        'Beginning of line
               !picPoint(1).Top = element.YStart * zoom - POINT_SIZE / 2
            Case "InfoBox", "Legend"
               !picPoint(1).Left = centerX                                      'Upper-left of Info
               !picPoint(1).Top = centerY
            Case Else                                                             'Lumps and braces
               !picPoint(1).Left = centerX - POINT_SIZE / 2
               !picPoint(1).Top = centerY - POINT_SIZE / 2
            End Select
         Case 2 '++++++++++++++++++++++++++++++++++++++Right Edge, General Sizing Or Width picPoint
                                     'Except for lines, in which case it is the arrow or second end
            Select Case element.objType
            Case "objLine"                                                             'End of line
               !picPoint(2).Left = element.xEnd * zoom - POINT_SIZE / 2
               !picPoint(2).Top = element.YEnd * zoom - POINT_SIZE / 2
            Case "objBrace"                     'Right or upper point of brace. Moving changes span
               Dim XOffset As Single, YOffset As Single
               Select Case element.Orientation
               Case TOP_BRACE
                  XOffset = element.Span / 2
                  YOffset = -(element.curvature * 2 + element.thickness)
               Case LEFT_BRACE
                  XOffset = element.curvature * 2 + element.thickness
                  YOffset = element.Span / 2
               Case RIGHT_BRACE
                  XOffset = -(element.curvature * 2 + element.thickness)
                  YOffset = element.Span / 2
               Case BOTTOM_BRACE
                  XOffset = element.Span / 2
                  YOffset = element.curvature * 2 + element.thickness
               End Select
               !picPoint(2).Left = element.centerX + XOffset * zoom - POINT_SIZE / 2
               !picPoint(2).Top = element.centerY - YOffset * zoom - POINT_SIZE / 2
            Case "Rectangle"
               X = element.wide * zoom / 2 * Cos(element.rotation)              'Rotate edge point,
               Y = element.wide * zoom / 2 * Sin(element.rotation)              'Y starts at 0
               !picPoint(2).Left = centerX + X - POINT_SIZE / 2
               !picPoint(2).Top = centerY + Y - POINT_SIZE / 2
            Case "Oval", "Arc"
               X = element.wide * zoom * Cos(element.rotation)                  'Rotate edge point,
               Y = element.wide * zoom * Sin(element.rotation)                  'Y starts at 0
               !picPoint(2).Left = centerX + X - POINT_SIZE / 2
               !picPoint(2).Top = centerY + Y - POINT_SIZE / 2
            Case "Brace"                        'Right or upper point of brace. Moving changes span
               Select Case element.rotation
               Case 0                                                                    'TOP_BRACE
                  X = element.wide / 2
                  Y = element.high                                     'Offset from center picPoint
               Case 1                                                                  'RIGHT_BRACE
                  X = -element.high
                  Y = element.wide / 2
               Case 2                                                                 'BOTTOM_BRACE
                  X = -element.wide / 2
                  Y = -element.high
               Case 3                                                                   'LEFT_BRACE
                  X = element.high
                  Y = -element.wide / 2
               End Select
               !picPoint(2).Left = centerX + X * zoom - POINT_SIZE / 2
               !picPoint(2).Top = centerY + Y * zoom - POINT_SIZE / 2
            Case "Gene"
               !picPoint(2).Left = centerX + element.wide / 2 * zoom - POINT_SIZE / 2
               !picPoint(2).Top = centerY - POINT_SIZE / 2
            Case "Vesicle"
               !picPoint(2).Left = centerX + element.wide * zoom - POINT_SIZE / 2
               !picPoint(2).Top = centerY - POINT_SIZE / 2
            Case "ProteinA"
               !picPoint(2).Left = centerX + element.wide * 1.7 / 1.1 * zoom - POINT_SIZE / 2
                          '1.7 because of circle overlap. 1.1, see objLump PROTEINA_ASPECT constant
               !picPoint(2).Top = centerY - POINT_SIZE / 2
            Case "ProteinB"
               !picPoint(2).Left = centerX + element.wide * 1.7 * zoom - POINT_SIZE / 2
                                                                     '1.7 because of circle overlap
               !picPoint(2).Top = centerY - POINT_SIZE / 2
            Case Else
               !picPoint(2).Left = centerX + element.wide / 2 * zoom - POINT_SIZE / 2
               !picPoint(2).Top = centerY - POINT_SIZE / 2
            End Select
         Case 3 '++++++++++++++++++++++++++++++++++++++++++++++++++++++ Lower Edge, Height picPoint
            Select Case element.objType
            Case "Rectangle"
               X = -element.high / 2 * zoom * Sin(element.rotation)             'Rotate edge point,
               Y = element.high / 2 * zoom * Cos(element.rotation)                   'X starts at 0
               !picPoint(3).Left = centerX + X - POINT_SIZE / 2
               !picPoint(3).Top = centerY + Y - POINT_SIZE / 2
            Case "Oval", "Arc"
               X = -element.high * zoom * Sin(element.rotation)                 'Rotate edge point,
               Y = element.high * zoom * Cos(element.rotation)                  'X starts at 0
               !picPoint(3).Left = centerX + X - POINT_SIZE / 2
               !picPoint(3).Top = centerY + Y - POINT_SIZE / 2
            End Select
         Case 4 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Rotation picPoint
            Select Case element.objType
            Case "Rectangle"                                                                'Corner
               X = -element.wide / 2 * Cos(element.rotation) + element.high / 2 _
                 * Sin(element.rotation)
               Y = -element.wide / 2 * Sin(element.rotation) - element.high / 2 _
                 * Cos(element.rotation)
            Case "Poly"
               X = -element.wide / 2 * Cos(element.rotation)                      'Y starts at zero
               Y = -element.wide / 2 * Sin(element.rotation)
            Case "Oval", "Arc"
               X = -element.wide * Cos(element.rotation)                          'Y starts at zero
               Y = -element.wide * Sin(element.rotation)
            Case "Brace"
               Select Case element.rotation
               Case 0                                                                    'TOP_BRACE
                  X = -element.wide / 2
                  Y = element.high                                     'Offset from center picPoint
               Case 1                                                                  'RIGHT_BRACE
                  X = -element.high
                  Y = -element.wide / 2
               Case 2                                                                 'BOTTOM_BRACE
                  X = element.wide / 2
                  Y = -element.high
               Case 3                                                                   'LEFT_BRACE
                  X = element.high
                  Y = element.wide / 2
               End Select
            Case "CellB"
               X = -element.wide * Cos(element.rotation)                          'Y starts at zero
               Y = -element.wide * Sin(element.rotation)
            End Select
            !picPoint(4).Left = centerX + X * zoom - POINT_SIZE / 2
            !picPoint(4).Top = centerY + Y * zoom - POINT_SIZE / 2
            
   '        X1 = X * Cos(rotation) - Y * Sin(rotation)
   '        Y1 = X * Sin(rotation) + Y * Cos(rotation)
         End Select
'         !picPoint(number).Line (0, 0)-(POINT_SIZE - 1, POINT_SIZE - 1), , B
'         !picPoint(number).Line (0, 0)-(POINT_SIZE - 1, POINT_SIZE - 1)
'         !picPoint(number).Line (POINT_SIZE - 1, 0)-(0, POINT_SIZE - 1)
         !picPoint(number).Enabled = True
         !picPoint(number).visible = True
      End If
   End With
End Sub
Function Exists(obj As Object) As Boolean '*************************** Test For Existence Of Object
   Dim typeString As Long                'Test for object type to force error on nonexistent object
   
   On Error GoTo ExistsError
   z = obj.objType 'Force error on nonexistent object
   Exists = True
   Exit Function
   
ExistsError:
   Exists = False
End Function
'Sub SetNewObject(obj As Object) '***************************** Set Type Of New Object To Be Dropped
'   If Not newObject Is Nothing Then '----------------------------------------------Unset Old Object
'      newObject.SelectMode = False
'      Select Case newObject.objType
'      Case "objLine"
'         If lineStarted Then
'            frmDrafter.ForeColor = vbWhite                                         'Delete X marker
'            frmDrafter.Line (XStart - 50, YStart)-Step(100, 0)
'            frmDrafter.Line (XStart, YStart - 50)-Step(0, 100)
'            frmDrafter.ForeColor = vbBlack
'            lineStarted = False
'         End If
'      Case Else
'      End Select
'      frmDrafter!sbrBar.Panels("Instructions").Text = ""                            'Any status bar entry erased
'      frmDrafter.MousePointer = vbDefault
'   End If
'   Set newObject = obj
'   If Not newObject Is Nothing Then '------------------------------------------------Set New Object
'      MultipleObjectDeselectAll
'      SetActiveObject Nothing                                         'Any active object turned off
'      newObject.SelectMode = True
'      frmDrafter.MousePointer = vbCrosshair
'      frmDrafter!sbrBar.Panels("Instructions").Text = newObject.statusBarPlaceText
'      frmDrafter.show
'   End If
'End Sub

Function OutsideBoard(obj As Variant) As Boolean '******************** Object Outside Current Board
   '  Must be called before any drawing takes place because object may be relocated
   '  All parameters in board coordinates
   '  Tests to see if object outside current board top, edge, width, and/or height
   '  All adjustments must coincide with grid coordinates
   Const fudge = 10                                 'To allow pixel overlap and approximation error
   
   If loading Then Exit Function                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   If creatingMappSet Then Exit Function
   
   With obj                                                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~With
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Object Too Wide Or High
   If .maxX - .minX > MAX_BOARD_WIDTH Or .maxY - .minY > MAX_BOARD_HEIGHT Then
      obj.MaxOnBoard
   End If
   
   If .minX < 0 Or .minY < 0 Then '++++++++++++++++++++++++++++++++++++++++++++ Object Before Board
         'Test this first because it may move object
         'If object outside minY or minX, center of object adjusted and true returned. Each object
         'must call this routine before drawing. If true returned, then object must
         'goto RelocateObject, recalculate, retest, and redraw
'      MsgBox "Object falls above or to the left of the Drafting Board. The object's location " _
'             & "is adjusted to keep it within the confines of the Board.", _
'             vbExclamation + vbOKOnly, "Board Parameter Problem"
      mappWindow.mouseIsDown = False
      If TypeName(obj) = "objLump" Then '-----------------------------------------------------Lumps
         If .minX < 0 Then .centerX = GridMin(.centerX - .minX)
         If .minY < 0 Then .centerY = GridMin(.centerY - .minY)
      ElseIf TypeName(obj) = "objLine" Then '-------------------------------------------------Lines
         If .minX < 0 Then
            .XStart = GridMin(.XStart - .minX + fudge)
            .xEnd = GridMin(.xEnd - .minX + fudge)
         End If
         If .minY < 0 Then
            .YStart = GridMin(.YStart - .minY + fudge)
            .YEnd = GridMin(.YEnd - .minY + fudge)
         End If
      ElseIf TypeName(obj) = "objLegend" Then '----------------------------------------------Legend
         If .minX < 0 Then .centerX = GridMin(.centerX - .minX)
         If .minY < 0 Then .centerY = GridMin(.centerY - .minY)
         .DrawObj
      ElseIf TypeName(obj) = "objInfo" Then '--------------------------------------------------Info
         If .minX < 0 Then .centerX = GridMin(.centerX - .minX)
         If .minY < 0 Then .centerY = GridMin(.centerY - .minY)
         .DrawObj
      ElseIf TypeName(obj) = "objSelectArea" Then '--------------------------------------SelectArea
         If .minX < 0 Then
            .maxX = .maxX - .minX
            .minX = 0
         End If
         If .minY < 0 Then
            .maxY = .maxY - .minY
            .minY = 0
         End If
         Screen.ActiveForm.shpSelected.Left = .minX
         Screen.ActiveForm.shpSelected.Top = .minY
      End If
      OutsideBoard = True
      dontClick = False
      If obj.editMode Then                                     'Move editpoints to changed location
         obj.SetEdit False
         obj.SetEdit True
      End If
      Exit Function     'Object must be moved; exit, relocate, and retest >>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If .maxX > MAX_BOARD_WIDTH Or .maxY > MAX_BOARD_HEIGHT Then '+++++++++++++++ Object Beyond Board
         'Test this first because it may move object
         'If object outside bottom or edge, center of object adjusted and true returned. Each
         'object must call this routine before drawing. If true returned, then object must
         'goto RelocateObject, recalculate, retest, and redraw
'      MsgBox "Object falls below or to the right of the Drafting Board. The object's location " _
'             & "is adjusted to keep it within the confines of the Board.", _
'             vbExclamation + vbOKOnly, "Board Parameter Problem"
      Screen.ActiveForm.mouseIsDown = False
      If TypeName(obj) = "objLump" Then '-----------------------------------------------------Lumps
         If .maxX > MAX_BOARD_WIDTH Then
            .centerX = GridMax(.centerX - (.maxX - MAX_BOARD_WIDTH) - fudge)
         End If
         If .maxY > MAX_BOARD_HEIGHT Then
            .centerY = GridMax(.centerY - (.maxY - MAX_BOARD_HEIGHT) - fudge)
         End If
      ElseIf TypeName(obj) = "objLine" Then '-------------------------------------------------Lines
         If .maxX > MAX_BOARD_WIDTH Then
            .XStart = GridMax(.XStart - (.maxX - MAX_BOARD_WIDTH) - fudge)
'            If .XEnd > .XStart Then
               .xEnd = GridMax(.xEnd - (.maxX - MAX_BOARD_WIDTH) - fudge)
'            Else
'               .XStart = GridMax(.XStart - (.maxX - MAX_BOARD_WIDTH) - fudge)
'            End If
            If Abs(.xEnd - .XStart) > MAX_BOARD_WIDTH Then '__________________________Line Too Long
               If .xEnd > .XStart Then                                           'Make it max width
                  .XStart = fudge
                  .xEnd = MAX_BOARD_WIDTH - fudge
               Else
                  .XStart = MAX_BOARD_WIDTH - fudge
                  .xEnd = fudge
               End If
            End If
         End If
         If .maxY > MAX_BOARD_HEIGHT Then
            .YStart = GridMax(.YStart - (.maxY - MAX_BOARD_HEIGHT) - fudge)
            .YEnd = GridMax(.YEnd - (.maxY - MAX_BOARD_HEIGHT) - fudge)
            If Abs(.YEnd - .YStart) > MAX_BOARD_HEIGHT Then '_________________________Line Too Long
               If .YEnd > .YStart Then                                           'Make it may width
                  .YStart = fudge
                  .YEnd = MAX_BOARD_HEIGHT - fudge
               Else
                  .YStart = MAX_BOARD_HEIGHT - fudge
                  .YEnd = fudge
               End If
            End If
         End If
      ElseIf TypeName(obj) = "objLegend" Or TypeName(obj) = "objInfo" Then  '----------Legend, Info
         If .maxX > MAX_BOARD_WIDTH Then
            .centerX = GridMax(.centerX - (.maxX - MAX_BOARD_WIDTH) - fudge)
         End If
         If .maxY > MAX_BOARD_HEIGHT Then
            .centerY = GridMax(.centerY - (.maxY - MAX_BOARD_HEIGHT) - fudge)
         End If
'         .DrawObj
      ElseIf TypeName(obj) = "objSelectArea" Then  '-------------------------------------SelectArea
         If .maxX > MAX_BOARD_WIDTH Then
            .minX = .minX - (.maxX - MAX_BOARD_WIDTH)
            .maxX = MAX_BOARD_WIDTH
         End If
         If .maxY > MAX_BOARD_HEIGHT Then
            .minY = .minY - (.maxY - MAX_BOARD_HEIGHT)
            .maxY = MAX_BOARD_HEIGHT
         End If
         Screen.ActiveForm.shpSelected.Left = .minX
         Screen.ActiveForm.shpSelected.Top = .minY
      End If
      OutsideBoard = True
      dontClick = False
      If obj.editMode Then                                     'Move editpoints to changed location
         obj.SetEdit False
         obj.SetEdit True
      End If
      Exit Function     'Object must be moved; exit, relocate, and retest >>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Object Beyond Current Board
   'If object beyond drawingBoard Width or Height then Width and/or Height
   'adjusted. No relocation of object occurs so false is returned.
   Dim boardAdjusted As Boolean
   
'   If .maxX > drawingBoard.Width Then '-----------------------------------------Object Beyond Board
   If .maxX > mappWindow.boardWidth Then '---------------------------------------Object Beyond Board
      mappWindow.boardWidth = .maxX + 1                     'Allow for single datatype approximation
      drawingBoard.Width = mappWindow.boardWidth * mappWindow.zoom
      boardAdjusted = True
      dontClick = False
      If TypeName(obj) = "objLegend" Or TypeName(obj) = "objInfo" Then
         OutsideBoard = True                                       'Force Legend and Info to redraw
      End If
   End If
'   If .maxY > drawingBoard.Height Then '----------------------------------------Object Beyond Board
   If .maxY > mappWindow.boardHeight Then '--------------------------------------Object Beyond Board
      mappWindow.boardHeight = .maxY + 1                    'Allow for single datatype approximation
      drawingBoard.Height = mappWindow.boardHeight * mappWindow.zoom
      boardAdjusted = True
      dontClick = False
      If TypeName(obj) = "objLegend" Or TypeName(obj) = "objInfo" Then
         OutsideBoard = True                                       'Force Legend and Info to redraw
      End If
   End If
   If boardAdjusted Then
      mappWindow.mouseIsDown = False
      mappWindow.ScrollBars
      dontClick = False
   End If
   End With                                                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End With
End Function

Sub CancelUndo() '******************************************************************* Disables Undo
'   frmDrafter.mnuUndo.Enabled = False
   Set oldObj = Nothing
End Sub

Sub HitRange(obj As Variant) '***************************** Graphically Shows Hit Range For Testing
   '  Currently only for lumps
   Dim X As Single, Y As Single, startX As Single, startY As Single, endX As Single, endY As Single
   Dim over As Single
   
   With Screen.ActiveForm
      .DrawWidth = 1
      If obj.objType = "objLine" Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++ Lines
         If obj.XStart < obj.xEnd Then
            startX = obj.XStart
            endX = obj.xEnd
         Else
            startX = obj.xEnd
            endX = obj.XStart
         End If
         If obj.YStart < obj.YEnd Then
            startY = obj.YStart
            endY = obj.YEnd
         Else
            startY = obj.YEnd
            endY = obj.YStart
         End If
         If obj.style = "Arc" Then
            over = Max(Abs(endX - startX) / CURVE_ASPECT / 2, Abs(endY - startY) / CURVE_ASPECT / 2)
         End If
      Else '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Lumps
         startX = obj.centerX - obj.wide
         startY = obj.centerY - obj.high
         endX = obj.centerX + obj.wide
         endY = obj.centerY + obj.high
      End If
      
      over = over + 100
      For X = startX - over To endX + over Step 10 '++++++++++++++++++++++++++++++++++ Draw Graphic
         For Y = startY - over To endY + over Step 10
            If .Point(X, Y) = vbBlack Then
               .foreColor = vbBlack
            ElseIf obj.CheckClick(X, Y) Then
               .foreColor = vbRed
            Else
               .foreColor = DEFAULT_NOCHANGE_COLOR
            End If
            Screen.ActiveForm.PSet (X, Y)
            DoEvents
         Next Y
      Next X
   End With
End Sub


