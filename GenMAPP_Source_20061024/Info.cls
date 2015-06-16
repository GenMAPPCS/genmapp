VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public editMode As Boolean 'local copy
Private mvarSelectMode As Boolean 'local copy

Public centerX As Single                                                         'Left for Info Box
Public centerY As Single                                                          'Top for Info Box
Public Action As String                                                '"Moving", "Width", "Height"
Public objType As String                                      'Not needed because only one Info Box
Public objKey As String                                            'Unique identifier for an object
   '  String because it is also the identifier in the objLumps collection. Goes in the MAPP
   '  database as long.
Public canvas As Object                                'Drawing surface on which originally created

Public html As String                                                        'The actual coded text
Public minY As Single, minX As Single
Public maxX As Single, maxY As Single                     'Farthest X and Y Board points for object
   '  Used for figuring minimum board size (autosize)

Public Sub Create(Optional X As Single = 200, Optional Y As Single = 200)
   '  Creates new html, such as when a change is made through frmMAPPInfo
   '  Created only on frmDrafter, can be drawn elsewhere

   If objType = "" Then                                                'Info Area Doesn't Exist Yet
      X = 0
      Y = 0
   End If
   objType = "InfoBox"
   html = "<font name=Arial, size=9>"
   With frmMAPPInfo
      If Dat(!txtAuthor) <> "" Then
         html = html & "<p><i><b>Author:</b></i> " & TextToHtml(Dat(!txtAuthor)) & "</p>"
      End If
      If Dat(!txtMaint) <> "" Then
         html = html & "<p><i><b>Maintained by:</b></i> " & TextToHtml(Dat(!txtMaint)) & "</p>"
      End If
      If Dat(!txtEMail) <> "" Then
'         html = html & "<p><i><b>E-mail:</b></i> <a href=""mailto:" & TextToHtml(Dat(!txtEMail)) & """>" & TextToHtml(Dat(!txtEMail)) & "</a></p>"
         html = html & "<p><i><b>E-mail:</b></i> " & TextToHtml(Dat(!txtEMail)) & "</p>"
      End If
      If Dat(!txtModify) <> "" Then
         html = html & "<p><i><b>Last modified:</b></i> " & TextToHtml(Dat(!txtModify)) & "</p>"
      End If
      If Dat(!txtRemarks) <> "" Then
         html = html & "<p>" & TextToHtml(Dat(!txtRemarks)) & "</p>"
      End If
      If Dat(!txtCopyright) <> "" Then
         html = html & "<p>Copyright &#169; " & TextToHtml(Dat(!txtCopyright)) & "</p>"
      End If
   End With
   centerX = X
   centerY = Y
   Set canvas = drawingBoard
   DrawObj
End Sub
Public Sub Delete()
   frmDrafter!picInfo.Cls
   frmDrafter!picInfo.Visible = False
   minX = 0
   minY = 0
   maxX = 0
   maxY = 0
End Sub
Public Sub DrawObj(Optional draw As Boolean = True, Optional outputObj As Object = Nothing)
'   Const LEFT_MARGIN = 0
'   Const TOP_MARGIN = 0
'   Dim obj As Object                    'Could be object within outputObj. Printer is just Printer.
'   '                                    'On frmDrafter or frmPreview, it is !picInfo.
   Dim cXzoom As Single, cYzoom As Single                'Center coordinates translated to screen
      '  This is divided by zoom. If zoom = 2, twice normal size, then all coordinates are doubled.
   Dim lineHeight As Single, lineLength As Single, index As Integer, endCode As Integer
   Dim maxLineLength As Single, totalHeight As Single
   Dim codeType As String, endCodeType As Integer, propertyName As String, propertyValue As String
   Dim char As String * 1
   Dim zoom As Single
'   Dim leftMargin As Single

   If canvas Is Nothing Then Exit Sub                      'No Info box >>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   zoom = canvas.container.zoom
   If outputObj Is Nothing Then Set outputObj = canvas
   If outputObj Is Printer Then
   End If
   
   cXzoom = centerX * zoom
   cYzoom = centerY * zoom

RelocateObject:
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Clear Out Any Old Info
   If minX <> maxX Then                                                            'Old info exists
      outputObj.Line (minX * zoom, minY * zoom)-(maxX * zoom, maxY * zoom), _
                      vbWhite, BF                            'Draw an opaque box where the Info was
'      RedrawArea minX, minY, maxX, maxY, outputObj   'Redraw any other object that was in this area
   End If
   If Not draw Or InStr(cfgLegend, "I") = 0 Then
      minX = 0
      minY = 0
      maxX = 0
      maxY = 0
      Exit Sub                                               '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up For Output Object
   With outputObj                                          '~~~~~~~~~~~~~~~~~~~~~~~~~With outputObj
   .Font.Underline = False
   .FontStrikethru = False
   .CurrentX = cXzoom                                  'cXzoom and cYzoom are top and left for Info
   .CurrentY = cYzoom
   maxLineLength = 0
   totalHeight = 0
   For index = 1 To Len(html) '=========================================================Decode HTML
      char = Mid(html, index, 1)
      Select Case char
      Case "<"
         endCode = InStr(index + 1, html, ">")
         endCodeType = InStr(index + 1, html, " ")
         If endCodeType = 0 Or endCodeType > endCode Then endCodeType = endCode
         codeType = Mid(html, index + 1, endCodeType - index - 1)
         index = endCodeType       'Leave index either beyond code type (ie font) or whole code (>)
         Select Case UCase(codeType)
         Case "I"
            .Font.Italic = True
         Case "/I"
            .Font.Italic = False
         Case "B"
            .Font.Bold = True
         Case "/B"
            .Font.Bold = False
         Case "FONT"
            Do While index < endCode
               index = index + 1 _
                     + NameValue(Mid(html, index + 1, endCode - index), propertyName, propertyValue)
               Select Case propertyName
               Case "NAME"
                  .Font.name = propertyValue
               Case "SIZE"
                  FontSizeFloor Val(propertyValue) * zoom, outputObj
               End Select
            Loop
         Case "P"
         Case "/P"
            .CurrentX = cXzoom
            .CurrentY = .CurrentY + lineHeight
            If lineLength > maxLineLength Then maxLineLength = lineLength
            lineLength = 0
            totalHeight = totalHeight + lineHeight
            lineHeight = 0
         End Select
      Case "&"                                                                   'Special character
         endCode = InStr(index, html, ";")
         If Mid(html, index + 1, 1) = "#" Then                                           'ANSI code
            char = Chr(Val(Mid(html, index + 2, endCode - index)))
         Else                                                              'Other Special Character
            Select Case UCase(Mid(html, index + 1, endCode - index - 1))
            Case "AMP"
               char = "&"
            Case "LT"
               char = "<"
            Case "GT"
               char = ">"
            Case Else
               char = "?"
            End Select
         End If
         index = endCode
         If .TextHeight(char) > lineHeight Then lineHeight = .TextHeight(char)
         lineLength = lineLength + .TextWidth(char)
         outputObj.Print char;
      Case Else                                                                   'Normal character
         If .TextHeight(char) > lineHeight Then lineHeight = .TextHeight(char)
         lineLength = lineLength + .TextWidth(char)
         outputObj.Print char;
      End Select
   Next index
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Boundary Properties
'   minX = cXzoom
'   minY = cYzoom
'   maxX = cXzoom + maxLineLength + 20                                            '20 to clear edges
'   maxY = cYzoom + totalHeight + 20
   minX = centerX
   minY = centerY
   maxX = centerX + (maxLineLength + 20) / zoom                                  '20 to clear edges
   maxY = centerY + (totalHeight + 20) / zoom
   If OutsideBoard(Me) Then GoTo RelocateObject                      'Only if minY or minX negative
   End With                                                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End With
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
   
   DrawObj False                                                                  'Erase old object
   Select Case movePoint
   Case 0 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ For Multiple Selections
      centerX = centerX + X - POINT_SIZE / 2                               'Move is always relative
      centerY = centerY + Y - POINT_SIZE / 2
   Case 1 '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Move
      centerX = X - POINT_SIZE / 2                             'Placing the point snaps to the grid
      centerY = Y - POINT_SIZE / 2
   End Select
   DrawObj
End Sub
Public Function CheckClick(X As Single, Y As Single) As Boolean '******************* See If Clicked
   '  X and Y passed in zoomed coordinates
   Dim zoom As Single
   
   zoom = mappWindow.zoom                                 'Info would only appear on the mappWindow
   
   X = X / zoom
   Y = Y / zoom
   
   If X >= minX And X <= maxX Then
      If Y >= minY And Y <= maxY Then
         CheckClick = True
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
'
'
'   '  activeObject variable always set somewhere else (SetActiveObject)
'   If edit Then
'      editMode = True
'      frmDrafter!picInfo.BorderStyle = vbFixedSingle                   'This indicates it is active
'      frmDrafter!picInfo.MousePointer = vbSizeAll
'      frmDrafter!sbrBar.Panels(1).Text = "Drag to move or drag bottom or left border to size"
'   Else
'      editMode = False
'      frmDrafter!picInfo.BorderStyle = vbBSNone
'      frmDrafter!picInfo.MousePointer = vbDefault
'      frmDrafter!sbrBar.Panels(1).Text = ""
'   End If
End Sub
Sub MaxOnBoard()
   '  Maintain compatibility with other objects
End Sub

Public Property Let SelectMode(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectMode = 5
    mvarSelectMode = vData
End Property


Public Property Get SelectMode() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectMode
    SelectMode = mvarSelectMode
End Property

'Public Property Let objKey(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.objKey = 5
'    mvarID = vData
'End Property
'
'
'Public Property Get objKey() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.objKey
'    objKey = mvarID
'End Property
'
