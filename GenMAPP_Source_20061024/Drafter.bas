Attribute VB_Name = "Drafter"
Option Explicit

Public Const TWIPS_CM = 567                                                          'Twips per cm.
Public Const TWIPS_PT = 24                                                    'Twips per font point
Public Const GRID_SIZE = 50
Public Const MAX_GENES_PER_SET = 35     'Genes allowable under one ID (e.g. Genbanks per SwissProt)
Public Const MAX_PRIMARY_TYPES = 30

Public Const INITIAL_WINDOW_WIDTH = 9000
Public Const INITIAL_WINDOW_HEIGHT = 6000

Public Const INITIAL_BOARD_WIDTH = ((40.1 * TWIPS_CM) \ GRID_SIZE) * GRID_SIZE  'Floored to nearest
Public Const INITIAL_BOARD_HEIGHT = ((30 * TWIPS_CM) \ GRID_SIZE) * GRID_SIZE      'grid coordinate
Public Const MIN_BOARD_WIDTH = 8 * TWIPS_CM
Public Const MIN_BOARD_HEIGHT = 6 * TWIPS_CM
Public Const MAX_BOARD_WIDTH = 57 * TWIPS_CM                                           '32319 twips
Public Const MAX_BOARD_HEIGHT = 57 * TWIPS_CM

Public Const GENE_MIN_HEIGHT = GRID_SIZE * 6       'Defined globally because Legend and Lump use it

'**************************************************************************** Dimensional Variables
Public boardWidthAdjust As Single                                    'Size of side borders in twips
Public boardHeightAdjust As Single                                    'Size of title, menu in twips
   '  0,0 on a window refers to the client area, inside the borders, title, and menu bar
   '  Width and Height are the dimensions of the entire window -- outside the borders. Dumb!
   '  These variables are set in the Main() based on the TwipsPerPixel, which varies according
   '  to the Windows setting for font size. Eg. with small fonts, this is 15. With small fonts
   '  (1.25x) this is 12.
'Public Const BOARD_WIDTH_ADJUST = 92
'Public Const BOARD_HEIGHT_ADJUST = 696

Public Type UndoType
   obj As Object       'Object that was moved
   X As Single         'Position before moving
   Y As Single         'Position after moving
End Type
Public newObj As Object, oldObj As Object

Public loading As Boolean            'Initial load of file or blank board. Resize doesn't set dirty
Public catalogName As String                      'Path and name for current open Table of Contents

'Public ScreenWidth As Single   'for testing
'Public ScreenHeight As Single

'Public Expression As String                                'Expression dataset file name (name.gex)
                                                       'If this is not "" then dbExpression is open
   '  Get rid of this variable. It always follows dbExpression and is always dbExpression.Name ????
'Public colorSet As String                  'ColorSet title (eg. 8-Week Average) from ColorSet table
Public GDBVersion As String                                            'Version of GenMAPP Database
Public commandLine As String                                   'Arguments coming in at command line
Public z As Variant                                                                  'Misc variable
Public callingFunction As String
Public dontClick As Boolean                 'Used with Delay timer to implement double-click events
Public mappWindow As Form                                      'The currently-being-used frmDrafter
   '  The mappWindow contains all the information about the current MAPP. It is always
   '  frmDrafter. It is set when a frmDrafter is activated. It need not be unset
   '  because some frmDrafter will always be active. mappWindow has all the info
   '  about objects, dbGene, dbExpression, MAPPName, rsColorSet, etc.
   '  The mappWindow itself is not always in focus, such as when making a choice from the
   '  toolbar or creating a MAPP set.
Public drawingBoard As Object                                     'The picDrafter on the mappWindow
   '  This is always the mappWindow.picDrafter, where stuff is being currently drawn.
   '  It is also set when a frmDrafter is activated.

'**************************************************************************** MAPP-Specific Globals
'These should be part of an instance of frmDrafter but VB doesn't allow Public declaration of
'arrays in object modules. Did not want to go through Let/Get routines if we are not going to set
'up MDI MAPP windows in VB.
Public colorIndexes(MAX_COLORSETS) As Integer    'Indexes in the Display table for chosen colorsets
   '  Eg: If colorIndexes(2) = 4 then that colorset uses Color4 column. This is also the value of
   '  the SetNo in ColorSet table.
   '  colorIndexes(0) is the number of colorsets
Public valueIndex As Integer       'Index in the Display table for the value to display. Eg: Value2
   '  This array and variable determine the sets of Color Sets and the value being displayed on a
   '  MAPP. From them, everything else can be determined for a particular Expression Dataset.
   

'*************************************************************** Recordset Of All Color Sets Chosen
Public Sub SetRsColorSet(Optional dataObj As Object = Nothing)
   '  Entry    dataObj   Object to which recordset belongs, typically a mappWindow
   '                     Defaults to current mappWindow
   '     globals needed:
   '     colorIndexes()  SetNo columns of chosen colorsets. colorIndexes(0) is number chosen
   '  Return   Recordset of chosen colorsets or Nothing for object, typically mappWindow.rsColorSet
   Dim sql As String
   
   If dataObj Is Nothing Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Default
      Set dataObj = mappWindow
   End If
   
   With dataObj
      If .dbExpression Is Nothing Then '++++++++++++++++++++++++++++++++++++++++++++++++ No Dataset
         Set .rsColorSet = Nothing
         Exit Sub                                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
      
      If colorIndexes(0) = 0 Then '+++++++++++++++++++++++++++++++++++++++++++ No Color Sets Chosen
         Set .rsColorSet = Nothing
      Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ One Or More Color Sets
         sql = colorIndexes(1)
         For i = 2 To colorIndexes(0)
            sql = sql & ", " & colorIndexes(i)
         Next i
         Set .rsColorSet = .dbExpression.OpenRecordset( _
               "SELECT * FROM ColorSet WHERE SetNo IN (" & sql & ")")
      End If
   End With
End Sub
Function SetFileName(fileType As String, Optional ByVal title As String = "*", _
                     Optional ByVal path As String = "")
   '  Opens or creates empty files to be written
   '  Enter:   fileType Sets extension in dialog
   '           title    Default title in dialog
   '           path     Default path or mruExportPath or mruMAPPPath if not passed
   '  Return:  File path and name after going through the common dialog
   '           "CANCEL" if user cancels

   Dim file As String, extension As String, dot As Integer
   
   If title = "" Then title = "*"
   dot = InStrRev(title, ".")
   If dot > 0 And dot > Len(title) - 5 Then
      title = Left(title, dot - 1)
   End If
   With Screen.ActiveForm                                  '~~~~~~~~~~~~~~~~~With Screen.ActiveForm
   Select Case fileType
   Case "HTML"
      If path = "" Then path = mruExportPath
      .dlgDialog.Filter = "HTML files (*.htm)|*.htm"
      .dlgDialog.InitDir = path
      .dlgDialog.FileName = title & ".htm"
      extension = ".htm"
   Case "BMP"
      If path = "" Then path = mruExportPath
      .dlgDialog.Filter = "Bitmap files (*.bmp)|*.bmp"
      .dlgDialog.InitDir = path
      .dlgDialog.FileName = title & ".bmp"
      extension = ".bmp"
   Case "JPEG"
      If path = "" Then path = mruExportPath
      .dlgDialog.Filter = "JPEG files (*.jpg)|*.jpg"
      .dlgDialog.InitDir = path
      .dlgDialog.FileName = title & ".jpg"
      extension = ".jpg"
   End Select
   If path = "" Then path = mruMappPath                                       'Default to MAPP path

On Error GoTo ErrorHandler

ReEnter:
   .dlgDialog.CancelError = True
   .dlgDialog.DialogTitle = "Export Destination"
   .dlgDialog.FLAGS = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist
   .dlgDialog.InitDir = path
   .dlgDialog.ShowSave
   file = .dlgDialog.FileName
   If InStr(file, ".") = 0 Then
      file = file & extension
   End If
   path = GetFolder(file)
   If fileType = "HTML" Then
      s = ValidHTMLName(GetFile(file))
      If s <> GetFile(file) Then
         .dlgDialog.FileName = s
         GoTo ReEnter
      End If
      file = path & s
   End If

   If UCase(Dir(file)) = UCase(Mid(file, InStrRev(file, "\") + 1)) Then
      Select Case MsgBox("Do you want to replace the current " & file & "?", _
                  vbYesNoCancel + vbQuestion, "Saving " & fileType)
      Case vbNo
         GoTo ReEnter                                      '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      Case vbCancel
         SetFileName = "CANCEL"
         GoTo ExitFunction                                 'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End Select
   End If
   End With                                                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End With
   
   Open file For Output As #19                     'Try opening file to see if drive can be written
   Close #19
   Kill file
   
   SetFileName = file
ExitFunction:
   Exit Function
   
ErrorHandler:
   Select Case Err.number
   Case 70
      MsgBox Err.Description & ". " & file & " possibly open in some other program.", _
            vbCritical, fileType & " Export Error"
      SetFileName = "ERROR"
   Case 75
      MsgBox Err.Description & ". " & file & " possibly a read-only drive.", _
            vbCritical, fileType & " Export Error"
      Resume ReEnter
   Case 76
      MsgBox Err.Description & ". " & file & " probably a nonexistent folder or drive.", _
            vbCritical, fileType & " Export Error"
      Resume ReEnter
   Case 32755                                                       'Not an error if just cancelled
      SetFileName = "CANCEL"
   Case Else
      MsgBox Err.Description, vbCritical, fileType & " Export Error"
      SetFileName = "ERROR"
   End Select
   On Error GoTo 0
   Resume ExitFunction
End Function

Sub RedrawArea(startX As Single, startY As Single, endX As Single, endY As Single, _
               Optional outputObj As Object = Nothing)
   '  Redraws any objects that fall within passed rectangle
   '  Entry    startX, startY    Upper-left corner of rectangle
   '           endX, endY        lower-right corner of rectangle
   '           outputObj         Object on which drawn, typically picDrafter
   
End Sub

Rem /////////////////////////////////////////////////////////////////////////// Grid Snap Functions
Function GridMin(Size As Single) As Single '*************************************** Fit Within Grid
   '  Return:  Even grid size at least as large as size passed
   Dim minimum As Single
   
   minimum = (Size \ GRID_SIZE) * GRID_SIZE                                               'To floor
   If Size > minimum Then minimum = minimum + GRID_SIZE                                    'To ceil
   GridMin = minimum
End Function
Function GridMax(Size As Single) As Single '*************************************** Fit Within Grid
   '  Return:  Even grid size <= size passed
   Dim maximum As Single
   
   maximum = (Size \ GRID_SIZE) * GRID_SIZE                                               'To floor
   GridMax = maximum
End Function
Function GridCoord(coord As Single) As Single '*************************************** Snap To Grid
   '  Entry:   coord    X or Y coordinate zoomed (typically comes from picDrafter)
   '  Return:  coord snapped to nearest grid intersection adjusted to zoom
   
   GridCoord = Round(coord / (GRID_SIZE * mappWindow.zoom)) * (GRID_SIZE * mappWindow.zoom)
End Function

Rem //////////////////////////////////////////////////////////////////////////////// Rotated Points
Function XPoint(originX As Single, distance As Single, ByVal angle As Single)
   angle = (angle - 90) * 3.14159 / 180
   XPoint = originX + Cos(angle) * distance
End Function
Function YPoint(originY As Single, distance As Single, ByVal angle As Single)
   angle = (angle - 90) * 3.14159 / 180
   YPoint = originY + Sin(angle) * distance
End Function

Sub BrokenLine(container As Object, XStart As Single, YStart As Single, _
               xEnd As Single, YEnd As Single)
   Dim dash As Single                                   'Length of solid piece. Open piece is equal
   Dim XLen As Single, YLen As Single
   Dim XCur As Single, YCur As Single, XStop As Single, YStop As Single
   
   If XStart = xEnd And YStart = YEnd Then Exit Sub        'No line >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   If container.DrawWidth = 1 Then
      container.DrawStyle = vbDot
      container.Line (XStart + XCur, YStart + YCur)-(xEnd, YEnd), container.foreColor
      container.DrawStyle = vbSolid
      Exit Sub             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   dash = container.DrawWidth * 25
   If XStart = xEnd Then
      XLen = 0
      YLen = dash * Sgn(YEnd - YStart)
   Else
      XLen = dash / (1 + Abs((YEnd - YStart) / (xEnd - XStart))) ^ 0.5 * Sgn(xEnd - XStart)
      YLen = XLen * (YEnd - YStart) / (xEnd - XStart)
   End If
   XStop = XStart + Abs(xEnd - XStart)  'Interpret all aspects of the line as positive to determine
   YStop = YStart + Abs(YEnd - YStart)  'how far to continue
   Do While XStart + Abs(XCur) + Abs(XLen) <= XStop And YStart + Abs(YCur) + Abs(YLen) <= YStop
      container.Line (XStart + XCur, YStart + YCur)-Step(XLen, YLen), container.foreColor
      XCur = XCur + 2 * XLen
      YCur = YCur + 2 * YLen
   Loop
   If XStart + Abs(XCur) < XStop Or YStart + Abs(YCur) < YStop Then
      container.Line (XStart + XCur, YStart + YCur)-(xEnd, YEnd), container.foreColor
   End If
End Sub
Function InvalidChrForDB(str As Variant) As String '********************* Checks Chrs For Databases
   '  Receive: str   String destined for DB as gene ID, expression column title, or criterion
   '  Return:  First invalid character or empty string
   For i = 1 To Len(str)
      Select Case Mid(str, i, 1)
      Case "'", Chr(34), ",", "|", "`", "[", "]", "!", ".", "$"
         InvalidChrForDB = Mid(str, i, 1)
         Exit Function                                      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End Select
   Next i
End Function
Sub RemoveLocalGenes()
'   If MsgBox("Are you sure you want to remove all the local genes from GenMAPP.gdb?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'   dbGene.Execute "DELETE FROM GenMAPP WHERE Local IS NOT NULL"
'   dbGene.Execute "DELETE FROM GenBank WHERE Local IS NOT NULL"
'   dbGene.Execute "DELETE FROM SwissProt WHERE Local IS NOT NULL"
'   dbGene.Execute "DELETE FROM SwissNo WHERE Local IS NOT NULL"
'   dbGene.Execute "DELETE FROM Other WHERE Local IS NOT NULL"
'   End
End Sub

Sub MakeMAPPBank()
'   Dim dbGene As Database
'   Dim rsMAPPBank As Recordset
'   Dim rsGenMAPP As Recordset
'   Dim rsSwissProt As Recordset
'   Dim rsGenBank As Recordset
'   Dim rs As Recordset
'
'   Set dbGene = OpenDatabase(GenMAPPPath)
'   Set rs = dbGene.OpenRecordset("SELECT GenMAPP, GenBank FROM GenMAPP, GenBank WHERE GenMAPP.SwissNo = GenBank.SwissNo")
'   Do Until rs.EOF
'      dbGene.Execute "INSERT INTO MAPPBank(GenMAPP, GenBank) VALUES ('" & rs!GenMAPP & "', '" & rs!GenBank & "')"
'      rs.MoveNext
'   Loop
End Sub
