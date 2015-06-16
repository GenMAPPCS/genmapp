Attribute VB_Name = "MiscFuncs"
Option Explicit
Public Const PI = 3.14159265358979
Public Const vbGray = &HE0E0E0

Public i As Integer, j As Integer, s As String, b As Boolean               'Miscellaneous variables
Public d As Double, l As Long, v As Variant

Function LineIntersection(aX1 As Single, aY1 As Single, aX2 As Single, aY2 As Single, bX1 As Single, bY1 As Single, bX2 As Single, bY2 As Single, X As Single, Y As Single) As Boolean
   Dim aX As Single, aY As Single, a As Single, bX As Single, bY As Single, b As Single, denom As Single
   
   aX = aX1 - aX2
   aY = aY2 - aY1
   a = aX2 * aY1 - aX1 * aY2
   bX = bX1 - bX2
   bY = bY2 - bY1
   b = bX2 * bY1 - bX1 * bY2
   denom = aY * bX - bY * aX
   If denom = 0 Then
      LineIntersection = False
   Else
      X = (aX * b - bX * a) / denom
      Y = (bY * a - aY * b) / denom
      LineIntersection = True
   End If
End Function

'*************************************************************************** Command Line Arguments
Function CommandLineArg(commandLine As String, key As String) As String
   '  12/27/04
   '  Entry    commandLine
   '           key         Character string in command line that identifies argument.
   '                       key:  Argument follows. Eg: ...xxx" "key: argument" "xxx...
   '                       .key  Argument is path to file, eg: ...xxx" "C:\file.key" "xxx..."
   '  Return   The argument or empty string if agrument doesn't exist.
   '           commandLine Passed command line with argument removed.
   Dim argument As String, quotes As Boolean
   Dim begArg As Integer, endArg As Integer
   
   If InStr(commandLine, key) = 0 Then '++++++++++++++++++++++++++++++++++++ Argument Doesn't Exist
      CommandLineArg = ""
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   commandLine = Trim(commandLine)
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++ Find Beginning And End And Extract Argument
   If InStr(key, ".") Then '=======================================================Argument Is Path
      endArg = InStr(commandLine, key) + Len(key) - 1                    'Before the quote or space
      If Mid(commandLine, endArg + 1, 1) = """" Then
         quotes = True
      End If
      If quotes Then '--------------------------------------------------------------Quoted Argument
         begArg = InStrRev(commandLine, """", endArg) + 1                          'After the quote
      Else '----------------------------------------------------------------------Unquoted Argument
         '  Must be space delimited
         begArg = InStrRev(commandLine, " ", endArg) + 1                           'After the space
            '  This could be the beginning of the line if InStrRev = 0
      End If
      CommandLineArg = Mid(commandLine, begArg, endArg - begArg + 1)
   Else '====================================================================Argument Is Key: Value
      begArg = InStr(commandLine, key)                                    'After the quote or space
      If begArg > 1 Then '---------------------------------------------------------Check for quotes
         If Mid(commandLine, begArg - 1, 1) = """" Then
            quotes = True
         End If
      End If
      If quotes Then '--------------------------------------------------------------Quoted Argument
         endArg = InStr(begArg, commandLine, """") - 1                            'Before the quote
      Else '----------------------------------------------------------------------Unquoted Argument
         '  Must be space delimited
         endArg = InStr(begArg + Len(key) + 1, commandLine, " ") - 1              'Before the space
         If endArg = -1 Then '________________________________________________Last Argument On Line
            endArg = Len(commandLine)
         End If
      End If
      CommandLineArg = Trim(Mid(commandLine, begArg + Len(key), endArg - begArg - Len(key) + 1))
   End If
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Remove Argument From Command Line
   If quotes Then '=================================================================Quoted Argument
      If begArg = 2 Then '---------------------------------------------------First Argument In Line
         commandLine = Trim(Mid(commandLine, endArg + 2))
      Else '-------------------------------------------------------------Argument In Middle Of Line
         commandLine = Trim(Left(commandLine, begArg - 2) & Mid(commandLine, endArg + 2))
            '  Trim takes care of trailing space if last argument
         If Mid(commandLine, begArg - 2, 2) = "  " Then            'Clean up double space if exists
            commandLine = Left(commandLine, begArg - 2) & Mid(commandLine, begArg)
         End If
      End If
   Else '=========================================================================Unquoted Argument
      If begArg = 1 Then '---------------------------------------------------First Argument In Line
         commandLine = Trim(Mid(commandLine, endArg + 2))           'Beyond space or at end of line
      Else '-------------------------------------------------------------Argument In Middle Of Line
         commandLine = Trim(Left(commandLine, begArg - 1) & Mid(commandLine, endArg + 1))
            '  Starts after the beginning space delimiter and leaves ending space delimiter
            '  in line if not the last argument.
            '  Trim takes care of trailing space if last argument
      End If
   End If
End Function

'******************************************************************** See If File Can Be Written To
Function FileWritable(file As String, Optional existingFile As Boolean = False) As String
   '  11/11/04
   '  Entry    file           File to test.
   '           existingFile   True if file must exist, ie. it is to be rewritten.
   '  Return   Empty string if file can be written to, error message if not.
   
On Error GoTo NotWritable
   Select Case DriveCheck(file) '++++++++++++++++++++++++++++++++++++++++++++++++ Drive Writability
   Case "MISSING"
      FileWritable = "The file" & vbCrLf & vbCrLf & file & vbCrLf & vbCrLf _
                   & "Cannot be written to because the drive does not exist or has no media."
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   Case "RO"
      FileWritable = "The file" & vbCrLf & vbCrLf & file & vbCrLf & vbCrLf _
                   & "Cannot be written to because it is on a non-writable drive such as a CD-ROM."
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End Select
   If Dir(GetFolder(file) & "WriteTest.$tm") <> "" Then '+++++++++++++++++++++ Get Rid Of Test File
      Kill GetFolder(file) & "WriteTest.$tm"
   End If
   If Dir(file) = "" Then '++++++++++++++++++++++++++++++++++++++++++++++++++++ File does Not Exist
      If existingFile Then '========================================File Must Exist And Be Writable
         FileWritable = "The file" & vbCrLf & vbCrLf & file & vbCrLf & vbCrLf _
                & "does not exist."
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      Else '======================================================================Try Creating File
         Open file For Output As #FILE_TEMP '---------------------Will Trigger Error If Unsucessful
         Close #FILE_TEMP
         Kill file
      End If
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ File Exists
      Open file For Append As #FILE_TEMP '------------------------------------Try Opening For Write
      Close #FILE_TEMP
   End If
'   Name file As GetFolder(file) & "WriteTest.$tm"
'   Name GetFolder(file) & "WriteTest.$tm" As file
   FileWritable = ""
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

NotWritable:
   FileWritable = "The file" & vbCrLf & vbCrLf & file & vbCrLf & vbCrLf _
                & "Cannot be written to. It may be open somewhere else, set to read only, be " _
                & "on a non-writable drive such as a CD."
End Function
Function ValidPathName(path As String) As Boolean
   If InChrs(path, "/*?""<>|") Then Exit Function             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   If InStr(path, "\\") Then Exit Function                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ValidPathName = True
End Function
'******************************************************** Produces File Of All Files In Folder Tree
Sub BuildFileTree(root As String, pattern As String, Optional subfolders As Boolean = True, _
                  Optional level As Integer = 0, Optional treeFile As String = "")
   '  Entry    root        Folder at root of tree, Eg: "C:\Folder\Sub\"
   '           pattern     File pattern to look for.  Eg: "*.mapp"
   '           subfolders  If true, include subfolders
   '           level       Depth in folder tree.
   '           treeFile    File in which tree is stored. Only needed for first call.
   Dim file As String, fileIndex As Integer, dirIndex As Integer
   
   If treeFile = "" Then treeFile = App.path & "\TreeFile.$tm"
   
   If level = 0 Then
      Open treeFile For Output As #FILE_TREE
   End If
   
   file = Dir(root & pattern, vbReadOnly)                                  'Include read-only files
   Do Until file = ""
      fileIndex = fileIndex + 1                    'Keep track of where we are in current directory
'      If Right(sourceFile, Len(ext) - 2) <> Right(ext, Len(ext) - 2) Then GoTo NextFile
         '  Unfortunately, the Dir function returns files that simply begin with the extension.
         '  For example, the extension "*.gex" will also return the file "whatever.gexz".
         '  Real dumb!
      Print #FILE_TREE, root & file
      file = Dir
   Loop
   
   If subfolders Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++ Find Directories Next
   '  Find the next directory and recursively call BuildFileTree to find files and subdirectories
      file = Dir(root & "\", vbDirectory)
      Do Until file = ""
         dirIndex = dirIndex + 1                   'Keep track of where we are in current directory
         If file <> "." And file <> ".." Then
            If (GetAttr(root & file) And vbDirectory) = vbDirectory Then
               BuildFileTree root & file & "\", pattern, subfolders, level + 1
               file = Dir(root, vbDirectory)        'Return to dir entry where we left off
               For i = 1 To dirIndex - 1            'because calling Dir again in BuildFileTree
                  file = Dir                        'will lose our place ("." and ".." are always
                                                    'first 2 directory entries)
               Next i
            End If
         End If
         file = Dir
      Loop
   End If
   If level = 0 Then Close #FILE_TREE
End Sub

Function FileAbbrev(file As String, Optional chrs As Integer = 60)
   '  5/28/03
   Dim abbrev As String, slash As Integer, shortFile As String
   
   If Len(file) > chrs Then
      slash = InStr(file, "\")
      shortFile = Left(file, slash) & "..."
      shortFile = shortFile & Mid(file, Len(file) - chrs + 1 + Len(shortFile))
      FileAbbrev = shortFile
'      slash = InStrRev(file, "\")
'      If Len(file) - slash < chrs - 3 Then
'         FileAbbrev = Left(file, chrs - 3 - Len(abbrev)) & "..." & Mid(file, slash)
'      Else
'         FileAbbrev = "..." & Right(file, chrs - 3)
'      End If
   Else
      FileAbbrev = file
   End If
End Function

'************************************************************************** Checks For Invalid Chrs
Function InvalidChr(str As Variant, place As String, Optional chrs As String = """$") As Boolean
   '  Entry:   str      String to check
   '           place    Where string is used (also verbiage for error message)
   '           chrs     The invalid characters, defaults to " and $
   '  Return:  True     if invalid chr found
   '           place    Invalid characters for Gene IDs, each followed by a space
   '     For Expression Datasets:
   '           str      Corrected heading for column
   Dim message As String
   
   Select Case place
   Case "criterion label"
      For i = 1 To Len(str)
         Select Case Mid(str, i, 1)
         Case Chr(34), "$", "|"
            message = message & Mid(str, i, 1) & " "
         End Select
      Next i
   Case "gene ID", "char data"
      '  Returns invalid characters in place variable.
      '  Never uses message so never prints error and InvalidChr must be set to True here.
      place = ""
      For i = 1 To Len(str)
         Select Case Mid(str, i, 1)
         Case Chr(34), "$", ",", "'"
            place = place & Mid(str, i, 1) & " "
         End Select
      Next i
      If place <> "" Then InvalidChr = True
   Case "column heading"
      '  This case asks for input of the correction and loops until it is clean.
      '  New title returned in str.
      '  Returns True if any invalid chrs were found
      Do
         message = ""
         For i = 1 To Len(str)
            Select Case Mid(str, i, 1)
            Case Chr(34), "`", "!", "[", "]", ".", ",", "$", "|"
               message = message & Mid(str, i, 1) & " "
            End Select
         Next i
         If message <> "" Then
            str = InputBox("Invalid character(s) " & message & "found in " & place & ".", _
                           "Invalid Column Heading", str)
            InvalidChr = True
         End If
      Loop While message <> ""
   Case "Gene Table Name"
      For i = 1 To Len(str)
         Select Case Mid(str, i, 1)
         Case "A" To "Z", "a" To "z", "0" To "9", "_"
         Case Else
            message = message & Mid(str, i, 1) & " "
         End Select
      Next i
   Case "remarks"
      '  Allow anything so that HTML is accepted
   Case Else
      '  "gene label", "gene identification", "backpage heading", "remarks", "color set name", etc
      For i = 1 To Len(str)
         If InStr(chrs, Mid(str, i, 1)) Then
            message = message & Mid(str, i, 1) & " "
         End If
      Next i
   End Select
   If message <> "" Then
      MsgBox "Invalid character(s) " & message & "found in " & place & ".", _
             vbOKOnly + vbExclamation, "Invalid Characters"
      InvalidChr = True
   End If
End Function

'**************************************************************** Removes Quotes Next To Delimiters
Function RemoveQuotes(ByVal str As String, Optional delimiter As String = "") As String
   '8/9/02
   '  Entry:   str         Delimited string, possibly with quotes around individual elements
   '           delimiter   If blank, search for tab, else default to comma
   '  Return:  String with quotes next to delimiters removed
   Dim delimPos As Integer
   
   If InStr(str, """") = 0 Then                                      'There are no quotes to remove
      RemoveQuotes = str
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If delimiter = "" Then
      If InStr(str, vbTab) Then
         delimiter = vbTab
      Else
         delimiter = ","
      End If
   End If
   
   If Left(str, 1) = """" Then str = Mid(str, 2)                                 'Starts with quote
   delimPos = InStr(1, str, delimiter)
   Do While delimPos '+++++++++++++++++++++++++++++++++++++++++ Check Either Side Of Each Delimiter
      If Mid(str, delimPos + 1, 1) = """" Then     'Beyond delimiter first so delimPos doesn't move
         str = Left(str, delimPos) & Mid(str, delimPos + 2)
      End If
      If delimPos > 1 Then                                              'In case of missing gene ID
         If Mid(str, delimPos - 1, 1) = """" Then                                 'Before delimiter
            str = Left(str, delimPos - 2) & Mid(str, delimPos)
            delimPos = delimPos - 1
         End If
      End If
      delimPos = InStr(delimPos + 1, str, delimiter)
   Loop
   If Right(str, 1) = """" Then str = Left(str, Len(str) - 1)                      'Ends with quote
   RemoveQuotes = str
End Function

Function EmbedLinks(ByVal str As String) As String '****************** Find And Embed Links In Text
   '  Entry:   str   A text string with possible internet links in it
   '  Return   The string with the links expanded into HTML
   '  Finds both mail- and http-looking links in text and embeds the HTML links
   '  HTTP and FTP links must not have @ in them
   Dim at As Integer, begLink As Integer, endLink As Integer, dot As Integer
   Dim link As String, protocol As String
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Mail Links
   at = InStr(str, "@")
   Do While at
      '===========================================================================Check End Of Link
      endLink = InvalidLinkChr(Mid(str, at + 1), 1, "MAIL") + at - 1
      If endLink = at - 1 Then                                    'No invalid chrs to end of string
         endLink = Len(str)
      End If
      If Mid(str, endLink, 1) = "." Then endLink = endLink - 1                 'Won't end in period
      If Mid(str, endLink, 1) = "_" Then                                   'Can't end in underscore
         GoTo InvalidMail                                  'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
      If InStr(Mid(str, at + 1, endLink - at - 1), ".") = 0 Then   'Must have dot somewhere after @
         GoTo InvalidMail                                  'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
      
      '=====================================================================Check Beginning Of Link
      begLink = InvalidLinkChr(Left(str, at - 1), -1, "MAIL") + 1
      If begLink = at Then                                                      'Can't start with @
         GoTo InvalidMail                                  'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
      If InStr(Mid(str, begLink, at - begLink), ".") Then                     'Can't have dot in it
         GoTo InvalidMail                                  'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
      
      '==================================================================================Embed Link
      link = "<a href=""mailto:" & Mid(str, begLink, endLink - begLink + 1) & """>" _
           & Mid(str, begLink, endLink - begLink + 1) & "</a>"
      str = Left(str, begLink - 1) & link & Mid(str, endLink + 1)
      endLink = begLink + Len(link) - 1

InvalidMail:
      at = InStr(endLink + 1, str, "@")
'Debug.Print str
   Loop
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ HTTP Links
   dot = InStr(str, ".")
   Do While dot
      '===========================================================================Check End Of Link
      If dot = Len(str) Then
         endLink = Len(str)
         GoTo InvalidLink
      End If
      endLink = InvalidLinkChr(Mid(str, dot + 1)) + dot - 1
      If endLink = dot Then                                'Invalid character immediately after dot
         GoTo InvalidLink                                  'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
      If endLink = dot - 1 Then                                   'No invalid chrs to end of string
         endLink = Len(str)
      End If
      If Mid(str, endLink, 1) = "." Then endLink = endLink - 1                 'Won't end in period
      If Mid(str, endLink, 1) = "_" Then                                   'Can't end in underscore
         GoTo InvalidLink                                  'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
      
      '=====================================================================Check Beginning Of Link
      begLink = InvalidLinkChr(Left(str, dot - 1), -1) + 1
      If begLink = dot Then                                                     'Can't start with .
         GoTo InvalidLink                                  'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End If
      If begLink > 1 Then
         If Mid(str, begLink - 1, 1) = "@" Then                                 'Can't have @ in it
            GoTo InvalidLink                               'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
         End If
      End If
      protocol = "http://"
      If begLink >= 8 Then
         If UCase(Mid(str, begLink - 7, 7)) = "HTTP://" Then
            begLink = begLink - 7
            protocol = ""                                                'Protocol part of the link
         End If
      ElseIf begLink >= 7 Then
         If UCase(Mid(str, begLink - 6, 6)) = "FTP://" Then
            begLink = begLink - 6
            protocol = ""                                                'Protocol part of the link
         End If
      End If
      
      '==================================================================================Embed Link
      link = "<a href=""" & protocol & Mid(str, begLink, endLink - begLink + 1) & """>" _
           & Mid(str, begLink, endLink - begLink + 1) & "</a>"
      str = Left(str, begLink - 1) & link & Mid(str, endLink + 1)
      endLink = begLink + Len(link) - 1

InvalidLink:
      dot = InStr(endLink + 2, str, ".")                             'Not sure why +2, but it works
'Debug.Print str
   Loop
   
   EmbedLinks = str
End Function
Function InvalidLinkChr(chrs As String, Optional direction As Integer = 1, Optional linkType As String = "HTTP") As Integer
   Dim i As Integer, beg As Integer, fin As Integer
   
   If direction = 1 Then
      beg = 1
      fin = Len(chrs)
   Else
      beg = Len(chrs)
      fin = 1
   End If
   For i = beg To fin Step direction
      Select Case Mid(chrs, i, 1)
      Case "A" To "Z", "a" To "z", "_", ".", "?", "=", "+", "&", "-", "0" To "9"
      Case "/"
         If linkType = "MAIL" Then
            InvalidLinkChr = i
            Exit Function                                  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
         If direction = -1 Then                         'HTTP link can't have / anywhere before dot
            InvalidLinkChr = i
            Exit Function                                  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      Case Else
         InvalidLinkChr = i
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End Select
   Next i
   InvalidLinkChr = 0
End Function

Function GetDrive(path As String) As String '********************************** Get Drive From Path
   ' 6/17/03
   '  Entry
   '     path     Full path, eg: C:\Whatever\file.abc
   '  Return
   '     Drive eg: C:
   Dim drive As String
   
   drive = Left(path, InStr(path, ":"))
   If drive = "" Then drive = "C:"
   GetDrive = drive
End Function
Function GetFolder(path As String) As String '******************************** Get Folder From Path
   ' 6/17/03
   '  Entry
   '     path     Full path, eg: C:\Whatever\file.abc
   '  Return
   '     Folder with trailing \, eg: C:\Whatever\
   Dim folder As String
   
   folder = Left(path, InStrRev(path, "\"))
   If folder = "" Then folder = "C:\"
   GetFolder = folder
End Function
Function GetFile(path As String) As String '************************************ Get File From Path
   ' 7/16/02
   '  Entry
   '     path     Full path, eg: C:\Whatever\file.abc
   '  Return
   '     File name including extension or nothing. Eg: file.abc
   
   GetFile = Mid(path, InStrRev(path, "\") + 1)
End Function
Function ReverseSlashes(ByVal s As String) As String '********* Backslashes To Slashes & Vice Versa
   '7/4/02
   '  If it finds a backslash anywhere in the string, it changes backslashes to forward slashes.
   '  Presumably, a legal HTML or UNIX string will not have any backslashes, so if it finds no
   '  backslashes, it then looks for forward slashes to change to backslashes.
   Dim slashChr As String                                                 '\ or /, look for \ first
   Dim replaceChr As String, i As Integer, slash As Integer
   Dim newString As String
   
   If InStr(s, "\") Then
      slashChr = "\"
      replaceChr = "/"
   Else
      slashChr = "/"
      replaceChr = "\"
   End If
   
   newString = s
   slash = InStr(newString, slashChr)
   Do While slash
      Mid(newString, slash, 1) = replaceChr
      slash = InStr(slash + 1, newString, slashChr)
   Loop
   ReverseSlashes = newString
End Function
Function HtmlHexColor(number As Long) As String '************** Converts VB Color Into Hex For HTML
   Dim html As String
   
   html = Right("000000" & Hex(number), 6)                                            'Form: 0D12C5
   HtmlHexColor = "#" & Right(html, 2) & Mid(html, 3, 2) & Left(html, 2)
                                                                       'Reverse RGB for web: C5120D
End Function
Function Div(Num As Variant, Den As Variant) '******************************* Allows Divide By Zero
   '5/16/02
   If Den = 0 Then Div = 0 Else Div = Num / Den
End Function
Function ElapsedTime(Optional time As Variant = -1, Optional returnOnly As Boolean) As Double
   '4/24/02
   '  Enter:   time        Ending time. If missing it is Now, ending the elapsed time.
   '                       Zero resets time and does not display message box.
   '           returnOnly  If true, does not print the message box with the elapsed time.
   '  Return:  Elapsed time in approximate seconds.
   Static beginTime As Double
   Dim elapsed As Double
   
   If time = 0 Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Reset Clock
      beginTime = Now
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Figure Elapsed Time
      If time = -1 Then
         time = CDbl(Now)
      ElseIf VarType(time) <> vbDouble Then
         time = CDbl(time)
      End If
      elapsed = (time - beginTime) * 86400                                     'Approximate seconds
      If Not returnOnly Then
         MsgBox "Elapsed time approximately " & Format(elapsed, "0.0##") & " seconds.", _
                vbInformation + vbOKOnly, "Elapsed Time"
      End If
      ElapsedTime = elapsed
   End If
End Function

'************************************************************************ Inputs LF-Delimited Lines
Function InputUnixLine(fileNo As Integer, Optional bytes As Long) As String
   '3/12/02
   '  Enter:   FileNo   Number of a file open for Binary
   '           bytes    If -1, resets byte count, returns empty line
   '  Return:  Line from file without line-ending characters or "**eof***"
   '           bytes    Current byte count
   
   Dim lin As String, char As Byte
   
   If bytes = -1 Then
      bytes = 0
      InputUnixLine = ""
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   lin = ""
   Get #fileNo, , char                                                'EOF not true until Get fails
   If EOF(fileNo) Then
      InputUnixLine = "**eof**"
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   Do Until char = 10 Or EOF(fileNo)                        'UNIX convention. Lines end in ASCII 10
      If char <> 13 Then                                     'Ignore the CR of a CR-LF if it occurs
         lin = lin & Chr(char)
      End If
      Get #fileNo, , char
   Loop
   bytes = Seek(fileNo)
   InputUnixLine = lin
End Function
'************************************************************* Separate Values In Delimited Strings
Function SeparateValues(values() As String, lin As String, Optional delimiter As String = ",") _
         As Integer
   '  5/22/03
   '  Entry    lin         String to separate
   '           delimiter   Character(s) separating individual values. Defaults to ","
   '  Return   Number of values found (does not exceed variables in values array)
   '           values()    Individual values, quotes trimmed. Zero-based
   Dim maxValues As Integer, delim As Integer, prevDelim As Integer, delimLen As Integer
   Dim totalValues As Integer       'Number of values actually found, one-based. Return of function
   
   maxValues = UBound(values)                                   'Don't need totalValues beyond this
   
'   ReDim values(maxValues)
   For i = 0 To maxValues
      values(i) = ""
   Next i
   If Len(lin) = 0 Or maxValues = 0 Then
      SeparateValues = 0
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   totalValues = 0                                                                       'One based
   delimLen = Len(delimiter)
   If Left(lin, delimLen) = delimiter Then                                   'Starts with delimiter
      prevDelim = delimLen + 1          'Beyond end of previous delimiter (beginning of next value)
   Else
      prevDelim = 1                     'Beyond end of previous delimiter (beginning of next value)
   End If
   delim = InStr(prevDelim, lin, delimiter)
   If delim = 0 Then delim = Len(lin) + 1
   Do While prevDelim < Len(lin) And totalValues <= maxValues
      totalValues = totalValues + 1
      values(totalValues - 1) = Mid(lin, prevDelim, delim - prevDelim)                  'Zero-based
      '---------------------------------------------------------------------------------Trim Quotes
         If Left(values(totalValues - 1), 1) = """" Then
            values(totalValues - 1) = Mid(values(totalValues - 1), 2)
         End If
         If Right(values(totalValues - 1), 1) = """" Then
            values(totalValues - 1) = Left(values(totalValues - 1), Len(values(totalValues - 1)) - 1)
         End If
      prevDelim = delim + delimLen
      delim = InStr(prevDelim, lin, delimiter)
      If delim = 0 Then delim = Len(lin) + 1
   Loop
   SeparateValues = totalValues
End Function
'************************************************************** Pipe- To Delimiter-Separated String
Function SeparatePipes(str As String, Optional delimiter As String = ", ")
   '  8/31/02
   '  Enter:   str         string with pipe-separated values. Eg: |P05534|P30448|P30449|Q95355|
   '           delimiter   Characters to separate return string with. Eg: vbCrLf
   '                       If not given, defaults to comma and space.
   '  Return:  Delimiter-separated. Eg: P05534, P30448, 30449, Q95355
   Dim pipe As Integer, nextPipe As Integer, strOut As String
   
   If Left(str, 1) = "|" Then
      pipe = 1
   Else
      pipe = 0
   End If
   Do While pipe < Len(str)
      nextPipe = InStr(pipe + 1, str, "|")
      If nextPipe = 0 Then nextPipe = Len(str) + 1
      If strOut <> "" Then
         strOut = strOut & delimiter & Mid(str, pipe + 1, nextPipe - pipe - 1)
      Else
         strOut = strOut & Mid(str, pipe + 1, nextPipe - pipe - 1)
      End If
      pipe = nextPipe
   Loop
   SeparatePipes = strOut
End Function
Sub DeleteIndexes(db As Database, table As String) '*************** Deletes All Indexes For A Table
   '  2/6/02
   '  Enter:   db    An open database
   '           table Table to delete indexes from
   
   i = db.TableDefs(table).Indexes.count
   For j = 0 To i - 1
                                             'This statement is ugly. There's gotta be a better way
      db.TableDefs(table).Indexes.Delete db.TableDefs(table).Indexes(0).name
   Next j
End Sub
Sub Delay(Seconds) '**************************************************** Delays process for Seconds
   '  8/16/01
   Dim beginTime As Date
   
   beginTime = Now
   DoEvents
   Do Until DateDiff("s", beginTime, Now) > Seconds
      DoEvents
   Loop
End Sub
Function DriveCheck(ByVal drive As String) As String '********************* Check Validity Of Drive
   '  7/15/03
   '  Entry    drive    Path string. Process looks at only first character.
   '  Return   Blank if successful, "MISSING" if it doesn't exist, "RO" if read-only
   '  Process  Requires GetDriveType() API
   
On Error GoTo ErrorHandler
   drive = Left(drive, 1) & ":\"
   If GetDriveType(drive) <= 1 Then '++++++++++++++++++++++++++++++++++++++++++ See If Drive Exists
      DriveCheck = "MISSING"
   Else '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Try Writing
      If Dir(drive & "DriveCheck.$tm") <> "" Then
         Kill drive & "DriveCheck.$tm"
      End If
      Open drive & "DriveCheck.$tm" For Output As #99
      Close #99
      Kill drive & "DriveCheck.$tm"
   End If
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
ErrorHandler:
   Select Case Err.number
   Case 76, 52
      DriveCheck = "MISSING"
   Case 75
      DriveCheck = "RO"
   End Select
End Function

'************************************************************************** Checks Validity Of Path
Function PathCheck(ByVal path As String, Optional readOnly As Boolean = False) As String
   '  7/15/03
   '  Entry    Path     Path not including file name to check for validity
   '                    Eg: C:\GenMAPP\Data\
   '                    If path contains a file name, file name is separated and also tested for writing
   '           readOnly If true, only need to read so don't check for writing
   '  Return   Blank if successful, "MISSING" if it doesn't exist, "RO" if read-only and check.
   '  Process  Attempt to write to test file in path if readOnly is false.
   '
   Dim slash As Integer, colon As Integer, testFile As String, s As String, file As String
   
   Select Case DriveCheck(path)
   Case "MISSING"
      MsgBox "Drive doesn't exist or has no media.", vbExclamation + vbOKOnly, "Checking Path"
      PathCheck = "MISSING"
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   Case "RO"
      If Not readOnly Then
         MsgBox "Cannot write to drive.", vbExclamation + vbOKOnly, "Checking Path"
         PathCheck = "RO"
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Case Else
   End Select
   
   If Dir(path, vbDirectory) = "" Then
      MsgBox "Path not found.", vbExclamation + vbOKOnly, "Checking Path"
      PathCheck = "MISSING"
      Exit Function
   ElseIf Not readOnly Then
On Error GoTo ErrorHandler
      testFile = path & "\drvtst.$tm"
      If Dir(testFile) <> "" Then Kill testFile
      Open testFile For Output As #99
      Close 99
      Kill testFile
   End If
   PathCheck = ""
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ErrorHandler:
   Select Case Err.number
   Case 76
      MsgBox Err.Description
      PathCheck = "MISSING"
   Case Else
      MsgBox "Cannot write to this path.", vbExclamation + vbOKOnly, "Checking Path"
      PathCheck = "RO"
   End Select
End Function

Sub FatalError(where As String, what As String, Optional logFile As String = "")
   '  6/19/02
   '  Enter    where    Spot in the program that calls this procedure
   '           what     Err.Description (or other indication)
   '  Sample call:
   '     FatalError "frmExpression:OpenExpressionDataset", Err.Description, "ConvertError.log"

   If logFile = "" Then
      logFile = App.path & "\" & App.EXEName & "Error.log"
   End If
   If InStr(logFile, ".") = 0 Then
      logFile = logFile & ".log"
   End If
   If InStr(logFile, "\") = 0 Then
      logFile = App.path & "\" & logFile
   End If
   MsgBox "Unexpected error in " & where & "." & vbCrLf & vbCrLf & what & vbCrLf & vbCrLf _
          & "Whoops! We hoped you would never have to see one of these but obviously, we " _
          & "screwed up. Please help us by trying to remember exactly how you got here and " _
          & "report this to genmapp@gladstone.ucsf.edu. Attach the file " & vbCrLf & vbCrLf _
          & logFile & vbCrLf & vbCrLf _
          & "and, if you can, any other files you used " _
          & "in this execution. We and the GenMAPP users appreciate your cooperation.", _
          vbCritical + vbOKOnly, "FATAL ERROR"
   Open logFile For Output As #40
   Print #40, "Build: "; BUILD
   Print #40, where
   Print #40, what
   Close #40
   End
End Sub
Function DisplayAmpersands(ByVal s As String) As String '************************** Makes & Into &&
   '  2/2/01
   Dim amp As Integer
   
   amp = InStr(s, "&")
   Do While amp
      s = Left(s, amp) & "&" & Mid(s, amp + 1)
      amp = InStr(amp + 2, s, "&")
   Loop
   DisplayAmpersands = s
End Function

Sub FontSizeFloor(spec As Single, obj As Object) '*************************** Next Lowest Font Size
   '  10/24/03
   '  True Type and most other PC fonts have discrete point sizes such as 9.75 or 8.25.
   '  This procedure sets font size for the obj that is less than or equal to the
   '  font-size specification.
   Dim Size As Integer, initialSize As Single, newSize As Single
Dim temp As Single
temp = spec
   
'   obj.Font.Size = 1.25
'   obj.Font.Size = spec
'   Do While obj.Font.Size > spec
'      spec = spec - 0.1
'      obj.Font.Size = spec
'   Loop

   If spec < 1.5 Then
      obj.Font.Size = 1.5
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   obj.Font.Size = spec
   Do While obj.Font.Size > spec
      spec = spec - 0.1
      obj.Font.Size = spec
   Loop
   
'
'   Size = spec * 10
'frmAbout.fontSize = ((Size \ 6) * 6) / 10
'frmAbout.fontSize = 8
'
'   FontSizeFloor = ((Size \ 6) * 6) / 10
'   FontSizeFloor = outputObj.Font.Size
End Sub
Function AddFolder(path As String) As String '****************************** Adds Folder To Storage
   '  11/20/00
   '  Entry:
   '     path  Path to be added to directory structure. May or may not end in \
   '  Return:
   '     Part of path that already exists. To be used to remove added path later if
   '        not needed
   '  For example: Path to be added is
   '        C:\Large\Medium\Small
   '     If C:\Large already existed, C:\Large is returned
   Dim root As String                                            'Part of path that already existed
   Dim partialPath As String, drive As String
   Dim slash As Integer, nextSlash As Integer
         
On Error GoTo ErrorHandler
   path = Dat(path)
   If InStr(path, ":") = 0 Then                                                           'No drive
      If Left(path, 1) = "\" Then                                                'Add current drive
         path = Left(CurDir, InStr(CurDir, ":")) & path
      Else
         path = Left(CurDir, InStr(CurDir, ":")) & "\" & path
      End If
   End If
   slash = InStr(path, "\")
   Do While slash < Len(path)
      nextSlash = InStr(slash + 1, path, "\")
      If nextSlash = 0 Then nextSlash = Len(path) + 1
      partialPath = Left(path, nextSlash - 1)
      If Dir(partialPath, vbDirectory) = "" Then
         MkDir partialPath
         If root = "" Then
            root = Left(path, slash - 1)
         End If
      End If
      slash = nextSlash
   Loop
ExitFunction:
   AddFolder = root
   Exit Function                                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

ErrorHandler:
   MsgBox Err.Description & ". Folder not created", vbCritical + vbOKOnly, "Creating folder"
   root = "<ERROR>"
   Resume ExitFunction
End Function
Function TextToHtml(txt As String) As String '************************** Makes Text HTML Compatible
   Dim index As Integer
   Dim html As String, char As String * 1
   
   For index = 1 To Len(txt)
      char = Mid(txt, index, 1)
      Select Case char
      Case "&"
         html = html & "&amp;"
      Case "<"
         html = html & "&lt;"
      Case ">"
         html = html & "&gt;"
      Case Else
         html = html & char
      End Select
   Next index
   TextToHtml = html
End Function

Function TextToFileName(ByVal name As String) As String
   '  10/31/00
   '  Returns name with all characters that are illegal for a file name changed to "_"
   Dim i As Integer
   
   For i = 1 To Len(name)
      Select Case Mid(name, i, 1)
      Case "\", "/", ":", "*", "?", "<", ">", "|"
         Mid(name, i, 1) = "_"
      End Select
   Next i
   TextToFileName = name
End Function

Function CheckForString(v As Variant) As Boolean '************************ True If Value NonNumeric
   Dim i As Integer, char As String * 1, value As String
   
   value = Dat(v)
   For i = 1 To Len(value)
      char = Mid(value, i, 1)
      If (char < "0" Or char > "9") And char <> "." Then                      'Character nonnumeric
         If i = 1 Or i = Len(value) Then                                 'In first or last position
            If char <> "+" And char <> "-" Then                                        'Allow signs
               CheckForString = True
               Exit Function                          '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            End If
         Else
            CheckForString = True
            Exit Function                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         End If
      End If
   Next i
End Function

'*************************************************************** Displays Value In chars Characters
Function ValueDisplay(number As Variant, chars As Integer, Optional Form As String = "Limit") _
         As String
   '  11/1/00
   '  Returns nulls as empty strings
   '  Returns first chars characters of strings
   '  Returns number as a string of chars characters including decimal point and minus sign
   '     For now, minus sign not counted as one of the chars
   '  form:
   '     Limit    Return is maximum of chars characters
   '        i.e. as 5-char output, 5 returns 5
   '     Round    Return is chars characters
   '        i.e. as 5-char output, 5 returns 5.000
   '  Doesn't print decimal point if not needed
   '  Exceeds chars characters if needed to express numeric value
   '     i.e. 4567 as a 3-character output returns 4567
   '  Doesn't display trailing zeros
   
   Dim exponent As Integer, formatString As String, Point As String
   Dim whole As Integer            'Number of chars to left of decimal point (including minus sign)
   Dim fraction As Integer                               'Number of chars to right of decimal point
   Dim Display As String, endDisplay As Integer
   
   If VarType(number) = vbNull Or VarType(number) = vbEmpty Then
      ValueDisplay = ""
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If VarType(number) = vbString Then
      ValueDisplay = Left(number, chars)
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   whole = 1                                             'Must have at least one, 0 if nothing else
   Do While Abs(number) >= 10 ^ whole
      whole = whole + 1
   Loop
   If number < 0 Then whole = whole + 1                                   'Minus sign adds one char
   If (number > 0 And (number >= 1000000 Or number < 0.0001)) Or (number < 0 And (number <= -100000 Or number > -0.001)) Then
                         '------------------------------------------------------Shift to E-Notation
      Display = Format(number, "#.#e-#")
   Else '-------------------------------------------------------------------------Standard Notation
      fraction = chars - whole - 1                                   'Left over for fractional part
      If whole >= chars - 1 Then              'No room for fraction after considering decimal point
         Point = ""
         fraction = 0
      Else
         Point = "."
      End If
      If fraction < 0 Then fraction = 0
      Display = Format(number, "0" & Point & String(fraction, "0"))
'      display = Format(number, String(whole, "0") & Point & String(fraction, "0"))
      If Form = "Limit" Then '................................Dump trailing zeros and decimal point
         endDisplay = Len(Display)
         If InStr(Display, ".") Then
            Do While Mid(Display, endDisplay, 1) = "0"
               endDisplay = endDisplay - 1
            Loop
         End If
         If Mid(Display, endDisplay, 1) = "." Then endDisplay = endDisplay - 1
         Display = Left(Display, endDisplay)
      End If
      
   End If
   ValueDisplay = Display
End Function

Function InCircle(angle As Single) '*********************** Returns Angle Between 0 and 359 Degrees
   '  8/16/00
   '  requires constant PI
   '  Angle received and returned in radians, 0 to < 2*PI radians (0 to 359.999999 degrees)
   If angle >= 2 * PI Then
      InCircle = angle - 2 * PI
   ElseIf angle < 0 Then
      InCircle = angle + 2 * PI
   Else
      InCircle = angle
   End If
End Function

Function PickColor() As Long '********************************************** Pick Color From Dialog
   '  7/27/00
   '  If cancelled, returns -1, else returns color
   
On Error GoTo CancelError
   With Screen.ActiveForm
      .dlgDialog.CancelError = True
      .dlgDialog.FLAGS = cdlCCRGBInit
      .dlgDialog.ShowColor
      PickColor = .dlgDialog.color
   End With
   Exit Function
   
CancelError:
   PickColor = -1
End Function

' ******************************************************************** Decode HTML Name/Value Pairs
Function NameValue(html As String, propertyName As String, propertyValue As String)
   '  7/5/00
   '  Returns the number of characters used up in name/value pair (not including > or ,)
   '     Position of code that ends name/value pair (> or ,)
   '  Returns next name/value pair in propertyName and propertyValue
   '  Returns empty name if at end of pairs.
   '     Also, NameValue will be 0 (no characters used)
   
   Dim index As Integer, equal As Integer, endPair As Integer
   
   index = 1
   Do While Mid(html, index, 1) = " "                                          'Skip leading spaces
      index = index + 1
   Loop
   endPair = InStr(html, ",")
   If endPair = 0 Then endPair = InStr(html, ">")
   NameValue = endPair - 1
   If endPair = 1 Then
      propertyName = ""
      NameValue = 1
      Exit Function               'At end of name/value pairs >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   equal = InStr(html, "=")
   If equal = 0 Then equal = endPair                                       'No value, only name
   propertyName = UCase(Trim(Mid(html, index, equal - index)))
   propertyValue = UCase(Trim(Mid(html, equal + 1, endPair - equal - 1)))
End Function
Function Dat(ByVal z As Variant) As String
   Rem************************************************************************
   Rem  CONVERTS VARIANT, PARTICULARLY DATABASE FIELD, TO STRING *************
   Rem************************************************************************
   On Error GoTo DatError
   If VarType(z) <> vbNull Then
      Dat = Trim(z)
   Else
      Dat = ""
   
   End If
DatContinue:
   Exit Function

DatError:
   Dat = ""
   Resume DatContinue
End Function
Function ToDB(ByVal z As Variant) As Variant
   '  Returns empty strings as nulls to assign to database fields
   '  Handy for both strings and dates (null strings enter null dates in DB)
   '  Checkboxes should work without modification
   If VarType(z) = vbString Then
      If Trim(z) = "" Then
         ToDB = Null
      Else
         ToDB = z
      End If
   Else
      ToDB = z
   End If
End Function
Function InChrs(S1 As String, S2 As String) As Integer '**************** Any Of 2nd String In First
   '  6/11/02
   '  Returns position of first occurrence of a character in S2 within S1 or zero if not found
   '  Requires Dat() function
   
   Dim index As Integer

   InChrs = 0
   If Dat(S1) = "" Then Exit Function                      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   If Dat(S2) = "" Then Exit Function                      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   For index = 1 To Len(S2)
      If InStr(S1, Mid(S2, index, 1)) Then
         InChrs = InStr(S1, Mid(S2, index, 1))
         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      End If
   Next index
End Function
Function Fmt(value As Variant, Form As String) As String '********************** General Formatting
   '  5/30/00
   '  form that begins with B, blank zeros
   '  form that begins with S, shortens to actual length
   '  form with number multiplies @ or # in form
   '     "!3@" = "!@@@"
   '     'O' used for zero character to multiply
   '        "4O" = "0000"
   
   Dim shorten As Boolean, begMult As Integer, endMult As Integer, length As Integer
   Dim zS As String
   
   On Error GoTo FmtErr

   If UCase(Left(Form, 1)) = "B" Then                                                  'Blank zeros
      If Val(value) = 0 Then
         Fmt = Space(Len(Form) - 1)
         Exit Function
      End If
      Form = Mid(Form, 2, 500)
   End If

   If UCase(Left(Form, 1)) = "S" Then                           'Shorten to length of actual number
      Form = Mid(Form, 2, 500)
      shorten = True
   Else
      shorten = False
   End If

   begMult = InChrs(Form, "123456789")                                         'Look for multiplier
   If begMult Then
      endMult = begMult + 1
      Do While Mid(Form, endMult, 1) >= "0" And Mid(Form, endMult, 1) <= "9"
         endMult = endMult + 1
      Loop
      If Mid(Form, endMult, 1) = "O" Then Mid(Form, endMult, 1) = "0"
      Form = Left(Form, begMult - 1) _
           & String(Val(Mid(Form, begMult, endMult - begMult)), Mid(Form, endMult, 1)) _
           & Mid(Form, endMult + 1)
   End If
   zS = Format(value, Form)
   length = Len(Form)
   If InStr(Form, "!") Then length = length - 1
   If shorten Then
      Fmt = Dat(zS)
   ElseIf length > Len(zS) Then                                          'Don't chop off characters
      Fmt = Right(Space(length) + zS, length)
   Else
      If InStr(Form, "@") Then                                 'VB puts extra chrs at end of format
         Fmt = Left(zS, length)
      Else
         Fmt = zS
      End If
   End If
FmtExit:
   Exit Function

FmtErr:
   Fmt = "<Error" & str$(Err) & ">"
   Resume FmtExit
End Function
Function NullZero(v As Variant) As Variant
   Rem************************************************************************
   Rem  RETURNS ZERO FOR NULL VARIABLES OR EMPTY STRINGS  ********************
   Rem************************************************************************

   If VarType(v) = vbNull Then
      NullZero = 0
   ElseIf VarType(v) = vbString Then
      If Dat(v) = "" Then
         NullZero = 0
      Else
         NullZero = -1
      End If
   Else
      NullZero = v
   End If
End Function
Function NVL(v As Variant, Optional substitute As Variant) As Variant

   If IsMissing(substitute) Then substitute = 0
   
   If VarType(v) = vbNull Then
      NVL = substitute
   Else
      NVL = v
   End If
End Function
Sub RotatePoint(X As Single, Y As Single, rotate As Single) '************************ Rotated Point
   '  6/2/00
   '  Rotate the point X,Y around 0,0 by rotate radians
   '  Result returned in and modifies the passed parameters X and Y
   '  Do we need this now that we have a better pair of formulas??????????????????????????????????
   
   Dim rad As Single                                        'Distance from origin to point (radius)
   Dim angle As Single                               'Angle of incoming point in relation to origin
   
   If X > 0 And Y >= 0 Then                                                             'Quadrant 1
      angle = Atn(Y / X)
   ElseIf X < 0 And Y >= 0 Then                                                         'Quadrant 2
      angle = Atn(Y / X) + PI
   ElseIf X < 0 And Y <= 0 Then                                                         'Quadrant 3
      angle = Atn(Y / X) + PI
   ElseIf X > 0 And Y <= 0 Then                                                         'Quadrant 4
      angle = 2 * PI + Atn(Y / X)
   ElseIf X = 0 And Y >= 0 Then                                                         '90 degrees
      angle = 0.5 * PI
   Else                                                                                '180 degrees
      angle = 1.5 * PI
   End If
   rad = (X ^ 2 + Y ^ 2) ^ 0.5
   X = Cos(angle + rotate) * rad
   Y = Sin(angle + rotate) * rad
End Sub
Function Max(a, b) '***************************************** RETURNS THE LARGER OF THE TWO NUMBERS
   If a > b Then Max = a Else Max = b
End Function
Function Min(a, b) '**************************************** RETURNS THE SMALLER OF THE TWO NUMBERS
   If a < b Then Min = a Else Min = b
End Function
Function ValidTableTitle(ByVal title As String) '************************ Title Valid For SQL Table
   '  Entry    Any set of characters
   '  Return   Set of only alphanumerics and underscore. Any other characters are
   '           converted to underscores.
   Dim i As Integer
   
   For i = 1 To Len(title)
      Select Case Mid(title, i, 1)
      Case "A" To "Z", "a" To "z", "0" To "9", "_"
      Case Else
         Mid(title, i, 1) = "_"
      End Select
   Next i
   ValidTableTitle = title
End Function

Function TextToSql(txt As Variant) As Variant '************************** Makes Text SQL Compatible
   '7/6/05
   Dim index As Integer
   Dim sql As String
   
'If VarType(txt) = 1 Then Stop
   If VarType(txt) <> vbString Or VarType(txt) = vbNull Then
      '  For some reason, unless I also check for vbNull (1) it passes this test sometimes
      TextToSql = txt
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   index = 1 '+++++++++++++++++++++++++++++++++++++++++++ Convert Single Quote To Two Single Quotes
   Do While index <= Len(txt)                                   'Must reevaluate Len(txt) each time
      Select Case Mid(txt, index, 1)
      Case "'", Chr(146)                                                      'Convert single quote
         '  Also convert typographer's quote used in previous versions of the databases
         txt = Left(txt, index) & "'" & Mid(txt, index + 1)                   'To two single quotes
         index = index + 1
      Case Else
      End Select
      index = index + 1
   Loop
   
'   For index = 1 To Len(txt) '+++++++++++++++++++++++++ Convert Single Quote To Typographer's Quote
'      Select Case Mid(txt, index, 1)
'      Case "'"                                                                'Convert single quote
'         Mid(txt, index, 1) = Chr(146)                                      'To typographer's quote
'      Case Else
'      End Select
'   Next index
'
   TextToSql = txt
End Function
Function SqlToText(sql As String) As String '************* Converts SQL Compatible To Straight Text
   Dim index As Integer
   Dim txt As String
   
   txt = sql
   For index = 1 To Len(sql)
      If Mid(sql, index, 1) = Chr(146) Then      'Convert typo's close single quote to single quote
         Mid(txt, index, 1) = "'"
      End If
   Next index
   SqlToText = txt
End Function
Function ValidHTMLName(file As String, Optional reportError As Boolean = True) As String
   '  7/4/02
   '  Enter:   file           File name to check. It is changed to conform if user consents to
   '                          corrections.
   '           reportError    Report nonconformances to user for confirmation if true. Otherwise,
   '                          just make appropriate changes.
   '  Return   Changed file name after user-approved error corrections or "" if unapproved
   '           or unable to correct
   Dim i As Integer, InvalidChr As Boolean, slash As Integer
   Dim newFile As String
   
   If Len(file) > 254 And reportError Then
      MsgBox "File path too long. It must not exceed 254 characters.", vbExclamation + vbOKOnly, _
             "File Name Check"
      ValidHTMLName = ""
      Exit Function                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   newFile = file

'   slash = InStrRev(file, "\")
'   If Len(Mid(file, slash + 1)) > 14 And reportError Then
'      If MsgBox("File name exceeds the 14-character limit imposed by some UNIX systems. Proceed?", _
'                vbExclamation + vbOKCancel, "File Name Check") = vbCancel Then
'         Exit Function                                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'      End If
'   End If
   
   For i = 1 To Len(newFile)
      If InStr("&*|[]{}$<>()#?'"";^!~% ", Mid(newFile, i, 1)) Then
         Mid(newFile, i, 1) = "_"
         InvalidChr = True
      End If
   Next i
   
   If InvalidChr And reportError Then
      MsgBox "File contained character(s) that were not valid across all systems. These " _
             & "characters have been changed to underscores." & vbCrLf & vbCrLf _
             & "(The illegal characters are: &*|[]{}$<>()#?'"";^!~% and space)", _
             vbInformation + vbOKOnly, "File Name Check"
'      ValidHTMLName = ""
'   Else
   End If
   ValidHTMLName = newFile
End Function

