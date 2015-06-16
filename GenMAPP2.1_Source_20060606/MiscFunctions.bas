Attribute VB_Name = "MiscFunctions"
Sub Assn(Dest As TextBox, V As Variant)
   Rem************************************************************************
   Rem  ASSIGNS VALUE TO VARIABLE IF VARIABLE EMPTY **************************
   Rem************************************************************************

   If Dat(Dest) = "" Then Dest = V
End Sub
Sub CheckBox(Check As CheckBox)
   If Check = vbChecked Then
      Check = vbUnchecked
   Else
      Check = vbChecked
   End If
End Sub
Function Dat(ByVal Z As Variant) As String '***************************************** Cleans String
   '  11/19/00
   '  Entry:
   '     Z     A string in almost any form
   '  Return:
   '     Usable string
   '  Process:
   '     Trims string
   '     Turns nulls into empty strings
   '     Turns any error condition into empty string
   '  Particularly useful for database fields that might be null
   
   On Error GoTo ErrorHandler
   If VarType(Z) <> vbNull Then
      Dat = Trim(Z)
   Else
      Dat = ""
   
   End If
ExitFunction:
   Exit Function

ErrorHandler:
   Dat = ""
   Resume ExitFunction
End Function
Function ToDB(ByVal Z As Variant) As Variant
   '  Returns empty strings as nulls to assign to database fields
   '  Handy for both strings and dates (null strings enter null dates in DB)
   '  Checkboxes should work without modification
   If VarType(Z) = vbString Then
      If Trim(Z) = "" Then
         ToDB = Null
      Else
         ToDB = Z
      End If
   Else
      ToDB = Z
   End If
End Function
Function DateTime(DTString$)
'  Returns numeric value of date given a date and time string

On Error GoTo DateTimeErr
   DateTime = DateValue(DTString$) + TimeValue(DTString$)
DateTimeRes:
   Exit Function
DateTimeErr:
   DateTime = 0
   Resume DateTimeRes
End Function
Sub Delay(Seconds)                                                          '
   Rem  *****************************************  Delays process for Seconds
   BeginTime = Now           '30 second delay to be sure fax process finishes
   DoEvents
   Do Until DateDiff("s", BeginTime, Now) > Seconds
      DoEvents
   Loop
End Sub
Function Dflt(Var, dfltVal)
   Rem************************************************************************
   Rem  dflts Var# TO dfltVal# IF Var# IS 0  ***************************
   Rem************************************************************************

   If Var = 0 Or Var = "" Then Dflt = dfltVal Else Dflt = Var

End Function
Function Div(Num As Variant, Den As Variant)
   If Den = 0 Then Div = 0 Else Div = Num / Den
End Function
Function DollarFmt(value As Variant) As String
   DollarFmt = ""
   If VarType(value) <> vbNull Then
      If Dat(value) = "-" Then
         DollarFmt = "           -"
         Exit Function
      End If
      If VarType(value) = V_STRING Then
         value = LTrim$(value)
         If InStr(value, " ") Then
            value = Left(value, InStr(value, " ") - 1)
         End If
      End If
      If Val(value) Then
         ZZ$ = Format(Val(value), "########0.00")
         DollarFmt = Space$(12 - Len(ZZ$)) & ZZ$
      End If
   End If
End Function
Function BoxFmt(ByVal value As Variant) As Variant
   '  Returns 0.00 formats for right-justified text boxes and labels for numbers
   '  Returns m/d/y format or dates
   '  Returns trimmed string
   '  Returns checked or unchecked for boolean for checkboxes
   '  Null values return empty strings
   If VarType(value) = vbNull Then                                                  'Handle nulls
      BoxFmt = ""
      Exit Function '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   If VarType(value) = vbBoolean Then                                          'Handle checkboxes
      If value Then
         BoxFmt = vbChecked
      Else
         BoxFmt = vbUnchecked
      End If
      Exit Function '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   If VarType(value) = vbDate Then                                                  'Handle Dates
      BoxFmt = Format(value, "Short Date")
      Exit Function '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   If VarType(value) = vbString Then                                            'Clean up strings
      BoxFmt = Trim(value)
      Exit Function '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   If Val(value) Then
      BoxFmt = Format(value, "0.00")
   End If
End Function
Sub FileToPrinter(fileName As String, pageHeader As String, docHeader As String)
   'File must have possible break lines preceeded by ASCII 128
   'Break is before that line
   'pageHeader prints at the beginning of each page
   '  If preceded by \, doesn't print on first page
   'docHeader prints at beginning of document
   Dim pageHeaderOnFirstpage As Boolean
   
   If Left(pageHeader, 1) = "\" Then
      pageHeaderOnFirstpage = False
      pageHeader = Mid(pageHeader, 2)                                                     'Dump \
   Else
      pageHeaderOnFirstpage = True
   End If
   Open fileName For Append As #29
   If LOF(29) < 5 Then                                                 'File empty or nonexistent
      Close #29
      Kill fileName
      Exit Sub                                       '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   Else
      Close #29
      Open fileName For Input As #29
   End If
   
   ReDim lin$(100)
   Page% = 1
   lines% = 0
   DocHeadLines% = 0
   HeaderLines% = 1
   MaxLines% = 60                                                         'Maximum lines per page
   FirstLine% = 1

   Rem  *****************************************************************  DECODE DOCUMENT HEADER
   ResetPrinter$ = Chr(27) & "E" & Chr(27) & "&k0S" & Chr(27) & "&16D" & Chr(27) & "&l4H"
   DocHead$ = ""
   lastComma% = 0
   comma% = InStr(docHeader, ",")
   If comma% = 0 Then comma% = Len(docHeader) + 1
   Do While comma% > lastComma%
      Select Case UCase(Mid(docHeader, lastComma% + 1, comma% - lastComma% - 1))
      Case "DUPLEX"                                          'Long-edge duplex
         DocHead$ = DocHead$ & Chr(27) & "&l1S"
      Case "LANDSCAPE"
         DocHead$ = DocHead$ & Chr(27) & "&l1O"
         MaxLines% = 45
      Case "LETTER"
         DocHead$ = DocHead$ & Chr(27) & "&f3y3X"
         MaxLines% = 53
      Case "UPPER"
         DocHead$ = DocHead$ & Chr(27) & "&l1H"
      Case Else
         zS$ = Mid(docHeader, lastComma% + 1, comma% - lastComma% - 1)
         lineEnd% = InStr(zS$, vbCrLf)
         Do While lineEnd%                                 'Count line endings
            DocHeadLines% = DocHeadLines% + 1
            lineEnd% = InStr(lineEnd% + 1, zS$, vbCrLf)
         Loop
         DocHead$ = DocHead$ & zS$
      End Select
      lastComma% = comma%
      comma% = InStr(lastComma% + 1, docHeader, ",")
      If comma% = 0 Then comma% = Len(docHeader) + 1
   Loop
   DocHead$ = ResetPrinter$ & DocHead$

   LastCr% = InStr(pageHeader, Chr(10))
   Do While LastCr%
      HeaderLines% = HeaderLines% + 1
      LastCr% = InStr(LastCr% + 1, pageHeader, Chr(10))
   Loop
   LastLine% = MaxLines% - 2 - HeaderLines%                      '2 for footer
'   Open "c:\windows\desktop\temp.$tm" For Output As #30                             'For testing
   Open LASER_PRINTER For Output As #30
   Print #30, DocHead$;
   If pageHeaderOnFirstpage Then Print #30, pageHeader
   Do Until EOF(29)
      If lines% > LastLine% - DocHeadLines% Then '------------------page break
                                            'Go back to last break possibility
         BreakLine% = lines%
         Do Until Left(lin$(lines%), 1) = Chr(128) Or lines% <= FirstLine%
            lines% = lines% - 1
         Loop
         If lines% <= FirstLine% Then                          'No break point
            lines% = BreakLine%                'Go back to original break line
         End If
         For Z% = FirstLine% To lines% - 1             'Print up to break line
            Print #30, lin$(Z%)
         Next Z%
         For Z% = lines% To LastLine% - DocHeadLines%    'Go to bottom of page
            Print #30, ""
         Next Z%
         Print #30, ""
         Print #30, "Page"; Page%; Chr(12); Chr(13);             'Break page
         DocHeadLines% = 0                      'Doc header only on first page
         Page% = Page% + 1
         Print #30, pageHeader
         FirstLine% = 1
         For lines% = lines% To LastLine% + 1
            Print #30, lin$(lines%)
            FirstLine% = FirstLine% + 1
         Next lines%
         lines% = FirstLine% - 1                            'Incremented below
      End If
      lines% = lines% + 1
      Line Input #29, lin$(lines%)
   Loop
   For Z% = FirstLine% To lines%                       'Finish unprinted lines
      Print #30, lin$(Z%)
   Next Z%
   For Z% = lines% + 1 To LastLine%                      'Go to bottom of page
      Print #30, ""
   Next Z%
   Print #30, ""
   If Page% > 1 Then Print #30, "Page"; Page%;
   Print #30, Chr(12); ResetPrinter$                    'Break page and reset
   Close #29, #30
End Sub

Function Fmt(value As Variant, Form As String) As String '********************** General Formatting
   '  2/16/04
   '  Calls InChrs() function
   '  Calls Dat() function
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
   Fmt = "<Error" & Str$(Err) & ">"
   Resume FmtExit
End Function
Function InChrs%(S1, S2)
   Rem************************************************************************
   Rem  TRUE IF ANY CHARACTERS EXIST IN BOTH STRINGS  ************************
   Rem************************************************************************
   '  Returns position of first occurrence of a character in S2 within S1

   InChrs% = 0
   If Dat(S1) = "" Then Exit Function
   If Dat(S2) = "" Then Exit Function
   For Z% = 1 To Len(S2)
      If InStr(S1, Mid(S2, Z%, 1)) Then
         InChrs% = InStr(S1, Mid(S2, Z%, 1))
         Exit Function
      End If
   Next Z%
End Function
Function InitCap(Tx As Variant) As String '************************ Returns value as initial caps
'  Must return value rather than accept string by reference because textboxes won't pass Text
'  property by reference
   Dim text As String, Char As Integer
   
   If VarType(Tx) <> V_STRING Or Tx = "" Then Exit Function
   
   text = LCase(Tx)
   Mid(text, 1, 1) = UCase(Mid(text, 1, 1))
   For Char = 2 To Len(text) - 1
      Select Case Mid(text, Char, 1)
      Case " ", "/", ",", ".", ":", ";", "-", "&", Chr(34)
         Mid(text, Char + 1, 1) = UCase(Mid(text, Char + 1, 1))
      End Select
   Next Char
   InitCap = text
End Function

Sub LineExpand(MaxWidth As Integer, Chars, zS() As String)
'*********************************************************** CHANGE AUTO LINE BREAKS TO REAL ONES
'  Line to break comes in in Chars
   '  Output in Zs()
   '  MaxWidth of zero means break only on CR-LF
   Dim lines As Integer, lineBeg As Integer, lineEnd As Integer
   
   If MaxWidth = 0 Then MaxWidth = 1000
   lines = 1               'Number of actual lines after counting auto breaks
   lineBeg = 1
   lineEnd = InStr(Chars, vbCrLf)                     'Last chrs actually in line
   If lineEnd = 0 Then lineEnd = Len(Chars) + 1
   Do While lineBeg <= lineEnd
      If lineEnd - lineBeg > MaxWidth Then                      'Find space
         lineEnd = lineBeg + MaxWidth - 1
         Do While Mid(Chars, lineEnd, 1) <> " " And lineEnd > lineBeg
            lineEnd = lineEnd - 1
         Loop
         If lineEnd <= lineBeg Then                                            'No place to break line
            zS(lines) = Mid(Chars, lineBeg, MaxWidth)                       'Go back to max width
            lineBeg = lineBeg + MaxWidth
         Else
            zS(lines) = Mid(Chars, lineBeg, lineEnd - lineBeg + 1)                         'w/spc
            lineBeg = lineEnd + 1                                                   'Beyond space
         End If
      Else                                                                         'Ends at CR-LF
         zS(lines) = Mid(Chars, lineBeg, lineEnd - lineBeg)                             'No CR-LF
         lineBeg = lineEnd + 2                                                      'Beyond CR-LF
      End If
      zS(lines) = RTrim(zS(lines))
      If zS(lines) = "" Then zS(lines) = " "
      lineEnd = InStr(lineBeg, Chars, vbCrLf)                                         'Next CR-LF
      If lineEnd = 0 Then lineEnd = Len(Chars) + 1
      lines = lines + 1
   Loop
   For lines = lines To UBound(zS)                                           'Blank rest of lines
      zS(lines) = ""
   Next lines
End Sub

Function ListingName(Credt, Debt) As String
   Rem************************************************************************
   Rem  BUILDS A LIMITED-LENGTH LISTING FROM CREDITOR AND DEBTOR  ************
   Rem************************************************************************

   Cred$ = Dat(Credt)                              'Don't change passed names
   Deb$ = Dat(Debt)
   MinDebtorLen% = 18
   MinCreditorLen% = 11
   If Len(Cred$) > MinCreditorLen% Then
      extra% = Max(0, MinDebtorLen% - Len(Deb$))
      Z1$ = Left(Cred$, MinCreditorLen% + extra%)
   Else
      Z1$ = Cred$
   End If
   If Len(Deb$) > MinDebtorLen% Then
      extra% = Max(0, MinCreditorLen% - Len(Cred$))
      Z2$ = Left(Deb$, MinDebtorLen% + extra%)
   Else
      Z2$ = Deb$
   End If
   ListingName = Z1$ + " v " + Z2$
End Function

Function Max(A, B) '***************************************** RETURNS THE LARGER OF THE TWO NUMBERS
   If A > B Then Max = A Else Max = B
End Function

Function Min(A, B) '**************************************** RETURNS THE SMALLER OF THE TWO NUMBERS
   If A < B Then Min = A Else Min = B
End Function
Function NVLd(Dflt As Variant, substitute As Double) As Double '*********** Null Value Substitute
   '  Works with Double data types
   If VarType(Dflt) = vbNull Then
      NVLd = substitute
   ElseIf Dflt = "" Then
      NVLd = substitute
   Else
      NVLd = Dflt
   End If
End Function
Function NullSub(Dflt As Variant, Subs As String) As String
   Rem************************************************************************
   Rem  IF NO DFLT$, SUBSTITUTE SUBS$  ***************************************
   Rem************************************************************************

   If VarType(Dflt) = vbNull Then
      NullSub = Subs
   ElseIf Dflt = "" Then
      NullSub = Subs
   Else
      NullSub = Dflt
   End If
End Function

Function NullZero(V As Variant) As Variant
   Rem************************************************************************
   Rem  RETURNS ZERO FOR NULL VARIABLES OR EMPTY STRINGS  ********************
   Rem************************************************************************

   If VarType(V) = vbNull Then
      NullZero = 0
   ElseIf VarType(V) = vbString Then
      If Dat(V) = "" Then
         NullZero = 0
      Else
         NullZero = -1
      End If
   Else
      NullZero = V
   End If
End Function
Function Para$(V)
   Rem************************************************************************
   Rem  SUBSTITUTES PARAGRAPH BREAKS FOR CrLfs OR \ IN STRING  *****************
   Rem************************************************************************
   
   Rem  **********************************************************  REPLACE \s
   PrevBrk% = 0
   Brk% = InStr(V, "\")
   Do While Brk%
      zS$ = zS$ & Mid(V, PrevBrk% + 1, Brk% - PrevBrk% - 1) & "\par "
      PrevBrk% = Brk%
      Brk% = InStr(PrevBrk% + 1, V, "\")
   Loop
   V = zS$ & Mid(V, PrevBrk% + 1, 500)
   
   Rem  *********************************************************  REPLACE vbcrlfs
   zS$ = ""
   PrevBrk% = -1
   Brk% = InStr(V, vbCrLf)
   Do While Brk%
      zS$ = zS$ & Mid(V, PrevBrk% + 2, Brk% - PrevBrk% - 2) & "\par "
      PrevBrk% = Brk%
      Brk% = InStr(PrevBrk% + 1, V, vbCrLf)
   Loop
   
   Para$ = zS$ & Mid(V, PrevBrk% + 2, 500)
End Function
Rem *************************************** SUBSTITUTES PARAGRAPH BREAKS FOR CrLfs OR \ IN STRING
Function zPara(rawText As String) As String
   Dim prevBreak As Integer, break As String                                    'Paragraph breaks
   Dim outText As String                                                'Processed text to output
   
   '================================================================================== REPLACE \s
   prevBreak = 0
   break = InStr(rawText, "\")
   Do While break
      outText = outText & Mid(rawText, prevBreak + 1, break - prevBreak - 1) & Chr(13)
      prevBreak = break
      break = InStr(prevBreak + 1, V, "\")
   Loop
   rawText = outText & Mid(rawText, prevBreak + 1)
   
   '============================================================================= REPLACE vbCrLfs
   outText = ""
   prevBreak = -1
   break = InStr(rawText, vbCrLf)
   Do While break
      outText = outText & Mid(rawText, prevBreak + 2, break - prevBreak - 2) & Chr(13)
      prevBreak = break
      break = InStr(prevBreak + 1, rawText, vbCrLf)
   Loop
   
   zPara = outText & Mid(rawText, prevBreak + 2, 500)
End Function
Function SearchKey(searchIn As Variant) As String '****************** SEARCH KEY OF 20 CHARACTERS
   '  Takes searchIn and returns only uppercase letters and a maximum of 20 characters
   '  The first character, no matter what it is, is taken
   Dim search As String, key As String, Char As String * 1
   Dim searchPos As Integer
   
   search = UCase(Trim(searchIn))
   If search = "" Then
      SearchKey = ""
      Exit Function '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   key = Left(search, 1)                                                'First chr no matter what
   For searchPos = 2 To Len(search)
      Char = Mid(search, searchPos, 1)
      If Char >= "A" And Char <= "Z" Then
         key = key & Char
      End If
   Next searchPos
   SearchKey = Left(key, 20)
End Function
Function value(ByVal Var As Variant) '************* RETURNS NUMERIC VALUES OF NUMBERS WITH COMMAS
   If VarType(Var) = vbNull Then
      value = 0
   ElseIf Dat(Var) = "" Then
      value = 0
   Else
      comma% = InStr(Var, ",")
      Do While comma%
         Var = Left(Var, comma% - 1) & Mid(Var, comma% + 1)
         comma% = InStr(Var, ",")
      Loop
      value = Val(Var)
   End If
End Function
Function EarlierDate(date1 As Variant, date2 As Variant) As Date '********** Returns Earlier Date
   'Returns earlier date or only valid date or zero if neither date valid
   If IsNull(date1) Then
      If IsNull(date2) Then
         EarlierDate = 0                                     'Error condition, neither date valid
      Else
         EarlierDate = CDate(date2)
      End If
   ElseIf IsNull(date2) Then
      EarlierDate = date1
   Else                                                                             'Neither null
      date1 = CDate(date1)
      date2 = CDate(date2)
      If date1 < date2 Then
         EarlierDate = date1
      Else
         EarlierDate = date2
      End If
   End If
End Function
