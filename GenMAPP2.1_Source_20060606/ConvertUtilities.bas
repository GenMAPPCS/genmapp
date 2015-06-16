Attribute VB_Name = "Utilities"
Option Explicit
'****************************************** Test To See If A Column Exists In A Table In A Database
Function ColumnExists(db As Database, table As String, column As String) As Boolean
   Dim rs As Recordset, tdf As TableDef, fld As Field
   
   If db Is Nothing Then Exit Function                     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   For Each tdf In db.TableDefs
      If tdf.name = table Then '++++++++++++++++++++++++++++++++++++++++++++++++++++++ Table Exists
         For Each fld In tdf.Fields
            If fld.name = column Then
               ColumnExists = True
               Exit Function                               '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            End If
         Next fld
      End If
   Next tdf
End Function
Sub Detail(Optional text As String = "")
   Screen.ActiveForm.lblDetail = text
   DoEvents
End Sub
Sub SetProgressBase(Optional bytes As Long = 0, Optional units As String = "bytes")
   '  6/20/03
   '  Requires lblPrgMax, lblPrgValue, and prgProgress on the ActiveForm

With Screen.ActiveForm
   For i = 0 To .count - 1
      If .Controls(i).name = "lblPrgMax" Then Exit For
   Next i
   If i >= .count Then Exit Sub                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      '  Could not find lblPrgMax on the Active Form, so cannot chart progress

   If bytes <> 0 Then
      .lblPrgMax = Format(bytes, "#,###") & " " & units & " total"
      .lblPrgMax.visible = True
      .lblPrgValue = ""
      .lblPrgValue.visible = True
      .prgProgress.Max = bytes
      .prgProgress.value = 0
      .prgProgress.visible = True
   Else                                      'Just SetProgressBase turns off the progress indicator
      .prgProgress.visible = False
      .lblPrgMax = ""
      .lblPrgMax.visible = False
      .lblPrgValue = ""
      .lblPrgValue.visible = False
   End If
   DoEvents
End With
End Sub
Sub History(Optional operation As String = "", Optional bytes As Long = -1)
   '  6/20/03
   '  Entry    operation   Description of operation
   '                       Change from currentOp writes current operation to MasterUpdate.log
   '                       If operation empty then just updates current byte count and does not
   '                          write to history file
   '                       If operation and bytes not given, makes operation invisible
   '           bytes       Number of bytes processed in current operation
   '                       If -1 then bytes don't print
   '                       If -2 then increment bytes
   '  Requires lblPrgMax, lblPrgValue, lblOperation, lblDetail, and prgProgress on the ActiveForm
   
   Static currentOp As String, currentBytes As String, fileOpen As Boolean
   
   
With Screen.ActiveForm
   For i = 0 To .count - 1
      If .Controls(i).name = "lblOperation" Then Exit For
   Next i
   If i >= .count Then Exit Sub                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      '  Could not find lblOperation on the Active Form, so cannot write History

   If operation = "" And bytes = -1 Then '+++++++++++++++++++++++++++++++++++++ Turn Everything Off
      .lblOperation.visible = False
      .lblDetail.visible = False
      DoEvents
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
      
   If Not fileOpen Then '+++++++++++++++++++++++++++++++++++++++++++++++++++ Open File At First Use
      Open App.path & "\" & App.EXEName & ".log" For Output As #29
      fileOpen = True
   End If
   
   If operation = "" And bytes <> -1 Then '++++++++++++++++++++++ Update Bytes In Current Operation
      currentBytes = bytes
      If .prgProgress.visible Then .prgProgress.value = Min(bytes, .prgProgress.Max)
         '  Min because we might have added a few records in some cases
      .lblPrgValue = Format(bytes, "#,###") & " processed"
      DoEvents
      Exit Sub                                             '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   End If
   
   If operation <> currentOp Then '++++++++++++++++++++++++++++++++++++++++++ Write To History File
      .lblOperation.visible = True
      If currentOp <> "" Then
         Print #29, currentOp;
         If currentBytes <> -1 Then
            Print #29, "  " & Format(currentBytes, "#,###")
         Else
            Print #29, ""
         End If
      End If
   End If
   .lblOperation = operation
   .lblDetail = ""
   currentOp = operation
   currentBytes = bytes
   DoEvents
End With
End Sub
