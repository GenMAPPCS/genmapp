Attribute VB_Name = "TOC_Old"
Option Explicit
Public dbCatalogExpression As Database, rsCatalogExpression As Recordset, rsCatalogColorSet As Recordset

Sub SetExpression()
   '  Setting the expression opens up the Color Set choices
'   With frmTOCOptions
'      If .txtExpression = "" Then
'         Set dbCatalogExpression = Nothing
'         .cmbColorSet.Clear
'         .cmbColorSet.Visible = False
'         .cmbCriterion.Clear
'         .cmbCriterion.Visible = False
'         Exit Sub                                             'No expression >>>>>>>>>>>>>>>>>>>>>>>>
'      End If
'
'   On Error GoTo ErrorHandler
'      .txtExpression.Tag = .txtExpression
'      Set dbCatalogExpression = OpenDatabase(.txtExpression, , True)
'      .cmbColorSet.Visible = True
'      .cmbColorSet.Clear
'      .cmbColorSet.AddItem "ANY"
'      Set rsCatalogColorSet = dbCatalogExpression.OpenRecordset("ColorSet", dbOpenTable)
'      Do Until rsCatalogColorSet.EOF
'         .cmbColorSet.AddItem rsCatalogColorSet!colorSet
'         rsCatalogColorSet.MoveNext
'      Loop
'      .cmbColorSet.ListIndex = 0                                                  'Default to "ANY"
'      rsCatalogColorSet.Close
'      Set rsCatalogColorSet = Nothing
'   End With
   
ExitSub:
   Exit Sub

ErrorHandler:
   If Err.number = 3024 Then
'      MsgBox "Expression dataset" & vbCrLf & vbCrLf & frmTOCOptions.txtExpression & vbCrLf & vbCrLf _
'             & "does not exist.", vbExclamation + vbOKOnly
      Set dbCatalogExpression = Nothing
'      With frmTOCOptions
'         .cmbColorSet.Visible = False
'         .cmbCriterion.Visible = False
'         .txtPercent.Visible = False
'         .txtExpression.SetFocus
'      End With
      Resume ExitSub
   Else
      FatalError "Catalog:SetExpression", Err.Description
   End If
End Sub
Sub SetColorSet()
   '  Setting the color set to other than ANY opens up the criterion choices
   
'   With frmtocOptions
'      If cmbColorSet.Text = "ANY" Then
'         cmbCriterion.Visible = False
'      Else
'         Set rsCatalogColorSet = dbCatalogExpression.OpenRecordset( _
'                             "SELECT * FROM ColorSet WHERE ColorSet = '" & cmbColorSet.Text & "'")
'         GetColorSet labels, criteria, colors, notFoundIndex, rsCatalogColorSet!criteria
'         cmbCriterion.Clear
'         cmbCriterion.AddItem "ANY"
'         For i = 1 To notFoundIndex - 2
'            cmbCriterion.AddItem labels(i)
'         Next i
'         cmbCriterion.ListIndex = 0
'         cmbCriterion.Visible = True
'      End If
'   End With
End Sub

