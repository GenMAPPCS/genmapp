VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPreview_Old 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Full-Page Preview"
   ClientHeight    =   8892
   ClientLeft      =   120
   ClientTop       =   720
   ClientWidth     =   10596
   Icon            =   "Preview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8892
   ScaleWidth      =   10596
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   240
      Top             =   7920
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbrBar 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   8640
      Visible         =   0   'False
      Width           =   10596
      _ExtentX        =   18690
      _ExtentY        =   445
      Style           =   1
      SimpleText      =   "Approximation of full-page printer output."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
   Begin VB.Menu mnuExport 
      Caption         =   "&Export"
      Begin VB.Menu mnuBMP 
         Caption         =   "to &BMP"
      End
      Begin VB.Menu mnuJPEG 
         Caption         =   "to &JPEG"
      End
      Begin VB.Menu mnuHTML 
         Caption         =   "to &HTML"
      End
   End
End
Attribute VB_Name = "frmPreview_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbGene As Database, dbExpression As Database
Public callingForm As Form
Private Sub Form_Activate()
   Set dbGene = callingForm.dbGene
   Set dbExpression = callingForm.dbExpression
   If htmlHome <> "" Then '+++++++++++++++++++++++++++++++++++++++++++++++++ Generate HTML Page Set
      HTMLExport Left(htmlHome, InStrRev(htmlHome, "\")), Mid(htmlHome, InStrRev(htmlHome, "\") + 1)
      Hide
   End If
End Sub

Private Sub mnuClose_Click()
   Hide
End Sub

Public Sub mnuHTML_Click() '****************************************************** Produce Web Site
   Dim file As String
   
   file = SetFileName("HTML", frmMAPPInfo.txtTitle, mruExportPath)
   If file = "ERROR" Or file = "CANCEL" Then Exit Sub      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   file = Left(file, InStrRev(file, ".") - 1)                                            'Dump .htm
   mruExportPath = Left(file, InStrRev(file, "\"))                   'Change configured export path
   HTMLExport mruExportPath, Mid(file, InStrRev(file, "\") + 1)
End Sub
Sub HTMLExport(folder As String, mapp As String) '*************************** Create Web Site Pages
   '  Enter:   folder   Main HTML page. Backpages will be in subfolder  Eg. C:\GenMAPP\Exports\
   '           mapp     Name of MAPP file and HTML page, without extension
   Dim graphic As Picture, JPEGfile As String
   Dim supportPath As String                            'Path for support files: JPGs and backpages
                                                        'Name of MAPP & /_Support/
   Dim element As Object
   Dim backpageFile As String, adjust As Single, file As String, backpageHead As String
   
   MousePointer = vbHourglass
'   mapp = Left(file, InStrRev(file, ".") - 1)
   If Dir(folder, vbDirectory) = "" Then
      MkDir folder
   End If
   If Dir(folder & "_Support", vbDirectory) = "" Then
      MkDir folder & "_Support"
   End If
   If Dir(folder & "_Support\" & mapp, vbDirectory) = "" Then
      MkDir folder & "_Support\" & mapp
   End If
'   If Dir(folder & mapp, vbDirectory) = "" Then
'      MkDir folder & mapp
'   End If
   supportPath = "_Support/" & mapp & "/"                        '/ is Standard HTML path delimiter
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Set Up HTML File
   Open folder & mapp & htmlSuffix & ".htm" For Output As #30
   Print #30, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
   Print #30, "<html>"
   Print #30, "<head>"
   Print #30, "   <title>" & frmMAPPInfo.txtTitle & "</title>"
   Print #30, "   <meta name=""generator"" content=""GenMAPP 2.0"">"
   Print #30, "</head>"
   Print #30, ""
   Print #30, "<body>"
   Print #30, colorSetHTML
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Insert MAPP Graphic
   JPEGfile = folder & "_Support\" & mapp & "\" & mapp & htmlSuffix & ".jpg"
   CreateJPEG JPEGfile
   JPEGfile = Mid(JPEGfile, InStrRev(JPEGfile, "\") + 1)
   Print #30, "<img src=""" & supportPath & JPEGfile & """ alt = """ & JPEGfile _
              & """ usemap=""#MAPP"" border=0>"
   Print #30, "<map name=""MAPP"">"
   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Create Backpages And Image Map
   adjust = mappWindow.zoom / Screen.TwipsPerPixelX
   For Each element In mappWindow.objLumps '----------------------------------------------Each Gene
      If element.objType = "Gene" Then
         If element.head <> "" Then
            backpageHead = element.head
         Else
            backpageHead = element.title
         End If
         backpageFile = CreateBackpage(element.ID, element.systemCode, backpageHead, _
                                       callingForm.dbGene, callingForm.dbExpression, _
                                       element, folder & "_Support\" & mapp & "\")
         If backpageFile <> "" Then                                  'Backpage successfully created
            backpageFile = Mid(backpageFile, InStrRev(backpageFile, "\") + 1)    'Filename, no path
            Print #30, "   <area href=""" & supportPath & backpageFile & """"
            Print #30, "         shape=""rect"""
            Print #30, "         coords=""" & Int((element.CenterX - element.wide / 2) * adjust) _
                       & "," & Int((element.CenterY - element.high / 2) * adjust) & "," _
                       & Int((element.CenterX + element.wide / 2) * adjust) & "," _
                       & Int((element.CenterY + element.high / 2) * adjust) & """"
            Print #30, "         alt=""Click for backpage, shift-click for separate window"">"
         End If
      End If
   Next element
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ End HTML file
   Print #30, "</map>"
   Print #30, "</body>"
   Print #30, "</html>"
   Close #30
   MousePointer = vbDefault
End Sub

Private Sub mnuJPEG_Click()
   Dim file As String
   
'   If Dir(mruExportPath, vbDirectory) = "" Then                           'Export folder doesn't exist
'      AddFolder mruExportPath
'   End If
   file = SetFileName("JPEG", frmMAPPInfo.txtTitle, mruExportPath)
   If file = "ERROR" Or file = "CANCEL" Then Exit Sub      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   mruExportPath = Left(file, InStrRev(file, "\"))                   'Change configured export path
   
   CreateJPEG file
End Sub
Sub CreateJPEG(file As String)
   Dim dib As New cDIBSection
   Dim dot As Integer
   
   dot = InStrRev(file, ".")
   If dot = 0 Then
      dot = Len(file) + 1
   End If
   file = Trim(Left(file, dot - 1)) & Mid(file, dot)
      '  Intel's SaveJPG() crashes if the file name ends in space, eg. "file .jpg"
   MousePointer = vbHourglass
   DoEvents
   dib.CreateFromPicture CaptureClient(Me)     'Picture object works as well as a StdPicture object
   SaveJPG dib, file
   MousePointer = vbDefault
ExitSub:
   Exit Sub                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
ErrorHandler:
   Select Case Err.number
   End Select
End Sub
Private Sub mnuPrint_Click()
   frmDrafter.mnuPrint_Click
End Sub

Private Sub mnuBMP_Click()
   '  This is a slightly different process than mnuJPEG or mnuHTML. They are better.
   '  Will probably want to change this to agree.
   Dim graphic As Picture, graphicName As String
   
On Error GoTo ErrorHandler

ReEnter:
   dlgDialog.CancelError = True
   dlgDialog.Filter = "Graphic files (.bmp)|bmp"
   dlgDialog.FileName = mruMappPath & "*.bmp"
   dlgDialog.FLAGS = cdlOFNExplorer + cdlOFNHideReadOnly
   dlgDialog.ShowSave
   graphicName = dlgDialog.FileName
   If InStr(graphicName, ".") = 0 Then
      graphicName = graphicName & ".bmp"
   End If

   If UCase(Dir(graphicName)) = UCase(Mid(graphicName, InStrRev(graphicName, "\") + 1)) Then
      Select Case MsgBox("Do you want to replace the current " & graphicName & "?", _
                  vbYesNoCancel + vbQuestion, "Saving Graphic")
      Case vbNo
         GoTo ReEnter                                   '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      Case vbCancel
         GoTo ExitSub                                   'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
      End Select
   End If
   
   MousePointer = vbHourglass
   DoEvents
   Set graphic = CaptureClient(Me)           'Put active form (frmDrafter) bitmap in Picture object
   SavePicture graphic, graphicName

ExitSub:
   MousePointer = vbDefault
   Exit Sub
   
ErrorHandler:
   If Err = 70 Then
      MsgBox Err.Description & ". " & graphicName & " possibly open in some other program.", _
            vbCritical, "BMP Export Error"
   ElseIf Err <> 32755 Then                                         'Not an error if just cancelled
      MsgBox Err.Description, vbCritical, "BMP Export Error"
   End If
   On Error GoTo 0
   Resume ExitSub
End Sub


