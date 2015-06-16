VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objSelectArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public minX As Single                                                      'Limits of selected area
Public minY As Single
Public maxX As Single
Public maxY As Single
'  These dimensional parameters are always in nonzoomed coordinates for easy match to other objects
Public editMode As Boolean                               'Maintain compatibility with other objects

Public Function DrawObj()
   With mappWindow
   If .shpSelected.Visible = True Then
      .shpSelected.Left = minX * .zoom
      .shpSelected.Top = minY * .zoom
   End If
   End With
End Function
Sub MaxOnBoard()
   '  Maintain compatibility with other objects
End Sub
