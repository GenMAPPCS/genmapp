VERSION 5.00
Begin VB.Form frmBalloon 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2844
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   4116
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2844
   ScaleWidth      =   4116
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Display(message As String, X As Single, Y As Single)
   Dim maxWidth As Single, lin As String, prevSlash As Integer, slash As Integer
   
   If message = "" Then Exit Sub                           '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   Left = X + 50
   Top = Y + 50
   Width = 10000                     'Must be big at first so that characters don't run beyond form
   Height = 10000
   Cls
   CurrentY = 0
   prevSlash = 0
   slash = InStr(message, "\")
   Do While prevSlash < Len(message)
      slash = InStr(prevSlash + 1, message, "\")
      If slash = 0 Then slash = Len(message) + 1
      
      lin = Mid(message, prevSlash + 1, slash - prevSlash - 1)
      CurrentX = 10
      Print lin
      maxWidth = Max(maxWidth, TextWidth(lin) + 50)
      prevSlash = slash
   Loop
   Width = maxWidth
   Height = CurrentY + 30
   Show
End Sub

