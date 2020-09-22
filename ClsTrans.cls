VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'********************************************
' Project:  Trans Class   v.1.0             *
' Author:   Ali Sayed                       *
' E-Mail:   AliSayed_7@Yahoo.com            *
' Date:     04/03/2003                      *
' Copyright © 2003 Ali Sayed                *
' Please let me know if you like it.        *
' For more information mail me.             *
'********************************************
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private SubClassed As Long
'trans options
Public Enum TransMode
    Normal = RasterOpConstants.vbSrcAnd
    Invert = RasterOpConstants.vbNotSrcCopy
End Enum

'let the class handle all event
Sub SubClass(nForm As Object, nTransMode As TransMode)
'check if already subclassed
If nForm.hwnd = SubClassed Then Exit Sub
SubClassed = nForm.hwnd
'subclass from
OldwndProc = SetWindowLong(nForm.hwnd, GWL_WNDPROC, AddressOf WindowProc)
'global obj
Set mCls = Me
Set mForm = nForm
mMode = nTransMode
'start trans
Trans mForm, mMode
End Sub

'do Transparent as asked
Public Sub Trans(nForm As Object, nTransMode As TransMode)
Dim srcHANDLE As Long, srcDC As Long
Dim xWidth As Long, yHeight As Long
Dim xLeft As Long, yTop As Long
'set form autoredraw
nForm.AutoRedraw = True
'remove old pic
nForm.Cls
'''''''''''size'''''''''''''''''''''''''''''
'set edges spaces
If nForm.BorderStyle > 0 Then xLeft = 2

If nForm.BorderStyle = 0 Then
    yTop = 0
ElseIf nForm.BorderStyle >= 1 And nForm.BorderStyle <= 3 Then
    yTop = 23
Else
    yTop = 20
End If
'width ,height for form
xWidth = nForm.Width / Screen.TwipsPerPixelX
yHeight = nForm.Height / Screen.TwipsPerPixelY
''''''''''''''size'''''''''''''''''''''''''

'since we want the background picture so we hide
'the form for instant
nForm.Hide
    DoEvents
    'get a handle to screen
    srcHANDLE = GetDesktopWindow()
    'get device context To screen
    srcDC = GetWindowDC(srcHANDLE)
    'draw background oicture to form
    BitBlt nForm.hDC, 0, 0, xWidth, yHeight, srcDC, (nForm.Left / Screen.TwipsPerPixelX) + xLeft, (nForm.Top / Screen.TwipsPerPixelY) + yTop, mMode
    'show the form again with picture drawn
nForm.Show
'refresh it :)
nForm.Refresh
End Sub
