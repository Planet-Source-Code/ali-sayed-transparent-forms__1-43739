Attribute VB_Name = "ModSubClass"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal OldwndProc As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const GWL_WNDPROC = -4
Public OldwndProc As Long

Private Const WM_MOVE = &H3
Private Const WM_SIZE = &H5
Private Const WM_ACTIVATE = &H6

Public mCls As ClsTrans
Public mForm As Form
Public mMode As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal WindowMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

Select Case WindowMsg
    'if window pos changed
    Case WM_MOVE
        mCls.Trans mForm, mMode
    'if window size changed
    Case WM_SIZE
        mCls.Trans mForm, mMode
End Select

WindowProc = CallWindowProc(OldwndProc, hwnd, WindowMsg, wParam, lParam)
End Function



