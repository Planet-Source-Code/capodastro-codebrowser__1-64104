VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMouseEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2000 Karl E. Peterson, All Rights Reserved
'  Find this and more samples at <http://www.mvps.org/vb>
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal dwData As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' API structure definition for Rectangle
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' API structure definition for Mouse Coordinates
Private Type POINTAPI
    X As Long
    Y As Long
End Type

' Flags used with mouse_event
Private Const MOUSEEVENTF_ABSOLUTE = &H8000& ' absolute move
Private Const MOUSEEVENTF_LEFTDOWN = &H2     ' left button down
Private Const MOUSEEVENTF_LEFTUP = &H4       ' left button up
' middle button up
Private Const MOUSEEVENTF_MOVE = &H1         ' mouse move
' right button up

' GetSystemMetrics() codes
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

' A few module level variables...
Private m_ScreenWidth As Long
Private m_ScreenHeight As Long
Private m_ClickDelay As Long

' Virtual scaling applied to screen...
Private Const m_Scale As Long = &HFFFF&


Public Sub SetMousePos(xPos As Long, yPos As Long)


100 SetCursorPos xPos, yPos
End Sub
Public Function GetY() As Long


    Dim n As POINTAPI
100 GetCursorPos n
102 GetY = n.Y
End Function
Public Function GetX() As Long


    Dim n As POINTAPI
100 GetCursorPos n
102 GetX = n.X
End Function
' ***********************************************************
'  Initialize
' ***********************************************************
Private Sub Class_Initialize()

    ' Store screen dimensions in pixels
100 m_ScreenWidth = GetSystemMetrics(SM_CXSCREEN)
102 m_ScreenHeight = GetSystemMetrics(SM_CYSCREEN)
    ' Default duration for mousedown
104 m_ClickDelay = 0   '250 'milliseconds
End Sub



Public Sub Click()

    ' Click the mouse, with delay to simulate human timing.
100 Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
102 If m_ClickDelay Then
104     DoEvents ' allow down position to paint
106     Call Sleep(m_ClickDelay)
    End If
108 Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
End Sub

Public Sub ClickWindow(ByVal hWnd As Long)

    ' Move cursor to window
100 Call Me.MoveToWindow(hWnd)
    ' Click it
102 Call Me.Click
End Sub

' X/Y need to be passed as pixels!
Public Sub MoveTo(ByVal X As Long, ByVal Y As Long, Optional ByVal Absolute As Boolean = True)

    Dim meFlags As Long

100 If Absolute Then
        ' Map into same coordinate space used by mouse_event.
102     X = (X / m_ScreenWidth) * m_Scale
104     Y = (Y / m_ScreenHeight) * m_Scale
        ' Set flags
106     meFlags = MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE
    Else
        ' Set flags for relative movement
108     meFlags = MOUSEEVENTF_MOVE
    End If

    ' Move the cursor to destination.
110 Call mouse_event(meFlags, X, Y, 0, 0)
End Sub

Public Sub MoveToWindow(ByVal hWnd As Long)

    Dim X As Long, Y As Long
    Dim r As RECT

    ' Place origin in center of control.
100 Call GetWindowRect(hWnd, r)
102 X = r.Left + (r.Right - r.Left) \ 2
104 Y = r.Top + (r.Bottom - r.Top) \ 2
106 Call Me.MoveTo(X, Y)
End Sub




