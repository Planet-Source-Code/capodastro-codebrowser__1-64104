Attribute VB_Name = "modUtility"
'The Connect routines are used just by the addin's first start.
'After that, VBE doesn't recognize new sessions, reacting instead in terms of
'loading, unloading and renaming projects. However, by new sessions the argument
'"OldProject.Name" in the event "rename" is vbnullstring.
'The events show up in the following order, example for a 2 prj group:
'removed        , it was a single project
'renamed
'renamed
'added
'added
'At this point is easy to assess the start and end point by loading a new
'prj or group, intended as new session. It would be possible to wait until the
'loading process ends and then refresh the tv tree, however doing on the
'same time results to be faster and the end point is used just for starting
'the tv node selection routine, for user orientation purposes.
''''''''''''''''
'The classification of events and procedures is achived by retriving the
'information from the right combo box in the code pane. Such information is
'on disposal only if the combo has the focus. In addition, the best way to
'do that is by selecting the section "General".
''''''''''''''''''''''
'The hook module messages are sent to a "Change" procedure in a text box,
'an old and well known trick.
'''''''''''''''''''''''''''''''
'The arguments in the events routines come too late, that's the reason for
'the waiting loop.
''''''''''''''''''''''''''''''''

Option Explicit

'setredraw
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_FRAME = &H400

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public IsLoading As Long
'for generating unique procedure keys.
Public k&
Public ProcessMsg As Boolean
Public StopHistory As Boolean
Public NoCodepane As Boolean
Public MDIClientStopped As Boolean
Public HwndTv As Long
Public IsPrjMDI As Boolean
Public KeyNotUnique As Boolean
Public IsCompLoading As Boolean
Public AfterRun As Boolean
Public IsFirstStart As Boolean


Public Sub FreezeMDIClient(Freeze As Boolean)


100 Select Case Freeze
    Case True
102     If MDIClientStopped Then
            Exit Sub
        Else
104         MDIClientStopped = True
106         StopRedraw hWndMDIClient, True
            'StopRedraw IDEhwnd, True
108         StopRedraw HwndTv, True
        End If
110 Case False
112     If MDIClientStopped Then
114         MDIClientStopped = False
116         StopRedraw hWndMDIClient, False
            'StopRedraw IDEhwnd, False
118         StopRedraw HwndTv, False
        Else
            Exit Sub
        End If
    End Select

End Sub

Public Sub StopRedraw(hWnd As Long, LockUpdate As Boolean)
    On Error GoTo StopRedraw_Err
    Dim r As RECT

100 If LockUpdate = True Then
102     SendMessage hWnd, WM_SETREDRAW, 0&, 0&
    Else
104     SendMessage hWnd, WM_SETREDRAW, 1&, 0&

106     GetClientRect hWnd, r
108     If RedrawWindow(hWnd, r, 0&, RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_UPDATENOW Or RDW_ALLCHILDREN Or RDW_FRAME) = 0 Then
            'Debug.Print "Failure with RedrawWindow!"
        End If

    End If

    Exit Sub

StopRedraw_Err:
110 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modUtility.StopRedraw " & _
            "at line " & Erl
112 Resume Next
End Sub




