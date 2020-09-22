Attribute VB_Name = "modSubclass"
Option Explicit

Public VBInstance                   As VBIDE.VBE 'this has the instantiated application object
Public Connect                      As Connect

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const IDX_WINDOWPROC        As Long = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Const PropName              As String = "Hooked"

Private Declare Function LBItemFromPt Lib "comctl32" (ByVal hLB As Long, ByVal X As Long, ByVal Y As Long, ByVal bAutoScroll As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'the classes to work with
Private Const VBA_WINDOW          As String = "VbaWindow"           ' hide LnkWnds if mouse over
Private Const VBA_COMBOBOX          As String = "ComboBox"
Public Const VBA_NEWPROC          As String = "NewProc"

'the events to retrive by hook
Private Const WM_SETTEXT = &HC
Private Const WM_SETFOCUS           As Long = 7
Private Const WM_KILLFOCUS          As Long = 8
Private Const WM_MDIACTIVATE        As Long = &H222

'lsthistory
Private Declare Function SendMessagebyString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam As String) As Long

Private Const Vb_MDIChild = "MDIChild"

'working variables
Public hWndMDIClient                As Long
Public hWndCodePane                 As Long
Public IDEhwnd                      As Long
Public hWndCmbLeft                  As Long
Public hWndCmbRight                 As Long
Public hWndProperties               As Long
Public bClosingSession              As Boolean
Public bInRunMode                   As Boolean
Public hWndTextBox                  As Long
Public bControlChange              As Boolean
Public Function HIWORD(ByVal nValue&) As Integer

    ' returns the high 16-bit integer from a 32-bit long integer
100 CopyMemory HIWORD, ByVal VarPtr(nValue) + 2, 2&
End Function

Public Function LOWORD(ByVal dwValue As Long) As Integer

    ' Returns the low 16-bit integer from a 32-bit long integer
100 CopyMemory LOWORD, dwValue, 2&
End Function

Public Function FindActiveCodepane() As Long

    Dim oCodePane As CodePane

100 Set oCodePane = VBInstance.ActiveCodePane

102 If Not oCodePane Is Nothing Then
104     FindActiveCodepane = FindWindowEx(hWndMDIClient, 0, _
                VBA_WINDOW, oCodePane.Window.Caption)
    End If

End Function



Public Sub SendString(sParam As String)
    On Error GoTo eH
100 If hWndTextBox Then _
            Call SendMessagebyString(hWndTextBox, WM_SETTEXT, 0, sParam)

    Exit Sub
eH:
102 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modSubclass.SendString " & _
            "at line " & Erl
104 Resume Next
End Sub


Private Function CodePaneProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo eH

    'Window Procedure for the active IDE Codepane
100 CodePaneProc = CallWindowProc(GetProp(hWnd, PropName), hWnd, nMsg, wParam, lParam)

    Exit Function
eH:
102 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modSubclass.CodePaneProc " & _
            "at line " & Erl
104 Resume Next
End Function

Private Function CodePaneCmbLProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo eH

    'Window Procedure for the active IDE Codepane
100 CodePaneCmbLProc = CallWindowProc(GetProp(hWnd, PropName), hWnd, nMsg, wParam, lParam)

    Exit Function
eH:
102 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modSubclass.CodePaneCmbLProc " & _
            "at line " & Erl
104 Resume Next
End Function

Private Function CodePaneCmbRProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo eH
    Static fActivateCombo As Boolean

    'Window Procedure for the active IDE Codepane
100 CodePaneCmbRProc = CallWindowProc(GetProp(hWnd, PropName), hWnd, nMsg, wParam, lParam)

102 If Not bInRunMode Then
        'snap selections by combo
104     If nMsg = WM_SETFOCUS Then
106         fActivateCombo = True

108     ElseIf nMsg = WM_KILLFOCUS Then
110         fActivateCombo = False

        End If

112     If fActivateCombo Then
114         If nMsg = 641 And wParam = 0 Then 'the user selected a member from the combo
                'check the members in the module
116             SendString VBA_NEWPROC
            End If    'end snap selections by combo
        Else
            'snap selections by pane
118         If nMsg = 336 Then 'the user selected a member from the pane
                'check the members in the module
120             SendString VBA_NEWPROC
            End If    'end snap selections by pane
        End If
    End If

    Exit Function
eH:
122 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modSubclass.CodePaneCmbRProc " & _
            "at line " & Erl
124 Resume Next
End Function
Function fnHasCode(cMod As CodeModule) As Boolean
    On Error GoTo eH
    Dim i As Long

100 For i = 1 To cMod.CountOfLines
102     If LenB(Trim$(cMod.Lines(i, 1))) Then
104         fnHasCode = True
            Exit Function
        End If
    Next

    Exit Function
eH:
106 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modSubclass.fnHasCode " & _
            "at line " & Erl
108 Resume Next
End Function

Private Sub HookCodePane()
    On Error GoTo eH

100 With VBInstance
102     If .ActiveWindow Is .ActiveCodePane.Window Then
104         hWndCodePane = FindWindowEx(hWndMDIClient, 0, VBA_WINDOW, .ActiveWindow.Caption)

106         If hWndCodePane Then
108             If GetProp(hWndCodePane, PropName) = 0 Then
110                 SetProp hWndCodePane, PropName, GetWindowLong(hWndCodePane, IDX_WINDOWPROC)
112                 SetWindowLong hWndCodePane, IDX_WINDOWPROC, AddressOf CodePaneProc
114                 HookCodePaneCombos
                End If
            End If


        End If
    End With 'VBINSTANCE

    Exit Sub
eH:
116 Select Case Err.Number
    Case 91 'obiekt variable nicht festgelegt, by scratch start
118     Resume Next
120 Case 76 'Path not found
122     UnhookMainWindow
    Case 40036, 60061 'Method '~' of object '~' failed, load with error
        Resume Next
124 Case Else
126     MsgBox Err.Description & vbCrLf & _
                "in CodeBrowser.modSubclass.HookCodePane " & _
                "at line " & Erl & vbCrLf & _
                Err.Number
        'Resume Next NOOOOOOOOO!!!!
    End Select
End Sub
Public Sub HookCodePaneCombos()
    On Error GoTo eH
    Dim tmpHwnd&, tmpHwnd1&, wRct As RECT, wRct1 As RECT

100 If hWndCodePane Then
102     hWndCodePane = FindWindowEx(hWndMDIClient, 0, VBA_WINDOW, vbNullString)

104     If hWndCodePane Then
106         tmpHwnd = FindWindowEx(hWndCodePane, 0, VBA_COMBOBOX, vbNullString)

108         If tmpHwnd Then
110             GetWindowRect tmpHwnd, wRct
112             tmpHwnd1 = FindWindowEx(hWndCodePane, tmpHwnd, VBA_COMBOBOX, vbNullString)

114             If tmpHwnd1 Then
116                 GetWindowRect tmpHwnd1, wRct1

118                 If wRct.Left > wRct1.Left Then
120                     hWndCmbRight = tmpHwnd
122                     hWndCmbLeft = tmpHwnd1
                    Else
124                     hWndCmbRight = tmpHwnd1
126                     hWndCmbLeft = tmpHwnd
                    End If
                End If
            End If
        End If

128     If GetProp(hWndCmbRight, PropName) = 0 Then
130         SetProp hWndCmbRight, PropName, GetWindowLong(hWndCmbRight, IDX_WINDOWPROC)
132         SetWindowLong hWndCmbRight, IDX_WINDOWPROC, AddressOf CodePaneCmbRProc
        End If

134     If GetProp(hWndCmbLeft, PropName) = 0 Then
136         SetProp hWndCmbLeft, PropName, GetWindowLong(hWndCmbLeft, IDX_WINDOWPROC)
138         SetWindowLong hWndCmbLeft, IDX_WINDOWPROC, AddressOf CodePaneCmbLProc
        End If
    End If

    Exit Sub
eH:
140 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modSubclass.HookCodePaneCombos " & _
            "at line " & Erl
End Sub
Public Sub HookMDIClient()
'! Delayed error handler
    On Error Resume Next

100 If hWndMDIClient Then
102     If GetProp(hWndMDIClient, PropName) = 0 Then
104         SetProp hWndMDIClient, PropName, GetWindowLong(hWndMDIClient, IDX_WINDOWPROC)
106         SetWindowLong hWndMDIClient, IDX_WINDOWPROC, AddressOf MDIClientProc
108         HookCodePane
        End If
    End If

End Sub
Public Sub HookMainWindow()
'! Delayed error handler
    On Error Resume Next
    Dim hWndTmp&

100 If IDEhwnd Then
102     If GetProp(IDEhwnd, PropName) = 0 Then
104         SetProp IDEhwnd, PropName, GetWindowLong(IDEhwnd, IDX_WINDOWPROC)
106         SetWindowLong IDEhwnd, IDX_WINDOWPROC, AddressOf MainWindowProc
108         HookMDIClient
        End If
    End If


110 hWndTmp = FindWindowEx(IDEhwnd, 0, "wndclass_pbrs", vbNullString)
112 hWndProperties = FindWindowEx(hWndTmp, 0, "ListBox", vbNullString)
End Sub
Private Function MDIClientProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'! Delayed error handler
    'Window procedure for VB's MDIClient window
    On Error Resume Next

100 MDIClientProc = CallWindowProc(GetProp(hWnd, PropName), hWnd, nMsg, wParam, lParam)  'call the original winproc to do what has to be done

102 If Not bInRunMode Then
104     Select Case nMsg 'and now split on message type
        Case WM_KILLFOCUS 'this codepane just lost the focus (remember - the original procedure has already been performed)
106         UnhookCodePane
108     Case WM_MDIACTIVATE, WM_SETFOCUS 'another codepane has been (re)activated by the user
110         HookCodePane
        End Select
    End If

End Function
Private Function MainWindowProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Static MDIChildPropSelected As Integer

100 MainWindowProc = CallWindowProc(GetProp(hWnd, PropName), hWnd, nMsg, wParam, lParam)  'call the original winproc to do what has to be done

102 If Not bInRunMode And IsPrjMDI Then
104     If nMsg = 528 Then                  'some change-focus by click msg
106         If VBInstance.ActiveWindow.Type = vbext_wt_PropertyWindow Then 'if Properties window has focus
108             If MDIChildPropSelected > 0 Then
110                 Call SendMessagebyString(hWndTextBox, WM_SETTEXT, 0, Vb_MDIChild)
                Else
                    'hit test, MDIChild is the nr 26 by Alphabetic
112                 If LBItemFromPt(hWndProperties, LOWORD(lParam), HIWORD(lParam), False) > 0 Then '= 26 'hit test, MDIChild is the nr 26 by Alphabetic
114                     If MDIChildPropSelected = 0 Then MDIChildPropSelected = 1 'enable the check during the next clicks
                    End If
                End If
            Else
116             If MDIChildPropSelected > 0 Then
118                 MDIChildPropSelected = MDIChildPropSelected + 1
120                 If MDIChildPropSelected > 7 Then MDIChildPropSelected = 0 '7, a number like an other
                End If
            End If
        End If
    End If
End Function




Public Sub UnhookCodePane()
    On Error GoTo eH
100 UnhookCodePaneCombos

102 If hWndCodePane Then
104     SetWindowLong hWndCodePane, IDX_WINDOWPROC, GetProp(hWndCodePane, PropName)
106     RemoveProp hWndCodePane, PropName
    End If

    Exit Sub
eH:
108 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modSubclass.UnhookCodePane " & _
            "at line " & Erl
110 Resume Next
End Sub

Public Sub UnhookCodePaneCombos()
    On Error GoTo eH

100 If hWndCmbRight Then
102     SetWindowLong hWndCmbRight, IDX_WINDOWPROC, GetProp(hWndCmbRight, PropName)
104     RemoveProp hWndCmbRight, PropName
    End If

106 If hWndCmbLeft Then
108     SetWindowLong hWndCmbLeft, IDX_WINDOWPROC, GetProp(hWndCmbLeft, PropName)
110     RemoveProp hWndCmbLeft, PropName
    End If

    Exit Sub

eH:
112 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modSubclass.UnhookCodePaneCombos " & _
            "at line " & Erl
114 Resume Next
End Sub
Public Sub UnhookMainWindow()
'! Delayed error handler
    On Error Resume Next

100 If IDEhwnd Then
102     UnhookMDIClient
104     SetWindowLong IDEhwnd, IDX_WINDOWPROC, GetProp(IDEhwnd, PropName)
106     RemoveProp IDEhwnd, PropName 'remove the property
    End If

End Sub
Public Sub UnhookMDIClient()
    On Error GoTo eH

100 If hWndMDIClient Then
102     UnhookCodePane
104     SetWindowLong hWndMDIClient, IDX_WINDOWPROC, GetProp(hWndMDIClient, PropName)
106     RemoveProp hWndMDIClient, PropName 'remove the property
    End If

    Exit Sub
eH:
108 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modSubclass.UnhookMDIClient " & _
            "at line " & Erl
110 Resume Next
End Sub





