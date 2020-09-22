VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9675
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10815
   _ExtentX        =   19076
   _ExtentY        =   17066
   _Version        =   393216
   Description     =   "Code Browser allows easy navigation through your code."
   DisplayName     =   "Code Browser Add-In"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const Vb_lBracket = "("
Private Const Vb_rBrackett = ")"
Private Const Vb_Bksl = "\"
Private Const Vb_Sep = "|"

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private hModule As Long
Private bHostShutdown As Boolean

'find the combos in the CodePane window
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'Global reference to the running instance of VB
Public VBInstance As VBE

'Events handlers
Public nEvents2 As Events2 'Get IDE status
Private WithEvents eVBBuildEvents As VBBuildEvents 'Get IDE status
Attribute eVBBuildEvents.VB_VarHelpID = -1
Public WithEvents mobjFCEvts As FileControlEvents
Attribute mobjFCEvts.VB_VarHelpID = -1
Public WithEvents eVBProjectEvents As VBProjectsEvents
Attribute eVBProjectEvents.VB_VarHelpID = -1
Public WithEvents eVBComponentsEvents As VBComponentsEvents
Attribute eVBComponentsEvents.VB_VarHelpID = -1
Public WithEvents eVBControlsEvents As VBControlsEvents
Attribute eVBControlsEvents.VB_VarHelpID = -1

'Module-level extensibility objects
Public mWindow As Window
Public mobjDoc As docCodeBrowser

'Dockable add-in needs a GUID
Private DockingAddInGUID As String
Private objAddInInst As Object
Private Const Vb_vbp = ".vbp"
Private Const Vb_x = "x"
Sub UnloadAddin()
    On Error GoTo eH
    Dim oPrjExplorerWindow As Window

100 UnhookMainWindow

    ' show again the ExplorerWindow
102 For Each oPrjExplorerWindow In VBInstance.Windows
104     If oPrjExplorerWindow.Type = vbext_wt_ProjectWindow Then
106         oPrjExplorerWindow.Visible = True
            Exit For
        End If
108 Next oPrjExplorerWindow


    'Destroy the add-in window
110 Set mobjDoc = Nothing
112 Set mWindow = Nothing

    'Destroy the event handlers
114 Set eVBControlsEvents = Nothing
116 Set eVBProjectEvents = Nothing
118 Set eVBComponentsEvents = Nothing
120 Set mobjFCEvts = Nothing
122 Set eVBBuildEvents = Nothing
124 Set nEvents2 = Nothing

    'Destroy the VbIDE interface for the add-in window
126 Set modSubclass.VBInstance = Nothing
128 Set modSubclass.Connect = Nothing

    Exit Sub
eH:
130 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.UnloadAddin " & _
            "at line " & Erl
132 Resume Next
End Sub


Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    'if the user closes the ide or sends Alt F4, what is the same thing
    On Error GoTo eH

100 bClosingSession = True
102 UnloadAddin

    Exit Sub
eH:
104 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.AddinInstance_OnBeginShutdown " & _
            "at line " & Erl
106 Resume Next
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
On Error GoTo eH
DoEvents

If VBInstance.VBProjects.Count Then
    mobjDoc.FS_ClickSelected VBInstance.ActiveVBProject.Name & Vb_Sep
End If

StopRedraw IDEhwnd, False
IsFirstStart = False
    Exit Sub
eH:
104 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.AddinInstance_OnStartupComplete " & _
            "at line " & Erl
106 Resume Next
End Sub

Private Sub AddinInstance_Terminate()
    On Error GoTo eH

100 Set VBInstance = Nothing
102 If bHostShutdown Then FreeLibrary hModule

    Exit Sub
eH:
104 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.AddinInstance_Terminate " & _
            "at line " & Erl
106 Resume Next
End Sub

Private Sub eVBBuildEvents_EnterDesignMode()
    On Error GoTo eH

100 bInRunMode = False
    AfterRun = True
    Exit Sub
eH:
102 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBBuildEvents_EnterDesignMode " & _
            "at line " & Erl
104 Resume Next
End Sub


Private Sub eVBBuildEvents_EnterRunMode()
    On Error GoTo eH

100 bInRunMode = True

    Exit Sub
eH:
102 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBBuildEvents_EnterRunMode " & _
            "at line " & Erl
104 Resume Next
End Sub


Private Sub eVBComponentsEvents_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    'Debug.Print "eVBComponentsEvents_ItemAdded " & VBComponent.Name
    On Error GoTo eH

100 If bInRunMode Then Exit Sub
102 IsCompLoading = True

    Do
104     DoEvents
106 Loop While VBComponent.Type = 0

108 mobjDoc.EV_InsertNewComponent VBComponent.Collection.Parent.Name, VBComponent
110 IsCompLoading = False

    Exit Sub
eH:
112 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBComponentsEvents_ItemAdded " & _
            "at line " & Erl
114 Resume Next
End Sub


Private Sub eVBComponentsEvents_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
    'Debug.Print "eVBComponentsEvents_ItemReloaded " & VBComponent.Name
    On Error GoTo eH

100 If bInRunMode Then Exit Sub
102 mobjDoc.L_ResetLists

    Exit Sub
eH:
104 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBComponentsEvents_ItemReloaded " & _
            "at line " & Erl
106 Resume Next

End Sub


Private Sub eVBComponentsEvents_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    'Debug.Print "eVBComponentsEvents_ItemRemoved " & VBComponent.Name
    On Error GoTo eH
    Dim sKey$
    Const Vb_Sep = "|"

100 If bInRunMode Then Exit Sub
102 If ProcessMsg = True Then ProcessMsg = False

    Do
104     DoEvents
106 Loop While VBComponent.Type = 0

108 Select Case VBComponent.Type
    Case vbext_ct_ResFile, vbext_ct_RelatedDocument
110     sKey = VBComponent.Collection.Parent.Name & Vb_Sep & VBComponent.FileNames(1)
112 Case Else
114     sKey = VBComponent.Collection.Parent.Name & Vb_Sep & VBComponent.Name & Vb_Sep
    End Select

116 mobjDoc.EV_DeleteFromKey sKey
118 DoEvents
120 mobjDoc.L_ResetLists
122 DoEvents
124 If ProcessMsg = False Then ProcessMsg = True

    Exit Sub
eH:
126 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBComponentsEvents_ItemRemoved " & _
            "at line " & Erl
128 Resume Next
End Sub


Private Sub eVBComponentsEvents_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    'Debug.Print "eVBComponentsEvents_ItemRenamed " & OldName
    On Error GoTo eH

    Dim sOldKey$, CompName$, PrjName$, oComp As VBComponent
    Const Vb_Sep = "|"

100 If bInRunMode Then Exit Sub

    Do
102     DoEvents
104 Loop While VBComponent.Type = 0

106 Set oComp = VBComponent
108 CompName = VBComponent.Name
110 PrjName = oComp.Collection.Parent.Name

112 Select Case VBComponent.Type
    Case vbext_ct_ResFile, vbext_ct_RelatedDocument
        'handle in filewrite
        Exit Sub
114 Case Else
116     sOldKey = VBComponent.Collection.Parent.Name & Vb_Sep & OldName & Vb_Sep
    End Select

118 mobjDoc.EV_DeleteFromKey sOldKey
120 DoEvents
122 mobjDoc.EV_InsertNewComponent PrjName, oComp
124 DoEvents
126 mobjDoc.M_RefreshMembers PrjName & Vb_Sep, CompName
128 DoEvents
130 mobjDoc.L_ResetLists

    Exit Sub
eH:
132 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBComponentsEvents_ItemRenamed " & _
            "at line " & Erl
134 Resume Next
End Sub


Private Sub eVBControlsEvents_ItemAdded(ByVal VBControl As VBIDE.VBControl)
On Error GoTo eH

100 If bInRunMode Then Exit Sub
106 mobjDoc.bRefreshing = True
108 LockWindowUpdate GetDesktopWindow
110 bControlChange = True
112 mobjDoc.EV_Timer2
114 DoEvents
116 LockWindowUpdate 0
118 mobjDoc.bRefreshing = False

    Exit Sub
eH:
 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBControlsEvents_ItemAdded " & _
            "at line " & Erl & vbCrLf & Err.Number
 Resume Next
End Sub

Private Sub eVBControlsEvents_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
On Error GoTo eH

100 If bInRunMode Then Exit Sub
106 mobjDoc.bRefreshing = True
108 LockWindowUpdate GetDesktopWindow
110 bControlChange = True
112 mobjDoc.EV_Timer2
114 DoEvents
116 LockWindowUpdate 0
118 mobjDoc.bRefreshing = False

    Exit Sub
eH:
 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBControlsEvents_ItemRemoved " & _
            "at line " & Erl & vbCrLf & Err.Number
 Resume Next
End Sub


Private Sub eVBControlsEvents_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
On Error GoTo eH

100 If bInRunMode Then Exit Sub
106 mobjDoc.bRefreshing = True
108 LockWindowUpdate GetDesktopWindow
110 bControlChange = True
112 mobjDoc.EV_Timer2
114 DoEvents
116 LockWindowUpdate 0
118 mobjDoc.bRefreshing = False
    Exit Sub
eH:
 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBControlsEvents_ItemRenamed " & _
            "at line " & Erl & vbCrLf & Err.Number
 Resume Next
End Sub


Private Sub eVBProjectEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    On Error GoTo eH

100 If bClosingSession Or bInRunMode Then Exit Sub
102 Do While VBProject.Name = vbNullString
104     DoEvents
    Loop
    
106 IsLoading = IsLoading - 1

108 If IsLoading <= 0 Then
110     mobjDoc.FS_Insert_Components

112     mobjDoc.M_RefreshActiveModule
114     mobjDoc.L_ResetListByNewStart
        DoEvents
    End If

    Exit Sub
eH:
116 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBProjectEvents_ItemAdded " & _
            "at line " & Erl & vbCrLf & Err.Number
118 Resume Next
End Sub

Private Sub eVBProjectEvents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    On Error GoTo eH
    Dim sOldKey$
    Const Vb_Sep = "|"

100 If bClosingSession Or bInRunMode Then Exit Sub

    Do
102     DoEvents
104 Loop While VBProject.Name = vbNullString

106 sOldKey = VBProject.Name & Vb_Sep
108 mobjDoc.EV_DeleteFromKey sOldKey
110 DoEvents
112 mobjDoc.L_ResetLists

114 If IsPrjMDI Then
116     IsPrjMDI = False

118     If VBInstance.VBProjects.Count Then
            Dim objPrj As VBProject
            Dim objComp As VBComponent

120         For Each objPrj In VBInstance.VBProjects
122             For Each objComp In objPrj.VBComponents
124                 If objComp.Type = vbext_ct_VBMDIForm Then
126                     IsPrjMDI = True
                        Exit Sub
                    End If
128             Next objComp
130         Next objPrj
        End If
    End If

    Exit Sub
eH:
132 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBProjectEvents_ItemRemoved " & _
            "at line " & Erl & vbCrLf & Err.Number
134 Resume Next
End Sub

Private Sub eVBProjectEvents_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    On Error GoTo eH
    Dim sOldKey$
    Const Vb_Sep = "|"

100 If bClosingSession Or bInRunMode Then Exit Sub

    Do
102     DoEvents
104 Loop While VBProject.Name = vbNullString

    'by starting a new session OldName is vbNullString
106 If OldName = vbNullString Then
108     If k > 72000 Then k = 0 'reset the Dummy KeyMaker
        Exit Sub
    End If

110 sOldKey = OldName & Vb_Sep
112 mobjDoc.EV_DeleteFromKey sOldKey
114 mobjDoc.L_ResetListByNewStart
116 mobjDoc.Tv_Insert_Project VBProject

    Exit Sub
eH:
118 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.eVBProjectEvents_ItemRenamed " & _
            "at line " & Erl & vbCrLf & Err.Number
120 Resume Next
End Sub

Private Sub mobjFCEvts_AfterAddFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
    On Error GoTo eH
100 If bInRunMode Then Exit Sub

    Dim NewNodeText$
    Const Vb_lBracket = "("
    Const Vb_rBrackett = ")"
    Const Vb_Sep = "|"
    Const Vb_Backslash = "\"

102 NewNodeText = Vb_lBracket & Right$(FileName, LenB(FileName) / 2 - InStrRev(FileName, Vb_Backslash)) & Vb_rBrackett

104 Select Case FileType
    Case 6 '.res file
106     mobjDoc.Ev_ResClean VBProject.Name & Vb_Sep, NewNodeText
    End Select

    Exit Sub
eH:
108 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.docCodeBrowser.Ev_ResClean " & _
            "at line " & Erl
110 Resume Next
End Sub

Private Sub mobjFCEvts_AfterChangeFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal NewName As String, ByVal OldName As String)
    'MsgBox "AfterChangeFileName, VBProject: " & VBProject.Name & _
     ", FileType: " & str(FileType) & ", NewName: " & NewName & _
     ", OldName: " & OldName
    On Error GoTo eH  'not every comp. has a name prop.
100 If bInRunMode Then Exit Sub

    Dim sNewText$, sKey$, OldFileText$
    Dim objPrj As VBProject

102 For Each objPrj In VBInstance.VBProjects
104     If objPrj.FileName = NewName Then
106         sKey = _
                    objPrj.Name & _
                    Vb_Sep
108         sNewText = _
                    objPrj.Name & _
                    Vb_lBracket & _
                    Right$(NewName, Len(NewName) - InStrRev(NewName, Vb_Bksl)) & _
                    Vb_rBrackett
110         mobjDoc.EV_NodeNewText sKey, sNewText
            Exit Sub
        End If
112 Next objPrj

114 OldFileText = Vb_lBracket & _
            Right$(OldName, Len(OldName) - InStrRev(OldName, Vb_Bksl)) & _
            Vb_rBrackett
116 sKey = mobjDoc.H_SearchNodeText(VBProject.Name & Vb_Sep, OldFileText)
118 sNewText = mobjDoc.H_NodeKeyToCodeModName(sKey)

120 sNewText = sNewText & _
            Vb_lBracket & _
            Right$(NewName, Len(NewName) - InStrRev(NewName, Vb_Bksl)) & _
            Vb_rBrackett
122 mobjDoc.EV_NodeNewText sKey, sNewText

    Exit Sub
eH:
124 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.docCodeBrowser.H_ShowDesigner " & _
            "at line " & Erl
126 Resume Next
End Sub

Sub Show()
    On Error GoTo eH

100 If Not VBInstance Is Nothing Then
        'Variablen im Userdokument setzen
102     Set mobjDoc.VBInstance = VBInstance
104     Set mobjDoc.Connect = Me

        'Fenster anzeigen
106     mWindow.Visible = True
    Else
108     MsgBox "no active VB Instance resolved"
    End If

    Exit Sub
eH:
110 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.Show " & _
            "at line " & Erl
112 Resume Next
End Sub


Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    'This event fires when the user loads the add-in from the Add-In Manager
    On Error GoTo eH

    'Retain handle to the current instance of Visual Basic for later use
100 Set VBInstance = Application
102 Set eVBProjectEvents = VBInstance.Events.VBProjectsEvents
104 Set eVBComponentsEvents = VBInstance.Events.VBComponentsEvents(Nothing)
106 Set eVBControlsEvents = VBInstance.Events.VBControlsEvents(Nothing, Nothing)
108 Set mobjFCEvts = VBInstance.Events.FileControlEvents(Nothing)
110 Set nEvents2 = VBInstance.Events
112 Set eVBBuildEvents = nEvents2.VBBuildEvents
114 Set modSubclass.VBInstance = VBInstance
116 Set modSubclass.Connect = Connect
IsFirstStart = True
    'Relevant hWnds
118 IDEhwnd = VBInstance.MainWindow.hWnd 'Not listed but endorsed!
120 hWndMDIClient = FindWindowEx(VBInstance.MainWindow.hWnd, 0, "MDIClient", vbNullString)
122 hModule = GetModuleHandle("CodeBrowser.dll")

124 If GetSetting(App.Title, "Settings", "DockingAddInGUID", "0") = "0" Then
        'freie GUID ermittel, wenn noch keine vorhanden
126     DockingAddInGUID = GUIDGen
128     SaveSetting App.Title, "Settings", "DockingAddInGUID", DockingAddInGUID
    Else
        'GUID laden
130     DockingAddInGUID = GetSetting(App.Title, "Settings", "DockingAddInGUID", "0")
    End If

132 Set objAddInInst = AddInInst

    'Convert the ActiveX document into a dockable tool window in the VB IDE
134 Set mWindow = VBInstance.Windows.CreateToolWindow(objAddInInst, "CodeBrowser.docCodeBrowser", "Code Browser", DockingAddInGUID, mobjDoc)
136 HookMainWindow

    'build UI
138 Me.Show

    'if we start with the IDE,
    'but we need to have first a loaded project.
140 If ConnectMode = ext_cm_AfterStartup Then
142     If VBInstance.VBProjects.Count Then
144         mobjDoc.FS_Insert_Components
146         mobjDoc.M_RefreshActiveModule
148         mobjDoc.L_ResetListByNewStart
        End If
    End If
    Exit Sub

eH:
150 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.AddinInstance_OnConnection " & _
            "at line " & Erl & vbCrLf & _
            Err.Number
152 Resume Next
End Sub
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    'This event fires when the user explicitly unloads the add-in
    On Error GoTo eH

100 UnloadAddin
102 Unload Me

104 If RemoveMode = vbext_dm_HostShutdown Then
106     bHostShutdown = True
    End If

    Exit Sub
eH:
108 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.Connect.AddinInstance_OnDisconnection " & _
            "at line " & Erl
110 Resume Next
End Sub










Private Sub mobjFCEvts_BeforeLoadFile(ByVal VBProject As VBIDE.VBProject, FileNames() As String)

    Dim i&

100 If Right$(FileNames(i), 1) = Vb_x Then Exit Sub

102 For i = 0 To UBound(FileNames)
104     If Right$(FileNames(i), 4) = Vb_vbp Then
106         IsLoading = IsLoading + 1
            Exit Sub
        End If
    Next

End Sub






