Attribute VB_Name = "modGenerateGUID"
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" _
        (pguid As GUID) As Long

Private Declare Function StringFromGUID2 Lib "ole32.dll" _
        (rguid As Any, ByVal lpstrClsId As Long, _
        ByVal cbMax As Long) As Long

Public Function GUIDGen() As String
    On Error GoTo eH
    Dim uGUID As GUID
    Dim sGUID As String
    Dim tGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim RetVal As Long
100 lLen = 40
102 bGUID = String(lLen, 0)
104 CoCreateGuid uGUID
106 RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
108 sGUID = bGUID
110 If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1
112 tGUID = Left$(sGUID, RetVal)
114 GUIDGen = Mid$(tGUID, 2, Len(tGUID) - 2)

    Exit Function
eH:
116 MsgBox Err.Description & vbCrLf & _
            "in CodeBrowser.modGenerateGUID.GUIDGen " & _
            "at line " & Erl
118 Resume Next
End Function





