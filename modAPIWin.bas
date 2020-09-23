Attribute VB_Name = "APIWin"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
  X As Long
  Y As Long
End Type


'APIs : WHERE THE REAL POWER IS
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" ( _
    ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal HWnd As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal HWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetParent Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
        ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal HWnd As Long, lpRect As RECT) As Long

Public Const WM_GETTEXT = &HD                   'Getting text of child window
Public Const WM_GETTEXTLENGTH = &HE
Public Const CB_FINDSTRINGEXACT = &H158

Private mhBrowser As Long
  

Public Sub cboAddDictinctItem(mvaroCBO As VB.ComboBox, Text As String, Optional Index As Long = -1)
    On Error Resume Next
    With mvaroCBO
        If Len(Text) = 0 Then Exit Sub ' Don't add nulls
        If SendMessageAny(.HWnd, CB_FINDSTRINGEXACT, -1, ByVal Text) = -1 Then
            If Index < 0 Then
                .AddItem Text
            Else
                .AddItem Text, Index
            End If
        End If
    End With
End Sub

Public Function GetPointHandle() As Long
    
    Dim CursorPos As POINTAPI
    GetCursorPos CursorPos

    GetPointHandle = WindowFromPoint(CursorPos.X, CursorPos.Y)
End Function

Public Sub DoSleep(Optional WaitMiliSeconds As Long = 70)
    DoEvents
    APIWin.Sleep WaitMiliSeconds
End Sub

Public Function GetClassNameX(ByVal HWnd As Long) As String
        
    Dim ret As Long
    Dim str As String * 255
    ret = GetClassName(HWnd, str, Len(str))
    GetClassNameX = TrimNull(str)
End Function

Public Function TrimNull(ByVal str As String) As String
    Dim Pos As Long
    Pos = InStr(1, str, Chr$(0))
    If Pos Then
        TrimNull = Mid$(str, 1, Pos - 1)
    End If
End Function

Public Function GetCaption(ByVal HWnd As Long) As String
    Dim Textlen As Long
    Dim Text As String

    Textlen = SendMessage(HWnd, WM_GETTEXTLENGTH, 0, 0)
    If Textlen = 0 Then Exit Function
    Textlen = Textlen + 1
    If Textlen > 260 Then Textlen = 260
    Text = Space$(Textlen)
    Textlen = SendMessage(HWnd, WM_GETTEXT, Textlen, ByVal Text)
    'The 'ByVal' keyword is necessary or you'll get an invalid page fault
    'and the app crashes, and takes VB with it.
    GetCaption = Left$(Text, Textlen)
End Function

Public Property Get TopParent(obj As Object) As Object
    Dim objParent As Object
    
    Set objParent = obj.Parent
    If objParent Is Nothing Then
        Exit Property
    End If
    
    Do While Not TypeOf objParent Is VB.Form
        Set objParent = objParent.Parent
    Loop

    Set TopParent = objParent
    Set objParent = Nothing
    
End Property

Public Function GetTopParent(HWnd As Long) As Long

    Dim lngParentHWnd As Long
    lngParentHWnd = HWnd
    Do
        lngParentHWnd = GetParent(lngParentHWnd)
        If lngParentHWnd = 0 Then Exit Do
        GetTopParent = lngParentHWnd
    Loop
End Function

'-------------------------------------------------------------------------------
' COMMENT: Get ListView Header Handle
'-------------------------------------------------------------------------------
Public Function GetWebBrowserHWnd(hWndParent As Long) As Long
    
    Dim ret As Long
    Dim hBrowser As Long
    mhBrowser = 0
    ret = EnumChildWindows(hWndParent, AddressOf WndEnumWebBrowser, hBrowser)
    GetWebBrowserHWnd = mhBrowser
    
End Function

'Callback function
Public Function WndEnumWebBrowser(ByVal HWnd As Long, ByVal lParam As Long) As Long
    
    'Debug.Print GetClassNameX(hWnd)
    
    If GetClassNameX(HWnd) = "Internet Explorer_Server" Then
        mhBrowser = HWnd
        WndEnumWebBrowser = 0
    Else
        WndEnumWebBrowser = 1
    End If
End Function

'-------------------------------------------------------------------------------
' Get ListView Header Handle
'-------------------------------------------------------------------------------
Public Function WndEnumListViewHeader(ByVal HWnd As Long, ByVal lParam As ListView) As Long
    Dim bRet As Long
    Dim myStr As String * 50
    
    bRet = GetClassName(HWnd, myStr, 50)
    If InStr(1, myStr, "msvb_lib_header") > 0 Then
        lParam.Tag = HWnd
        WndEnumListViewHeader = 0
    Else
        WndEnumListViewHeader = 1
    End If

End Function


