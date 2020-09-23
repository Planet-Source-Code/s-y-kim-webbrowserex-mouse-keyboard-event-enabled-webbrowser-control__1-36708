Attribute VB_Name = "MouseHooks"
' Contains all the code required to install a thread- level Mouse Hook procedure.

' IMPORTANT USEFUL INFO!!!
' NOTE: You can make a thread- level mouse hook dll using this module
' and CMouseHook class module.
'
Option Explicit

'GetAsyncKeyState constants
Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_MBUTTON = &H4
Private Const KEY_PRESSED = -32768
Private Const KEY_NOT_PRESSED = 0

'Constant used with GetSystemMetrics to determine if the left and right
'mouse buttons have been swapped.
Private Const SM_SWAPBUTTON = 23

'Constants for the Code parameter passed to
'the Mouse hook function
Private Const HC_ACTION = 0
Private Const HC_NOREMOVE = 3

'Mouse Hook-Type constant
Private Const WH_MOUSE = 7

'Mouse Message Constants
Private Const WM_MOUSEMOVE = &H200
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_NCLBUTTONDBLCLK = &HA3
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const WM_NCMBUTTONDBLCLK = &HA9
Private Const WM_NCMBUTTONDOWN = &HA7
Private Const WM_NCMBUTTONUP = &HA8
Private Const WM_NCRBUTTONDBLCLK = &HA6
Private Const WM_NCRBUTTONDOWN = &HA4
Private Const WM_NCRBUTTONUP = &HA5

'Point structure for mouse coordinates.
Type POINTAPI
    X As Long
    Y As Long
End Type

'Structure passed in the LParam of the MouseProc.
Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    HWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

'Installs the hook.
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'Calls the next hook in the hook chain.
Private Declare Function CallNextMouseHookEx Lib "user32" Alias "CallNextHookEx" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
'Removes the hook.
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'Copies a block of memory.
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal lngBytes As Long)
'Verifies that a pointer is valid (readable).
Private Declare Function IsBadCodePtr Lib "Kernel32" (ByVal lpfn As Long) As Long
'Determines if the left and right mouse buttons have been swapped.
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Returns state information about a key or mouse button.
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'Handle of window over which the mouse button was last pressed.
Private mlngMouseDownHWnd As Long

'Handle of the Mouse hook.
Private mlngMouseHook As Long

'Count of hooks.
Private mlngHookCount As Long

'Collection of MouseHook object addresses
Public gcolMouseEventObjects As New Collection


'-------------------------------------------------------------------------------
'  Install MouseHook
'-------------------------------------------------------------------------------
'  Call this procedure to install the MouseHook function
'-------------------------------------------------------------------------------
Public Sub InstallMouseHook()
    If mlngMouseHook = 0 Then
        If mlngHookCount = 0 Then
            mlngMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0, App.ThreadID)
        End If
    End If
    mlngHookCount = mlngHookCount + 1
End Sub

'-------------------------------------------------------------------------------
' Remove MouseHook
'-------------------------------------------------------------------------------
'  Call this procedure to remove the MouseHook function
Public Sub RemoveMouseHook()
    If mlngHookCount = 0 Then
        Exit Sub
    End If
    
    If mlngMouseHook <> 0 Then
        mlngHookCount = mlngHookCount - 1
        If mlngHookCount = 0 Then
            UnhookWindowsHookEx mlngMouseHook
            mlngMouseHook = 0
        End If
    End If
End Sub

'-------------------------------------------------------------------------------
'  CallDblClick method
'-------------------------------------------------------------------------------
Private Sub CallDblClick(ByRef MH As MOUSEHOOKSTRUCT, ByRef Cancel As Boolean)
    Dim EventObject As CMouseHook
    Dim lngAddress As Variant
    
    For Each lngAddress In gcolMouseEventObjects
        If Not IsBadCodePtr(lngAddress) Then
            Set EventObject = GetObjectFromAddress(lngAddress)
            EventObject.RaiseDblClick MH.HWnd, Cancel
        End If
    Next
    Set EventObject = Nothing
End Sub

'-------------------------------------------------------------------------------
'  CallMouseDown method
'-------------------------------------------------------------------------------
Private Sub CallMouseDown(ByRef MH As MOUSEHOOKSTRUCT, ByRef Cancel As Boolean, Button As Integer)
    Dim EventObject As CMouseHook
    Dim lngAddress As Variant
    
    For Each lngAddress In gcolMouseEventObjects
        If Not IsBadCodePtr(lngAddress) Then
            Set EventObject = GetObjectFromAddress(lngAddress)
            EventObject.RaiseMouseDown MH.HWnd, _
                                        Button, _
                                        GetKeyShift(), _
                                       (MH.pt.X * Screen.TwipsPerPixelX), _
                                       (MH.pt.Y * Screen.TwipsPerPixelY), _
                                       Cancel
        End If
    Next
    Set EventObject = Nothing
End Sub


'-------------------------------------------------------------------------------
'  CallMouseUp method
'-------------------------------------------------------------------------------
Private Sub CallMouseUp(ByRef MH As MOUSEHOOKSTRUCT, ByRef Cancel As Boolean, Button As Integer)
    Dim EventObject As CMouseHook
    Dim lngAddress As Variant
    
    For Each lngAddress In gcolMouseEventObjects
        If Not IsBadCodePtr(lngAddress) Then
            Set EventObject = GetObjectFromAddress(lngAddress)
            EventObject.RaiseMouseUp MH.HWnd, _
                                Button, _
                                GetKeyShift(), _
                                (MH.pt.X * Screen.TwipsPerPixelX), _
                                (MH.pt.Y * Screen.TwipsPerPixelY), _
                                Cancel
        End If
    Next
    Set EventObject = Nothing
End Sub


'-------------------------------------------------------------------------------
'  CallMouseMove method
'-------------------------------------------------------------------------------
Private Sub CallMouseMove(ByRef MH As MOUSEHOOKSTRUCT, ByRef Cancel As Boolean)
    Dim EventObject As CMouseHook
    Dim lngAddress As Variant
    
    For Each lngAddress In gcolMouseEventObjects
        If Not IsBadCodePtr(lngAddress) Then
            Set EventObject = GetObjectFromAddress(lngAddress)
            EventObject.RaiseMouseMove MH.HWnd, _
                                       GetMouseButtons(), _
                                       GetKeyShift(), _
                                       (MH.pt.X * Screen.TwipsPerPixelX), _
                                       (MH.pt.Y * Screen.TwipsPerPixelY), _
                                       Cancel
        End If
    Next
    Set EventObject = Nothing
End Sub

'-------------------------------------------------------------------------------
' Public Mouse Hook Procedure
'-------------------------------------------------------------------------------
'  This procedure intercepts all Mouse messages sent to the
'  current thread.
'-------------------------------------------------------------------------------
Private Function MouseProc(ByVal lngCode As Long, ByVal WP As Long, ByRef LP As MOUSEHOOKSTRUCT) As Long
    Dim Cancel As Boolean
    'Prevent recursion.
    Static blnInMouseProc As Boolean
    
    If blnInMouseProc Then
        MouseProc = CallNextMouseHookEx(mlngMouseHook, lngCode, WP, LP)
        Exit Function
    End If
    
    blnInMouseProc = True
    
    'If lngCode < 0 then we must pass the message
    'to the next hook procedure in the chain,
    'return it's value and exit the function.
    If lngCode < 0 Then
        MouseProc = CallNextMouseHookEx(mlngMouseHook, lngCode, WP, LP)
        blnInMouseProc = False
        Exit Function
    End If
    
    If lngCode = HC_ACTION Then
        
        'Check which message we are processing and raise the appropriate event.
        Select Case WP
        
            Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK, _
                    WM_NCLBUTTONDBLCLK, WM_NCMBUTTONDBLCLK, WM_NCRBUTTONDBLCLK
        
                    'Raise the double-click event.
                    CallDblClick LP, Cancel
                    
        
            Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, _
                    WM_NCLBUTTONDOWN, WM_NCMBUTTONDOWN, WM_NCRBUTTONDOWN
                    
                    'Raise the mouse down event.
                    CallMouseDown LP, Cancel, GetMouseButton(WP)
        
        
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP, _
                    WM_NCLBUTTONUP, WM_NCMBUTTONUP, WM_NCRBUTTONUP
                    
                    'Raise the mouse up event.
                    CallMouseUp LP, Cancel, GetMouseButton(WP)
        
        
            Case WM_MOUSEMOVE, WM_NCMOUSEMOVE
            
                    'Raise the mouse move event.
                    CallMouseMove LP, Cancel
        
        End Select
        
        If Not Cancel Then
            MouseProc = CallNextMouseHookEx(mlngMouseHook, lngCode, WP, LP)
        Else
            MouseProc = 1
        End If
        
        blnInMouseProc = False
        Exit Function
    End If 'If lngCode = HC_ACTION Then
    
    MouseProc = CallNextMouseHookEx(mlngMouseHook, lngCode, WP, LP)
    blnInMouseProc = False

End Function


'-------------------------------------------------------------------------------
' GetKeyShift
'-------------------------------------------------------------------------------
'  This procedure reads the state of the Control, Alt, and Shift
'  keys.  It is called by the KeyboardProc to obtain a VB-Friendly
'  version of this information.
'-------------------------------------------------------------------------------
Private Function GetKeyShift() As Integer
    Dim Shift As Integer
    
    If (GetAsyncKeyState(VK_CONTROL) And KEY_PRESSED) = KEY_PRESSED Then _
        Shift = Shift + vbCtrlMask
        
    If (GetAsyncKeyState(VK_SHIFT) And KEY_PRESSED) = KEY_PRESSED Then _
        Shift = Shift + vbShiftMask
    
    If (GetAsyncKeyState(VK_MENU) And KEY_PRESSED) = KEY_PRESSED Then _
        Shift = Shift + vbAltMask
        
    GetKeyShift = Shift
End Function


'-------------------------------------------------------------------------------
' GetMouseButton/GetMouseButtons
'-------------------------------------------------------------------------------
'  These procedures read the state of the left, right and middle
'  mouse buttons.  It is called by the MouseProc to obtain a VB-Friendly
'  version of this information.
'-------------------------------------------------------------------------------
Private Function GetMouseButton(ButtonMsg As Long) As Integer
    Dim blnSwap As Boolean
    blnSwap = GetSystemMetrics(SM_SWAPBUTTON)
    Dim Button As Integer
    
    Select Case ButtonMsg
    Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_NCLBUTTONDOWN, WM_NCLBUTTONUP
        GetMouseButton = IIf(Not blnSwap, vbLeftButton, vbRightButton)
    Case WM_RBUTTONDOWN, WM_RBUTTONUP, WM_NCRBUTTONDOWN, WM_NCRBUTTONUP
        GetMouseButton = IIf(Not blnSwap, vbRightButton, vbLeftButton)
    Case Else
        GetMouseButton = vbMiddleButton
    End Select
End Function

Private Function GetMouseButtons() As Integer
    Dim Button As Integer
    Dim blnSwap As Boolean
    
    blnSwap = GetSystemMetrics(SM_SWAPBUTTON)
    
    If (GetAsyncKeyState(VK_LBUTTON) And KEY_PRESSED) = KEY_PRESSED Then _
        Button = Button + IIf(Not blnSwap, vbLeftButton, vbRightButton)
        
    If (GetAsyncKeyState(VK_RBUTTON) And KEY_PRESSED) = KEY_PRESSED Then _
        Button = Button + IIf(Not blnSwap, vbRightButton, vbLeftButton)
        
    If (GetAsyncKeyState(VK_MBUTTON) And KEY_PRESSED) = KEY_PRESSED Then _
        Button = vbMiddleButton
    
    GetMouseButtons = Button
End Function


'-------------------------------------------------------------------------------
' GetObjectFromAddress
'-------------------------------------------------------------------------------
'  This procedure returns an object instance from its address
'  and implicitly calls IDispatch->AddRef on the object.
'-------------------------------------------------------------------------------
Public Function GetObjectFromAddress(ByVal lngAddress As Long) As Object
    Dim obj As Object
    If Not lngAddress = 0 Then
        If Not IsBadCodePtr(lngAddress) Then
            CopyMemory obj, lngAddress, 4
            Set GetObjectFromAddress = obj
            CopyMemory obj, 0&, 4
        End If
    End If
End Function



