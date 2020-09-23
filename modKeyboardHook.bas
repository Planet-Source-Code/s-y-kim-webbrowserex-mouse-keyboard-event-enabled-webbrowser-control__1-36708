Attribute VB_Name = "KeyboardHooks"

' Contains all the code required to install a thread- level Keyboard Hook procedure.

' IMPORTANT USEFUL INFO!!!
' NOTE: You can make a thread- level keyboard hook dll using this module
' and CKeyboard Hook class module.

Option Explicit

'GetAsyncKeyState constants
Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12
Private Const KEY_PRESSED = -32768
Private Const KEY_NOT_PRESSED = 0

'Constants for the Code parameter passed to
'the Keyboard hook function
Private Const HC_ACTION = 0
Private Const HC_NOREMOVE = 3

'Keyboard Hook-Type constant
Private Const WH_KEYBOARD = 2

'Installs the hook.
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'Calls the next hook in the hook chain.
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Removes the hook.
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'Returns state information about a key or mouse button.
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'Copies a block of memory.
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal lngBytes As Long)
'Verifies that a pointer is valid (readable).
Private Declare Function IsBadCodePtr Lib "Kernel32" (ByVal lpfn As Long) As Long


'Handle of the keyboard hook.
Private mlngKBHook As Long
Private mlngHookCount As Long

'Collection of KeyboardEvent object addresses
Public gcolKeyboardEventObjects As New Collection

'-------------------------------------------------------------------------------
'  InstallKeyboardHook
'-------------------------------------------------------------------------------
'  Call this procedure to install the KeyboardHook function
'-------------------------------------------------------------------------------
Public Sub InstallKeyboardHook()
    If mlngKBHook = 0 Then
        If mlngHookCount = 0 Then
            mlngKBHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, 0, App.ThreadID)
        End If
    End If
    mlngHookCount = mlngHookCount + 1
End Sub

'-------------------------------------------------------------------------------
' Public Keyboard Hook Procedure
'-------------------------------------------------------------------------------
'  This procedure intercepts all keyboard messages sent to the
'  current thread.
'-------------------------------------------------------------------------------
Private Function KeyboardProc(ByVal code As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim KeyCode As Integer
    Dim Shift As Integer
    
    'Prevent recursion.
    Static blnInKeyboardProc As Boolean
    
    If blnInKeyboardProc Then
        KeyboardProc = CallNextHookEx(mlngKBHook, code, wParam, ByVal lParam)
        Exit Function
    End If
    
    blnInKeyboardProc = True
    
    'If code < 0 then we must pass the message
    'to the next hook procedure in the chain,
    'return it's value and exit the function.
    If code < 0 Then
        KeyboardProc = CallNextHookEx(mlngKBHook, code, wParam, ByVal lParam)
        blnInKeyboardProc = False
        Exit Function
    End If
    
    'HC_ACTION means a key has been pressed or released.
    If code = HC_ACTION Then
    
        KeyCode = wParam
        
        'Get the state of the shift, control, and Alt keys.
        'and convert them to the format that VB uses (vbAltMask + vbCtrlMask + vbShiftMask)
        Shift = GetKeyShift()
        
        'lParam >= 0 means the key was pressed (Key Down)
        If lParam >= 0 Then  'Key Down
        
            'Raise the KeyDown event.
            CallKeyDown KeyCode, Shift
            
            If KeyCode <> 0 Then
                KeyboardProc = CallNextHookEx(mlngKBHook, code, wParam, ByVal lParam)
            Else
                KeyboardProc = 1
            End If
            
            blnInKeyboardProc = False
            Exit Function
            
        Else  'Key Up
            
            'Raise the KeyUp event.
            CallKeyUp KeyCode, Shift
            
            If KeyCode <> 0 Then
                KeyboardProc = CallNextHookEx(mlngKBHook, code, wParam, ByVal lParam)
            Else
                KeyboardProc = 1
            End If
            
            blnInKeyboardProc = False
            Exit Function
        End If
        
    End If 'If code = HC_ACTION Then
    
    KeyboardProc = CallNextHookEx(mlngKBHook, code, wParam, ByVal lParam)
    blnInKeyboardProc = False
End Function

'-------------------------------------------------------------------------------
' Remove Keyboard Hook
'-------------------------------------------------------------------------------
'  Call this procedure to remove the KeyboardHook function
'-------------------------------------------------------------------------------
Public Sub RemoveKeyboardHook()
    If mlngHookCount = 0 Then
        Exit Sub
    End If
    
    If mlngKBHook <> 0 Then
        mlngHookCount = mlngHookCount - 1
        If mlngHookCount = 0 Then
            UnhookWindowsHookEx mlngKBHook
            mlngKBHook = 0
        End If
    End If
End Sub

'-------------------------------------------------------------------------------
'  CallKeyDown method
'-------------------------------------------------------------------------------
'  This procedure raises a KeyDown event on all the
'  KeyboardEvent objects contained in the gcolKeyboardEventObects
'  collection when a key is pressed.
'-------------------------------------------------------------------------------
Private Sub CallKeyDown(KeyCode As Integer, Shift As Integer)
    Dim EventObject As CKeyboardHook
    Dim lngAddress As Variant
    
    For Each lngAddress In gcolKeyboardEventObjects
        If Not IsBadCodePtr(lngAddress) Then
            Set EventObject = GetObjectFromAddress(lngAddress)
            EventObject.RaiseKeyDown KeyCode, Shift
        End If
    Next
    Set EventObject = Nothing
End Sub

'-------------------------------------------------------------------------------
' CallKeyUp method
'-------------------------------------------------------------------------------
'  This procedure raises a KeyUp event on all the
'  KeyboardEvent objects contained in the gcolKeyboardEventObects
'  collection when a key is released.
'-------------------------------------------------------------------------------
Private Sub CallKeyUp(KeyCode As Integer, Shift As Integer)
    Dim EventObject As CKeyboardHook
    Dim lngAddress As Variant
    
    For Each lngAddress In gcolKeyboardEventObjects
        If Not IsBadCodePtr(lngAddress) Then
            Set EventObject = GetObjectFromAddress(lngAddress)
            EventObject.RaiseKeyUp KeyCode, Shift
        End If
    Next
    Set EventObject = Nothing
End Sub

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



