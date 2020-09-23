VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKeyboardHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)


'-------------------------------------------------------------------------------
' Class_Initialize
'-------------------------------------------------------------------------------
'  Adds this object's address to the gcolKeyboardEventObjects
'  collection and activates the keyboard hook.
Private Sub Class_Initialize()
  
    gcolKeyboardEventObjects.Add ObjPtr(Me), CStr(ObjPtr(Me))
    InstallKeyboardHook

End Sub

'-------------------------------------------------------------------------------
' Class_Terminate
'-------------------------------------------------------------------------------
'  Removes this object's address from the gcolKeyboardEventObjects
'  collection and deactivates the keyboard hook.
'
Private Sub Class_Terminate()
  
    gcolKeyboardEventObjects.Remove CStr(ObjPtr(Me))
    RemoveKeyboardHook

End Sub

'-------------------------------------------------------------------------------
' RaiseKeyDown
'-------------------------------------------------------------------------------
'Raises the KeyDown event.
Friend Sub RaiseKeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'-------------------------------------------------------------------------------
' RaiseKeyUp
'-------------------------------------------------------------------------------
'Raises the KeyUp event.
Friend Sub RaiseKeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub




