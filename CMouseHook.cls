VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMouseHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event MouseDown(ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
Event MouseMove(ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
Event MouseUp(ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
Event DblClick(ByVal HWnd As Long, Cancel As Boolean)

'-------------------------------------------------------------------------------
' Class_Initialize
'-------------------------------------------------------------------------------
'  Adds the address of this object to the gcolMouseEventObjects collection
'  and activates the mouse hook.

Private Sub Class_Initialize()
    'Add the address of the object to the collection, using the
    'address as a key.
    gcolMouseEventObjects.Add ObjPtr(Me), CStr(ObjPtr(Me))
    InstallMouseHook
End Sub

'-------------------------------------------------------------------------------
' Class_Terminate
'-------------------------------------------------------------------------------
'  Removes this object from the gcolMouseEventObjects collection
'  and deactivates the keyboard hook.

Private Sub Class_Terminate()
    gcolMouseEventObjects.Remove CStr(ObjPtr(Me))
    RemoveMouseHook
End Sub

'-------------------------------------------------------------------------------
' RaiseDblClick
'-------------------------------------------------------------------------------
'  Raises the DblClick event.

Friend Sub RaiseDblClick(ByVal HWnd As Long, Cancel As Boolean)
    RaiseEvent DblClick(HWnd, Cancel)
End Sub

'-------------------------------------------------------------------------------
' RaiseMouseDown
'-------------------------------------------------------------------------------
'  Raises the MouseDown event.
'
Friend Sub RaiseMouseDown(ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    RaiseEvent MouseDown(HWnd, Button, Shift, X, Y, Cancel)
End Sub

'-------------------------------------------------------------------------------
' RaiseMouseMove
'-------------------------------------------------------------------------------
'  Raises the MouseMove event.
'
Friend Sub RaiseMouseMove(ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    RaiseEvent MouseMove(HWnd, Button, Shift, X, Y, Cancel)
End Sub

'-------------------------------------------------------------------------------
' RaiseMouseUp
'-------------------------------------------------------------------------------
'  Raises the MouseUp event.
'
Friend Sub RaiseMouseUp(ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    RaiseEvent MouseUp(HWnd, Button, Shift, X, Y, Cancel)
End Sub


