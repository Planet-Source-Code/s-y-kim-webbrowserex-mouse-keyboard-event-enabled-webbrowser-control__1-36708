VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl WebBrowserEx 
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   4035
   ScaleWidth      =   4800
   ToolboxBitmap   =   "WebBrowserEx.ctx":0000
   Begin VB.CommandButton cmdGo 
      Caption         =   "ÀÌµ¿"
      Height          =   300
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboURL 
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3495
      ExtentX         =   6165
      ExtentY         =   3625
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.StatusBar sbMain 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3780
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "WebBrowserEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private WithEvents mobjWebDoc As MSHTML.HTMLDocument
Attribute mobjWebDoc.VB_VarHelpID = -1
Private WithEvents MouseEvent As CMouseHook
Attribute MouseEvent.VB_VarHelpID = -1
Private WithEvents KeyboardEvent As CKeyboardHook
Attribute KeyboardEvent.VB_VarHelpID = -1
Private WithEvents frmTopParent As VB.Form
Attribute frmTopParent.VB_VarHelpID = -1
Private m_Documents As Collection 'HTML Documents collection
Private m_Frames As Collection 'HTML Frames collection
'-------------------------------------------------------------------------------
' Webbrowser naviagtion events
'-------------------------------------------------------------------------------
'
Event IntializeBeforeGoHome(Cancel As Boolean)
Event StatusTextChange(ByVal Text As String)
Event TitleChange(ByVal Text As String)
Event NewDocumentStart(ByVal WebDoc As HTMLDocument, ByVal URL As String, ByVal IsTargetedToFrame As Boolean, ByVal TargetFrameName As String, Cancel As Boolean)
Event NewDocumentComplete(ByVal WebDoc As HTMLDocument, ByVal URL As String, ByVal IsTargetedToFrame As Boolean, ByVal TargetFrameName As String)
Event BeforeNavigate2(ByVal WebDoc As HTMLDocument, ByVal URL As String, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Event NavigateComplete2(ByVal WebDoc As HTMLDocument, ByVal URL As String)
Event DocumentComplete(ByVal WebDoc As HTMLDocument, ByVal URL As String)
Event BeforeNewWindow2(ByVal URL As String, NewBrowser As Object, Cancel As Boolean)

'-------------------------------------------------------------------------------
' User control-wide mouse events
'-------------------------------------------------------------------------------
'
Event UserControlMouseUp(ByVal Control As Object, ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
Event UserControlMouseMove(ByVal Control As Object, ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
Event UserControlMouseDown(ByVal Control As Object, ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)

'-------------------------------------------------------------------------------
' WebBrowser mouse events
'-------------------------------------------------------------------------------
'
Event WebBrowserDblClick(Cancel As Boolean)
Event WebBrowserMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
Event WebBrowserMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
Event WebBrowserMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
Event WebBrowserMouseDownContextMenu(ByVal IsMouseOnLink As Boolean, ByVal URL As String, ByVal SelText As String, Cancel As Boolean)
Event WebBrowserMouseUpContextMenu(ByVal IsMouseOnLink As Boolean, ByVal URL As String, ByVal SelText As String, Cancel As Boolean)

'-------------------------------------------------------------------------------
' WebBrowser keyboard events
'-------------------------------------------------------------------------------
'
Event WebBrowserKeyDown(KeyCode As Integer, Shift As Integer)
Event WebBrowserKeyUp(KeyCode As Integer, Shift As Integer)

'-------------------------------------------------------------------------------
' Go button keyboard events
'-------------------------------------------------------------------------------
'
Event GoButtonKeyDown(KeyCode As Integer, Shift As Integer)
Event GoButtonKeyUp(KeyCode As Integer, Shift As Integer)

'-------------------------------------------------------------------------------
' Address bar keyboard events (combo box)
'-------------------------------------------------------------------------------
'
Event AddressBarContextMenu(Cancel As Boolean)
Event AddressBarKeyDown(KeyCode As Integer, Shift As Integer)
Event AddressBarKeyUp(KeyCode As Integer, Shift As Integer)

'-------------------------------------------------------------------------------
' Statusbar keyboard events
'-------------------------------------------------------------------------------
'
Event StatusBarKeyDown(KeyCode As Integer, Shift As Integer)
Event StatusBarKeyUp(KeyCode As Integer, Shift As Integer)

'-------------------------------------------------------------------------------
' Statusbar mouse events
'-------------------------------------------------------------------------------
'
Event StatusBarPanelClick(Panel As MSComctlLib.Panel)
Event StatusBarPanelDblClick(Panel As MSComctlLib.Panel)
Event StatusBarMouseMove(Panel As MSComctlLib.Panel, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event StatusBarMouseDown(Panel As MSComctlLib.Panel, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event StatusBarMouseUp(Panel As MSComctlLib.Panel, Button As Integer, Shift As Integer, X As Single, Y As Single)


Private mstrStatusText As String 'Webrrowser status text
Private mlngWBHwnd As Long 'Webbrowser handle
Private mHwndComboEdit  As Long 'Combo edit box handle
Private mobjTopParent As VB.Form 'Reference to Top parent form

Private mbRunMode As Boolean 'Run/Develope mode detection variable

Private newX As Single 'Variables that contained the converted x, y coordinates
Private newY As Single

Private mhwndTopParent As Long 'Top parent form handle
Private mstrClickedLinkURL As String 'Click-on url
Private mstrNavigate2URL As String 'Navigation start url
Private mstrNavigatedURL As String 'Navigated url
Private mstrTargetFrameName As String 'Naviation target frame


Const m_def_PopupWindowAllowed = True
Const m_def_OpenHomePageAtStart = True
Const m_def_AddressBarVisible = True
Const m_def_StatusBarVisible = True
Const m_def_MouseEventEnabled = True
Const m_def_KeyboardEventEnabled = True

Private m_AddressBarVisible As Boolean
Private m_PopupWindowAllowed As Boolean
Private m_OpenHomePageAtStart As Boolean
Private m_StatusBarVisible As Boolean
Private m_MouseEventEnabled As Boolean
Private m_KeyboardEventEnabled As Boolean

'-------------------------------------------------------------------------------
' TopParent
'-------------------------------------------------------------------------------
' Get the top parent form of the user control
Public Property Get TopParent() As Object
    
    On Error Resume Next
    
    If mobjTopParent Is Nothing Then
        
        Dim objParent As Object
        
        Set objParent = UserControl.Parent
        Do While Not TypeOf objParent Is VB.Form
            Set objParent = objParent.Parent
        Loop
        Set mobjTopParent = objParent
        Set objParent = Nothing
    
    End If
    
    Set TopParent = mobjTopParent
    
End Property

Private Sub GetTopParentHandle()
    Dim objParent As Object
    
    Set objParent = UserControl.Parent
    Do While Not TypeOf objParent Is VB.Form
        Set objParent = objParent.Parent
    Loop
    mhwndTopParent = objParent.HWnd
    Set frmTopParent = objParent
    Set objParent = Nothing
    
End Sub

Public Property Get TargetFrameName() As String
    TargetFrameName = mstrTargetFrameName
End Property

Public Property Get PopupWindowAllowed() As Boolean
    PopupWindowAllowed = m_PopupWindowAllowed
End Property

Public Property Let PopupWindowAllowed(ByVal New_PopupWindowAllowed As Boolean)
    m_PopupWindowAllowed = New_PopupWindowAllowed
    PropertyChanged "PopupWindowAllowed"
End Property


Private Function GetTopParent(HWnd As Long) As Long
    Dim lngParentHWnd As Long
    lngParentHWnd = HWnd
    Do
        lngParentHWnd = GetParent(lngParentHWnd)
        If lngParentHWnd = 0 Then Exit Do
        GetTopParent = lngParentHWnd
    Loop
End Function

Private Sub cboURL_Click()
    cmdGo_Click
End Sub

Private Sub cboURL_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdGo_Click
    End If
End Sub

Private Sub cmdGo_Click()
    On Error Resume Next
    If Len(cboURL.Text) Then
        WB.Navigate cboURL.Text
    End If
End Sub

Public Property Get GoButton() As Object
    Set GoButton = cmdGo
End Property

Public Property Get AddressBox() As Object
    Set AddressBox = cboURL
End Property

Public Property Get ClickedURL() As String
    ClickedURL = mstrClickedLinkURL
End Property

Public Property Get AddressBarVisible() As Boolean
    AddressBarVisible = m_AddressBarVisible
End Property

Public Property Let AddressBarVisible(ByVal New_AddressBarVisible As Boolean)
    m_AddressBarVisible = New_AddressBarVisible
    cmdGo.Visible = New_AddressBarVisible
    cboURL.Visible = New_AddressBarVisible
    UserControl_Resize
    PropertyChanged "AddressBarVisible"
End Property

Public Property Get StatusBarVisible() As Boolean
    StatusBarVisible = m_StatusBarVisible
End Property

Public Property Let StatusBarVisible(ByVal New_StatusBarVisible As Boolean)
    m_StatusBarVisible = New_StatusBarVisible
    UserControl_Resize
    PropertyChanged "StatusBarVisible"
End Property

Public Sub Navigate2ClickedLink()
    WB.Navigate mstrClickedLinkURL
End Sub

Private Sub frmTopParent_Activate()
    Static bFirstActivated As Boolean
    If Not bFirstActivated Then
        Dim Cancel As Boolean
        RaiseEvent IntializeBeforeGoHome(Cancel)
        If Not Cancel Then Me.GoHome
        '-----------------------------------------------------------------------
        ' HOOK: Mouse & Keyboard Event
        '-----------------------------------------------------------------------
        If Me.MouseEventEnabled Then
            Set MouseEvent = New CMouseHook
        End If
        If Me.KeyboardEventEnabled Then
            Set KeyboardEvent = New CKeyboardHook
        End If
        bFirstActivated = True
    End If
End Sub

'-------------------------------------------------------------------------------
' Keyboard event
'-------------------------------------------------------------------------------
'
Private Sub KeyboardEvent_KeyDown(KeyCode As Integer, Shift As Integer)
    If TopParent.ActiveControl Is Extender Then
        On Error Resume Next
        
        Dim lngTemp As Long
        lngTemp = UserControl.ActiveControl.HWnd
        If lngTemp = 0 Then
            
            Select Case APIWin.GetPointHandle
            
            Case sbMain.HWnd
                RaiseEvent StatusBarKeyDown(KeyCode, Shift)
                
            Case Me.hWndBrowser
                RaiseEvent WebBrowserKeyDown(KeyCode, Shift)
            
            End Select
        
        End If
    End If
End Sub

Private Sub KeyboardEvent_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If TopParent.ActiveControl Is Extender Then
        
        On Error Resume Next
        
        Select Case UserControl.ActiveControl.HWnd
        
        Case cmdGo.HWnd
            RaiseEvent GoButtonKeyUp(KeyCode, Shift)
        
        Case cboURL.HWnd
            RaiseEvent AddressBarKeyUp(KeyCode, Shift)
        
        Case sbMain.HWnd
            RaiseEvent StatusBarKeyUp(KeyCode, Shift)
        
        Case Else
            RaiseEvent WebBrowserKeyUp(KeyCode, Shift)
        
        End Select
    
    End If
    
End Sub

'-------------------------------------------------------------------------------
' Mouse events
'-------------------------------------------------------------------------------
'
Private Sub MouseEvent_DblClick(ByVal HWnd As Long, Cancel As Boolean)
    Select Case HWnd
    
    Case Me.hWndBrowser
        RaiseEvent WebBrowserDblClick(Cancel)
    
    End Select
End Sub


Private Sub MouseEvent_MouseDown(ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    
    'Convert app-based position to usercontrol-based position
    Dim newX As Single, newY As Single, Control As Object
    
    If Not IsMouseInMe(HWnd, X, Y, newX, newY, Control) Then Exit Sub
   
    RaiseEvent UserControlMouseDown(Control, HWnd, Button, Shift, newX, newY, Cancel)
    If Cancel Then Exit Sub
        
    Select Case HWnd
    
        Case Me.hWndBrowser
            RaiseEvent WebBrowserMouseDown(Button, Shift, newX, newY, Cancel)
            
            'This is the proper time to raise context menu event
            If (Not Cancel) And Button = vbRightButton Then '
                RaiseEvent WebBrowserMouseDownContextMenu(Me.IsURLClicked, mstrStatusText, Me.SelText, Cancel)
            End If
            
        Case mHwndComboEdit 'if is the combobox(address bar)'s edit menu
            'If right button is clicked
            If Button = vbRightButton Then
                RaiseEvent AddressBarContextMenu(Cancel)
            End If
    
    End Select
    
End Sub

Private Sub MouseEvent_MouseMove(ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    
    'Convert app-based position to usercontrol-based position
    Dim newX As Single, newY As Single, Control As Object
    
    If Not IsMouseInMe(HWnd, X, Y, newX, newY, Control) Then Exit Sub
    RaiseEvent UserControlMouseMove(Control, HWnd, Button, Shift, newX, newY, Cancel)
    If Cancel Then Exit Sub
    
    Select Case HWnd
    
    Case Me.hWndBrowser
        
        RaiseEvent WebBrowserMouseMove(Button, Shift, newX, newY, Cancel)
        
    End Select
    
End Sub

Private Sub MouseEvent_MouseUp(ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    
    Dim newX As Single, newY As Single, Control As Object
    
    If Not IsMouseInMe(HWnd, X, Y, newX, newY, Control) Then Exit Sub
    RaiseEvent UserControlMouseUp(Control, HWnd, Button, Shift, newX, newY, Cancel)
    If Cancel Then Exit Sub
    
    Select Case HWnd
    
    Case Me.hWndBrowser
        
        RaiseEvent WebBrowserMouseUp(Button, Shift, newX, newY, Cancel)
        
        'This is another proper time to raise context menu event
        If (Not Cancel) And Button = vbRightButton Then
            RaiseEvent WebBrowserMouseUpContextMenu(Me.IsURLClicked, mstrStatusText, Me.SelText, Cancel)
        End If
        
    End Select
    
End Sub


'Check whether a position is in the area of user control
Private Function IsMouseInMe(HWnd As Long, X As Single, Y As Single, _
    newX As Single, newY As Single, Control As Object) As Boolean
    
    If mhwndTopParent <> GetTopParent(HWnd) Then
        Exit Function
    End If
    
    CalculateNewPosition X, Y, newX, newY
    
    With UserControl
        IsMouseInMe = newX > 0 And newX < .Width And newY > 0 And newY < .Height
    End With
    
    Select Case HWnd
    Case Me.hWndBrowser: Set Control = WB
    Case mHwndComboEdit: Set Control = cboURL
    Case sbMain.HWnd: Set Control = sbMain
    Case cmdGo.HWnd: Set Control = cmdGo
    End Select
    
End Function

'Calcuate new position starting from the left, top of user control
Private Sub CalculateNewPosition(X As Single, Y As Single, newX As Single, newY As Single)
    Dim lpRect As RECT
    Call GetWindowRect(UserControl.HWnd, lpRect)
    newX = X - lpRect.Left * Screen.TwipsPerPixelX
    newY = Y - lpRect.Top * Screen.TwipsPerPixelY
End Sub

'-------------------------------------------------------------------------------
' Determine whether the mouse is on a url or a url is clicked.
'-------------------------------------------------------------------------------
Public Function IsURLClicked() As Boolean
    If Len(Mid$(Me.StatusText, 1, 3)) > 0 Then
        IsURLClicked = InStr(1, "ftp;file;htt;www", Mid(Me.StatusText, 1, 3)) > 0
        If IsURLClicked Then mstrClickedLinkURL = Me.StatusText
    End If
End Function

'-------------------------------------------------------------------------------
' Status bar events
'-------------------------------------------------------------------------------
'
Private Sub sbMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    RaiseEvent StatusBarMouseDown(GetPanel(X), Button, Shift, X, Y)
End Sub

Private Sub sbMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent StatusBarMouseMove(GetPanel(X), Button, Shift, X, Y)
End Sub

Private Sub sbMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent StatusBarMouseUp(GetPanel(X), Button, Shift, X, Y)
End Sub

Private Sub sbMain_PanelClick(ByVal Panel As MSComctlLib.Panel)
    RaiseEvent StatusBarPanelClick(Panel)
End Sub

Private Sub sbMain_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    RaiseEvent StatusBarPanelDblClick(Panel)
End Sub

'Determine panel using the position
Private Function GetPanel(X As Single) As MSComctlLib.Panel

    On Error Resume Next
    
    If sbMain.Style = sbrNormal Then
        If sbMain.Panels.Count = 1 Then
            Set GetPanel = sbMain.Panels(1)
        Else
            Dim sWidth As Single
            Dim i As Long, iPanel As Long
            For i = 1 To sbMain.Panels.Count
                sWidth = sWidth + sbMain.Panels(i).Width
                If X <= sWidth Then
                    iPanel = i: Exit For
                End If
            Next
            If iPanel > 0 Then
                Set GetPanel = sbMain.Panels(i)
            End If
        End If
    End If
    
End Function

'-------------------------------------------------------------------------------
' Enabling/disabling mouse hooking
'-------------------------------------------------------------------------------
'
Public Property Get MouseEventEnabled() As Boolean
    MouseEventEnabled = m_MouseEventEnabled
    
End Property

Public Property Let MouseEventEnabled(ByVal New_MouseEventEnabled As Boolean)
    m_MouseEventEnabled = New_MouseEventEnabled
    PropertyChanged "MouseEventEnabled"
    
    On Error Resume Next
    If mbRunMode Then
        If New_MouseEventEnabled Then
            If MouseEvent Is Nothing Then
                Set MouseEvent = New CMouseHook
            End If
'            RefreshMouseHooking
        Else
            Set MouseEvent = Nothing
        End If
    End If
End Property

'-------------------------------------------------------------------------------
' Enabling/disabling keyboard hooking
'-------------------------------------------------------------------------------
'
Public Property Get KeyboardEventEnabled() As Boolean
    KeyboardEventEnabled = m_KeyboardEventEnabled
End Property

Public Property Let KeyboardEventEnabled(ByVal New_KeyboardEventEnabled As Boolean)
    m_KeyboardEventEnabled = New_KeyboardEventEnabled
    PropertyChanged "KeyboardEventEnabled"
    
    On Error Resume Next
    If mbRunMode Then
        If New_KeyboardEventEnabled Then
            If KeyboardEvent Is Nothing Then
                Set KeyboardEvent = New CKeyboardHook
            End If
        Else
            Set KeyboardEvent = Nothing
        End If
    End If
End Property

Private Sub UserControl_Initialize()
'    On Error Resume Next
End Sub

'»ç¿ëÀÚ Á¤ÀÇ ÄÁÆ®·Ñ¿¡ ´ëÇÑ ¼Ó¼ºÀ» ÃÊ±âÈ­ÇÕ´Ï´Ù.
Private Sub UserControl_InitProperties()
    m_PopupWindowAllowed = m_def_PopupWindowAllowed
    m_OpenHomePageAtStart = m_def_OpenHomePageAtStart
    m_AddressBarVisible = m_def_AddressBarVisible
    m_StatusBarVisible = m_def_StatusBarVisible
    m_MouseEventEnabled = m_def_MouseEventEnabled
    m_KeyboardEventEnabled = m_def_KeyboardEventEnabled
End Sub

'ÀúÀå¼Ò¿¡¼­ ¼Ó¼º°ªÀ» ·ÎµåÇÕ´Ï´Ù.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    On Error Resume Next
    
    m_PopupWindowAllowed = PropBag.ReadProperty("PopupWindowAllowed", m_def_PopupWindowAllowed)
    m_OpenHomePageAtStart = PropBag.ReadProperty("OpenHomePageAtStart", m_def_OpenHomePageAtStart)
    m_AddressBarVisible = PropBag.ReadProperty("AddressBarVisible", m_def_AddressBarVisible)
    m_StatusBarVisible = PropBag.ReadProperty("StatusBarVisible", m_def_StatusBarVisible)
    mstrNavigate2URL = ""
    m_MouseEventEnabled = PropBag.ReadProperty("MouseEventEnabled", m_def_MouseEventEnabled)
    m_KeyboardEventEnabled = PropBag.ReadProperty("KeyboardEventEnabled", m_def_KeyboardEventEnabled)
    
    If Ambient.UserMode Then
        mbRunMode = True 'Flags that we are in Run mode
        GetTopParentHandle 'Obtain the top parent handle
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("PopupWindowAllowed", m_PopupWindowAllowed, m_def_PopupWindowAllowed)
    Call PropBag.WriteProperty("OpenHomePageAtStart", m_OpenHomePageAtStart, m_def_OpenHomePageAtStart)
    Call PropBag.WriteProperty("AddressBarVisible", m_AddressBarVisible, m_def_AddressBarVisible)
    Call PropBag.WriteProperty("StatusBarVisible", m_StatusBarVisible, m_def_StatusBarVisible)
    Call PropBag.WriteProperty("MouseEventEnabled", m_MouseEventEnabled, m_def_MouseEventEnabled)
    Call PropBag.WriteProperty("KeyboardEventEnabled", m_KeyboardEventEnabled, m_def_KeyboardEventEnabled)
End Sub



'-------------------------------------------------------------------------------
' Resize controls
'-------------------------------------------------------------------------------
Public Sub EnsureResize()
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    WB.ZOrder 0
    
    ResizeStatusBar
    
    If Me.AddressBarVisible Then
        cmdGo.Move Width - cmdGo.Width, 0
        cboURL.Move 0, 0, cmdGo.Left - 50
        If Me.StatusBarVisible Then
            WB.Move 0, cboURL.Height + 50, Width, Height - (cboURL.Height + 50) - sbMain.Height
        Else
            WB.Move 0, cboURL.Height + 50, Width, Height - (cboURL.Height + 50)
        End If
    Else
        If Me.StatusBarVisible Then
            WB.Move 0, 0, Width, Height - sbMain.Height
        Else
            WB.Move 0, 0, Width, Height
        End If
    End If
    
End Sub

'Resize status bar panels
Private Sub ResizeStatusBar()
    
    If sbMain.Visible = False Then Exit Sub
    
    sbMain.Move 0, Height - sbMain.Height, Width, sbMain.Height
    If sbMain.Panels.Count = 1 Then
       sbMain.Panels(1).Width = sbMain.Width
    Else
        Dim sglTotal As Single
        Dim i As Long
        For i = 2 To sbMain.Panels.Count
            If sbMain.Panels(i).Visible Then
                sglTotal = sglTotal + sbMain.Panels(i).Width
            End If
        Next
        sbMain.Panels(1).Width = sbMain.Width - sglTotal
    End If

End Sub

Private Sub UserControl_Terminate()
    
    On Error Resume Next
    
    Set frmTopParent = Nothing
    Set m_Documents = Nothing
    Set m_Frames = Nothing
    
    If mbRunMode Then 'if run mode
        
        If Me.MouseEventEnabled Then
            Set MouseEvent = Nothing
        End If
        If Me.KeyboardEventEnabled Then
            Set KeyboardEvent = Nothing
        End If
        
    End If
    
End Sub

Private Sub WB_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    On Error Resume Next
    
    
    
    Dim objWebDoc As MSHTML.HTMLDocument
    Set objWebDoc = pDisp.Document
    
    
    If Len(mstrNavigate2URL) = 0 Then 'The previous document was completley open and
                                                   'a new page is requested
    
        mstrNavigatedURL = "" 'Set the navigated url to empty
        mstrNavigate2URL = CStr(URL) 'Save the requested url
        
        mstrTargetFrameName = CStr(TargetFrameName) 'Save the target frame
        
        If pDisp Is WB.object Then 'is a top page
        
            Me.AddressBox.Text = mstrNavigate2URL 'Show the url on the address bar
            
            RaiseEvent NewDocumentStart(objWebDoc, mstrNavigate2URL, _
                        False, mstrTargetFrameName, Cancel)
            
        
        Else 'The page is wil be open on a frame, do not show the url on the address bar
        
            RaiseEvent NewDocumentStart(objWebDoc, mstrNavigate2URL, _
                        True, mstrTargetFrameName, Cancel)
        End If
        
        If Cancel Then
            Set objWebDoc = Nothing: Exit Sub
        End If
        
        RaiseEvent BeforeNavigate2(objWebDoc, CStr(URL), Flags, TargetFrameName, PostData, Headers, Cancel)
        
    Else
        
        RaiseEvent BeforeNavigate2(objWebDoc, CStr(URL), Flags, TargetFrameName, PostData, Headers, Cancel)

    End If
    
    
    Set objWebDoc = Nothing
    
End Sub


Private Sub WB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    
    On Error Resume Next
    
    'Get the browser handle to ensure the  proper mouse hooking
    GetBrowserHandle
        
    Dim objWebDoc As MSHTML.HTMLDocument
    Set objWebDoc = pDisp.Document
    
    RaiseEvent DocumentComplete(objWebDoc, CStr(URL))
   
    If CStr(URL) = mstrNavigate2URL Then 'A new page was completely open.
    
        mstrNavigate2URL = ""
        mstrNavigatedURL = CStr(URL)
        cboAddDictinctItem cboURL, mstrNavigatedURL, 0 'Add the url to the addres box
        
        If pDisp Is WB.object Then
            RaiseEvent NewDocumentComplete(objWebDoc, mstrNavigatedURL, False, mstrTargetFrameName)
        Else
            RaiseEvent NewDocumentComplete(objWebDoc, mstrNavigatedURL, True, mstrTargetFrameName)
        End If
    End If
    
End Sub

Private Sub WB_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    
    On Error Resume Next

    Dim objWebDoc As MSHTML.HTMLDocument
    Set objWebDoc = pDisp.Document
    
    RaiseEvent NavigateComplete2(objWebDoc, CStr(URL))
    Set objWebDoc = Nothing
    
End Sub


'Enables the user to select whether to allow  a new window open
Private Sub WB_NewWindow2(ppDisp As Object, Cancel As Boolean)
       
    Dim NewBrowser As SHDocVwCtl.Webbrowser
    
    RaiseEvent BeforeNewWindow2(Me.StatusText, NewBrowser, Cancel)
    
    If Not Cancel Then 'if not canceled
    
        If NewBrowser Is Nothing Then 'if a new browser is not designated
            
            Cancel = Not PopupWindowAllowed 'Decide according to the PopupWindowAllowed setting
        
        Else 'if a new browser is  designated, open the new window with it.
            
            NewBrowser.RegisterAsBrowser = True
            Set ppDisp = NewBrowser.object
            
        End If

    End If
    
    Set NewBrowser = Nothing
    
End Sub

Public Property Get StatusText() As String
    StatusText = mstrStatusText
End Property

Public Property Let StatusText(Text As String)
    mstrStatusText = Text
    If sbMain.Style = sbrNormal Then
        sbMain.Panels(1).Text = Text
    Else
        sbMain.SimpleText = Text
    End If
End Property

Public Property Get StatusBar() As Object
    Set StatusBar = sbMain
End Property

Public Property Get StatusBarPanel(Index) As MSComctlLib.Panel
    On Error Resume Next
    Set StatusBarPanel = sbMain.Panels(Index)
End Property

Private Sub WB_StatusTextChange(ByVal Text As String)
    Me.StatusText = Text
    IsURLClicked 'Save if the text is url.
    RaiseEvent StatusTextChange(Text)
End Sub

Public Property Get Webbrowser() As Object
    Set Webbrowser = WB
End Property

Public Property Get Document() As MSHTML.HTMLDocument
    Set Document = WB.Document
End Property

Public Property Get Busy() As Boolean
    Busy = WB.Busy
End Property

'Navigation
Public Sub GoBack()
    On Error Resume Next
    WB.GoBack
End Sub

Public Sub GoForward()
    On Error Resume Next
    WB.GoForward
End Sub

Public Sub GoHome()
    On Error Resume Next
    mstrNavigate2URL = ""
    WB.GoHome
End Sub

Public Sub GoSearch()
    On Error Resume Next
    WB.GoSearch
End Sub

Public Property Get LocationName() As String
    On Error Resume Next
    LocationName = WB.LocationName
End Property

Public Property Get LocationTitle() As String
    On Error Resume Next
    LocationTitle = WB.Document.Title
End Property

Public Property Get LocationURL() As String
    On Error Resume Next
    LocationURL = WB.LocationURL
End Property

Public Sub Navigate(ByVal URL As String, Optional Flags As Variant, Optional TargetFrameName As Variant, Optional PostData As Variant, Optional Headers As Variant)
    On Error Resume Next
    mstrNavigate2URL = ""
    WB.Navigate URL, Flags, TargetFrameName, PostData, Headers
End Sub

Public Sub Navigate2(URL As Variant, Optional Flags As Variant, Optional TargetFrameName As Variant, Optional PostData As Variant, Optional Headers As Variant)
    On Error Resume Next
    mstrNavigate2URL = ""
    WB.Navigate2 URL, Flags, TargetFrameName, PostData, Headers
End Sub

Public Sub StopNavigation()
    On Error Resume Next
    WB.Stop
End Sub

Public Sub Refresh()
    On Error Resume Next
    mstrNavigate2URL = ""
    WB.Refresh
End Sub

Public Sub Refresh2(Optional Level As Variant)
    On Error Resume Next
    mstrNavigate2URL = ""
    WB.Refresh2 Level
End Sub


Public Property Get SelText() As String
    
    Dim Prev As String
    Prev = Clipboard.GetText
    Copy
    SelText = Clipboard.GetText
    Clipboard.SetText Prev

End Property

Public Function CopyLink()
    
    If IsURLClicked Then
        Clipboard.SetText mstrStatusText
    End If

End Function

Public Function Copy()
    
    Clipboard.Clear
    WB.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER

End Function

Public Function Cut()
    
    Clipboard.Clear
    WB.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DONTPROMPTUSER

End Function

Public Function Paste()
    
    WB.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DONTPROMPTUSER

End Function

Public Function SelectAll()
    
    WB.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER

End Function

Public Sub FindFiles()
    
    On Error Resume Next 'Spits an error without this line
    WB.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DONTPROMPTUSER

End Sub

Public Sub FindText()
    
    'Mimics ctl+f keys being pressed to bring up the
    'find text dialog
    WB.SetFocus
    SendKeys "^f", True

End Sub

Public Sub Save()
    
    On Error GoTo Oops
    If WB.LocationURL = "" Then Exit Sub
    WB.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
Oops:

End Sub

'Return the document collection
Public Property Get Documents() As Collection
    
    If m_Documents Is Nothing Then RefreshDocuments
    Set Documents = m_Documents

End Property

'Refresh the document collection
Public Sub RefreshDocuments()
    
    Set m_Documents = New Collection
    CollectDocuments WB.Document

End Sub

'Return the designated document using index or Url
Public Property Get DocumentX(Index) As MSHTML.HTMLDocument
    
    On Error GoTo Bye
    
    If VarType(Index) = vbString Then
        Dim doc As MSHTML.HTMLDocument
        For Each doc In Documents
            If doc.URL = Index Then
                Set DocumentX = doc: Exit For
            End If
        Next
    Else
        Set DocumentX = m_Documents(Index)
    End If
    
    Exit Property

Bye:
End Property

'collect document objects to the documents collection
Private Function CollectDocuments(WebDocument As MSHTML.HTMLDocument) As Long
    
    On Error GoTo ExitHere
     
    'Count this document
    CollectDocuments = 1
    
    'Add this document to collection
    m_Documents.Add WebDocument
    
    With WebDocument
        If .Frames.length > 0 Then
            Dim i As Long  'Search through frames
            For i = 0 To .Frames.length - 1
                CollectDocuments = CollectDocuments + CollectDocuments(.Frames(i).Document)
            Next
        End If
    End With
    
ExitHere:
    Exit Function
End Function

'Frames collection on the webbrowser
Public Property Get Frames() As Collection
    If m_Frames Is Nothing Then RefreshFrames
    Set Frames = m_Frames
End Property

Public Sub RefreshFrames()
    
    Set m_Frames = New Collection
    CollectFrames WB.Document

End Sub

'Return the designated Frame using index or Url
Public Property Get Frame(Index) As MSHTML.HTMLFrameElement
    
    On Error GoTo Bye
    
    If VarType(Index) = vbString Then
        
        Dim oFrame As MSHTML.HTMLFrameElement
        
        For Each oFrame In Frames
            If oFrame.Name = Index Then
                Set Frame = oFrame: Exit For
            End If
        Next
        Set oFrame = Nothing
    Else
        Set Frame = m_Frames(Index)
    End If
    
    Exit Property
Bye:
    Err.Raise Err.Number
End Property


Private Sub CollectFrames(WebDocument As MSHTML.HTMLDocument)

    With WebDocument
        If .Frames.length > 0 Then
            Dim i As Long  'Search through frames
            For i = 0 To .Frames.length - 1
                m_Frames.Add .Frames(i)
                CollectFrames .Frames(i).Document
            Next
        End If
    End With
    
End Sub

'-------------------------------------------------------------------------------
' All Links  on the web browser
'-------------------------------------------------------------------------------
Public Property Get Links(Optional GetUniqueLinks As Boolean = True) As Collection
    
    Dim colLinks As New Collection
    Dim doc As MSHTML.HTMLDocument
    Dim Link As MSHTML.HTMLLinkElement
    
    On Error Resume Next
    Dim i As Long
    
    For Each doc In Me.Documents
        For Each Link In doc.Links
            If GetUniqueLinks Then
                  AddDistrinctLinkObj colLinks, Link
            Else
                 AddLinkObj colLinks, Link
            End If
        Next
    Next
    
    Set Links = colLinks
    Set colLinks = Nothing
    Set doc = Nothing
    Set Link = Nothing
End Property

Private Sub AddDistrinctLinkObj(colLinks As Collection, Link As Object)
    On Error GoTo Bye
    colLinks.Add Link, Link.href
Bye:
End Sub

Private Sub AddLinkObj(colLinks As Collection, Link As Object)
    On Error GoTo Bye
    colLinks.Add Link
Bye:
End Sub

'Get all of the distinct liinks on the web browser
Public Function GetDistinctLinks(colLink As Collection, colLinkText As Collection) As Long
    Dim doc As MSHTML.HTMLDocument
    Dim Link As MSHTML.HTMLLinkElement
    
    If colLink Is Nothing Then Set colLink = New Collection
    If colLinkText Is Nothing Then Set colLinkText = New Collection

    Dim i As Long
    For Each doc In Me.Documents
        For Each Link In doc.Links
            GetDistinctLinks = GetDistinctLinks + AddLink(colLink, colLinkText, Link)
        Next
    Next
End Function

Private Function AddLink(colLink As Collection, colLinkText As Collection, Link As Object) As Long
    On Error GoTo Bye
    colLink.Add Link.href, Link.href
    colLinkText.Add Link.innerText
    AddLink = 1
Bye:
End Function


'-------------------------------------------------------------------------------
' Images collection on the web browser
'-------------------------------------------------------------------------------
Public Property Get Images(Optional GetUniqueImages As Boolean = True) As Collection
    
    Dim colImages As New Collection
    Dim doc As MSHTML.HTMLDocument
    Dim img As MSHTML.HTMLImg
    
    On Error Resume Next
    
    For Each doc In Me.Documents
        For Each img In doc.Links
            If GetUniqueImages Then
                colImages.Add img, img.href
            Else
                colImages.Add img
            End If
        Next
    Next
    
    Set Images = colImages
    Set colImages = Nothing
    Set doc = Nothing
    Set img = Nothing
End Property


Public Property Get OpenHomePageAtStart() As Boolean
    OpenHomePageAtStart = m_OpenHomePageAtStart
End Property

Public Property Let OpenHomePageAtStart(ByVal New_OpenHomePageAtStart As Boolean)
    m_OpenHomePageAtStart = New_OpenHomePageAtStart
    PropertyChanged "OpenHomePageAtStart"
End Property

Private Sub WB_TitleChange(ByVal Text As String)
    RaiseEvent TitleChange(Text)
End Sub

Private Sub GetBrowserHandle()
    On Error Resume Next
    
    'Static hBrowser As Long
    Dim hBrowser As Long
    
    If Not mbRunMode Then Exit Sub
    'If hBrowser <> 0 Then Exit Sub 'Tracking already started
    If mlngWBHwnd <> 0 Then Exit Sub 'Tracking already started
    hBrowser = GetWebBrowserHWnd(UserControl.HWnd)
    If hBrowser = 0 Then Exit Sub
    mlngWBHwnd = hBrowser
    
    If mHwndComboEdit <> 0 Then Exit Sub
    'get the handle to the edit portion
    'of the combo control
    mHwndComboEdit = APIWin.FindWindowEx(cboURL.HWnd, 0&, vbNullString, vbNullString)
    
End Sub



Public Property Get HWnd()
    HWnd = UserControl.HWnd
End Property

Public Property Get hWndBrowser() As Long
   If mlngWBHwnd = 0 Then RefreshHandles
   hWndBrowser = mlngWBHwnd
End Property

Private Sub RefreshHandles()
    On Error Resume Next
    mlngWBHwnd = GetWebBrowserHWnd(UserControl.HWnd)
    mHwndComboEdit = APIWin.FindWindowEx(cboURL.HWnd, 0&, vbNullString, vbNullString)
End Sub

Public Property Get hWndAddressBar() As Long
   hWndAddressBar = cboURL.HWnd
End Property

Public Property Get hWndGoButton() As Long
   hWndGoButton = cmdGo.HWnd
End Property

Public Property Get hWndStatusBar() As Long
   hWndStatusBar = sbMain.HWnd
End Property
