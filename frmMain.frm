VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Mouse/Keyboard Event Enabled Webbrowser"
   ClientHeight    =   4245
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6300
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  '¾Æ·¡ ¸ÂÃã
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3990
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10610
         EndProperty
      EndProperty
   End
   Begin WBEventEnabled.WebBrowserEx WB 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
      PopupWindowAllowed=   0   'False
   End
   Begin VB.Menu mnuWebPopup 
      Caption         =   "Web"
      Begin VB.Menu mnuWeb 
         Caption         =   "&Go"
         Index           =   10
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Open &New"
         Index           =   11
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Go &Back"
         Index           =   13
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Go &Forward"
         Index           =   14
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Go &Home"
         Index           =   15
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "&Refresh"
         Index           =   16
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "&Insert..."
         Index           =   25
         Begin VB.Menu mnuWebInsert 
            Caption         =   "downloader"
            Index           =   0
         End
         Begin VB.Menu mnuWebInsert 
            Caption         =   "context menu disable"
            Index           =   1
         End
         Begin VB.Menu mnuWebInsert 
            Caption         =   "convert database file"
            Index           =   2
         End
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "-"
         Index           =   29
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Copy &Link"
         Index           =   30
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Copy &Text"
         Index           =   31
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "-"
         Index           =   39
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Show &Address Bar"
         Index           =   90
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Show &Status Bar"
         Index           =   91
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "-"
         Index           =   101
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Enable &Mouse Event"
         Index           =   200
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Enable &Keyboard Event"
         Index           =   201
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "-"
         Index           =   299
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Block Address Bar Context Menu (combo box)"
         Checked         =   -1  'True
         Index           =   300
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    DoFormLoad
End Sub

Public Sub DoFormLoad()
    Me.Width = 11380
    Me.Height = 8000
    
    'Initialize browser
    WB.AddressBox.AddItem "https://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?chkCode3rdPartyReview=on&optSort=DateDescending&cmSearch=Search&blnResetAllVariables=TRUE&txtMaxNumberOfEntriesPerPage=50&txtCriteria=&chkCodeTypeZip=on&chkCodeTypeText=on&chkCodeTypeArticle=on&chkCodeDifficulty=1%2C+2%2C+3%2C+4&lngWId=1"
    mnuWeb(90).Checked = WB.AddressBarVisible
    mnuWeb(91).Checked = WB.StatusBarVisible
    
    With WB.StatusBar 'Add new panels to the statusbar control of the webbrowser control
        Dim oPanel As MSComctlLib.Panel
        Set oPanel = .Panels.Add
        oPanel.Width = 2000
        oPanel.AutoSize = sbrNoAutoSize
        Set oPanel = .Panels.Add
        oPanel.Width = 2000
        oPanel.AutoSize = sbrNoAutoSize
        Set oPanel = Nothing
        'WB.StatusBarVisible = False
    End With

    'Enable the mouse event for the webbrowser control
    WB.MouseEventEnabled = True
    Me.mnuWeb(200).Checked = WB.MouseEventEnabled
    
    'Enable the keyboard event for the webbrowser control
    WB.KeyboardEventEnabled = True
    Me.mnuWeb(201).Checked = WB.KeyboardEventEnabled
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    sbMain.Visible = True
    WB.Move 0, 0, ScaleWidth, ScaleHeight - sbMain.Height
    WB.EnsureResize

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'Ensure to release event objects
    WB.MouseEventEnabled = False
    WB.KeyboardEventEnabled = False
End Sub

Private Sub mnuWeb_Click(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
    
        Case 10: WB.Navigate2ClickedLink 'Just go to the slected link
        
        Case 11: OpenNew 'Open the selected link with a new instance of this form
        
        Case 13: WB.GoBack 'Go back
        
        Case 14: WB.GoForward 'Go forward
        
        Case 15: WB.GoHome 'Go home
        
        Case 16: WB.Refresh 'Refresh
        
        Case 30: WB.CopyLink 'Copy the selected link
        
        Case 31: WB.Copy 'Copy the selected text
        
        Case 90: 'Show/Hide Address Bar of the browser control
            mnuWeb(Index).Checked = Not mnuWeb(Index).Checked
            WB.AddressBarVisible = mnuWeb(Index).Checked
        
        Case 91: 'Show/Hide Status Bar of the browser control
            mnuWeb(Index).Checked = Not mnuWeb(Index).Checked
            WB.StatusBarVisible = mnuWeb(Index).Checked
            
        Case 200: 'Enable/disable mouse event for the browser control
            Me.mnuWeb(Index).Checked = Not Me.mnuWeb(Index).Checked
            WB.MouseEventEnabled = Me.mnuWeb(Index).Checked
            
        Case 201: 'Enable/disable keyboard event for the browser control
            Me.mnuWeb(Index).Checked = Not Me.mnuWeb(Index).Checked
            WB.KeyboardEventEnabled = Me.mnuWeb(Index).Checked
        
        Case 300: 'Enable/disable address bar (combo box) default context menu
            Me.mnuWeb(Index).Checked = Not Me.mnuWeb(Index).Checked
    End Select
    
End Sub

'Open a new link with  a new instance of this form
Private Sub OpenNew()
    Dim frm As frmMain
    Set frm = New frmMain
    
    'Set the tag so that the control knows that it is a new browser instance and
    'does not show the Home page blindly.
    'We should also set the 'Cancel' parameter  in the "WB_IntializeBeforeGoHome" event
    'so that the new browser open a new window with the currenlty given url.
    
    frm.WB.Tag = "OpenNew"
    frm.WB.OpenHomePageAtStart = False
    frm.WB.Webbrowser.RegisterAsBrowser = True 'Required - register as browser
    frm.WB.Navigate WB.ClickedURL 'Open the selected link

    frm.Show 'show the new form
    
    'Inherit mouse/keyboard-event enable status
    frm.WB.MouseEventEnabled = Me.WB.MouseEventEnabled
    frm.WB.KeyboardEventEnabled = Me.WB.KeyboardEventEnabled
End Sub

Private Sub mnuWebInsert_Click(Index As Integer)
    'Insert the selected text (menu caption) to the web browser field.
    
    WB.Webbrowser.SetFocus
    Clipboard.SetText mnuWebInsert(Index).Caption
    SendKeys "^v", True
    
End Sub

'-------------------------------------------------------------------------------
' Browser control - UserControl- wide mouse events
' these events fire before individual events for webbrowser,
' address box (combo box), Go button and status bar fire.
'-------------------------------------------------------------------------------
Private Sub WB_UserControlMouseDown(ByVal Control As Object, ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
'
End Sub
Private Sub WB_UserControlMouseMove(ByVal Control As Object, ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    
    'We receives the UserControl-wide mouse event messages.
    'We can handle messages in these "UserControlMouseMove, UserControlMouseDown, ..." events
    'instead using the individual event for webbrowser, address box (combo box), Go button and status bar.
    
    On Error Resume Next
    'hWnd can be identfied with
    'WB.hWnd, WB.hWndAddressBar, WB.hWndBrowser, WB.hWndGoButton, and  WB.hWndStatusBar parameters
    
    WB.StatusBarPanel(2).Text = "x=" & X & " y=" & Y
    WB.StatusBarPanel(3).Text = Me.Name & "." & WB.Name & "." & Control.Name
    'WB.StatusBarPanel(3).Text = Me.Name & "." & WB.hWnd & "." & Control.hWnd
End Sub

Private Sub WB_UserControlMouseUp(ByVal Control As Object, ByVal HWnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
'
End Sub

'-------------------------------------------------------------------------------
' Browser control - Address bar control (combo box)
'-------------------------------------------------------------------------------
Private Sub WB_AddressBarContextMenu(Cancel As Boolean)
    'To disable the combo url box (addressBar)'s context menu
    'Set the Cancel parameter to True.
    
    Cancel = Me.mnuWeb(300).Checked
    If Cancel Then
        MsgBox "Default context menu of the combo box is disabled." & vbCrLf & _
                        "You can use your own custom menu hear."
    End If
    
End Sub
Private Sub WB_AddressBarKeyDown(KeyCode As Integer, Shift As Integer)
'
End Sub

Private Sub WB_AddressBarKeyUp(KeyCode As Integer, Shift As Integer)
'
End Sub

'-------------------------------------------------------------------------------
' Browser control - WebBrowser Keyboard Event
'-------------------------------------------------------------------------------
Private Sub WB_WebBrowserKeyDown(KeyCode As Integer, Shift As Integer)
    
    StatusBar = "Browser KeyDown:: KeyCode=" & KeyCode & " Shift=" & Shift
    
    Select Case Shift
    Case vbAltMask
        Select Case KeyCode
        Case vbKeyC:
        End Select
     
     'To cancel a keystroke, just set the keycode to 0 as in the standard VB way
    Case vbCtrlMask
        Select Case KeyCode
        Case vbKeyZ: WB.GoBack: KeyCode = 0
        Case vbKeyX: WB.GoForward: KeyCode = 0
        Case vbKeyG: WB.Navigate2ClickedLink: KeyCode = 0 'Go
        Case vbKeyN: Call OpenNew: KeyCode = 0 'Open with a new window
        Case vbKeyB: WB.GoBack: KeyCode = 0  'Go bacjk
        Case vbKeyF: WB.GoForward: KeyCode = 0  'Go forward
        Case vbKeyL: WB.CopyLink: KeyCode = 0  'Copy link
        Case vbKeyC: WB.Copy: KeyCode = 0  'Copy text
        End Select
        
    Case vbShiftMask
    Case vbAltMask + vbCtrlMask
    Case vbShiftMask + vbCtrlMask
    Case vbAltMask + vbShiftMask
    End Select
    
    StatusBar = "Browser KeyDown:: KeyCode=" & KeyCode & " Shift=" & Shift
End Sub

Private Sub WB_WebBrowserKeyUp(KeyCode As Integer, Shift As Integer)
    StatusBar = "Browser KeyUp:: KeyCode=" & KeyCode & " Shift=" & Shift
End Sub

'-------------------------------------------------------------------------------
' Browser control - WebBrowser Mouse Events
'-------------------------------------------------------------------------------
Private Sub WB_WebBrowserMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    StatusBar = "Browser MouseDown:: Button(" & Button & ") Shift=" & Shift & " x=" & X & " y=" & Y & " Cancel=" & Cancel
End Sub

Private Sub WB_WebBrowserMouseDownContextMenu(ByVal IsMouseOnLink As Boolean, ByVal URL As String, ByVal SelText As String, Cancel As Boolean)
    StatusBar = "Browser MouseDown ContextMenu:: IsOnLink=" & _
            IsMouseOnLink & " Link=" & URL & " SelText=" & SelText & " Cancel=" & Cancel

    'Cancel to display the default context menu
    Cancel = True
    'Show the custom context menu
    mnuWeb(10).Enabled = IsMouseOnLink 'Go
    mnuWeb(11).Enabled = IsMouseOnLink 'Open new
    mnuWeb(25).Enabled = Not IsMouseOnLink 'Inset text to a field
    mnuWeb(30).Enabled = IsMouseOnLink 'Copy link
    mnuWeb(31).Enabled = (Len(SelText) > 0) 'Copy text
    
    Me.PopupMenu Me.mnuWebPopup
    
    StatusBar = "Browser MouseDown ContextMenu:: IsOnLink=" & _
            IsMouseOnLink & " Link=" & URL & " SelText=" & SelText & " Cancel=" & Cancel

End Sub

Private Sub WB_WebBrowserMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    StatusBar = "Browser MouseMove:: Button(" & Button & ") Shift=" & Shift & " x=" & X & " y=" & Y & " Cancel=" & Cancel
End Sub

Private Sub WB_WebBrowserMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
    StatusBar = "Browser MouseUp:: Button(" & Button & ") Shift=" & Shift & " x=" & X & " y=" & Y & " Cancel=" & Cancel
End Sub

'-------------------------------------------------------------------------------
' Browser control - WebBrowser Navigation Event
'-------------------------------------------------------------------------------
Private Sub WB_BeforeNewWindow2(ByVal URL As String, NewBrowser As Object, Cancel As Boolean)
    'Now window is started to be open"
    
    'Create new instance of this form
    Dim frm As frmMain
    Set frm = New frmMain
    
    'Set the tag so that the control knows that it is a new browser instance and
    'does not show the Home page blindly.
    'We should also set the 'Cancel' parameter  in the "WB_IntializeBeforeGoHome" event
    'so that the new browser open a new window with the currenlty given url.
    
    frm.WB.Tag = "OpenNew"
    Set NewBrowser = frm.WB.Webbrowser
    frm.Show
    
    'Inherit mouse/keyboard event enable status
    frm.WB.MouseEventEnabled = Me.WB.MouseEventEnabled
    frm.WB.KeyboardEventEnabled = Me.WB.KeyboardEventEnabled
End Sub

'"WB_IntializeBeforeGoHome" event
Private Sub WB_IntializeBeforeGoHome(Cancel As Boolean)
    
    StatusBar = "New Window will open - before GoHome"
    
    'If this form is first open, WB's tag is set to Empty string.
    'Otherwise, we have set the tag to "OpenNew" in the "WB_BeforeNewWindow2" event
    Cancel = (WB.Tag = "OpenNew")
    
End Sub

Private Sub WB_NewDocumentStart(ByVal WebDoc As MSHTML.HTMLDocument, ByVal URL As String, ByVal IsTargetedToFrame As Boolean, ByVal TargetFrameName As String, Cancel As Boolean)
    StatusBar = "New Document Start:: " & URL
End Sub

Private Sub WB_BeforeNavigate2(ByVal WebDoc As MSHTML.HTMLDocument, ByVal URL As String, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    StatusBar = "Before Navigate2:: " & URL
End Sub

Private Sub WB_NavigateComplete2(ByVal WebDoc As MSHTML.HTMLDocument, ByVal URL As String)
    StatusBar = "Navigate Complete2:: " & URL
End Sub

Private Sub WB_DocumentComplete(ByVal WebDoc As MSHTML.HTMLDocument, ByVal URL As String)
    StatusBar = "Document Complete:: " & URL
End Sub

Private Sub WB_NewDocumentComplete(ByVal WebDoc As MSHTML.HTMLDocument, ByVal URL As String, ByVal IsTargetedToFrame As Boolean, ByVal TargetFrameName As String)
    StatusBar = "New Document Complete:: " & URL
    Me.Caption = WB.LocationTitle
End Sub

'-------------------------------------------------------------------------------
' Browser control - StatusBar Mouse Event
'-------------------------------------------------------------------------------
Private Sub WB_StatusBarMouseMove(Panel As MSComctlLib.Panel, Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar = "StatusBar MouseMove::Panel(" & Panel.Index & ") Button(" & Button & ") Shift=" & Shift & " x=" & X & " y=" & Y
End Sub

Private Sub WB_StatusBarMouseDown(Panel As MSComctlLib.Panel, Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar = "StatusBar MouseDown::Panel(" & Panel.Index & ") Button(" & Button & ") Shift=" & Shift & " x=" & X & " y=" & Y
End Sub

Private Sub WB_StatusBarMouseUp(Panel As MSComctlLib.Panel, Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar = "StatusBar MouseUp::Panel(" & Panel.Index & ") Button(" & Button & ") Shift=" & Shift & " x=" & X & " y=" & Y
End Sub

Private Sub WB_StatusBarPanelClick(Panel As MSComctlLib.Panel)
    StatusBar = "StatusBar Click::Panel(" & Panel.Index & ")"
End Sub

Private Sub WB_StatusBarPanelDblClick(Panel As MSComctlLib.Panel)
    StatusBar = "StatusBar DblClick::Panel(" & Panel.Index & ")"
End Sub

'-------------------------------------------------------------------------------
' StatusText/Title Change Event
'-------------------------------------------------------------------------------
Public Property Let StatusBar(strText As String)
    Me.sbMain.Panels(1).Text = strText
End Property

Private Sub WB_StatusTextChange(ByVal Text As String)
'
End Sub

Private Sub WB_TitleChange(ByVal Text As String)
    Me.Caption = Text
    Me.Caption = WB.LocationTitle
End Sub
