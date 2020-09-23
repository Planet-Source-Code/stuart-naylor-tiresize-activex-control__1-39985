Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

     'API Declarations used for subclassing.
      Public Declare Sub CopyMemory _
         Lib "kernel32" Alias "RtlMoveMemory" _
            (pDest As Any, _
            pSrc As Any, _
            ByVal ByteLen As Long)

      Public Declare Function SetWindowLong _
         Lib "user32" Alias "SetWindowLongA" _
            (ByVal hWnd As Long, _
            ByVal nIndex As Long, _
            ByVal dwNewLong As Long) As Long

      Public Declare Function GetWindowLong _
         Lib "user32" Alias "GetWindowLongA" _
            (ByVal hWnd As Long, _
            ByVal nIndex As Long) As Long

      Public Declare Function CallWindowProc _
         Lib "user32" Alias "CallWindowProcA" _
            (ByVal lpPrevWndFunc As Long, _
            ByVal hWnd As Long, _
            ByVal Msg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
            
        Public Declare Function ChildWindowFromPoint _
            Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, _
            ByVal yPoint As Long) As Long


' You can find more o these (lower) in the API Viewer.  Here
' they are used only for resizing the left and right
Public Const HTCLIENT = 1
Public Const HTCAPTION = 2
Public Const HTSYSMENU = 3
Public Const HTGROWBOX = 4
Public Const HTMENU = 5
Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7
Public Const HTMINBUTTON = 8
Public Const HTMAXBUTTON = 9
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTBORDER = 18
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_EXITSIZEMOVE = 562
Public Const WM_SYSCOMMAND = &H112


      'Constants for GetWindowLong() and SetWindowLong() APIs.
        Public Const GWL_WNDPROC = (-4)
        Public Const GWL_USERDATA = (-21)
        Public Const WM_MENUSELECT = &H11F
        Public Const WM_PARENTNOTIFY = &H210
        Public Const WM_MOUSEACTIVATE = &H21
        Public Const WM_NOTIFY As Long = &H4E&
        Public Const WM_HSCROLL = &H114
        Public Const WM_VSCROLL = &H115
        Public Const NM_RCLICK = -5
        Public Const WM_LBUTTONDBLCLK = &H203
        Public Const WM_LBUTTONDOWN = &H201
        Public Const WM_LBUTTONUP = &H202
        Public Const WM_MBUTTONDBLCLK = &H209
        Public Const WM_MBUTTONDOWN = &H207
        Public Const WM_MBUTTONUP = &H208
        Public Const WM_RBUTTONDBLCLK = &H206
        Public Const WM_RBUTTONDOWN = &H204
        Public Const WM_RBUTTONUP = &H205
        Public Const WM_MOUSEFIRST = &H200

    Public Type Subclass
        hWnd As Long
        ProcessId As Long
    End Type
    'Used to hold a reference to the control to call its procedure.
      'NOTE: "UserControl1" is the UserControl.Name Property at
      '      design-time of the .CTL file.
      '      ('As Object' or 'As Control' does not work)
      Dim ctlShadowControl As TiResize

      'Used as a pointer to the UserData section of a window.
      Public mWndSubClass(1) As Subclass
      
      'Used as a pointer to the UserData section of a window.
      Dim ptrObject As Long

      'The address of this function is used for subclassing.
      'Messages will be sent here and then forwarded to the
      'UserControl's WindowProc function. The HWND determines
      'to which control the message is sent.
      Public Function SubWndProc( _
         ByVal hWnd As Long, _
         ByVal Msg As Long, _
         ByVal wParam As Long, _
         ByVal lParam As Long) As Long

         On Error Resume Next

         'Get pointer to the control's VTable from the
         'window's UserData section. The VTable is an internal
         'structure that contains pointers to the methods and
         'properties of the control.
         ptrObject = GetWindowLong(mWndSubClass(0).hWnd, GWL_USERDATA)

         'Copy the memory that points to the VTable of our original
         'control to the shadow copy of the control you use to
         'call the original control's WindowProc Function.
         'This way, when you call the method of the shadow control,
         'you are actually calling the original controls' method.
         CopyMemory ctlShadowControl, ptrObject, 4

         'Call the WindowProc function in the instance of the UserControl.
         SubWndProc = ctlShadowControl.WindowProc(hWnd, Msg, _
            wParam, lParam)

         'Destroy the Shadow Control Copy
         CopyMemory ctlShadowControl, 0&, 4
         Set ctlShadowControl = Nothing
      End Function



 



