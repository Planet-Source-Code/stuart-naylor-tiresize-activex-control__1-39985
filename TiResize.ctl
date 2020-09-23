VERSION 5.00
Begin VB.UserControl TiResize 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "TiResize.ctx":0000
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3150
      Top             =   2340
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3915
      Top             =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   1590
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   0
      Left            =   15
      Top             =   420
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   1
      Left            =   435
      Top             =   420
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   2
      Left            =   885
      Top             =   420
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   3
      Left            =   1290
      Top             =   435
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   4
      Left            =   1725
      Top             =   435
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   5
      Left            =   2145
      Top             =   435
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   6
      Left            =   2610
      Top             =   450
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   7
      Left            =   3015
      Top             =   450
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   0
      Left            =   15
      MousePointer    =   8  'Size NW SE
      TabIndex        =   7
      Top             =   0
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   1
      Left            =   435
      MousePointer    =   7  'Size N S
      TabIndex        =   6
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   2
      Left            =   885
      MousePointer    =   6  'Size NE SW
      TabIndex        =   5
      Top             =   15
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   3
      Left            =   1290
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   4
      Left            =   1740
      MousePointer    =   8  'Size NW SE
      TabIndex        =   3
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   5
      Left            =   2145
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   0
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   6
      Left            =   2580
      MousePointer    =   6  'Size NE SW
      TabIndex        =   1
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   7
      Left            =   3015
      MousePointer    =   9  'Size W E
      TabIndex        =   0
      Top             =   30
      Width           =   360
   End
End
Attribute VB_Name = "TiResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This is the first version of the TiResize Control and it is at the 80 /20 stage so it will do for now
'The code has been extracted from various sources from MSDN to PSC so thanks to the shoulders of giants and all that
'If I was going to provide move on from the 80 / 20 then the subclassing routine would be changed for message filtering
'Also the move method needs to be rehashed, as it is it works
'This is the original draft of the code so really it also could do with a spring clean and comments
'The Ocx is activated in a mouse down proc ...
'Set Me.TiResize1.ResizeControl = Text1(Index)
'Me.TiResize1.Visible = True
' Stuart Naylor Industrial Technology

Option Explicit
      'mWndProcOrg holds the original address of the
      'Window Procedure for this window. This is used to
      'route messages to the original procedure after you
      'process them.
      Private mWndProcOrg As Long
      Private IsSubClassed As Boolean
      'Handle (hWnd) of the subclassed window.
      Private mHWndSubClassed As Long
Private mActiveControl As Object
     Private Sub Subclass()
         '-------------------------------------------------------------
         'Initiates the subclassing of this UserControl's window (hwnd).
         'Records the original WinProc of the window in mWndProcOrg.
         'Places a pointer to the object in the window's UserData area.
         '-------------------------------------------------------------

         'Exit if the window is already subclassed.
         If mWndSubClass(0).ProcessId Then Exit Sub

            'Redirect the window's messages from this control's default
            'Window Procedure to the SubWndProc function in your .BAS
            'module and record the address of the previous Window
            'Procedure for this window in mWndProcOrg.
            
            mWndSubClass(0).ProcessId = SetWindowLong(hWnd, GWL_WNDPROC, _
                                         AddressOf SubWndProc)

            'Record your window handle in case SetWindowLong gave you a
            'new one. You will need this handle so that you can unsubclass.
            mWndSubClass(0).hWnd = hWnd

            'Store a pointer to this object in the UserData section of
            'this window that will be used later to get the pointer to
            'the control based on the handle (hwnd) of the window getting
            'the message.
            Call SetWindowLong(hWnd, GWL_USERDATA, ObjPtr(Me))
            
            mWndSubClass(1).ProcessId = SetWindowLong(mActiveControl.hWnd, GWL_WNDPROC, _
                                         AddressOf SubWndProc)
            mWndSubClass(1).hWnd = mActiveControl.hWnd
            IsSubClassed = True
            
            
      End Sub

      Private Sub UnSubClass()
         '-----------------------------------------------------------
         'Unsubclasses this UserControl's window (hwnd), setting the
         'address of the Windows Procedure back to the address it was
         'at before it was subclassed.
         '-----------------------------------------------------------

         'Ensures that you don't try to unsubclass the window when
         'it is not subclassed.
         If mWndSubClass(0).ProcessId = 0 Then Exit Sub

         'Reset the window's function back to the original address.
         SetWindowLong mWndSubClass(0).hWnd, GWL_WNDPROC, mWndSubClass(0).ProcessId
         '0 Indicates that you are no longer subclassed.
         mWndSubClass(0).ProcessId = 0

         'Ensures that you don't try to unsubclass the window when
         'it is not subclassed.
         If mWndSubClass(1).ProcessId = 0 Then Exit Sub

         'Reset the window's function back to the original address.
         SetWindowLong mWndSubClass(1).hWnd, GWL_WNDPROC, mWndSubClass(1).ProcessId
         '0 Indicates that you are no longer subclassed.
         mWndSubClass(1).ProcessId = 0
      End Sub

      Friend Function WindowProc(ByVal hWnd As Long, _
         ByVal uMsg As Long, ByVal wParam As Long, _
         ByVal lParam As Long) As Long
         '--------------------------------------------------------------
         'Process the window's messages that are sent to your UserControl.
         'The WindowProc function is declared as a "Friend" function so
         'that the .BAS module can call the function but the function
         'cannot be seen from outside the UserControl project.
         '--------------------------------------------------------------
        If hWnd = mWndSubClass(1).hWnd Then
        Select Case uMsg
        Case WM_EXITSIZEMOVE
        Debug.Print "WM_EXITSIZEMOVE"
        Timer1.Enabled = True
        Case WM_LBUTTONDOWN
        Debug.Print "WM_LBUTTONDOWN"
        Case WM_LBUTTONUP
        Timer2.Enabled = False
        Timer3.Enabled = True
        Debug.Print "WM_LBUTTONUP"
        Case WM_MOUSEFIRST
        
        Debug.Print "WM_MOUSEFIRST"
        End Select
        Debug.Print uMsg
        End If
        
        
         'Start Demo Code: Changes the color of the UserControl each
         'time the control is clicked in design-time from red to blue
         'or from blue to red.
              'End Demo Code.

            'Forwards the window's messages that came in to the original
            'Window Procedure that handles the messages and returns
            'the result back to the SubWndProc function.
            If hWnd = mWndSubClass(0).hWnd Then
            WindowProc = CallWindowProc(mWndSubClass(0).ProcessId, hWnd, _
                          uMsg, wParam, ByVal lParam)
            End If
            
            If hWnd = mWndSubClass(1).hWnd Then
            WindowProc = CallWindowProc(mWndSubClass(1).ProcessId, hWnd, _
                          uMsg, wParam, ByVal lParam)
            End If
            
      End Function '

Private Sub Timer1_Timer()
Timer1.Enabled = False
CheckSameContainer
MoveUserControl
End Sub
Private Sub Timer2_Timer()
Timer2.Enabled = False
SendMessage mActiveControl.hWnd, WM_SYSCOMMAND, 61458, 0
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
mActiveControl.Visible = True
CheckSameContainer
MoveUserControl
SetControlOnTop
End Sub

Private Sub UserControl_Resize()
PlaceAnchors
End Sub
Public Property Set ResizeControl(ByVal ActiveControl As Object)


    Set mActiveControl = ActiveControl
    PropertyChanged "ActiveControl"

    If ActiveControl Is Nothing Then

    Else

    ReleaseCapture
    If ActiveControl.hWnd <> mWndSubClass(1).hWnd And IsSubClassed = True Then
    UnSubClass
    IsSubClassed = False
    End If

    Subclass
    Timer2.Enabled = True

    

    End If

End Property

Private Sub CheckSameContainer()
'Make sure the usercontrol is in the same container as the bound control
Dim I As Long
Dim OBJ As Object
On Error Resume Next
For I = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hWnd = UserControl.ParentControls.Item(I).hWnd Then
    Set OBJ = UserControl.ParentControls.Item(I)
        Exit For
    End If
Next I
If mActiveControl Is Nothing Then Exit Sub
If OBJ Is Nothing Then Exit Sub
    If mActiveControl.Container <> OBJ.Container Then
        Set OBJ.Container = mActiveControl.Container
    End If
End Sub
Private Sub MoveUserControl()
'find this control and move it on the parent container
Dim I As Long
Dim OBJ As Object
On Error Resume Next
For I = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hWnd = UserControl.ParentControls.Item(I).hWnd Then
    Set OBJ = UserControl.ParentControls.Item(I)
        OBJ.Move mActiveControl.Left - 100, mActiveControl.Top - 100, mActiveControl.Width + 200, mActiveControl.Height + 200
        Exit For
    End If
Next I

End Sub

Private Sub SetControlOnTop()
Dim I As Long
Dim OBJ As Object

On Error Resume Next
For I = 0 To UserControl.ParentControls.Count - 1
    'get this usercontrol from the parent
    If UserControl.hWnd = UserControl.ParentControls.Item(I).hWnd Then
        Set OBJ = UserControl.ParentControls.Item(I)
        OBJ.ZOrder 0
        Exit For
    End If
Next I


End Sub


Private Sub PlaceAnchors()

Label1(0).Move 0, 0, 100, 100
Shape1(0).Move 0, 0, 100, 100
Label1(1).Move (UserControl.Width / 2) - 50, 0, 100, 100
Shape1(1).Move (UserControl.Width / 2) - 50, 0, 100, 100
Label1(2).Move UserControl.Width - 100, 0, 100, 100
Shape1(2).Move UserControl.Width - 100, 0, 100, 100
Label1(3).Move UserControl.Width - 100, (UserControl.Height / 2) - 50, 100, 100
Shape1(3).Move UserControl.Width - 100, (UserControl.Height / 2) - 50, 100, 100
Label1(4).Move UserControl.Width - 100, UserControl.Height - 100, 100, 100
Shape1(4).Move UserControl.Width - 100, UserControl.Height - 100, 100, 100
Label1(5).Move (UserControl.Width / 2) - 50, UserControl.Height - 100, 100, 100
Shape1(5).Move (UserControl.Width / 2) - 50, UserControl.Height - 100, 100, 100
Label1(6).Move 0, UserControl.Height - 100, 100, 100
Shape1(6).Move 0, UserControl.Height - 100, 100, 100
Label1(7).Move 0, (UserControl.Height / 2) - 50, 100, 100
Shape1(7).Move 0, (UserControl.Height / 2) - 50, 100, 100
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wParam As Long

If Button = 1 Then
Select Case Index
Case 0 'NorthWest
wParam = HTTOPLEFT
Case 1 'North
wParam = HTTOP
Case 2 'NorthEast
wParam = HTTOPRIGHT
Case 3 'East
wParam = HTRIGHT
Case 4 'SouthEast
wParam = HTBOTTOMRIGHT
Case 5 'South
wParam = HTBOTTOM
Case 6 'SouthWest
wParam = HTBOTTOMLEFT
Case 7 'West
wParam = HTLEFT
End Select

ReleaseCapture
SendMessage mActiveControl.hWnd, WM_NCLBUTTONDOWN, wParam, 0

End If
End Sub

Private Sub UserControl_Terminate()
UnSubClass
End Sub
