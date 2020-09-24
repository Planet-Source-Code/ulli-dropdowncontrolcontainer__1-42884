Attribute VB_Name = "mSubclass"
Option Explicit

'for subclassing
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const IDX_WINDOWPROC        As Long = -4

'for determining whether the mouse cursor has definitely left the control
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type
Private Rc              As RECT

Private Type POINTAPI
    X                   As Long
    Y                   As Long
End Type
Private Pt              As POINTAPI

'for mouse tracking
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TrackMouseEvent) As Long
Private Type TrackMouseEvent
    cbSize              As Long
    dwFlags             As Long
    hWndTrack           As Long
    dwHoverTime         As Long
End Type
Private MouseTrackStruct      As TrackMouseEvent
Private Const WM_MOUSEHOVER   As Long = &H2A1 'windows message identifier
Private Const WM_MOUSELEAVE   As Long = &H2A3 'windows message identifier
Private Const TME_HOVER       As Long = 1 'flag bit
Private Const TME_LEAVE       As Long = 2 'flag bit

'miscelaneous (this module may have to serve several instances of the DDCC)
Private CurrentDDCC     As DDCC
Private TrackingIsOn    As Boolean
Private OldProcPtr      As Long
Private OldhWnd         As Long

Private Function CatchMessages(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case nMsg
      Case WM_MOUSELEAVE
        'note that this msg is also received when the mouse exits the DDCC to
        'enter one of the controls within so we have to check whether the mouse
        'is still within the boundaries of the DDCC
        GetWindowRect hWnd, Rc
        With Rc 'correct values for border
            .Top = .Top + 2
            .Left = .Left + 2
            .Bottom = .Bottom - 2
            .Right = .Right - 2
        End With 'RC
        GetCursorPos Pt
        With Pt
            If PtInRect(Rc, .X, .Y) = 0 Then 'the mouse is ouside the control
                If CurrentDDCC.OpenOn = MouseIn Then
                    CurrentDDCC.Collapse
                End If
            End If
        End With 'PT
        TrackingIsOn = False 'once we got a tracking msg
      Case WM_MOUSEHOVER
        CurrentDDCC.RaiseHoverEvent
        TrackingIsOn = False 'once we got a tracking msg
    End Select
    CatchMessages = CallWindowProc(OldProcPtr, hWnd, nMsg, wParam, lParam) 'call chain to previous proc

End Function

Public Sub StartMouseTracking(DDCC As DDCC, ByVal hWnd As Long, ByVal HoverTime As Long)

    If Not TrackingIsOn Then
        TrackingIsOn = True
        StopMouseTracking 'unhook previous window if any
        OldProcPtr = GetWindowLong(hWnd, IDX_WINDOWPROC) 'hook this window
        OldhWnd = hWnd
        SetWindowLong hWnd, IDX_WINDOWPROC, AddressOf CatchMessages
        Set CurrentDDCC = DDCC 'save reference to current DDCC
        With MouseTrackStruct
            .cbSize = Len(MouseTrackStruct)
            .dwFlags = TME_LEAVE Or TME_HOVER
            .hWndTrack = hWnd
            .dwHoverTime = HoverTime
        End With 'MOUSETRACKSTRUCT
        TrackMouseEvent MouseTrackStruct 'set mouse tracking on
    End If

End Sub

Public Sub StopMouseTracking()

    If OldProcPtr Then 'unhook unless already unhooked
        SetWindowLong OldhWnd, IDX_WINDOWPROC, OldProcPtr
        OldProcPtr = 0
    End If

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Feb-02 13:26) 45 + 62 = 107 Lines
