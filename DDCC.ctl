VERSION 5.00
Begin VB.UserControl DDCC 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ControlContainer=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2400
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   2100
   ToolboxBitmap   =   "DDCC.ctx":0000
   Begin VB.Line Ln 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   583.333
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  '2D
      BackColor       =   &H80000004&
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   390
   End
End
Attribute VB_Name = "DDCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'DropDown ControlContainer DDCC   OCX Control

'Property names
Private Const pnExpandedHeight As String = "ExpandedHeight"
Private Const pnTitle          As String = "Title"
Private Const pnTitleAlign     As String = "TitleAlign"
Private Const pnHoverTime      As String = "HoverTime"
Private Const pnBackColor      As String = "BackColor"
Private Const pnTitleBackColor As String = "TitleBackColor"
Private Const pnTitleForeColor As String = "TitleForekColor"
Private Const pnOpenOn         As String = "OpenOn"

'My property variables
Private myExpandedHeight        As Long
Private myCollapsed             As Boolean
Private myTitle                 As String
Private myHoverTime             As Long
Private myOpenOn                As OpenOnBits
Private Const HOVER_DEFAULT     As Long = -1

Private Const vbRunMode         As Boolean = True
Private Const vbDesignMode      As Boolean = False

Private HeightExtra             As Long 'for border

'Events
Public Event Expand()
Public Event Collapse()
Public Event Hover()

'Public Enumerations
Public Enum TitleAlign
    TitleAlignLeft = vbLeftJustify
    TitleAlignRight = vbRightJustify
    TitleAlignCenter = vbCenter
End Enum

Public Enum OpenOnBits
    MouseIn = 1
    Titleclick = 2
End Enum

Public Sub Collapse()

    If Not myCollapsed Then
        Do Until UserControl.Height <= lblTitle.Height + HeightExtra
            UserControl.Height = UserControl.Height - Screen.TwipsPerPixelY
            DoEvents
        Loop
        UserControl.Height = lblTitle.Height + HeightExtra
        myCollapsed = True
        RaiseEvent Collapse
    End If

End Sub

Public Property Get Collapsed() As Boolean

    Collapsed = myCollapsed

End Property

Public Property Get CollapsedHeight() As Long

    CollapsedHeight = lblTitle.Height + HeightExtra

End Property

Public Sub Expand()

    If myCollapsed Then
        RaiseEvent Expand
        If Ambient.UserMode = vbRunMode Then
            StartMouseTracking Me, UserControl.hWnd, HoverTime
        End If
        Do Until UserControl.Height >= myExpandedHeight
            UserControl.Height = UserControl.Height + Screen.TwipsPerPixelY
            DoEvents
        Loop
        UserControl.Height = myExpandedHeight
        myCollapsed = False
    End If

End Sub

Public Property Get Expanded() As Boolean

    Expanded = Not myCollapsed

End Property

Public Property Get ExpandedHeight() As Long
Attribute ExpandedHeight.VB_Description = "Sets / returns the expanded height."

    ExpandedHeight = myExpandedHeight

End Property

Public Property Let ExpandedHeight(ByVal newHeight As Long)

    myExpandedHeight = newHeight
    PropertyChanged pnExpandedHeight

End Property

Public Property Get HoverTime() As Long
Attribute HoverTime.VB_Description = "Sets / returns the hover time, before the Hover Event is fired."

    HoverTime = myHoverTime

End Property

Public Property Let HoverTime(ByVal newHovertime As Long)

    myHoverTime = newHovertime
    PropertyChanged pnHoverTime

End Property

Private Sub lblTitle_Click()

    If myOpenOn = Titleclick Then
        If myCollapsed Then
            Expand
          Else 'MYCOLLAPSED = FALSE/0
            Collapse
        End If
    End If

End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If myOpenOn = MouseIn Then
        Expand
    End If

End Sub

Public Property Let OpenOn(ByVal newOpenOn As OpenOnBits)

    If newOpenOn = Titleclick Or newOpenOn = MouseIn Then
        myOpenOn = newOpenOn
        PropertyChanged pnOpenOn
      Else 'NOT NEWOPENON...
        Err.Raise 380
    End If

End Property

Public Property Get OpenOn() As OpenOnBits

    OpenOn = myOpenOn

End Property

Friend Sub RaiseHoverEvent()

    RaiseEvent Hover

End Sub

Public Property Let Title(newTitle As String)
Attribute Title.VB_Description = "Sets / returns the title caption."

    myTitle = newTitle
    Select Case TitleAlign
      Case TitleAlignLeft
        lblTitle = " " & myTitle
      Case TitleAlignCenter
        lblTitle = myTitle
      Case TitleAlignRight
        lblTitle = myTitle & " "
    End Select
    PropertyChanged pnTitle

End Property

Public Property Get Title() As String

    Title = myTitle

End Property

Public Property Let TitleAlign(ByVal TitleAlignment As TitleAlign)
Attribute TitleAlign.VB_Description = "Sets / returns the title alignment."

    If TitleAlignment = TitleAlignCenter Or TitleAlignment = TitleAlignLeft Or TitleAlignment = TitleAlignRight Then
        lblTitle.Alignment = TitleAlignment
        Title = myTitle
        PropertyChanged pnTitleAlign
      Else 'NOT TITLEALIGNMENT...
        Err.Raise 380
    End If

End Property

Public Property Get TitleAlign() As TitleAlign

    TitleAlign = lblTitle.Alignment

End Property

Public Property Let TitleBackColor(ByVal newColor As OLE_COLOR)
Attribute TitleBackColor.VB_Description = "Sets / returns the title backcolor."

    lblTitle.BackColor = newColor
    PropertyChanged pnTitleBackColor

End Property

Public Property Get TitleBackColor() As OLE_COLOR

    TitleBackColor = lblTitle.BackColor

End Property

Public Property Let TitleForeColor(ByVal newColor As OLE_COLOR)
Attribute TitleForeColor.VB_Description = "Sets / returns the title forecolor."

    lblTitle.ForeColor = newColor
    PropertyChanged pnTitleForeColor

End Property

Public Property Get TitleForeColor() As OLE_COLOR

    TitleForeColor = lblTitle.ForeColor

End Property

Public Property Let DDCCBackColor(ByVal newColor As OLE_COLOR)
Attribute DDCCBackColor.VB_Description = "Sets / returns the dropdown container backcolor."

    BackColor = newColor
    PropertyChanged pnBackColor

End Property

Public Property Get DDCCBackColor() As OLE_COLOR

    DDCCBackColor = BackColor

End Property

Private Sub UserControl_Initialize()

    myExpandedHeight = UserControl.Height

End Sub

Private Sub UserControl_InitProperties()

    myTitle = Ambient.DisplayName
    myOpenOn = MouseIn
    Title = myTitle
    HoverTime = HOVER_DEFAULT
    lblTitle.Alignment = TitleAlignLeft

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Ambient.UserMode = vbRunMode Then
        StartMouseTracking Me, UserControl.hWnd, HoverTime
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Title = .ReadProperty(pnTitle, "")
        TitleAlign = .ReadProperty(pnTitleAlign, TitleAlignLeft)
        myExpandedHeight = .ReadProperty(pnExpandedHeight, UserControl.Height)
        myOpenOn = .ReadProperty(pnOpenOn, MouseIn)
        HoverTime = .ReadProperty(pnHoverTime, HOVER_DEFAULT)
        BackColor = .ReadProperty(pnBackColor, vbButtonFace)
        lblTitle.BackColor = .ReadProperty(pnTitleBackColor, vbMenuBar)
        lblTitle.ForeColor = .ReadProperty(pnTitleForeColor, vbMenuText)
    End With 'PROPBAG

End Sub

Private Sub UserControl_Resize()

    lblTitle.Width = ScaleWidth 'adjust widht of title
    Ln.X2 = ScaleWidth
    If Ambient.UserMode = vbDesignMode Then
        myExpandedHeight = Height
    End If

End Sub

Private Sub UserControl_Show()

    If Ambient.UserMode = vbRunMode Then
        HeightExtra = 4 * Screen.TwipsPerPixelY
        UserControl.Height = lblTitle.Height + HeightExtra
        myCollapsed = True
    End If

End Sub

Private Sub UserControl_Terminate()

    StopMouseTracking

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty pnTitle, myTitle, ""
        .WriteProperty pnTitleAlign, lblTitle.Alignment
        .WriteProperty pnExpandedHeight, myExpandedHeight
        .WriteProperty pnOpenOn, myOpenOn
        .WriteProperty pnHoverTime, HoverTime
        .WriteProperty pnBackColor, BackColor
        .WriteProperty pnTitleBackColor, lblTitle.BackColor
        .WriteProperty pnTitleForeColor, lblTitle.ForeColor
    End With 'PROPBAG

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Feb-02 13:26) 43 + 278 = 321 Lines
