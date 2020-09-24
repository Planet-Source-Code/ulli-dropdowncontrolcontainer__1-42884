VERSION 5.00
Object = "*\ADDCC.vbp"
Begin VB.Form fDemo 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "DDCC-Demo"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin DDCCVBX.DDCC DDCC2 
      Height          =   2265
      Left            =   3465
      TabIndex        =   8
      Top             =   315
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   3995
      Title           =   "DDCC2"
      TitleAlign      =   0
      ExpandedHeight  =   2265
      OpenOn          =   2
      HoverTime       =   -1
      BackColor       =   -2147483633
      TitleBackColor  =   -2147483644
      TitleForekColor =   -2147483641
   End
   Begin DDCCVBX.DDCC DDCC3 
      Height          =   1380
      Left            =   3465
      TabIndex        =   9
      Top             =   2265
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2434
      Title           =   "DDCC3"
      TitleAlign      =   0
      ExpandedHeight  =   1380
      OpenOn          =   2
      HoverTime       =   -1
      BackColor       =   -2147483633
      TitleBackColor  =   -2147483644
      TitleForekColor =   -2147483641
   End
   Begin DDCCVBX.DDCC DDCC1 
      Height          =   3330
      Left            =   120
      TabIndex        =   0
      Top             =   315
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   5874
      Title           =   "DDCC1"
      TitleAlign      =   0
      ExpandedHeight  =   3330
      OpenOn          =   1
      HoverTime       =   -1
      BackColor       =   -2147483633
      TitleBackColor  =   -2147483644
      TitleForekColor =   -2147483641
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   1380
         TabIndex        =   7
         Top             =   870
         Width           =   1080
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1005
         Left            =   225
         TabIndex        =   4
         Top             =   2025
         Width           =   2250
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   1140
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   270
            Left            =   255
            TabIndex        =   5
            Top             =   270
            Width           =   1140
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   225
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1590
         Width           =   2220
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   225
         TabIndex        =   2
         Top             =   870
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   225
         Left            =   225
         TabIndex        =   1
         Top             =   435
         Width           =   1665
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   2970
      Top             =   2835
   End
End
Attribute VB_Name = "fDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'just playing a little with the events and properties

Private Sub Command1_Click()

    DDCC1.Collapse

End Sub

Private Sub Command2_Click()

    DDCC2.Expand

End Sub

Private Sub Timer1_Timer()

    DDCC2.Collapse

End Sub

Private Sub ddcc1_Hover()

    Beep

End Sub

Private Sub ddcc2_Collapse()

    Command2.Enabled = True
    Timer1.Enabled = False

End Sub

Private Sub ddcc2_Expand()

    Timer1.Enabled = True
    Command2.Enabled = False

End Sub

Private Sub ddcc3_Collapse()

    DDCC3.OpenOn = Titleclick

End Sub

Private Sub ddcc3_Expand()

    DDCC3.OpenOn = MouseIn 'as it is now expanded it in fact closes on MouseOut

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Feb-02 13:26) 2 + 52 = 54 Lines
