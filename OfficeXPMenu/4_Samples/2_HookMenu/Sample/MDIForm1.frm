VERSION 5.00
Object = "*\A..\..\..\..\..\..\DOWNLO~1\OFFICE~1\OFFICE~1\4_SAMP~1\2_HOOK~1\HookMenu.vbp"
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   3840
      Top             =   1800
      _ExtentX        =   900
      _ExtentY        =   900
      SelectDisabled  =   0   'False
      MenuBackColor   =   16105118
      MenuGradientColor=   16112323
      MenuForeColor   =   -2147483640
      MenuBorderColor =   9841920
      MenuBackSelectColor=   8508412
      MenuForeSelectColor=   -2147483640
      MenuGradientSelectColor=   12775167
      MenuBorderSelectColor=   3693887
      PopupBorderColor=   9841920
      PopupBorderColor=   -2147483640
      PopupBackSelectColor=   8508412
      PopupForeSelectColor=   -2147483640
      PopupBorderSelectColor=   3693887
      SideBarColor    =   15118989
      SideBarGradientColor=   16707040
      CheckBackColor  =   7323903
      CheckForeColor  =   9841920
      CheckBackSelectColor=   4096254
      CheckForeSelectColor=   9841920
      ShadowColor     =   3693887
      OfficeMenuStyle =   1
      OfficeMenuTheme =   0
      MenuBarGradient =   -1  'True
      MenuGradientBehaviour=   1
      BmpCount        =   4
      Mask:1          =   16711935
      Key:1           =   "#mnuNew"
      Bmp:2           =   "MDIForm1.frx":0000
      Key:2           =   "#mnuOpen"
      Bmp:3           =   "MDIForm1.frx":0428
      Mask:3          =   6382695
      Key:3           =   "#mnuSaveProject"
      Bmp:4           =   "MDIForm1.frx":0850
      Key:4           =   "#mnuExit"
      UseSystemFont   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuNew 
         Caption         =   "Office 2003 Menu"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Custom Theme"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuMinus1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Disable Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuForm2 
         Caption         =   "Checkbox Item"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMinus2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveProject 
         Caption         =   "Save Icon Item"
      End
      Begin VB.Menu mnuSaveProjectAs 
         Caption         =   "Normal Item"
      End
      Begin VB.Menu mnuMinus 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditMinus1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPasye 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      NegotiatePosition=   3  'Right
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()
'    MDIForm1.ValidateControls
    'DoEvents
End Sub

Private Sub MDIForm_Load()
    MDIForm1.Width = Screen.Width
    MDIForm1.Top = 0
    MDIForm1.Left = 0
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuForm1_Click()
    Form1.Show
End Sub

Private Sub mnuForm2_Click()
    Dim f As New Form1
    f.Show
End Sub

