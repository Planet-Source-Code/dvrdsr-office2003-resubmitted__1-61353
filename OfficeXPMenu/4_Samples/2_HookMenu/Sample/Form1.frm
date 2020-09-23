VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E4ECE9&
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   2550
   ClientTop       =   1980
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   5910
   Begin VB.CommandButton Command1 
      Caption         =   "Test Right To Left"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1200
      List            =   "Form1.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   720
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0051
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0163
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0275
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":03E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":04F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0605
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0717
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0829
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":093B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C71
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D83
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Test"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Another test"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Cancel"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   1104
      Left            =   360
      TabIndex        =   1
      Text            =   "There are NO MORE issues with TextBox context menus :-))"
      Top             =   2280
      Width           =   4968
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3960
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto columns"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5295
      Picture         =   "Form1.frx":1137
      Top             =   750
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Right click for context menu"
      Height          =   264
      Left            =   1512
      TabIndex        =   0
      Top             =   336
      Width           =   3036
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
         Begin VB.Menu mnuOpen 
            Caption         =   "&Mail"
            Index           =   0
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "&Note"
            Index           =   1
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "Memo"
            Index           =   2
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "Appointment"
            Index           =   3
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "00"
            Index           =   4
            Begin VB.Menu mnuOpen00 
               Caption         =   "00-00"
               Index           =   0
            End
            Begin VB.Menu mnuOpen00 
               Caption         =   "00-01"
               Index           =   1
            End
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "99"
            Index           =   5
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "88"
            Index           =   6
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "77"
            Index           =   7
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "66"
            Index           =   8
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "55"
            Index           =   9
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "44"
            Index           =   10
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "33"
            Index           =   11
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "22"
            Index           =   12
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "11"
            Index           =   13
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print..."
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print Preview"
         Checked         =   -1  'True
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Exit	Alt+F4"
         Index           =   7
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu mnuEdit 
         Caption         =   "Undo"
         Index           =   0
         Begin VB.Menu mnuUndo 
            Caption         =   "1111"
            Index           =   0
         End
         Begin VB.Menu mnuUndo 
            Caption         =   "2222"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Add menu"
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Icon size"
      Index           =   2
      Begin VB.Menu mnuSize 
         Caption         =   "16x16 px"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuSize 
         Caption         =   "20x20 px"
         Index           =   1
      End
      Begin VB.Menu mnuSize 
         Caption         =   "24x24 px"
         Index           =   2
      End
      Begin VB.Menu mnuSize 
         Caption         =   "28x28 px"
         Index           =   3
      End
      Begin VB.Menu mnuSize 
         Caption         =   "32x32 px"
         Index           =   4
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "popup"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Properties"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "New"
         Index           =   1
         Begin VB.Menu mnuPopupOpen 
            Caption         =   "Mail"
            Index           =   0
         End
         Begin VB.Menu mnuPopupOpen 
            Caption         =   "Appointement"
            Index           =   1
         End
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Cancel"
         Index           =   3
      End
   End
   Begin VB.Menu mnuApperance 
      Caption         =   "Apperance"
      Begin VB.Menu mnuStyle 
         Caption         =   "Style"
         Begin VB.Menu mnuOfficeXP 
            Caption         =   "Office XP"
         End
         Begin VB.Menu mnuOffice2003 
            Caption         =   "Office 2003"
         End
      End
      Begin VB.Menu mnuTheme 
         Caption         =   "Theme"
         Begin VB.Menu mnuBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnuOlive 
            Caption         =   "Olive Green"
         End
         Begin VB.Menu mnuSilver 
            Caption         =   "Silver"
         End
         Begin VB.Menu mnuSepApperance 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSystem 
            Caption         =   "System"
         End
         Begin VB.Menu mnuCustom 
            Caption         =   "Custom"
         End
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Test && Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum UcsFileMenu
    ucsFileNew = 0
    ucsFileSave = 2
    ucsFilePrintPreview = 5
    ucsFileExit = 7
    ucsEditUndo = 0
    ucsEditCut = 2
    ucsEditAddMenu = 6
    ucsEditSep = 7
    ucsMainPopup = 3
End Enum

Private Sub Combo1_Click()
    MDIForm1.ctxHookMenu1.AutoColumn = Me.Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Test Right To Left" Then
    MDIForm1.ctxHookMenu1.RightToLeft = True
    Command1.Caption = "Normal"
Else
    Command1.Caption = "Test Right To Left"
    MDIForm1.ctxHookMenu1.RightToLeft = False
End If
End Sub

Private Sub Form_Activate()
    'Call ctxHookMenu1.SetBitmap(mnuFile(ucsFileNew), Image1.Picture, &HC0C0C0)
End Sub

Private Sub ctxHookMenu1_DrawItemFont(Font As stdole.StdFont, Caption As String, ForeColour As stdole.OLE_COLOR)

    If Caption = "Normal" Then
        Font.Bold = True
        Font.Underline = True
        ForeColour = vbRed
    ElseIf Caption = "Exit" Then
        Font.Bold = True
        Font.Size = Font.Size + 2
        Font.Underline = True
        Font.Italic = True
        ForeColour = vbRed
    ElseIf Caption = "Office 2003" Then
        Font.Bold = False
        Font.Italic = True
        ForeColour = vbBlue
    ElseIf Caption = "Properties" Then
        Font.Bold = True
        Font.Italic = True
        ForeColour = &H40C0&
    Else
        Font.Bold = False
        Font.Underline = False
        ForeColour = vbBlack
    End If

End Sub

Private Sub ctxHookMenu1_DrawItemHoverFont(SelectedFont As stdole.StdFont, Caption As String, SelectedForeColour As stdole.OLE_COLOR, SelectedBackColour As stdole.OLE_COLOR, SelectedBorderColour As stdole.OLE_COLOR)
    If Caption = "&Open" Then
        SelectedFont.Bold = True
        SelectedFont.Italic = True
        SelectedForeColour = vbGreen
        SelectedBackColour = vbYellow
        SelectedBorderColour = vbGreen
    End If
End Sub


'Private Sub Form_Activate()
'    Call ctxHookMenu1.SetBitmap(mnuFile(ucsFileNew), Image1.Picture, &HC0C0C0)
'    Me.Combo1 = "Not Set"
    
'End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuMain(ucsMainPopup), , , , mnuPopup(0)
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub mnuBlue_Click()
    ctxHookMenu1.OfficeMenutheme = Blue
End Sub

Private Sub mnuCustom_Click()
    ctxHookMenu1.OfficeMenutheme = Custom
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Select Case Index
    Case ucsEditUndo
        Call ctxHookMenu1.SetBitmap(mnuFile(ucsFileSave), Image1.Picture, &HC0C0C0)
    Case ucsEditCut
        mnuFile(ucsFileNew).Caption = mnuFile(ucsFileNew).Caption & "1"
    Case ucsEditAddMenu
        mnuEdit(ucsEditSep).Visible = True
        Load mnuEdit(mnuEdit.Count)
        mnuEdit(mnuEdit.UBound).Caption = "Test - " & Timer
    End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
    Case ucsFileNew
        Dim f As New Form1
        f.Show
        mnuFile(ucsFileNew).Checked = Not mnuFile(ucsFileNew).Checked
    Case ucsFileExit
        Unload Me
    Case ucsFilePrintPreview
        mnuFile(ucsFilePrintPreview).Checked = Not mnuFile(ucsFilePrintPreview).Checked
    End Select
End Sub

Private Sub mnuOffice2003_Click()
    ctxHookMenu1.OfficeMenuStyle = Office2003
End Sub

Private Sub mnuOfficeXP_Click()
    ctxHookMenu1.OfficeMenuStyle = OfficeXP
End Sub

Private Sub mnuOlive_Click()
    ctxHookMenu1.OfficeMenutheme = Olive
End Sub

Private Sub mnuOpen_Click(Index As Integer)
    If Index < 4 Then
        MDIForm1.Show
    End If
End Sub

Private Sub mnuSilver_Click()
    ctxHookMenu1.OfficeMenutheme = Silver
End Sub

Private Sub mnuSize_Click(Index As Integer)
    Dim lI As Long
    ctxHookMenu1.BitmapSize = 16 + Index * 4
    For lI = mnuSize.LBound To mnuSize.UBound
        mnuSize(lI).Checked = Index = lI
    Next
End Sub

Private Sub mnuSystem_Click()
    ctxHookMenu1.OfficeMenutheme = System
End Sub
