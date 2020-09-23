VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   5160
   ClientLeft      =   2055
   ClientTop       =   2700
   ClientWidth     =   6465
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   6465
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
    ctxHookMenu1.OfficeMenuStyle = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    Set ctxHookMenu1.Font = Command1.Font
End Sub

Private Sub Form_Load()
   Me.WindowState = vbMaximized
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuExit_Click()
    Unload MDIForm1
End Sub

Private Sub mnuNew_Click()
    Dim f As New Form2
    f.Show
End Sub

