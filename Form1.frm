VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "amopad"
   ClientHeight    =   7905
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1320
      Top             =   5640
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   5160
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin Project1.ctrBMPText Text1 
      Height          =   3855
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   6255
      _extentx        =   11033
      _extenty        =   6800
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuSeperator0 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuImportFont 
      Caption         =   "&Import Font"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilename As String

Private Sub Form_Resize()
Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub mnuNew_Click()
Text1.Text = ""
strFilename = ""
Caption = "amopad"
End Sub

Private Sub mnuOpen_Click()
On Error Resume Next

cmd.DialogTitle = "Open File"
cmd.Filter = "(*.txt;*.inf;*.nfo;*.cfg;*.log;*.bat)|*.txt;*.inf;*.nfo;*.cfg;*.log;*.bat|(*.*)|*.*"
cmd.Flags = cdlOFNFileMustExist
cmd.ShowOpen
If Err Then Exit Sub

Text1.loadTextFile cmd.Filename

If Err Then
   MsgBox Err.Description
   Err.Clear
Else
   strFilename = cmd.Filename
   Caption = "amopad - " & cmd.FileTitle
End If
End Sub

Private Sub mnuSave_Click()
If Len(strFilename) = 0 Then
   mnuSaveas_Click
Else
   On Error Resume Next
   Text1.saveTextFile strFilename
   If Err Then
      MsgBox Err.Description
      Err.Clear
   End If
End If
End Sub

Private Sub mnuSaveas_Click()
On Error Resume Next

cmd.DialogTitle = "Save File"
cmd.Filter = "*.*"
cmd.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
cmd.ShowSave
If Err Then Exit Sub

Text1.saveTextFile cmd.Filename
If Err Then
   MsgBox Err.Description
   Err.Clear
Else
   strFilename = cmd.Filename
   Caption = "amopad - " & cmd.FileTitle
End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuImportFont_Click()
On Error Resume Next

cmd.DialogTitle = "Import Font"
cmd.Filter = "(*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif|(*.*)|*.*"
cmd.Flags = cdlOFNFileMustExist
cmd.ShowOpen
If Err Then Exit Sub

Text1.importFont cmd.Filename
If Err Then
   MsgBox Err.Description
   Err.Clear
End If
End Sub

