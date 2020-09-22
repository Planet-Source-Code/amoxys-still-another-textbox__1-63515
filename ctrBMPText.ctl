VERSION 5.00
Begin VB.UserControl ctrBMPText 
   BackColor       =   &H00808080&
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   Begin VB.PictureBox pctHScroll 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   343
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
      Width           =   5175
   End
   Begin VB.PictureBox pctVScroll 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   5160
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox pctFont 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   2940
      Left            =   5880
      Picture         =   "ctrBMPText.ctx":0000
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   112
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.PictureBox pctOut 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      MousePointer    =   3  'I-Cursor
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDelimiter1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "ctrBMPText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long

Dim charWidth As Long, charHeight As Long
Dim lngSX As Long, lngSY As Long, lngEX As Long, lngEY As Long
Dim sW As Long, sH As Long
Dim strLines() As String, cLf As Long, maxLineLength As Long

Dim offsetX As Long, offsetY As Long
Dim maxOffsetX As Long, maxOffsetY As Long

Dim bSelecting As Boolean

Dim bPChar As Byte, bLocked As Boolean


Private Sub UserControl_Initialize()
charWidth = pctFont.ScaleWidth / 16
charHeight = pctFont.ScaleHeight / 14

ReDim strLines(0)
cLf = 0
maxLineLength = 0
maxOffsetX = 0
maxOffsetY = 0

refreshScrollbars
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
pctOut.Move 0, 0, UserControl.ScaleWidth - pctVScroll.Width, UserControl.ScaleHeight - pctHScroll.Height
pctVScroll.Move pctOut.Width, 0, pctVScroll.Width, pctOut.Height
pctHScroll.Move 0, pctOut.Height, pctOut.Width
If Err Then Exit Sub

sW = pctOut.ScaleWidth \ charWidth - 1
sH = pctOut.ScaleHeight \ charHeight - 1

drawText
refreshScrollbars
End Sub

Private Sub mnuCut_Click()
If Not bLocked Then
   Clipboard.Clear
   Clipboard.SetText SelText
   SelText = ""
End If
End Sub

Private Sub mnuCopy_Click()
Clipboard.Clear
Clipboard.SetText SelText
End Sub

Private Sub mnuPaste_Click()
If Not bLocked Then
   SelText = Clipboard.GetText
End If
End Sub

Private Sub mnuDelete_Click()
If Not bLocked Then
   SelText = ""
End If
End Sub

Private Sub mnuSelectAll_Click()
setSel 0, 0, Len(strLines(cLf)), cLf
End Sub

Private Sub pctOut_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SX As Long, SY As Long, EX As Long, EY As Long

If Shift = vbCtrlMask Then
   Select Case KeyCode
      Case vbKeyC
         mnuCopy_Click
      Case vbKeyA
         mnuSelectAll_Click
      Case vbKeyX
         mnuCut_Click
      Case vbKeyV
         mnuPaste_Click
   End Select
Else
   Select Case KeyCode
      Case vbKeyPageUp To vbKeyDown
         EX = lngEX
         EY = lngEY
         Select Case KeyCode
            Case vbKeyPageUp
               EY = EY - sH - 1
            Case vbKeyPageDown
               EY = EY + sH + 1
            Case vbKeyEnd
               EX = Len(strLines(EY))
            Case vbKeyHome
               EX = 0
            Case vbKeyLeft
               If EX = 0 Then
                  If EY > 0 Then
                     EY = EY - 1
                     EX = Len(strLines(EY))
                  End If
               Else
                  EX = EX - 1
               End If
            Case vbKeyUp
               EY = EY - 1
            Case vbKeyRight
               If EX = Len(strLines(EY)) Then
                  If EY < cLf Then
                     EY = EY + 1
                     EX = 0
                  End If
               Else
                  EX = EX + 1
               End If
            Case vbKeyDown
               EY = EY + 1
         End Select
         If Shift And vbShiftMask Then
            SX = lngSX
            SY = lngSY
         Else
            SX = EX
            SY = EY
         End If
         setSel SX, SY, EX, EY
      Case vbKeyDelete
         If Not bLocked Then
            If (lngSY = lngEY) And (lngSX = lngEX) Then
               If lngEX = Len(strLines(lngEY)) Then
                  If lngEY < cLf Then
                     setSel lngSX, lngSY, 0, lngEY + 1
                  End If
               Else
                  setSelEndX lngEX + 1
               End If
            End If
            SelText = ""
         End If
   End Select
End If
End Sub

Private Sub pctOut_KeyPress(KeyAscii As Integer)
If bLocked Then Exit Sub

Select Case KeyAscii
   Case Is >= 32
      SelText = Chr$(KeyAscii)
   Case vbKeyBack
      If Not bLocked Then
         If (lngSY = lngEY) And (lngSX = lngEX) Then
            If lngEX = 0 Then
               If lngEY > 0 Then
                  setSel Len(strLines(lngSY - 1)), lngSY - 1, 0, lngSY
               End If
            Else
               setSel lngSX - 1, lngSY, lngSX, lngSY
            End If
         End If
         SelText = ""
      End If
   Case vbKeyReturn
      SelText = vbCrLf
End Select

refreshCaret
End Sub

Private Sub pctOut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
   Case 1
      X = Int(X / charWidth + 0.5) + offsetX
      Y = Y \ charHeight + offsetY
      setSel X, Y, X, Y
      bSelecting = True
End Select

refreshCaret
End Sub

Private Sub pctOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bSelecting Then
   setSel lngSX, lngSY, Int(X / charWidth + 0.5) + offsetX, Y \ charHeight + offsetY
End If
End Sub

Private Sub pctOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bSelecting = False
If Button = 2 Then UserControl.PopupMenu mnuEdit, 0, X, Y
End Sub

Private Sub pctVScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call pctVScroll_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pctVScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then Exit Sub
Dim hgt As Double

hgt = pctOut.ScaleHeight \ charHeight
hgt = hgt / (cLf + 1)
hgt = hgt * pctVScroll.ScaleHeight

Y = Y - hgt / 2

hgt = 1 / (cLf + 1)
hgt = hgt * pctVScroll.ScaleHeight

Y = Y / hgt

offsetY = Y

refreshScrollbars
drawText
End Sub

Private Sub pctVScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctOut.SetFocus
End Sub

Private Sub pctHScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call pctHScroll_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pctHScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then Exit Sub
Dim hgt As Double

hgt = pctOut.ScaleWidth \ charWidth
hgt = hgt / (maxLineLength + 1)
hgt = hgt * pctHScroll.ScaleWidth

X = X - hgt / 2

hgt = 1 / (maxLineLength + 1)
hgt = hgt * pctHScroll.ScaleWidth

X = X / hgt

offsetX = X

refreshScrollbars
drawText
End Sub

Private Sub pctHScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctOut.SetFocus
End Sub

Private Sub pctOut_GotFocus()
refreshCaret
End Sub

Private Sub pctOut_LostFocus()
DestroyCaret
End Sub

Private Sub drawText()
Dim Char As Byte, X As Long, Y As Long
Dim toX As Long, toY As Long

pctOut.Cls

toY = offsetY + pctOut.ScaleHeight \ charHeight + 1
If cLf < toY Then toY = cLf

For Y = 0 To toY - offsetY
   toX = offsetX + pctOut.ScaleWidth \ charWidth + 1
   If Len(strLines(offsetY + Y)) < toX Then toX = Len(strLines(offsetY + Y))
   For X = 1 To toX - offsetX
      If bPChar Then
         Char = bPChar
      Else
         Char = Asc(Mid$(strLines(offsetY + Y), offsetX + X, 1))
      End If
      If Char < 32 Then Char = 32
      BitBlt pctOut.hDC, (X - 1) * charWidth, Y * charHeight, charWidth, charHeight, pctFont.hDC, (Char Mod 16) * charWidth, (Char \ 16 - 2) * charHeight, vbSrcCopy
   Next X
Next Y

invertSelection
pctOut.Refresh
End Sub

Private Sub invertSelection()
Dim Y As Long
Dim SX As Long, SY As Long, EX As Long, EY As Long

getSortedSel SX, SY, EX, EY

If SY = EY Then
   BitBlt pctOut.hDC, (SX - offsetX) * charWidth, (SY - offsetY) * charHeight, (EX - SX) * charWidth, charHeight, 0, 0, 0, vbDstInvert
Else
   BitBlt pctOut.hDC, (SX - offsetX) * charWidth, (SY - offsetY) * charHeight, (Len(strLines(SY)) - SX) * charWidth, charHeight, 0, 0, 0, vbDstInvert
   For Y = SY + 1 To EY - 1
      BitBlt pctOut.hDC, -offsetX * charWidth, (Y - offsetY) * charHeight, Len(strLines(Y)) * charWidth, charHeight, 0, 0, 0, vbDstInvert
   Next Y
   BitBlt pctOut.hDC, -offsetX * charWidth, (EY - offsetY) * charHeight, EX * charWidth, charHeight, 0, 0, 0, vbDstInvert
End If
End Sub

Private Sub getSortedSel(Optional ByRef SX As Long, Optional ByRef SY As Long, Optional ByRef EX As Long, Optional ByRef EY As Long)
SX = lngSX
SY = lngSY
EX = lngEX
EY = lngEY
If (EY < SY) Or ((EY = SY) And (EX < SX)) Then
   SX = SX Xor EX
   EX = SX Xor EX
   SX = SX Xor EX
   
   SY = SY Xor EY
   EY = SY Xor EY
   SY = SY Xor EY
End If
End Sub

Private Sub refreshCaret()
CreateCaret pctOut.hwnd, 0, 1, charHeight
ShowCaret pctOut.hwnd
SetCaretPos (lngEX - offsetX) * charWidth, (lngEY - offsetY) * charHeight
End Sub

Public Property Get Text() As String
Text = Join(strLines, vbCrLf)
End Property

Public Property Let Text(ByVal vNewValue As String)
Dim X As Long, Y As Long, Char As String, tempLine As String

If vNewValue = "" Then
   ReDim strLines(0)
   cLf = 0
Else
   strLines = Split(vNewValue, vbCrLf)
   cLf = UBound(strLines)
End If

refreshLineLength

setSel 0, 0, 0, 0
refreshScrollbars
drawText
End Property

Private Sub setSelStartX(ByRef newSX As Long)
invertSelection
lngSX = newSX
If lngSX < 0 Then lngSX = 0
If lngSX > Len(strLines(lngSY)) Then lngSX = Len(strLines(lngSY))
invertSelection
pctOut.Refresh
End Sub

Private Sub setSelStartY(ByRef newSY As Long)
invertSelection
lngSY = newSY
If lngSY < 0 Then lngSY = 0
If lngSY > cLf Then lngSY = cLf
invertSelection
pctOut.Refresh
End Sub

Private Sub setSelEndX(ByRef newEX As Long)
lngEX = newEX
If lngEX < 0 Then lngEX = 0
If lngEX > Len(strLines(lngEY)) Then lngEX = Len(strLines(lngEY))
scrollToCaret
End Sub

Private Sub setSelEndY(ByRef newEY As Long)
lngEY = newEY
If lngEY < 0 Then lngEY = 0
If lngEY > cLf Then lngEY = cLf
scrollToCaret
End Sub

Private Sub setSel(ByVal SelStartX As Long, ByVal SelStartY As Long, ByVal SelEndX As Long, ByVal SelEndY As Long)
lngSY = SelStartY
If lngSY < 0 Then
   lngSY = 0
ElseIf lngSY > cLf Then
   lngSY = cLf
End If

lngSX = SelStartX
If lngSX < 0 Then
   lngSX = 0
ElseIf lngSX > Len(strLines(lngSY)) Then
   lngSX = Len(strLines(lngSY))
End If


lngEY = SelEndY
If lngEY < 0 Then
   lngEY = 0
ElseIf lngEY > cLf Then
   lngEY = cLf
End If

lngEX = SelEndX
If lngEX < 0 Then
   lngEX = 0
ElseIf lngEX > Len(strLines(lngEY)) Then
   lngEX = Len(strLines(lngEY))
End If

scrollToCaret
End Sub

Public Sub loadTextFile(Filename As String)
Dim k As Long, b() As Byte

k = FreeFile
Open Filename For Binary As #k
ReDim b(LOF(k))
Get #k, , b
Close #k

Text = StrConv(b, vbUnicode)
End Sub

Public Function saveTextFile(Filename As String) As Boolean
Dim k As Long

k = FreeFile
Open Filename For Output As #k
Print #k, Text
Close #k
End Function

Private Sub scrollToCaret()
Dim pX As Long, pY As Long

pX = lngEX - offsetX
pY = lngEY - offsetY

If pX < 0 Then
   offsetX = offsetX + pX
ElseIf pX > sW Then
   offsetX = offsetX + pX - sW
End If

If pY < 0 Then
   offsetY = offsetY + pY
ElseIf pY > sH Then
   offsetY = offsetY + pY - sH
End If

refreshScrollbars
drawText
refreshCaret
End Sub

Private Sub refreshScrollbars()
Dim scrTick As Double, scrSize As Long

pctHScroll.Cls

maxOffsetX = maxLineLength + 1 - pctOut.ScaleWidth \ charWidth
If maxOffsetX < 0 Then maxOffsetX = 0
If offsetX > maxOffsetX Then offsetX = maxOffsetX
If offsetX < 0 Then offsetX = 0

scrTick = pctHScroll.ScaleWidth / (maxLineLength + 1)
scrSize = pctOut.ScaleWidth \ charWidth
BitBlt pctHScroll.hDC, scrTick * offsetX, 0, scrTick * scrSize, pctHScroll.ScaleHeight, 0, 0, 0, vbDstInvert




pctVScroll.Cls

maxOffsetY = cLf + 1 - pctOut.ScaleHeight \ charHeight
If maxOffsetY < 0 Then maxOffsetY = 0
If offsetY > maxOffsetY Then offsetY = maxOffsetY
If offsetY < 0 Then offsetY = 0

scrTick = pctVScroll.ScaleHeight / (cLf + 1)
scrSize = pctOut.ScaleHeight \ charHeight
BitBlt pctVScroll.hDC, 0, scrTick * offsetY, pctVScroll.ScaleWidth, scrTick * scrSize, 0, 0, 0, vbDstInvert
End Sub

Private Sub refreshLineLength()
Dim i As Long, t As Long
maxLineLength = 0
For i = 0 To cLf
   t = Len(strLines(i))
   If t > maxLineLength Then
      maxLineLength = t
   End If
Next i
End Sub

Public Sub importFont(Filename As String)
Dim tPic As StdPicture
On Error Resume Next

Set tPic = LoadPicture(Filename)
If Err Then
   Err.Raise Err.Number
   Exit Sub
End If

If tPic.Width < 16 Or tPic.Height < 14 Then
   Err.Raise vbObjectError, , "Font is too small."
   Exit Sub
End If

pctFont = tPic

charWidth = pctFont.ScaleWidth / 16
charHeight = pctFont.ScaleHeight / 14

sW = pctOut.ScaleWidth \ charWidth - 1
sH = pctOut.ScaleHeight \ charHeight - 1

refreshScrollbars

drawText
End Sub

Public Sub exportFont(ByRef Filename As String)
On Error Resume Next

SavePicture pctFont.Picture, Filename
If Err Then
   Err.Raise Err.Number
   Exit Sub
End If

End Sub

Public Property Get SelStart() As Long
Dim SX As Long, SY As Long
Dim i As Long

getSortedSel SX, SY

SelStart = SY * 2 + SX
For i = 0 To SY - 1
   SelStart = SelStart + Len(strLines(i))
Next i
End Property

Public Property Let SelStart(ByVal newStart As Long)
Dim SX As Long, SY As Long

newStart = newStart + 2
For SY = 0 To cLf
   newStart = newStart - Len(strLines(SY)) - 2
   If newStart <= 0 Then Exit For
Next SY

If newStart <= 0 Then
   SX = Len(strLines(SY)) + newStart
Else
   SX = Len(strLines(cLf))
End If

setSel SX, SY, SX, SY
End Property

Public Property Get SelText() As String
Dim SX As Long, SY As Long, EX As Long, EY As Long
Dim strText As String
Dim Y As Long

getSortedSel SX, SY, EX, EY

If SY = EY Then
   If SX = SY Then Exit Property
   strText = Mid$(strLines(SY), SX + 1, EX - SX)
Else
   strText = Mid$(strLines(SY), SX + 1) & vbCrLf
   For Y = SY + 1 To EY - 1
      strText = strText & strLines(Y) & vbCrLf
   Next Y
   strText = strText & Left$(strLines(EY), EX)
End If

SelText = strText
End Property

Public Property Let SelText(ByVal strString As String)
Dim SX As Long, SY As Long, EX As Long, EY As Long
Dim arString() As String, cAr As Long, cCng As String
Dim lX As Long, lY As Long
Dim i As Long

arString = Split(strString, vbCrLf)
cAr = UBound(arString)
If cAr < 0 Then
   cAr = 0
   ReDim arString(0)
End If

getSortedSel SX, SY, EX, EY

lX = Len(arString(cAr))
If cAr = 0 Then lX = lX + SX
lY = SY + cAr

arString(cAr) = arString(cAr) & Mid$(strLines(EY), EX + 1)
strLines(SY) = Left$(strLines(SY), SX) & arString(0)

cCng = cAr + SY - EY
If cCng > 0 Then
   ReDim Preserve strLines(cLf + cCng)
   For i = cLf To EY + 1 Step -1
      strLines(i + cCng) = strLines(i)
   Next i
Else
   For i = EY + 1 To cLf
      strLines(i + cCng) = strLines(i)
   Next i
   ReDim Preserve strLines(cLf + cCng)
End If
   
For i = 1 To cAr
   strLines(SY + i) = arString(i)
Next i

cLf = UBound(strLines)

refreshLineLength
setSel lX, lY, lX, lY
End Property

Public Property Get SelLength() As Long
Dim SX As Long, SY As Long, EX As Long, EY As Long
Dim i As Long

getSortedSel SX, SY, EX, EY

SelLength = (EY - SY) * 2 + EX - SX
For i = SY To EY - 1
   SelLength = SelLength + Len(strLines(i))
Next i
End Property

Public Property Let SelLength(ByVal newLength As Long)
Dim SX As Long, SY As Long, EX As Long, EY As Long
If newLength < 0 Then
   Err.Raise 380
   Exit Property
End If

getSortedSel SX, SY, EX, EY
newLength = newLength + SX

newLength = newLength + 2
For EY = SY To cLf
   newLength = newLength - Len(strLines(EY)) - 2
   If newLength <= 0 Then Exit For
Next EY

If newLength <= 0 Then
   EX = Len(strLines(EY)) + newLength
Else
   EX = Len(strLines(cLf))
End If

setSel SX, SY, EX, EY
End Property

Public Property Get PasswordChar() As String
If bPChar Then
   PasswordChar = Chr$(bPChar)
Else
   bPChar = ""
End If
End Property

Public Property Let PasswordChar(ByVal newChar As String)
If Len(newChar) > 0 Then
   bPChar = Asc(newChar)
Else
   bPChar = 0
End If

drawText
End Property

Public Property Get Locked() As Boolean
Locked = bLocked
End Property

Public Property Let Locked(ByVal bNewValue As Boolean)
bLocked = bNewValue
mnuPaste.Enabled = bLocked
mnuCut.Enabled = bLocked
mnuDelete.Enabled = bLocked
End Property
