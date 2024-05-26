VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Left            =   2880
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   1200
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Menu BtnClear 
      Caption         =   "清除(&C)"
   End
   Begin VB.Menu BtnImport 
      Caption         =   "导入(&I)"
   End
   Begin VB.Menu BtnExport 
      Caption         =   "导出(&E)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Board(99, 99) As Byte, CurMouseY#, CurMouseX#
Dim List&(10000)

Private Sub BtnClear_Click()
Dim i&, j&
For i = 0 To 99
    For j = 0 To 99
        Board(i, j) = 0
    Next
Next
Timer1 = True
End Sub

Private Sub BtnExport_Click()
Dim Str As String * 2500
Dim i&, j&
For i = 0 To 99
    For j = 0 To 24
        Mid(Str, i * 25 + j + 1, 1) = Chr(Board(i, j * 4) + Board(i, j * 4 + 1) * 3 + Board(i, j * 4 + 2) * 9 + Board(i, j * 4 + 3) * 27 + 35)
    Next
Next
Clipboard.Clear
Clipboard.SetText Str
MsgBox "已导出到剪贴板"
End Sub

Private Sub BtnImport_Click()
Dim Str$, i&, j&, e&
Str = Clipboard.GetText
If Len(Str) <> 2500 Then
    MsgBox "存档不合法"
    Exit Sub
End If
For i = 0 To 99
    For j = 0 To 24
        e = Asc(Mid(Str, i * 25 + j + 1, 1)) - 35
        If e < 0 Then e = 0
        If e > 80 Then e = 80
        Board(i, j * 4) = e Mod 3
        e = e \ 3
        Board(i, j * 4 + 1) = e Mod 3
        e = e \ 3
        Board(i, j * 4 + 2) = e Mod 3
        e = e \ 3
        Board(i, j * 4 + 3) = e Mod 3
    Next
Next
Timer1 = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case vbKeySpace
    Step
Case vbKeyA, vbKeyA + 32
    Timer2.Interval = 0
Case vbKeyS, vbKeyS + 32
    Timer2.Interval = 1000
Case vbKeyD, vbKeyD + 32
    Timer2.Interval = 500
Case vbKeyF, vbKeyF + 32
    Timer2.Interval = 200
Case vbKeyG, vbKeyG + 32
    Timer2.Interval = 100
Case vbKeyH, vbKeyH + 32
    Timer2.Interval = 1
End Select
End Sub

Private Sub Form_Load()
Dim CurrentWidth#
CurrentWidth = 15 * 800
Width = Width - ScaleWidth + CurrentWidth
Height = Height - ScaleHeight + CurrentWidth
Scale (0, 0)-(100, 100)
Redraw
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 0 Or Y < 0 Or X >= 100 Or Y >= 100 Then Exit Sub
If Button = 1 And Board(Int(Y), Int(X)) <> 0 Then
    Fill Int(X), Int(Y), 3 Xor Board(Int(Y), Int(X))
Else
    Form_MouseMove Button, Shift, X, Y
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i#, j&, k&
If X < 0 Or Y < 0 Or X >= 100 Or Y >= 100 Then Exit Sub
If Button = 2 Then
    For i = 0 To 1 Step 0.01 * (1 + Shift * 2)
        For j = Int(Y) - Shift To Int(Y) + Shift
            For k = Int(X) - Shift To Int(X) + Shift
                If Y >= 0 And Y < 100 And X >= 0 And X < 100 Then
                    Board(j, k) = 0
                End If
            Next
        Next
    Next
    Timer1 = True
ElseIf Button = 1 Then
    For i = 0 To 1 Step 0.01
        Fill Int(X * i + CurMouseX * (1 - i)), Int(Y * i + CurMouseY * (1 - i)), 1
    Next
End If
CurMouseY = Y
CurMouseX = X
End Sub

Function IsBridge(ByVal X&, ByVal Y&) As Boolean
If Board(X, Y) Then Exit Function
If X = 0 Or X = 99 Then Exit Function
If Y = 0 Or Y = 99 Then Exit Function
IsBridge = Board(X - 1, Y) * Board(X + 1, Y) * Board(X, Y - 1) * Board(X, Y + 1)
End Function

Sub Fill(ByVal X&, ByVal Y&, ByVal Color&)
Dim ListLen&, CurPos&
Board(Y, X) = Color
More:
If Y > 0 Then
    If Board(Y - 1, X) = 3 - Color Then
        Board(Y - 1, X) = Color
        List(ListLen) = (Y - 1) * 65536 + X
        ListLen = ListLen + 1
    ElseIf IsBridge(Y - 1, X) Then
        If Board(Y - 2, X) = 3 - Color Then
            Board(Y - 2, X) = Color
            List(ListLen) = (Y - 2) * 65536 + X
            ListLen = ListLen + 1
        End If
    End If
End If
If Y < 99 Then
    If Board(Y + 1, X) = 3 - Color Then
        Board(Y + 1, X) = Color
        List(ListLen) = (Y + 1) * 65536 + X
        ListLen = ListLen + 1
    ElseIf IsBridge(Y + 1, X) Then
        If Board(Y + 2, X) = 3 - Color Then
            Board(Y + 2, X) = Color
            List(ListLen) = (Y + 2) * 65536 + X
            ListLen = ListLen + 1
        End If
    End If
End If
If X > 0 Then
    If Board(Y, X - 1) = 3 - Color Then
        Board(Y, X - 1) = Color
        List(ListLen) = (Y) * 65536 + X - 1
        ListLen = ListLen + 1
    ElseIf IsBridge(Y, X - 1) Then
        If Board(Y, X - 2) = 3 - Color Then
            Board(Y, X - 2) = Color
            List(ListLen) = (Y) * 65536 + X - 2
            ListLen = ListLen + 1
        End If
    End If
End If
If X < 99 Then
    If Board(Y, X + 1) = 3 - Color Then
        Board(Y, X + 1) = Color
        List(ListLen) = (Y) * 65536 + X + 1
        ListLen = ListLen + 1
    ElseIf IsBridge(Y, X + 1) Then
        If Board(Y, X + 2) = 3 - Color Then
            Board(Y, X + 2) = Color
            List(ListLen) = (Y) * 65536 + X + 2
            ListLen = ListLen + 1
        End If
    End If
End If
If ListLen > CurPos Then
    X = List(CurPos) And 65535
    Y = List(CurPos) \ 65536
    CurPos = CurPos + 1
    GoTo More
End If
Timer1 = True
End Sub

Sub Redraw()
Debug.Print 1, Timer
Dim i&, j&
For i = 0 To 99
    For j = 0 To 99
        Picture1.PSet (j, i), Choose(1 + Board(i, j), IIf(i + j And 1, &HFFCCFF, &HCCFFCC), vbBlack, vbRed)
    Next
Next
Debug.Print 2, Timer
Me.PaintPicture Picture1.Image, 0, 0, ScaleWidth, ScaleHeight
Debug.Print 3, Timer
End Sub

Private Sub Timer1_Timer()
Timer1 = False
Redraw
End Sub

Function NeighborC%(ByVal Y&, ByVal X&)
If Y > 0 Then If Board(Y - 1, X) Then NeighborC = NeighborC + 1
If X > 0 Then If Board(Y, X - 1) Then NeighborC = NeighborC + 2
If Y < 99 Then If Board(Y + 1, X) Then NeighborC = NeighborC + 4
If X < 99 Then If Board(Y, X + 1) Then NeighborC = NeighborC + 8
End Function

Sub Step()
Dim Y&, X&, Yl&, Xl&
Dim Val&(99, 99)
Dim ListLen&, CurPos&, Sum&
For Y = 0 To 99
    For X = 0 To 99
        If Board(Y, X) = 0 Then
            Select Case NeighborC(Y, X)
            Case 14
                Val(Y + 1, X) = Val(Y + 1, X) + IIf(Board(Y, X - 1) = Board(Y, X + 1), -32768, 32768)
            Case 13
                Val(Y, X + 1) = Val(Y, X + 1) + IIf(Board(Y - 1, X) = Board(Y + 1, X), -32768, 32768)
            Case 11
                Val(Y - 1, X) = Val(Y - 1, X) + IIf(Board(Y, X - 1) = Board(Y, X + 1), -32768, 32768)
            Case 7
                Val(Y, X - 1) = Val(Y, X - 1) + IIf(Board(Y - 1, X) = Board(Y + 1, X), -32768, 32768)
            End Select
            Val(Y, X) = 1
        End If
    Next
Next
For Y = 0 To 99
    For X = 0 To 99
        If Val(Y, X) - 1 And 1 Then
            ListLen = 0
            CurPos = 0
            List(0) = Y * 65536 + X
            Yl = Y: Xl = X
            Sum = Val(Yl, Xl) + (Board(Yl, Xl) And 2) - 1
            Val(Yl, Xl) = 1
More:           If Yl > 0 Then
                If Board(Yl - 1, Xl) Then
                    If Val(Yl - 1, Xl) <> 1 Then
                        Sum = Sum + Val(Yl - 1, Xl) + (Board(Yl - 1, Xl) And 2) - 1
                        Val(Yl - 1, Xl) = 1
                        ListLen = ListLen + 1
                        List(ListLen) = (Yl - 1) * 65536 + Xl
                    End If
                ElseIf IsBridge(Yl - 1, Xl) Then
                    If Val(Yl - 2, Xl) <> 1 Then
                        Sum = Sum + Val(Yl - 2, Xl) + (Board(Yl - 2, Xl) And 2) - 1
                        Val(Yl - 2, Xl) = 1
                        ListLen = ListLen + 1
                        List(ListLen) = (Yl - 2) * 65536 + Xl
                    End If
                End If
            End If
            If Yl < 99 Then
                If Board(Yl + 1, Xl) Then
                    If Val(Yl + 1, Xl) <> 1 Then
                        Sum = Sum + Val(Yl + 1, Xl) + (Board(Yl + 1, Xl) And 2) - 1
                        Val(Yl + 1, Xl) = 1
                        ListLen = ListLen + 1
                        List(ListLen) = (Yl + 1) * 65536 + Xl
                    End If
                ElseIf IsBridge(Yl + 1, Xl) Then
                    If Val(Yl + 2, Xl) <> 1 Then
                        Sum = Sum + Val(Yl + 2, Xl) + (Board(Yl + 2, Xl) And 2) - 1
                        Val(Yl + 2, Xl) = 1
                        ListLen = ListLen + 1
                        List(ListLen) = (Yl + 2) * 65536 + Xl
                    End If
                End If
            End If
            If Xl > 0 Then
                If Board(Yl, Xl - 1) Then
                    If Val(Yl, Xl - 1) <> 1 Then
                        Sum = Sum + Val(Yl, Xl - 1) + (Board(Yl, Xl - 1) And 2) - 1
                        Val(Yl, Xl - 1) = 1
                        ListLen = ListLen + 1
                        List(ListLen) = (Yl) * 65536 + Xl - 1
                    End If
                ElseIf IsBridge(Yl, Xl - 1) Then
                    If Val(Yl, Xl - 2) <> 1 Then
                        Sum = Sum + Val(Yl, Xl - 2) + (Board(Yl, Xl - 2) And 2) - 1
                        Val(Yl, Xl - 2) = 1
                        ListLen = ListLen + 1
                        List(ListLen) = (Yl) * 65536 + Xl - 2
                    End If
                End If
            End If
            If Xl < 99 Then
                If Board(Yl, Xl + 1) Then
                    If Val(Yl, Xl + 1) <> 1 Then
                        Sum = Sum + Val(Yl, Xl + 1) + (Board(Yl, Xl + 1) And 2) - 1
                        Val(Yl, Xl + 1) = 1
                        ListLen = ListLen + 1
                        List(ListLen) = (Yl) * 65536 + Xl + 1
                    End If
                ElseIf IsBridge(Yl, Xl + 1) Then
                    If Val(Yl, Xl + 2) <> 1 Then
                        Sum = Sum + Val(Yl, Xl + 2) + (Board(Yl, Xl + 2) And 2) - 1
                        Val(Yl, Xl + 2) = 1
                        ListLen = ListLen + 1
                        List(ListLen) = (Yl) * 65536 + Xl + 2
                    End If
                End If
            End If
            If CurPos < ListLen Then
                CurPos = CurPos + 1
                Yl = List(CurPos) \ 65536
                Xl = List(CurPos) And 65535
                GoTo More
            End If
            Sum = IIf(Sum > 0, 5, 3)
            For CurPos = 0 To ListLen
                Val(List(CurPos) \ 65536, List(CurPos) And 65535) = Sum
            Next
        End If
    Next
Next
For Y = 0 To 99
    For X = 0 To 99
        Board(Y, X) = Val(Y, X) \ 2
    Next
Next
Timer1 = True
End Sub













Private Sub Timer2_Timer()
Timer2 = False
Step
Timer2 = True
End Sub
