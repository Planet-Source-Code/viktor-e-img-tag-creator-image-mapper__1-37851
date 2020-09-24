VERSION 5.00
Begin VB.Form fim 
   AutoRedraw      =   -1  'True
   Caption         =   "Image"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3090
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3045
      Left            =   0
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   0
      ToolTipText     =   "Right-click to clear map"
      Top             =   0
      Width           =   4545
   End
End
Attribute VB_Name = "fim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub p_Click()
If ImgSrc = "" Then MsgBox "No image selected": Exit Sub
If ToPage = "" Then MsgBox "No page selected": Exit Sub
End Sub

Private Sub p_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    Canceled = True
    cb = False
    Form1.ClearAREA
    Form1.EraseLastAREAbyRightClick
    Exit Sub
End If
Canceled = False
If Form1.oshape(0).Value = True Then 'circle
    If cb = False Then 'circle closed
        cx = X: cy = Y
    Else
        CoordArray = cx & ", " & cy & ", " & r
        Form1.RewriteMAP
    End If
    cb = Not cb
End If
If Form1.oshape(1).Value = True Then 'rectangle
    If f = 0 Then Form1.ClearAREA
    p.PSet (X, Y)
    ReDim Preserve xs(f): ReDim Preserve ys(f)
    xs(f) = X: ys(f) = Y
    If f > 0 Then
        p.Line (xs(f), ys(f))-(xs(f - 1), ys(f))
        p.Line (xs(f - 1), ys(f))-(xs(f - 1), ys(f - 1))
        p.Line (xs(f), ys(f - 1))-(xs(f - 1), ys(f - 1))
        p.Line (xs(f), ys(f - 1))-(xs(f), ys(f))
        CoordArray = xs(0) & ", " & ys(0) & ", " & xs(1) & ", " & ys(1)
        Form1.RewriteMAP
        f = 0: Erase xs: Erase ys
        Exit Sub
    End If
    f = f + 1
End If
If Form1.oshape(2).Value = True Then 'polygon
    p.PSet (X, Y)
    f = f + 1
    ReDim Preserve xs(f): ReDim Preserve ys(f)
    xs(f) = X: ys(f) = Y
    If f = 1 Then
        CoordArray = xs(f) & ", " & ys(f)
    End If
    If f > 1 Then
        CoordArray = CoordArray & ", " & xs(f) & ", " & ys(f)
        p.Line (xs(f), ys(f))-(xs(f - 1), ys(f - 1))
    End If
    Form1.bcp.Enabled = IIf((f > 2), True, False)
End If
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Canceled = False Then
If Form1.oshape(0).Value = True Then
    'draw circle
    If cb = True Then
        p.Cls
        p.Circle (cx, cy), Abs(X - cx)
        r = Abs(X - cx)
        Form1.lcc.Caption = "Center: (" & cx & ", " & cy & ")    Radius: " & r
    End If
End If
If Form1.oshape(1).Value = True Then
    'draw rectangle
    If f > 0 Then
    p.Cls
    p.Line (X, Y)-(xs(f - 1), Y)
    p.Line (xs(f - 1), Y)-(xs(f - 1), ys(f - 1))
    p.Line (X, ys(f - 1))-(xs(f - 1), ys(f - 1))
    p.Line (X, ys(f - 1))-(X, Y)
    Form1.lcc.Caption = "(" & xs(0) & "->" & X & ", " & ys(0) & "->" & Y & ")    W=" & Abs(xs(f - 1) - X) & ", H=" & Abs(ys(f - 1) - Y)
    End If
End If
End If
End Sub
