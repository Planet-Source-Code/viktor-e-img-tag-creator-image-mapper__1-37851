VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HTML image MAPper"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6240
   ScaleWidth      =   3915
   Begin VB.CommandButton bend 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   16
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton belast 
      Caption         =   "Erase last added area"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   15
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CommandButton badd 
      Caption         =   "Add this area to IMG header"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   90
      TabIndex        =   3
      Top             =   3330
      Width           =   2535
   End
   Begin VB.TextBox tareas 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3840
      Width           =   3735
   End
   Begin VB.TextBox talt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2800
      TabIndex        =   9
      Top             =   1485
      Width           =   980
   End
   Begin VB.CommandButton bcopy 
      Caption         =   "Copy IMG header"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2800
      TabIndex        =   4
      Top             =   3330
      Width           =   980
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1340
      Left            =   90
      TabIndex        =   13
      Top             =   800
      Width           =   2535
      Begin VB.CommandButton bcp 
         Caption         =   "Close poly"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   880
         Width           =   975
      End
      Begin VB.OptionButton oshape 
         Caption         =   "CIRCLE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton oshape 
         Caption         =   "RECT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton oshape 
         Caption         =   "POLYGON"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CommandButton bselpg 
      Caption         =   "Page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   2800
      TabIndex        =   2
      Top             =   240
      Width           =   980
   End
   Begin VB.CommandButton bselim 
      Caption         =   "Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdial 
      Left            =   2520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cblc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2520
      Width           =   3740
   End
   Begin VB.Label lcc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   17
      Top             =   2200
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALT text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2805
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Line colour"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Application for mapping image regions for HTML pages

'<AREA
'  COORDS = "coords"
'  Shape = "shape"
'  HREF = "location"
'  NOHREF
'  Target = "windowName"
'  ONMOUSEOUT = "outJScode"
'  ONMOUSEOVER = "overJScode"
'  Name = "areaName"
'  ALT = "No-display/Hover text"
'>

Private Sub Form_Load()
With Me
    .Left = Screen.Width - .Width - 20
    .Top = 20
End With
Colours = Array("White", "Yellow", "Lime", "Blue", "Red")
With cdial
    .Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    .InitDir = App.Path
End With
fim.p.ScaleMode = vbPixels
With cblc
    For c = 0 To UBound(Colours)
        SendMessage .hwnd, CB_ADDSTRING, 0, ByVal CStr(Colours(c))
    Next c
    .ListIndex = 1
End With
End Sub

Private Sub cblc_Click()
'Contour colour selection
With fim.p
Select Case cblc.ListIndex
Case 0
    .ForeColor = vbWhite
Case 1
    .ForeColor = vbYellow
Case 2
    .ForeColor = &HFF00&
Case 3
    .ForeColor = vbBlue
Case 4
    .ForeColor = vbRed
End Select
End With
End Sub

Private Sub badd_Click()
'Add last defined AREA to the IMG header
a = a + 1
ReDim Preserve Areas(a)
Areas(a) = "<AREA" & vbCrLf & " COORDS=""" & CoordArray & """" & vbCrLf & _
ShapeType & vbCrLf & ToPage & vbCrLf & " ALT=""" & ALTtext & """" & ">" & vbCrLf & "</AREA>"

tareas.Text = "<IMG SRC=""" & ImgSrc & """" & vbCrLf & " USEMAP = ""#GeneratedMAP"">" & vbCrLf & "</IMG>"
tareas.Text = tareas.Text & vbCrLf & "<MAP NAME=""GeneratedMAP"">"
AllAREAs = ""
For k = 1 To UBound(Areas)
    AllAREAs = AllAREAs & vbCrLf & Areas(k)
Next k
tareas.Text = tareas.Text & AllAREAs & vbCrLf & "</MAP>"
belast.Enabled = IIf((UBound(Areas) > 1), True, False)
End Sub

Private Sub belast_Click()
'Erase last defined AREA from the IMG header
If UBound(Areas) > 1 Then
    AllAREAs = ""
    ReDim Preserve Areas(UBound(Areas) - 1)
    tareas.Text = "<IMG SRC=""" & ImgSrc & """" & vbCrLf & " USEMAP = ""#GeneratedMAP"">" & vbCrLf & "</IMG>"
    tareas.Text = tareas.Text & vbCrLf & "<MAP NAME=""GeneratedMAP"">"
    For k = 1 To UBound(Areas)
        AllAREAs = AllAREAs & vbCrLf & Areas(k)
    Next k
    tareas.Text = tareas.Text & vbCrLf & AllAREAs & vbCrLf & "</MAP>"
    a = a - 1
    ClearAREA
End If
belast.Enabled = IIf((UBound(Areas) > 1), True, False)
End Sub

Public Sub ClearAREA()
'Clear AREA
Erase xs: Erase ys
f = 0
fim.p.Cls
t.Text = "": talt.Text = ""
lcc.Caption = ""
bcp.Enabled = False
CoordArray = ""
bcopy.Enabled = False
End Sub
Public Sub EraseLastAREAbyRightClick()
'If the user has just defined an AREA and then right-clicks on it,
'then the next AREA will take the place of this one in the tareas.Text
On Error Resume Next
If AREACompleted = True Then
    AllAREAs = ""
    If UBound(Areas) > 1 Then ReDim Preserve Areas(UBound(Areas) - 1)
    For k = 0 To UBound(Areas)
        AllAREAs = AllAREAs & vbCrLf & Areas(k)
    Next k
    If a > 0 Then a = a - 1
    AREACompleted = False
End If
End Sub

Private Sub oshape_Click(Index As Integer)
AREACompleted = False
ClearAREA 'clear anyway
Select Case Index
Case 0 'circle
    ShapeType = " SHAPE=CIRCLE"
    bcp.Enabled = False
Case 1 'rectangle
    ShapeType = " SHAPE=RECT"
    bcp.Enabled = False
Case 2 'polygon
    ShapeType = " SHAPE=POLYGON"
End Select
End Sub

Private Sub bselim_Click()
'Image selection
With cdial
    .Filter = "Common image file types|*.jpg;*.jpeg;*.gif;*.bmp;*.dib;*.pcx"
    .DialogTitle = "Select an image to map"
    .FileName = ""
    .ShowOpen
    If .FileName <> "" Then
        fim.Visible = True
        fim.p.Picture = LoadPicture(.FileName)
        ImgSrc = .FileName
        If ToPage <> "" Then Frame1.Enabled = True
    End If
End With
End Sub

Private Sub bselpg_Click()
'Page to point to
With cdial
    .Filter = "HTML|*.htm;*.html"
    .DialogTitle = "Select a page to point to"
    .FileName = ""
    .ShowOpen
    If .FileName <> "" Then
        ToPage = " HREF=""" & .FileName & """"
        If ImgSrc <> "" Then Frame1.Enabled = True
    End If
End With
End Sub

Private Sub bcp_Click()
'Close POLYGON
fim.p.Line (xs(f), ys(f))-(xs(1), ys(1))
CoordArray = CoordArray & ", " & xs(1) & ", " & ys(1)
RewriteMAP
End Sub

Public Sub RewriteMAP()
Select Case ShapeType
Case " SHAPE=CIRCLE", " SHAPE=RECT"
    RewriteMAP2
Case " SHAPE=POLYGON"
    If f > 2 Then
        RewriteMAP2
        f = 0
    End If
End Select
End Sub
Private Sub RewriteMAP2()
IMGMAPtext = "<IMG SRC=""" & ImgSrc & """" & vbCrLf & " USEMAP = ""#GeneratedMAP"">" & vbCrLf & "</IMG>"
IMGMAPtext = IMGMAPtext & vbCrLf & "<MAP NAME=""GeneratedMAP"">" & vbCrLf
LastArea = "<AREA" & vbCrLf & " COORDS=""" & CoordArray & """" & vbCrLf & _
ShapeType & vbCrLf & ToPage & vbCrLf & " ALT=""" & ALTtext & """" & ">" & vbCrLf & "</AREA>"
IMGMAPtext = IMGMAPtext & LastArea
IMGMAPtext = IMGMAPtext & vbCrLf & "</MAP>"
t.Text = IMGMAPtext
AREACompleted = True
lcc.Caption = ""
End Sub

Private Sub talt_Change()
'ALTernative text
ALTtext = talt.Text
End Sub
Private Sub t_Change()
'Locked
badd.Enabled = IIf((t.Text <> ""), True, False)
End Sub
Private Sub tareas_Change()
'Locked
bcopy.Enabled = IIf((tareas.Text <> ""), True, False)
End Sub

Private Sub bcopy_Click()
'Put IMG tag in Clipboard
Clipboard.Clear
Clipboard.SetText tareas.Text
End Sub

Private Sub bend_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
ClearAREA
Erase Areas
Erase Colours
AllAREAs = ""
fim.p.Picture = LoadPicture()
Unload fim
Set fim = Nothing
Set Form1 = Nothing
End
End Sub
