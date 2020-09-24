Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_ADDSTRING = &H143
Public xs() As Long, ys() As Long
Public f As Long, r As Long
Public Canceled As Boolean, cb As Boolean, cx As Long, cy As Long
Public ImgSrc As String, ToPage As String, ALTtext As String
Public CoordArray As String, ShapeType As String
Public IMGMAPtext As String
Public Areas() As String, a As Long
Public AllAREAs As String, LastArea As String
Public AREACompleted As Boolean
Public Colours()
