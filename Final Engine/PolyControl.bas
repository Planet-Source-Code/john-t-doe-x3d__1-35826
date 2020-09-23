Attribute VB_Name = "PolyControl"
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, _
                ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" _
                (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal HDC As Long, _
                ByVal hRgn As Long, ByVal hBrush As Long) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Const WINDING = 2

Private Type Point3D
    X As Single
    Y As Single
    Z As Single
End Type

Public Type Point
    X As Single
    Y As Single
End Type

Dim Points() As POINTAPI
Dim solidbrush As Long
Dim ccolor As Long
Dim rgn As Long
Dim maxpoints As Integer
Dim realmp
Sub StartPolygon(BrushColor As Long, PointNum As Integer)
Dim pointnumx As Integer
pointnumx = PointNum + 1
ReDim Points(pointnumx)
maxpoints = PointNum
realmp = 0
solidbrush = CreateSolidBrush(BrushColor)
End Sub
Sub SetPoint(PointNum As Integer, X As Single, Y As Single)
Points(PointNum).X = X
Points(PointNum).Y = Y
If PointNum > realmp Then realmp = PointNum
End Sub
Sub Render(HDC As Long)
'Form1.Caption = Points(0).X & " " & Points(0).Y
rgn = CreatePolygonRgn(Points(0), realmp + 1, WINDING)
FillRgn HDC, rgn, solidbrush
DeleteObject rgn
DeleteObject solidbrush
End Sub

