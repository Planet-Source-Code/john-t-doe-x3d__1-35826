Attribute VB_Name = "TDTypes"
Public Type Point3D
    X As Single
    Y As Single
    Z As Single
End Type
Public Type Polygon
    Points() As Point3D
    PointNum As Integer
    Color As Long
    NoFlag As Long
    AZ As Single
End Type
Public Type RPolygon
    Points() As Point
    PointNum As Integer
End Type
Public Type XObject
    Polygons() As Polygon
    X As Single
    Y As Single
    Z As Single
    PolyNum As Integer
End Type
Public Type Light
    X As Single
    Y As Single
    Z As Single
    Factor As Single
    Brightness As Byte
End Type
Public Type ColoredLight
    X As Single
    Y As Single
    Z As Single
    Factor As Single
    Color As Single
End Type
