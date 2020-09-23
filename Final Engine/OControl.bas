Attribute VB_Name = "OControl"
Dim World(1000) As XObject 'Objects
Dim WNum As Integer 'Number of Objects
Dim CurObject As XObject 'Object Currently Being Built
Dim CurPolygon As Polygon 'Polygon Currently Being Built
Dim maxpnum As Integer 'The Number of Polygons in a object
Dim maxnum As Integer
Dim HDC As Long 'The Rendering Target
Dim sX As Integer 'The width of Rendering Target
Dim sY As Integer 'The Height of Rendering Target
Dim Camera As Point3D 'The Camera
Dim Sun As Light 'The Sun(Main lightsource)
Dim BackColor As Long
Dim ViewDist As Single
'Set the properties of the sun
Sub SetSun(X As Single, Y As Single, Z As Single, RayLoss As Integer, Brightness As Integer)
Sun.X = X
Sun.Y = Y
Sun.Z = Z
Sun.Factor = RayLoss
Sun.Brightness = Brightness
End Sub
'Initialize the engine
Sub Init(Width As Integer, Height As Integer, TargetHDC As Long, BackColor As Long, ViewRange As Single)
sX = Width
sY = Height
HDC = TargetHDC
ViewDist = ViewRange
'OControl.ChangeScreenSettings 640, 480, 32
End Sub
'Start an Object
Sub StartObject(PolygonNum As Integer)
maxpnum = 0
ReDim CurObject.Polygons(PolygonNum)
CurObject.PolyNum = 0
End Sub
'Set the polygon of an object to the one currently built
Sub SetObjectPolygon(PolygonID As Integer)
CurObject.Polygons(PolygonID) = CurPolygon
If maxpnum < PolygonID Then
maxpnum = PolygonID
CurObject.PolyNum = maxnum
End If
End Sub
'Finilize an object and make it visible
Sub FinishObject(ObjectID As Integer)
World(ObjectID) = CurObject
If ObjectID > WNum Then
WNum = ObjectID
End If
End Sub
'Start a polygon
Sub StartPolygon(PointNum As Integer, Color As Long)
ReDim CurPolygon.Points(PointNum)
CurPolygon.PointNum = PointNum
CurPolygon.Color = Color
maxnum = 0
End Sub
'Set a point in a polygon
Sub SetPoint(PointNum As Integer, X As Single, Y As Single, Z As Single)
If PointNum > maxnum Then
maxnum = PointNum
CurPolygon.PointNum = maxnum
End If
With CurPolygon
    .Points(PointNum).X = X
    .Points(PointNum).Y = Y
    .Points(PointNum).Z = Z
End With
End Sub
'Move an object
Sub MoveObject(ObjectID, NewX As Single, NewY As Single, NewZ As Single)
With World(ObjectID)
    .X = NewX
    .Y = NewY
    .Z = NewZ
End With
End Sub
'Set the position of the camera
Sub SetCameraPos(X As Single, Y As Single, Z As Single)
Camera.X = X
Camera.Y = Y
Camera.Z = Z
End Sub
'Render the world to the viewport
Sub Render()
On Error Resume Next
Dim Obj() As XObject
ReDim Obj(WNum)
Dim pn As Integer
Dim Objects() As XObject
Dim PNUM As Single
ReDim Objects(WNum)
'Copy World onto Temparary World For Rendering
For R% = 0 To WNum
Objects(R%) = World(R%)
Next

'Move Points Based on Object Location
For o = 0 To WNum 'Current Object Loop
For pu = 0 To Objects(WNum).PolyNum 'Current Polygon Loop
For pn = 0 To Objects(WNum).Polygons(pu).PointNum 'Current Point Loop
With Objects(WNum).Polygons(pu).Points(pn) 'Point being modified
.X = .X + Objects(WNum).X 'Add Object Location X to point X
.Y = .Y + Objects(WNum).Y 'Add Object Location Y to point Y
.Z = .Z + Objects(WNum).Z 'Add Object Location Z to point Z
End With
Next
Next
Next


Dim distx As Single 'Distance Between Polygon and Sun
Dim disty As Single
Dim distz As Single
Dim dist As Single

Dim cx As Single 'The Location of the object used to find distance from the sun
Dim cy As Single
Dim cz As Single
Dim sbn As Long
'Shade Polygons
If Sun.Factor > 0 Then 'Check if the amount of light lost when moved from sun is greater then 0.  If not do not shade.  If it is then shade
For o = 0 To WNum 'Current Object Loop
For pu = 0 To Objects(WNum).PolyNum 'Current Polygon Loop
With Objects(o).Polygons(pu).Points(0) 'First point of polygon used to determine shading dist from sun.
cx = Objects(o).X + .X
cy = Objects(o).Y + .Y
cz = Objects(o).Z + .Z

If cx > Sun.X Then 'Detect if the X of the point is greater then the X of the sun as not to come up with a negitive number in distance.
distx = cx - Sun.X 'Find points distance from sun on X-Range
Else
distx = Sun.X - cx 'Find points distance from sun on X-Range
End If

If cy > Sun.Y Then 'Do the same with Y as X
disty = cy - Sun.Y 'Find points distance from sun on Y-Range
Else
disty = Sun.Y - cy 'Find points distance from sun on Y-Range
End If

If cz > Sun.Z Then 'Do the same with Z as X
distz = cz - Sun.Z 'Find points distance from sun on Z-Range
Else
distz = Sun.Z - cz 'Find points distance from sun on Z-Range
End If
dist = distx + disty + distz 'Find total distance from sun.
Form1.Caption = dist * SunFactor
sbn = RGB(Sun.Brightness - dist * Sun.Factor, Sun.Brightness - dist * Sun.Factor, Sun.Brightness - dist * Sun.Factor) 'Come up with color that should be subtracted from polygon color based on distance from the sun.
Objects(o).Polygons(pu).Color = SubtractRgb(Objects(o).Polygons(pu).Color, sbn) 'Shade the polygon
End With
Next
Next
End If
'Relocate World Based on Camera
For o = 0 To WNum 'Object loop
For pu = 0 To Objects(o).PolyNum 'Polygon Loop
For pn = 0 To 2 'Point loop
With Objects(o).Polygons(pu).Points(pn)
    .X = .X - Camera.X 'Relocate each point based on distance from Camera/Eye
    .Y = .Y - Camera.Y
    .Z = .Z - Camera.Z
End With
Next
Next
Next

'ZBuffering:
'ZBuffering works by drawing polygons in order:
'From the ones furthest from the camera to the ones closest.
'Its a simple consept.
Dim Z(2) As Integer
Dim AZ As Single
Dim curobj As XObject
Dim sortobj As XObject
Dim Polysurface() As Polygon
Dim ix As Long
For o = 0 To WNum
For pu = 0 To Objects(WNum).PolyNum
ix = ix + 1
Next
Next
ReDim Polysurface(ix)

'Find the average Z of every polygon
For o = 0 To WNum
For pu = 0 To Objects(WNum).PolyNum
For pn = 0 To 2
With Objects(o).Polygons(pu).Points(pn)
    Z(pn) = .Z
End With
Next
AZ = (Z(0) + Z(1) + Z(2)) / 3
Objects(o).Polygons(pu).AZ = AZ
If Objects(o).Polygons(pu).AZ > ViewDist Then
Objects(o).Polygons(pu).NoFlag = 1
End If
Next
Next
ix = 0
For o = 0 To WNum
For pu = 0 To Objects(WNum).PolyNum
Polysurface(ix) = Objects(o).Polygons(pu)
ix = ix + 1
Next
Next
Dim curaz As Single
Dim transpoly As Polygon

For i = 0 To ix - 1
curaz = Polysurface(i).AZ
For k% = 0 To ix - 1
If Polysurface(k%).AZ < curaz Then
transpoly = Polysurface(i)
Polysurface(i) = Polysurface(k%)
Polysurface(k%) = transpoly
End If
Next
Next


Dim ax As Single
Dim ay As Single
Dim a
Dim objs As XObject
Dim Polznum() As Long
Dim ozn() As Long
ReDim ozn(WNum)

For o = 0 To ix - 1
For pn = 0 To 2
With Polysurface(o).Points(pn) 'Current point being moved
   If .Z > 0 Then 'Only put point/polygon on screen if it is infront of camera
    DoEvents
   If .X <> 0 Then 'Make shure the X isn't zero to keep from getting Division by Zero
   If .Z / .X > 0 Then 'Prevent division by Zero.
    'The code below is hard to explain.
    'Pnum will be the Z of the point divided by the width of the viewport(draw area).
    'Then the X on the screen becomes the 3D Space X divided by PNUM
    'PNUM comes up with a ratio so that if the Z is one then if X is one it takes up the hole viewport
    'If Z is two then an X of one takes up half the view port.
    'This code works the same with all the other values(Y and Z).
    PNUM = (.Z / sX)
    .X = .X / PNUM
    End If
    End If
    If .Y <> 0 Then
    If .Z / .Y > 0 Then
    PNUM = (.Z / sY)
      .Y = .Y / PNUM
    End If
    End If
        Else
      Polysurface(o).NoFlag = 1 'If out of view do not display polygon.
    End If
End With
Next
Next


'Draw background color
Dim nsx As Single
Dim nsy As Single
nsx = sX
nsy = sY
PolyControl.StartPolygon BackColor, 5
PolyControl.SetPoint 0, 0, 0
PolyControl.SetPoint 1, 0, nsy
PolyControl.SetPoint 2, nsx, nsy
PolyControl.SetPoint 3, nsx, 0
PolyControl.Render HDC

'Display To Screen
For o = 0 To ix - 1 'Object Loop
If Polysurface(o).NoFlag = 0 Then
PolyControl.StartPolygon Polysurface(o).Color, 2   'Start a 2D Polygon
For pn = 0 To 2 'Point loop
'Check that polygon is in view range.
With Polysurface(o).Points(pn) 'Current point
    PolyControl.SetPoint pn, .X, .Y 'Add point to 2D polygon
End With
DoEvents
Next
PolyControl.Render HDC 'Draw the polygon to view area
End If
Next


End Sub
Function Degrees(Radians As Single)
Degrees = Radians * 180 / Pi
End Function
Function Radians(Degrees As Single)
Radians = Degrees * 180 / Pi
End Function
Function SubtractRgb(Number1 As Long, Number2 As Long) As Double
   Dim R1 As Byte, G1 As Byte, B1 As Byte, R2 As Byte, G2 As Byte, B2 As Byte
   Dim R As Byte, G As Byte, B As Byte
   ColorMunip.Long2RGB Number1, R1, G1, B1
   ColorMunip.Long2RGB Number2, R2, G2, B2
    If R1 > 255 - R2 Then
    R = R1 - (255 - R2)
    Else
    R = 0
    End If
    If G1 > 255 - G2 Then
    G = G2 - (255 - G2)
    Else
    G = 0
    End If
    If B1 > 255 - B2 Then
    B = B1 - (255 - B2)
    Else
    B = 0
    End If
   SubtractRgb = RGB(R, G, B)
End Function
