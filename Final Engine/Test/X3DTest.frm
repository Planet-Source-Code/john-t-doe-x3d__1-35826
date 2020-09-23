VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z As Single
Dim X3D As New X3DEngine
Private Sub Form_Click()
Do
If z < 3 Then
z = z + 0.1
X3D.MoveObject 2, 0, 0, z
Else
X3D.RotateObject 2, 0, 1, 0
End If
X3D.Render
Me.Caption = z
DoEvents
Loop
End Sub

Sub TDTri()

End Sub

Private Sub Form_DblClick()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()

X3D.SetSun 1, 1, 1, 10, 255
'Picture1.Visible = False
X3D.Init Form1.Width, Form1.Height, Form1.hDC, 0, 250

X3D.StartObject 1
X3D.StartPolygon 10, RGB(0, 0, 255)
X3D.SetPoint 0, 0.5, 0.5, 0
X3D.SetPoint 1, 0, 0.5, 0
X3D.SetPoint 2, 0.5, 1, 0
X3D.SetObjectPolygon 0
X3D.FinishObject 2
X3D.StartObject 1

X3D.StartPolygon 10, RGB(255, 0, 0)
X3D.SetPoint 0, 0.5, 0.5, 3
X3D.SetPoint 1, -1, 0.5, 3
X3D.SetPoint 2, 0.5, 2, 3
X3D.SetObjectPolygon 0
X3D.FinishObject 1
X3D.MoveObject 0, 0, 0, 3

X3D.SetCameraPos -0.5, 0, 0
End Sub


Function Degrees(Radians As Single)
Degrees = Radians * 180 / Pi
End Function
Function Radians(Degrees As Single)
Radians = Degrees * 180 / Pi
End Function

Private Sub Picture1_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Form_Click
End Sub

