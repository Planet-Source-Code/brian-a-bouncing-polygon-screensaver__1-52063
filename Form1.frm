VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Shattergon"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SPOLYGON
    XCenter As Integer
    YCenter As Integer
    XVertex(1 To 10) As Integer
    YVertex(1 To 10) As Integer
    DispX As Integer
    DispY As Integer
    Mass As Double
    Angle As Double
    RSpeed As Double
    AngleR As Double
    Color As Long
    NVertex As Byte
    DispVector As Double
    Displacement As Double
End Type

Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long) 'mouse visible(0 or 1)

Private Const PI As Double = 3.14159265358979

Private Shape(1 To 10) As SPOLYGON, a As Byte
Private MouseMoveInd As Byte, SX(0 To 254) As Integer, SY(0 To 254) As Integer
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'ScreenSaverStuff-------------------------------------
'/////////////////////////////////////////////////////
Private Sub Form_Activate()
    Randomize
    a = 1
    Show
    DoEvents
    SetItUp
    Dim n As Byte
    For n = 0 To 254
        SX(n) = Int(Rnd * Screen.Width)
        SY(n) = Int(Rnd * Screen.Height)
    Next
    Do While a = 1
        MoveShapes
        ShowShapes
        DoEvents
    Loop
    Unload Me
End Sub

Private Sub Form_Load()
    ShowCursor (0)
    If App.PrevInstance = True Then End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseMoveInd = MouseMoveInd + 1
    If MouseMoveInd >= 5 Then a = 254
End Sub

Private Sub Form_Click()
    a = 254
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    a = 254
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Actions---------------------------------------------------
'//////////////////////////////////////////////////////////
Private Sub MoveShapes()
    Dim n As Byte
    For n = 1 To 10
        With Shape(n)
        'screen boundaries
            If .XCenter + (.Mass * 2) >= Screen.Width Then
                .Angle = .Angle + 15#
                .DispX = .XCenter
                .DispY = .YCenter
                .Displacement = 0
            End If
            If .YCenter + (.Mass * 2) >= Screen.Height Then
                .Angle = .Angle + 15#
                .DispX = .XCenter
                .DispY = .YCenter
                .Displacement = 0
            End If
            If .XCenter - (.Mass * 2) <= 0 Then
                .Angle = .Angle + 15#
                .DispX = .XCenter
                .DispY = .YCenter
                .Displacement = 0
            End If
            If .YCenter - (.Mass * 2) <= 0 Then
                .Angle = .Angle + 15#
                .DispX = .XCenter
                .DispY = .YCenter
                .Displacement = 0
            End If
            '--
            If .Mass < 300 Then .Mass = .Mass + 0.025
            .Displacement = .Displacement + (.DispVector)
            .XCenter = Cos(Rad(.Angle)) * .Displacement + .DispX
            .YCenter = Sin(Rad(.Angle)) * .Displacement + .DispY
            Dim m As Byte
            For m = 1 To .NVertex
                .XVertex(m) = .Mass * 2 * Cos(Rad(.AngleR)) + .XCenter
                .YVertex(m) = .Mass * 2 * Sin(Rad(.AngleR)) + .YCenter
                .AngleR = .AngleR + (360 / .NVertex) + .RSpeed
                If .AngleR >= 360 Then .AngleR = .AngleR - 360
            Next
            If .Angle >= 360 Then .Angle = .Angle - 360
            For m = 1 To 10
                If m <> n Then
                    If Dist(.XCenter, .YCenter, Shape(m).XCenter, Shape(m).YCenter) <= (.Mass * 2) + (Shape(m).Mass * 2) Then ChangeShape n, m
                End If
            Next
            If .DispVector = 0 Then .DispVector = Rnd * 4 + 0.1
        End With
    Next
End Sub

Private Sub ChangeShape(n As Byte, m As Byte)
    Dim o As Byte
    With Shape(n)
        For o = 1 To .NVertex
            .XVertex(o) = Empty
            .YVertex(o) = Empty
        Next
        .NVertex = Int(Rnd * 8) + 3
        .Mass = Int(Rnd * 200) + 100
        .Angle = Rnd * 361
        .AngleR = .Angle
        .Displacement = 0
        .DispX = .XCenter
        .DispY = .YCenter
        .RSpeed = Rnd * 0.125 - 0.0625
        .DispVector = Rnd * 7 - 3.5
        If .DispVector = 0 Then .DispVector = Rnd * 4 + 0.1
        .Color = RGB(Int(Rnd * 200) + 56, Int(Rnd * 200) + 56, Int(Rnd * 200) + 56)
        For o = 1 To .NVertex
            .XVertex(o) = .Mass * 2 * Cos(Rad(.AngleR)) + .XCenter
            .YVertex(o) = .Mass * 2 * Sin(Rad(.AngleR)) + .YCenter
            .AngleR = .AngleR + (360 / .NVertex)
        Next
    End With
    With Shape(m)
        For o = 1 To .NVertex
            .XVertex(o) = Empty
            .YVertex(o) = Empty
        Next
        .NVertex = Int(Rnd * 8) + 3
        .Mass = Int(Rnd * 200) + 100
        .Angle = Rnd * 361
        .AngleR = .Angle
        .Displacement = 0
        .DispX = .XCenter
        .DispY = .YCenter
        .RSpeed = Rnd * 0.125 - 0.0625
        .DispVector = Rnd * 7 - 3.5
        If .DispVector = 0 Then .DispVector = Rnd * 4 + 0.1
        .Color = RGB(Int(Rnd * 200) + 56, Int(Rnd * 200) + 56, Int(Rnd * 200) + 56)
        For o = 1 To .NVertex
            .XVertex(o) = .Mass * 2 * Cos(Rad(.AngleR)) + .XCenter
            .YVertex(o) = .Mass * 2 * Sin(Rad(.AngleR)) + .YCenter
            .AngleR = .AngleR + (360 / .NVertex)
        Next
    End With
End Sub

Private Sub ShowShapes()
    Show
    Me.Cls
    Dim n As Byte
    For n = 0 To 254
        PSet (SX(n), SY(n)), RGB(255, 255, 255)
    Next
    For n = 1 To 10
        With Shape(n)
            If .YVertex(4) = Empty Then
                DrawPoly .Color, .XVertex(1), .YVertex(1), .XVertex(2), .YVertex(2), .XVertex(3), .YVertex(3)
            ElseIf .YVertex(5) = Empty Then
                DrawPoly .Color, .XVertex(1), .YVertex(1), .XVertex(2), .YVertex(2), .XVertex(3), .YVertex(3), .XVertex(4), .YVertex(4)
            ElseIf .YVertex(6) = Empty Then
                DrawPoly .Color, .XVertex(1), .YVertex(1), .XVertex(2), .YVertex(2), .XVertex(3), .YVertex(3), .XVertex(4), .YVertex(4), .XVertex(5), .YVertex(5)
            ElseIf .YVertex(7) = Empty Then
                DrawPoly .Color, .XVertex(1), .YVertex(1), .XVertex(2), .YVertex(2), .XVertex(3), .YVertex(3), .XVertex(4), .YVertex(4), .XVertex(5), .YVertex(5), .XVertex(6), .YVertex(6)
            ElseIf .YVertex(8) = Empty Then
                DrawPoly .Color, .XVertex(1), .YVertex(1), .XVertex(2), .YVertex(2), .XVertex(3), .YVertex(3), .XVertex(4), .YVertex(4), .XVertex(5), .YVertex(5), .XVertex(6), .YVertex(6), .XVertex(7), .YVertex(7)
            ElseIf .YVertex(9) = Empty Then
                DrawPoly .Color, .XVertex(1), .YVertex(1), .XVertex(2), .YVertex(2), .XVertex(3), .YVertex(3), .XVertex(4), .YVertex(4), .XVertex(5), .YVertex(5), .XVertex(6), .YVertex(6), .XVertex(7), .YVertex(7), .XVertex(8), .YVertex(8)
            ElseIf .YVertex(10) = Empty Then
                DrawPoly .Color, .XVertex(1), .YVertex(1), .XVertex(2), .YVertex(2), .XVertex(3), .YVertex(3), .XVertex(4), .YVertex(4), .XVertex(5), .YVertex(5), .XVertex(6), .YVertex(6), .XVertex(7), .YVertex(7), .XVertex(8), .YVertex(8), .XVertex(9), .YVertex(9)
            Else: DrawPoly .Color, .XVertex(1), .YVertex(1), .XVertex(2), .YVertex(2), .XVertex(3), .YVertex(3), .XVertex(4), .YVertex(4), .XVertex(5), .YVertex(5), .XVertex(6), .YVertex(6), .XVertex(7), .YVertex(7), .XVertex(8), .YVertex(8), .XVertex(9), .YVertex(9), .XVertex(10), .YVertex(10)
            End If
        End With
    Next
End Sub

Private Sub SetItUp()
    Dim n As Byte
    For n = 1 To 10
        With Shape(n)
            .XCenter = Int(Rnd * Screen.Width / 1.5) + (Screen.Width / 6)
            .YCenter = Int(Rnd * Screen.Height / 1.5) + (Screen.Height / 6)
            .NVertex = Int(Rnd * 8) + 3
            .Mass = Int(Rnd * 300) + 100
            .Angle = Rnd * 361
            .AngleR = .Angle
            .Displacement = 0
            .DispX = .XCenter
            .DispY = .YCenter
            .RSpeed = Rnd * 0.125 - 0.0625
            .DispVector = Rnd * 7 - 3.5
            If .DispVector = 0 Then .DispVector = Rnd * 4 + 0.1
            .Color = RGB(Int(Rnd * 200) + 56, Int(Rnd * 200) + 56, Int(Rnd * 200) + 56)
            Dim m As Byte
            For m = 1 To .NVertex
                .XVertex(m) = .Mass * 2 * Cos(Rad(.AngleR)) + .XCenter
                .YVertex(m) = .Mass * 2 * Sin(Rad(.AngleR)) + .YCenter
                .AngleR = .AngleR + (360 / .NVertex)
            Next
        End With
    Next
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Functions---------------------------------------------
'//////////////////////////////////////////////////////
Private Function Rad(Degrees As Double) As Double
    Rad = Degrees * PI / 180
End Function

Private Function Deg(Radians As Double) As Double
    Deg = Radians * 180 / PI
End Function

Function Dist(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional Z1 As Integer, Optional Z2 As Integer)
    If Z1 = Empty Then Dist = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2) Else: Dist = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2 + (Z2 - Z1) ^ 2)
End Function

Function DrawPoly(Color As Long, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, X3 As Integer, Y3 As Integer, Optional X4 As Integer, Optional Y4 As Integer, Optional X5 As Integer, Optional Y5 As Integer, Optional X6 As Integer, Optional Y6 As Integer, Optional X7 As Integer, Optional Y7 As Integer, Optional X8 As Integer, Optional Y8 As Integer, Optional X9 As Integer, Optional Y9 As Integer, Optional X10 As Integer, Optional Y10 As Integer)
    Line (X1, Y1)-(X2, Y2), Color
    Line (X2, Y2)-(X3, Y3), Color
    If Y4 = Empty Then
        Line (X3, Y3)-(X1, Y1), Color
    Else:
        Line (X3, Y3)-(X4, Y4), Color
        If Y5 = Empty Then
            Line (X4, Y4)-(X1, Y1), Color
        Else:
            Line (X4, Y4)-(X5, Y5), Color
            If Y6 = Empty Then
                Line (X5, Y5)-(X1, Y1), Color
            Else:
                Line (X5, Y5)-(X6, Y6), Color
                If Y7 = Empty Then
                    Line (X6, Y6)-(X1, Y1), Color
                Else:
                    Line (X6, Y6)-(X7, Y7), Color
                    If Y8 = Empty Then
                        Line (X7, Y7)-(X1, Y1), Color
                    Else:
                        Line (X7, Y7)-(X8, Y8), Color
                        If Y9 = Empty Then
                            Line (X8, Y8)-(X1, Y1), Color
                        Else:
                            Line (X8, Y8)-(X9, Y9), Color
                            If Y10 = Empty Then
                                Line (X9, Y9)-(X1, Y1), Color
                            Else:
                                Line (X9, Y9)-(X10, Y10), Color
                                Line (X10, Y10)-(X1, Y1), Color
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    ShowCursor (1)
End Sub
