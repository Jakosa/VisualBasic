VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   9600
   ClientLeft      =   165
   ClientTop       =   -285
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Balra"
      Height          =   615
      Left            =   9120
      TabIndex        =   3
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   9120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Jobbra"
      Height          =   615
      Left            =   10560
      TabIndex        =   0
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   3240
      X2              =   3240
      Y1              =   8880
      Y2              =   9360
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   8880
      Y2              =   9360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   2400
      X2              =   3240
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   2400
      X2              =   3240
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cél"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      UseMnemonic     =   0   'False
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Start"
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   3840
      UseMnemonic     =   0   'False
      Width           =   330
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   43
      X1              =   5040
      X2              =   4200
      Y1              =   5640
      Y2              =   6240
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   42
      X1              =   4200
      X2              =   3360
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   41
      X1              =   2880
      X2              =   2040
      Y1              =   7080
      Y2              =   6720
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   40
      X1              =   2040
      X2              =   720
      Y1              =   6720
      Y2              =   5520
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   39
      X1              =   5400
      X2              =   4680
      Y1              =   6720
      Y2              =   6960
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   38
      X1              =   5760
      X2              =   5400
      Y1              =   6000
      Y2              =   6720
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   37
      X1              =   4680
      X2              =   3720
      Y1              =   6960
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   36
      X1              =   2280
      X2              =   3360
      Y1              =   6000
      Y2              =   6480
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   35
      X1              =   1560
      X2              =   2280
      Y1              =   5400
      Y2              =   6000
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   34
      X1              =   3720
      X2              =   2880
      Y1              =   7200
      Y2              =   7080
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   33
      X1              =   5760
      X2              =   5400
      Y1              =   4200
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   32
      X1              =   5400
      X2              =   5040
      Y1              =   4920
      Y2              =   5640
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   31
      X1              =   6960
      X2              =   5760
      Y1              =   4320
      Y2              =   4200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   30
      X1              =   7200
      X2              =   6120
      Y1              =   4920
      Y2              =   5280
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   29
      X1              =   6120
      X2              =   5760
      Y1              =   5280
      Y2              =   6000
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   2
      X1              =   7680
      X2              =   6960
      Y1              =   3720
      Y2              =   4320
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   28
      X1              =   8040
      X2              =   7200
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   27
      X1              =   8400
      X2              =   8040
      Y1              =   3840
      Y2              =   4560
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   26
      X1              =   8400
      X2              =   8400
      Y1              =   3240
      Y2              =   3840
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   25
      X1              =   7680
      X2              =   7680
      Y1              =   3120
      Y2              =   3720
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   0
      X1              =   7680
      X2              =   4560
      Y1              =   3120
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   24
      X1              =   8400
      X2              =   4200
      Y1              =   2640
      Y2              =   840
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   23
      X1              =   4560
      X2              =   3720
      Y1              =   1920
      Y2              =   1560
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   22
      X1              =   4200
      X2              =   3360
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   21
      X1              =   3720
      X2              =   3360
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   20
      X1              =   3360
      X2              =   3240
      Y1              =   1800
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   19
      X1              =   3240
      X2              =   3000
      Y1              =   1920
      Y2              =   2520
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   18
      X1              =   2400
      X2              =   1800
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   17
      X1              =   1800
      X2              =   1560
      Y1              =   3240
      Y2              =   3600
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   16
      X1              =   1560
      X2              =   1560
      Y1              =   4800
      Y2              =   5400
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   15
      X1              =   1560
      X2              =   1560
      Y1              =   3600
      Y2              =   4200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   14
      X1              =   3000
      X2              =   2400
      Y1              =   2520
      Y2              =   3120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   13
      X1              =   960
      X2              =   720
      Y1              =   3120
      Y2              =   3720
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   12
      X1              =   2040
      X2              =   1440
      Y1              =   2520
      Y2              =   2640
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   11
      X1              =   2640
      X2              =   2400
      Y1              =   1680
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   10
      X1              =   2400
      X2              =   2040
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   9
      X1              =   3360
      X2              =   2880
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   8
      X1              =   8400
      X2              =   8400
      Y1              =   2640
      Y2              =   3240
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   7
      X1              =   1440
      X2              =   960
      Y1              =   2640
      Y2              =   3120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   6
      X1              =   1560
      X2              =   1560
      Y1              =   4200
      Y2              =   4800
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   5
      X1              =   2880
      X2              =   2640
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   4
      X1              =   720
      X2              =   720
      Y1              =   4320
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   3
      X1              =   720
      X2              =   720
      Y1              =   3720
      Y2              =   4320
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   1
      X1              =   720
      X2              =   720
      Y1              =   4920
      Y2              =   5520
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   840
      X2              =   840
      Y1              =   8880
      Y2              =   9480
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   8880
      Y2              =   9480
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   840
      X2              =   1560
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   840
      X2              =   1560
      Y1              =   8880
      Y2              =   8880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim x0 As Double
Dim y0 As Double
Dim ex As Double
Dim ey As Double
Dim VSzogTomb(1 To 30) As Double
Const xhossz = 80           ' Kocsi hossza.
Const yhossz = 60           ' Kocsi szélessége.

Private Sub Command1_Click()
    NextCoordinate True, False
    NextMove
End Sub

Private Sub Command2_Click()
    NextCoordinate False, False
    NextMove
End Sub

Private Sub Form_Load()
    Dim ciklus As Long, pontok As Long, vonal As Long
    ex = 0.6
    ey = -1
    x0 = 1100
    y0 = 4000
    
    For ciklus = 1 To 30
        VSzogTomb(ciklus) = Rnd * 0.1
    Next ciklus
End Sub

Private Sub Timer1_Timer()
    i = i + 1

    If i > 30 Then
        i = 1
    End If

    NextCoordinate True, True

    Dim xb As Single, yb As Single, xj As Single, yj As Single
    xb = x0 + xhossz * ex - yhossz * ey
    yb = y0 + xhossz * ey + yhossz * ex
    xj = x0 + xhossz * ex + yhossz * ey
    yj = y0 + xhossz * ey - yhossz * ex

    Dim distb As Single, distj As Single, ciklus As Integer
    distb = 1000000
    distj = 1000000

    For ciklus = Line5.LBound To Line5.UBound
        Dim db As Single, dj As Single
        db = Distance(xb, yb, Line5(ciklus).x1, Line5(ciklus).x2, Line5(ciklus).y1, Line5(ciklus).y2)
        dj = Distance(xj, yj, Line5(ciklus).x1, Line5(ciklus).x2, Line5(ciklus).y1, Line5(ciklus).y2)

        If distb > db Then
            distb = db
        End If

        If distj > dj Then
            distj = dj
        End If
    Next ciklus

    If distb - distj > 0 And distb - distj < 200 Then
        NextCoordinate False, False
        NextMove
    ElseIf distj - distb > 0 And distj - distb < 200 Then
        NextCoordinate True, False
        NextMove
    End If
    
    NextMove

    x0 = x0 + 50 * ex
    y0 = y0 + 50 * ey
    NextMove

    Me.Circle (x0, y0), 20, vbRed
End Sub

Private Function Distance(ByVal PointX As Single, ByVal PointY As Single, ByVal LineX1 As Single, ByVal LineX2 As Single, ByVal LineY1 As Single, ByVal LineY2 As Single) As Single
    Dim AA As Single, BB As Single, CC As Single, DD As Single
    Dim dot As Single, len_sq As Single, param As Single
    Dim xx As Single, yy As Single
    AA = PointX - LineX1
    BB = PointY - LineY1
    CC = LineX2 - LineX1
    DD = LineY2 - LineY1

    dot = AA * CC + BB * DD
    len_sq = CC * CC + DD * DD
    param = dot / len_sq

    If param < 0 Then
        xx = LineX1
        yy = LineY1
    ElseIf param > 1 Then
        xx = LineX2
        yy = LineY2
    Else
        xx = LineX1 + param * CC
        yy = LineY1 + param * DD
    End If

    Distance = Sqr(((PointX - xx) * (PointX - xx)) + ((PointY - yy) * (PointY - yy)))
End Function

Private Sub NextCoordinate(BIrany As Boolean, VSzog As Boolean)
    Dim bj As Double

    If BIrany Then
        bj = -0.15
    Else
        bj = 0.15
    End If

    If VSzog Then
        ex = Cos(VSzogTomb(i)) * ex - Sin(VSzogTomb(i)) * ey
        ey = Cos(VSzogTomb(i)) * ey + Sin(VSzogTomb(i)) * ex
    Else
        ex = Cos(bj) * ex - Sin(bj) * ey
        ey = Cos(bj) * ey + Sin(bj) * ex
    End If

    ex = ex / Sqr(ex * ex + ey * ey)
    ey = ey / Sqr(ex * ex + ey * ey)
End Sub

Private Sub NextMove()
    Line1.x1 = x0 + xhossz * ex - yhossz * ey
    Line1.y1 = y0 + xhossz * ey + yhossz * ex
    Line1.x2 = x0 - xhossz * ex - yhossz * ey
    Line1.y2 = y0 - xhossz * ey + yhossz * ex

    Line2.x1 = x0 + xhossz * ex + yhossz * ey
    Line2.y1 = y0 + xhossz * ey - yhossz * ex
    Line2.x2 = x0 - xhossz * ex + yhossz * ey
    Line2.y2 = y0 - xhossz * ey - yhossz * ex

    Line3.x1 = Line1.x1
    Line3.x2 = Line2.x1
    Line3.y1 = Line1.y1
    Line3.y2 = Line2.y1

    Line4.x1 = Line1.x2
    Line4.x2 = Line2.x2
    Line4.y1 = Line1.y2
    Line4.y2 = Line2.y2
End Sub
