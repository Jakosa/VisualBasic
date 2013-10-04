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
      Interval        =   80
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
Dim i As Long, j As Long
Dim x0 As Double
Dim y0 As Double
Dim ex As Double
Dim ey As Double
Dim A(1 To 30) As Double
Dim XV(1 To 10000) As Double, YV(1 To 10000) As Double
Const xhossz = 80
Const yhossz = 60
Const MaxPontok = 10

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
        A(ciklus) = Rnd * 0.1
    Next ciklus

    For vonal = Line5.LBound To Line5.UBound
        For ciklus = 0 To MaxPontok
            pontok = pontok + 1
            XV(pontok) = (Line5(vonal).X1 * ciklus + Line5(vonal).X2 * (MaxPontok - ciklus)) / MaxPontok
            YV(pontok) = (Line5(vonal).Y1 * ciklus + Line5(vonal).Y2 * (MaxPontok - ciklus)) / MaxPontok
        Next ciklus
    Next vonal
End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub Timer1_Timer()
    Dim xb As Double, yb As Double, xj As Double, yj As Double
    Dim mb As Double, mj As Double, ciklus As Long
    i = i + 1
    j = j + 1

    If i > 30 Then
        i = 1
    End If

    NextCoordinate True, True
    
    xb = x0 + 400 * ex - 300 * ey
    yb = y0 + 400 * ey + 300 * ex
    xj = x0 + 400 * ex + 300 * ey
    yj = y0 + 400 * ey - 300 * ex

    mb = 10000000
    mj = 10000000

    For ciklus = LBound(XV) To UBound(XV)
        If mb > (xb - XV(ciklus)) ^ 2 + (yb - YV(ciklus)) ^ 2 Then
            mb = (xb - XV(ciklus)) ^ 2 + (yb - YV(ciklus)) ^ 2
        End If

        If mj > (xj - XV(ciklus)) ^ 2 + (yj - YV(ciklus)) ^ 2 Then
            mj = (xj - XV(ciklus)) ^ 2 + (yj - YV(ciklus)) ^ 2
        End If
    Next ciklus

    Debug.Print mb - mj
    
    If mb - mj > 20000 Then
        NextCoordinate False, False
        NextMove

    ElseIf mb - mj < -20000 Then
        NextCoordinate True, False
        NextMove
    Else
        x0 = x0 - 55 * ex
        y0 = y0 - 55 * ey
    End If
    
    NextMove

    x0 = x0 + 50 * ex
    y0 = y0 + 50 * ey
    NextMove

    Me.Circle (x0, y0), 20, vbRed
End Sub

Private Sub NextCoordinate(BIrany As Boolean, ismeretlen As Boolean)
    Dim bj As Double

    If BIrany Then
        bj = -0.15
    Else
        bj = 0.15
    End If

    If ismeretlen Then
        ex = Cos(A(i)) * ex - Sin(A(i)) * ey
        ey = Cos(A(i)) * ey + Sin(A(i)) * ex
    Else
        ex = Cos(bj) * ex - Sin(bj) * ey
        ey = Cos(bj) * ey + Sin(bj) * ex
    End If

    ex = ex / Sqr(ex * ex + ey * ey)
    ey = ey / Sqr(ex * ex + ey * ey)
End Sub

Private Sub NextMove()
    Line1.X1 = x0 + xhossz * ex - yhossz * ey
    Line1.Y1 = y0 + xhossz * ey + yhossz * ex
    Line1.X2 = x0 - xhossz * ex - yhossz * ey
    Line1.Y2 = y0 - xhossz * ey + yhossz * ex

    Line2.X1 = x0 + xhossz * ex + yhossz * ey
    Line2.Y1 = y0 + xhossz * ey - yhossz * ex
    Line2.X2 = x0 - xhossz * ex + yhossz * ey
    Line2.Y2 = y0 - xhossz * ey - yhossz * ex

    Line3.X1 = Line1.X1
    Line3.X2 = Line2.X1
    Line3.Y1 = Line1.Y1
    Line3.Y2 = Line2.Y1

    Line4.X1 = Line1.X2
    Line4.X2 = Line2.X2
    Line4.Y1 = Line1.Y2
    Line4.Y2 = Line2.Y2
End Sub
