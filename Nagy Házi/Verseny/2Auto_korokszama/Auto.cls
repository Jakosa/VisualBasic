VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Timer0 As VB.Timer
Attribute Timer0.VB_VarHelpID = -1
Dim WithEvents KVonal1 As VB.Line
Attribute KVonal1.VB_VarHelpID = -1
Dim WithEvents KVonal2 As VB.Line
Attribute KVonal2.VB_VarHelpID = -1
Dim WithEvents KVonal3 As VB.Line
Attribute KVonal3.VB_VarHelpID = -1
Dim WithEvents KVonal4 As VB.Line
Attribute KVonal4.VB_VarHelpID = -1
Dim VSzogTomb(1 To 30) As Double
Dim Color As ColorConstants
Dim Nyomvonal As Boolean
Dim Started As Boolean
Dim x0 As Double
Dim y0 As Double
Dim ex As Double
Dim ey As Double
Dim i As Long
Const PontosabbKor = True   ' Ha "true" akkor tökéletesebben középen van az autó. "false" esetén viszont szerintem élethübb.
Const xhossz = 80           ' Kocsi hossza.
Const yhossz = 60           ' Kocsi szélessége.
Dim ACelvonal As Boolean       ' Átment a célvonalon vagy sem.
Dim KorokSzama As Byte

Public Property Get GetKorokSzama() As Byte
    GetKorokSzama = KorokSzama
End Property

Public Sub SetX0(X As Double)
    x0 = X
End Sub

Public Sub SetY0(Y As Double)
    y0 = Y
End Sub

Public Sub SetEX(X As Double)
    ex = X
End Sub

Public Sub SetEY(Y As Double)
    ey = Y
End Sub

Public Sub SetColor(Szin As String)
    Select Case Szin
        Case "Kék", "kék"
            Color = ColorConstants.vbBlue
        Case "Piros", "piros"
            Color = ColorConstants.vbRed
        Case "Zöld", "zöld"
            Color = ColorConstants.vbGreen
        Case "Sárga", "sárga"
            Color = ColorConstants.vbYellow
        Case Else
            Color = vbBlack
    End Select

    KVonal1.BorderColor = Color
    KVonal2.BorderColor = Color
    KVonal3.BorderColor = Color
    KVonal4.BorderColor = Color
End Sub

Public Sub SetVisible(visible As Boolean)
    KVonal1.visible = visible
    KVonal2.visible = visible
    KVonal3.visible = visible
    KVonal4.visible = visible
End Sub

Public Sub SetBorderWidth(Nagysag As Byte)
    KVonal1.BorderWidth = Nagysag
    KVonal2.BorderWidth = Nagysag
    KVonal3.BorderWidth = Nagysag
    KVonal4.BorderWidth = Nagysag
End Sub

Public Sub SetNyomvonal(nyom As Boolean)
    Nyomvonal = nyom
End Sub

Public Sub Load(Szam As Byte)
    Set KVonal1 = Palya.Controls.Add("VB.Line", "KocsiVonal" & Szam & "1", Palya)
    Set KVonal2 = Palya.Controls.Add("VB.Line", "KocsiVonal" & Szam & "2", Palya)
    Set KVonal3 = Palya.Controls.Add("VB.Line", "KocsiVonal" & Szam & "3", Palya)
    Set KVonal4 = Palya.Controls.Add("VB.Line", "KocsiVonal" & Szam & "4", Palya)

    SetVisible True

    Set Timer0 = Palya.Controls.Add("VB.Timer", "Timer" & Szam & "1", Palya)
    Timer0.Enabled = False
    Timer0.Interval = 1

    Dim ciklus As Long
    For ciklus = 1 To 30
        VSzogTomb(ciklus) = Rnd * 0.1
    Next ciklus
End Sub

Public Sub Start()
    If Started Then
        Exit Sub
    End If

    Started = True
    Timer0.Enabled = True
End Sub

Public Sub Stop_Kocsi()
    If Not Started Then
        Exit Sub
    End If

    Started = False
    Timer0.Enabled = False
End Sub

Private Sub Class_Initialize()
    ex = 0.6
    ey = -1
    x0 = 1100
    y0 = 4000
    KorokSzama = 0
    SetNyomvonal True
    ACelvonal = False
End Sub

Private Sub Class_Terminate()
    Set Timer0 = Nothing
End Sub

Private Sub Timer0_Timer()
    i = i + 1

    If i > 30 Then
        i = 1
    End If

    NextCoordinate True, True

    Dim xb As Single, yb As Single, xj As Single, yj As Single
    ' Ideálisan középen van a jármû ha 400*300-as a méret.
    If PontosabbKor Then
        xb = x0 + 400 * ex - 300 * ey
        yb = y0 + 400 * ey + 300 * ex
        xj = x0 + 400 * ex + 300 * ey
        yj = y0 + 400 * ey - 300 * ex
    Else
        xb = x0 + xhossz * ex - yhossz * ey
        yb = y0 + xhossz * ey + yhossz * ex
        xj = x0 + xhossz * ex + yhossz * ey
        yj = y0 + xhossz * ey - yhossz * ex
    End If

    Dim dist As Single, ciklus As Integer
    dist = 1000000

    For ciklus = Palya.Line5.LBound To Palya.Line5.UBound
        Dim d As Single
        d = Distance(xb, yb, Palya.Line5(ciklus).X1, Palya.Line5(ciklus).X2, Palya.Line5(ciklus).Y1, Palya.Line5(ciklus).Y2)

        If dist > d Then
            dist = d
        End If
    Next ciklus

    ' Ideálisan középen van a jármû ha 400*300-as a méret.
    If PontosabbKor Then
        If dist < 200 Then
            NextCoordinate True, False
            NextMove
        ElseIf dist > 200 Then
            NextCoordinate False, False
            NextMove
        Else
            x0 = x0 - 55 * ex
            y0 = y0 - 55 * ey
        End If
    Else
        If dist < 270 Then
            NextCoordinate True, False
            NextMove
        ElseIf dist > 270 Then
            NextCoordinate False, False
            NextMove
        Else
            x0 = x0 - 55 * ex
            y0 = y0 - 55 * ey
        End If
    End If
    
    NextMove

    x0 = x0 + 50 * ex
    y0 = y0 + 50 * ey
    NextMove

    dist = Distance(x0, y0, Palya.CelVonal.X1, Palya.CelVonal.X2, Palya.CelVonal.Y1, Palya.CelVonal.Y2)

    If dist < 100 And Not ACelvonal Then
        ACelvonal = True
        KorokSzama = KorokSzama + 1
        Palya.List1.AddItem "Beértem a célba: " & Color & " Korok száma: " & KorokSzama
    ElseIf dist > 100 And ACelvonal Then
        ACelvonal = False
    End If

    If Nyomvonal Then
        Palya.Circle (x0, y0), 10, Color
    End If
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
    KVonal1.X1 = x0 + xhossz * ex - yhossz * ey
    KVonal1.Y1 = y0 + xhossz * ey + yhossz * ex
    KVonal1.X2 = x0 - xhossz * ex - yhossz * ey
    KVonal1.Y2 = y0 - xhossz * ey + yhossz * ex

    KVonal2.X1 = x0 + xhossz * ex + yhossz * ey
    KVonal2.Y1 = y0 + xhossz * ey - yhossz * ex
    KVonal2.X2 = x0 - xhossz * ex + yhossz * ey
    KVonal2.Y2 = y0 - xhossz * ey - yhossz * ex

    KVonal3.X1 = KVonal1.X1
    KVonal3.X2 = KVonal2.X1
    KVonal3.Y1 = KVonal1.Y1
    KVonal3.Y2 = KVonal2.Y1

    KVonal4.X1 = KVonal1.X2
    KVonal4.X2 = KVonal2.X2
    KVonal4.Y1 = KVonal1.Y2
    KVonal4.Y2 = KVonal2.Y2
End Sub
