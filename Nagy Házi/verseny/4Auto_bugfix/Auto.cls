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
Dim KorokSzama As Byte
Dim EgyKorido As Date ' Egy kör
Dim StopKorido As Date
Dim BKorido As Boolean      ' Induláskor állítja csak be a köidõt az adott idõre.
Dim BStopKorido As Boolean
Dim EgySzektorido As Date  ' Egy szektor
Dim ASzektorvonal As Boolean ' Átment a szektoron.
Dim LegjobbSzektorido(1 To 3) As Single
Dim LegjobbKorido As Single
Dim LegjobbKoridoSzama As Byte
Dim AutoSzama As Byte
Dim nyvp As Long

Private Type NyomvonalPontok
    x0 As Double
    y0 As Double
End Type

Private Type Korok
    Korido As Single
    Szektorido(1 To 3) As Single
End Type

Dim KorokTomb() As Korok
Dim NyPontok(0 To 1000000) As NyomvonalPontok

Public Property Get GetKorokSzama() As Byte
    GetKorokSzama = KorokSzama
End Property

Public Property Get GetKorSzektoridok(ii As Integer) As Single
    GetKorSzektoridok = KorokTomb(KorokSzama).Szektorido(ii)
End Property

Public Property Get GetLegjobbKorido() As Single
    GetLegjobbKorido = LegjobbKorido
End Property

Public Property Get GetLegjobbKoridoSzama() As Byte
    GetLegjobbKoridoSzama = LegjobbKoridoSzama
End Property

Public Property Get GetLegjobbSzektoridok(ii As Integer) As Single
    GetLegjobbSzektoridok = LegjobbSzektorido(ii)
End Property

Public Property Get GetColor() As String
    Dim szin As String

    Select Case Color
        Case ColorConstants.vbBlue
            szin = "kék"
        Case ColorConstants.vbRed
            szin = "piros"
        Case ColorConstants.vbGreen
            szin = "zöld"
        Case ColorConstants.vbYellow
            szin = "sárga"
        Case Else
            szin = "fekete"
    End Select

    GetColor = szin
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

Public Sub SetColor(szin As String)
    Select Case szin
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

Public Sub Load(szam As Byte)
    AutoSzama = szam
    Set KVonal1 = Palya.Controls.Add("VB.Line", "KocsiVonal" & szam & "1", Palya)
    Set KVonal2 = Palya.Controls.Add("VB.Line", "KocsiVonal" & szam & "2", Palya)
    Set KVonal3 = Palya.Controls.Add("VB.Line", "KocsiVonal" & szam & "3", Palya)
    Set KVonal4 = Palya.Controls.Add("VB.Line", "KocsiVonal" & szam & "4", Palya)

    SetVisible True

    Set Timer0 = Palya.Controls.Add("VB.Timer", "Timer" & szam & "1", Palya)
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

    If BKorido Then
        EgyKorido = Now
        EgySzektorido = EgyKorido
        BKorido = False
    End If

    Started = True
    Timer0.Enabled = True
End Sub

Public Sub Stop_Kocsi()
    If Not Started Then
        Exit Sub
    End If

    Started = False
    StopKorido = Now
    BStopKorido = True
    Timer0.Enabled = False
End Sub

Private Sub Class_Initialize()
    Dim ciklus As Integer

    ex = 0.6
    ey = -1
    x0 = 1100
    y0 = 4000
    nyvp = 0
    BKorido = True
    BStopKorido = False
    KorokSzama = Palya.GetKezdokorErteke()
    SetNyomvonal True
    ASzektorvonal = False
    ReDim KorokTomb(Palya.GetKezdokorErteke() To Palya.GetMKorokSzama()) As Korok

    LegjobbKorido = 10000000   ' Kezdõérték

    For ciklus = LBound(LegjobbSzektorido) To UBound(LegjobbSzektorido)
        LegjobbSzektorido(ciklus) = 10000000  ' Kezdõérték
    Next ciklus
End Sub

Private Sub Class_Terminate()
    Set KVonal1 = Nothing
    Set KVonal2 = Nothing
    Set KVonal3 = Nothing
    Set KVonal4 = Nothing
    Set Timer0 = Nothing
End Sub

Public Sub Dispose()
    Dim i As Long
    For i = 1 To 4
        Palya.Controls.Remove "KocsiVonal" & AutoSzama & i
    Next i

    Palya.Controls.Remove "Timer" & AutoSzama & "1"

    If nyvp > 0 Then
        For i = LBound(NyPontok) To nyvp
            If NyPontok(i).x0 > 0 And NyPontok(i).y0 > 0 Then
                Palya.Circle (NyPontok(i).x0, NyPontok(i).y0), 10, &H8000000E
            End If
        Next i
    End If
End Sub

Public Sub Show()
    NextMove
End Sub

Private Sub Timer0_Timer()
    i = i + 1

    If i > 30 Then
        i = 1
    End If

    If BStopKorido Then
        Dim NowTime As Date
        NowTime = Now
        EgyKorido = DateAdd("s", DateDiff("s", StopKorido, NowTime), EgyKorido)
        EgySzektorido = DateAdd("s", DateDiff("s", StopKorido, NowTime), EgySzektorido)
        BStopKorido = False
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

    Szektorok

    If Nyomvonal And nyvp < UBound(NyPontok) Then
        NyPontok(nyvp).x0 = x0
        NyPontok(nyvp).y0 = y0
        nyvp = nyvp + 1
    
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

' Két szektoridõ mérése értelemszerüen távolság alapján ha elérte.
' 0-tól 1-ig tartanak a szektorvonalak mert a célvonal egyben a harmadik szektor.
Private Sub Szektorok()
    If KorokSzama > Palya.GetMKorokSzama() Then
        Exit Sub
    End If

    Dim NextKor As Boolean, ii As Byte
    For ii = LBound(KorokTomb(KorokSzama).Szektorido) To UBound(KorokTomb(KorokSzama).Szektorido)
        Dim dist As Single, ciklus As Byte, szam As Byte
        dist = 1000000

        For ciklus = Palya.SzektorVonal.LBound To Palya.SzektorVonal.UBound
            Dim d As Single
            d = Distance(x0, y0, Palya.SzektorVonal(ciklus).X1, Palya.SzektorVonal(ciklus).X2, Palya.SzektorVonal(ciklus).Y1, Palya.SzektorVonal(ciklus).Y2)

            If dist > d Then
                dist = d
                szam = ciklus + 1 ' Mert egyel magasabbon vannak
            End If
        Next ciklus

        If dist < 100 And Not ASzektorvonal Then
            Dim NowTime As Date
            NowTime = Now
            ASzektorvonal = True

            If szam = 3 Then
                NextKor = True
                KorokTomb(KorokSzama).Korido = DateDiff("s", EgyKorido, NowTime)

                If LegjobbKorido > KorokTomb(KorokSzama).Korido Then
                    LegjobbKorido = KorokTomb(KorokSzama).Korido
                    LegjobbKoridoSzama = KorokSzama
                End If

                KorokTomb(KorokSzama).Szektorido(szam) = KorokTomb(KorokSzama).Korido

                Dim j As Byte
                For j = LBound(KorokTomb(KorokSzama).Szektorido) To UBound(KorokTomb(KorokSzama).Szektorido) - 1 ' Mert utolsót nem akarjuk kivonni
                    KorokTomb(KorokSzama).Szektorido(szam) = KorokTomb(KorokSzama).Szektorido(szam) - KorokTomb(KorokSzama).Szektorido(j)
                Next j
            Else
                KorokTomb(KorokSzama).Szektorido(szam) = DateDiff("s", EgySzektorido, NowTime)
            End If

            If LegjobbSzektorido(szam) > KorokTomb(KorokSzama).Szektorido(szam) Then
                LegjobbSzektorido(szam) = KorokTomb(KorokSzama).Szektorido(szam)
            End If

            NowTime = Now

            If szam = 3 Then
                EgyKorido = NowTime
            End If

            EgySzektorido = NowTime
        ElseIf dist > 100 And ASzektorvonal Then
            ASzektorvonal = False
        End If
    Next ii

    If NextKor Then
        KorokSzama = KorokSzama + 1
    End If
End Sub
