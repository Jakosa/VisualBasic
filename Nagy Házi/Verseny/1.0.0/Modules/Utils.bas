Attribute VB_Name = "Utils"
' Fejl�c
' K�sz�tette: Jakosa Csaba �rp�d
' Fejl�c v�ge

Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As EnSetWindowPos) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function SetViewportOrgEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Const WM_PAINT As Long = &HF&

' L�trehozza a Config v�ltoz�t.
Public Config As New Config
' L�trehozza a Map v�ltoz�t.
Public Map As New Map
' L�trehozza a VegeredmenyMentese v�ltoz�t.
Public VegeredmenyMentese As New VegeredmenyMentese

' Megn�zi hogy a v�ltoz� "boolean"-e.
' A "value" t�rolja v�ltoz� �rt�k�t.
Public Function IsBoolean(ByVal value As Variant) As Boolean
    ' Ha hiba van akkor ugrik a hibar�szhez.
    On Error GoTo ErrorHandler

    ' b v�ltoz� l�trehoz�sa.
    Dim b As Boolean
    ' Megn�zi hogy a "value" �rt�ke igaz-e vagy hamis. Val�s�gban megpr�b�lja konvert�lni.
    b = CBool(value)
    ' Ha sikeres a konvert�l�s akkor nincs hiba ez�rt igaz �rt�kre �ll�tja a visszat�r�si �rt�ket.
    IsBoolean = True

    ' Hiba eset�n ide ugrik a program.
ErrorHandler:
    ' Hamis (False) �rt�kkel l�p ki.
End Function

' Megn�zi hogy a v�ltoz� "integer"-e.
' A "value" t�rolja v�ltoz� �rt�k�t.
Public Function IsInteger(ByVal value As Variant) As Boolean
    ' Ha hiba van akkor ugrik a hibar�szhez.
    On Error GoTo ErrorHandler

    ' i v�ltoz� l�trehoz�sa.
    Dim i As Boolean
    ' Megn�zi hogy a "value" �rt�ke igaz-e vagy hamis. Val�s�gban megpr�b�lja konvert�lni.
    i = CInt(value)
    ' Ha sikeres a konvert�l�s akkor nincs hiba ez�rt igaz �rt�kre �ll�tja a visszat�r�si �rt�ket.
    IsInteger = True

    ' Hiba eset�n ide ugrik a program.
ErrorHandler:
    ' Hamis (False) �rt�kkel l�p ki.
End Function

' Megn�zi hogy a v�ltoz� "byte"-e.
' A "value" t�rolja v�ltoz� �rt�k�t.
Public Function IsByte(ByVal value As Variant) As Boolean
    ' Ha hiba van akkor ugrik a hibar�szhez.
    On Error GoTo ErrorHandler

    ' b v�ltoz� l�trehoz�sa.
    Dim b As Boolean
    ' Megn�zi hogy a "value" �rt�ke igaz-e vagy hamis. Val�s�gban megpr�b�lja konvert�lni.
    b = CByte(value)
    ' Ha sikeres a konvert�l�s akkor nincs hiba ez�rt igaz �rt�kre �ll�tja a visszat�r�si �rt�ket.
    IsByte = True

    ' Hiba eset�n ide ugrik a program.
ErrorHandler:
    ' Hamis (False) �rt�kkel l�p ki.
End Function

' Kisz�molja a pont �s szakasz t�vols�g�t.
' A "PointX" t�rolja a pont X koordin�t�j�t.
' A "PointY" t�rolja a pont Y koordin�t�j�t.
' A "LineX1" t�rolja a szakasz X1 koordin�t�j�t.
' A "LineX2" t�rolja a szakasz X2 koordin�t�j�t.
' A "LineY1" t�rolja a szakasz Y1 koordin�t�j�t.
' A "LineY2" t�rolja a szakasz Y2 koordin�t�j�t.
Public Function Distance(ByVal PointX As Single, ByVal PointY As Single, ByVal LineX1 As Single, ByVal LineX2 As Single, ByVal LineY1 As Single, ByVal LineY2 As Single) As Single
    ' V�ltoz�k.
    Dim AA As Single, BB As Single, CC As Single, DD As Single
    ' V�ltoz�k.
    Dim dot As Single, len_sq As Single, param As Single
    ' V�ltoz�k.
    Dim xx As Single, yy As Single

    ' X pont �s egyenes X1 koordin�t�j�nak k�l�nbs�ge.
    AA = PointX - LineX1
    ' Y pont �s egyenes Y1 koordin�t�j�nak k�l�nbs�ge.
    BB = PointY - LineY1
    ' Egyenes X2 �s egyenes X1 koordin�t�j�nak k�l�nbs�ge.
    CC = LineX2 - LineX1
    ' Egyenes Y2 �s egyenes Y1 koordin�t�j�nak k�l�nbs�ge.
    DD = LineY2 - LineY1

    ' V�ltoz�k szorzata �s �sszege.
    dot = AA * CC + BB * DD
    ' V�ltoz�k szorzata �s �sszege.
    len_sq = CC * CC + DD * DD
    ' K�t v�ltoz� h�nyadosa.
    param = dot / len_sq

    ' Akkor fut le ha a param kisebb mint 0.
    If param < 0 Then
        ' xx be�ll�t�sa.
        xx = LineX1
        ' yy be�ll�t�sa.
        yy = LineY1
    ' Akkor fut le ha a param nagyobb mint 1.
    ElseIf param > 1 Then
        ' xx be�ll�t�sa.
        xx = LineX2
        ' yy be�ll�t�sa.
        yy = LineY2
    Else
        ' xx be�ll�t�sa.
        xx = LineX1 + param * CC
        ' yy be�ll�t�sa.
        yy = LineY1 + param * DD
    End If

    ' Visszadja a pontok t�vols�g�t.
    Distance = Sqr(((PointX - xx) * (PointX - xx)) + ((PointY - yy) * (PointY - yy)))
End Function

' Megny�tja �s be�ll�tja a hibaablak param�tereit.
' A "Title" az ablak c�me.
' A "Message" az ablakra ki�rt hiba�zenet.
' A "Leallas" pedig a program le�ll�t�s�t t�rolja. Ha sz�ks�ges.
Public Sub WarningWindow(ByVal Title As String, ByVal Message As String, ByVal Leallas As Boolean)
    ' Hiba�zenet elt�rol�sa.
    WarningForm.HibaUzenet = Message
    ' Le�ll�t�s elt�rol�sa.
    WarningForm.Leallitas = Leallas
    ' Ablak c�mm�nek elt�rol�sa.
    WarningForm.Caption = Title
    ' Megjeleniti a hiba ablakot. K�zben megg�tolja hogy a m�g�tte l�v� ablakra r� lehessen kattintani (1 a tilt�s).
    WarningForm.Show 1
End Sub
