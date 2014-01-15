VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1695
      Left            =   2400
      TabIndex        =   1
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Angolul voltak dokument·lva a kÛdsorok de majd mindent ·tÌrok magyarra + ÈrthetıvÈ teszem. Nekem nem volt akkor lÈnyeg mi hogyan van elnevezve :D

Dim q As New TokenQueue
Dim sk As New TokenStack

Private Sub Command1_Click()
    'makepolishform(recvData.Args.substr(recvData.firstSpace+1));
    'calculate(recvData.Channel);

    'Dim cl As CalcLexer
    'cl = New CalcLexer(s)
    'Set cl = New CalcLexer
    'cl.Init "1+1"

    'Dim ss() As String, i As Long
    'ss = cl.strlower("Abc")

    'For i = LBound(ss) To UBound(ss)
    '    Debug.Print "asdddd: " & ss(i)
    'Next i

    Set q = New TokenQueue
    Set sk = New TokenStack

    'makepolishform "1+1"
    'makepolishform "1+1+1+1-1+10+10+30+20"
    'makepolishform "(1+1)+1+1"
    'makepolishform "1,1+1*5"
    'makepolishform "4^4"
    'makepolishform "4**4"
    makepolishform "(1+1)*4"
    'makepolishform "(1+1)+(1+3)" 'innÈtıl hib·k vannak
    'makepolishform "(1+1*3)^6-(4-5)"
    'makepolishform "(1+1*3)^6+(5+4)*400"
    'makepolishform "((120+13)+(50*6))-(10-30)^4"
    'makepolishform Text1.Text
    calculate
    Set q = Nothing
    Set sk = Nothing
End Sub

Private Function getValue(t As Token) As Token
    'Debug.Print "getValue"

    If t.tkntype = TKN_NUMBER Then
        Set getValue = t
        'Debug.Print "getValue1"
    Else
        'Debug.Print "getValue2"
        't.tkntype = TKN_NUMBER
        't.number = varv.getVar(t.str);
    '    Set getValue = t
    End If
End Function

Private Sub makepolishform(szam As String)
    Dim s As String

    s = szam
    Dim cl As New CalcLexer
    'cl = New CalcLexer(s)
    'Set cl = New CalcLexer
    cl.Init s

    Dim t As New Token

    Do
        'If t.tkntype = TKN_NEXT_LINE Then
            'cl = CalcLexer(s)
        '    Set cl = Nothing
        '    Set cl = New CalcLexer
        '    cl.Init s
        'End If

        Set t = cl.getNextToken

        If t.tkntype = TKN_NUMBER Or t.tkntype = TKN_NAME Then
            q.Tesz t
        ElseIf t.tkntype = TKN_OPEN_PAR Then
            sk.Tesz t
        ElseIf t.tkntype = TKN_CLOSE_PAR Then
            Do While Not sk.Ures
                If sk.Keresgel.tkntype = TKN_OPEN_PAR Then
                    Exit Do
                End If

                q.Tesz sk.Kap
            Loop

            sk.Kap
        ElseIf t.tkntype = TKN_OPERATOR Then
            Do While Not sk.Ures
                If getpriority(sk.Keresgel) > getpriority(t) Then
                    Exit Do
                End If

                q.Tesz sk.Kap
            Loop

            sk.Tesz t
        End If
    Loop While Not t.tkntype = TKN_TERMINATION

    Do While Not sk.Ures
        q.Tesz sk.Kap
    Loop
End Sub

Private Sub calculate()
    Set sk = Nothing
    Set sk = New TokenStack

    Do While Not q.Ures
        Dim t As Token
        'q.Lista
        Set t = q.Kap
        'Debug.Print "qqqqq0: " & t.tkntype
        'Debug.Print "qqqqq2: " & t.number

        If t.tkntype = TKN_NUMBER Or t.tkntype = TKN_NAME Then
            sk.Tesz t
        ElseIf t.tkntype = TKN_OPERATOR Then
            Dim o1 As New Token, o2 As New Token, e As New Token

            If t.str = "+" Then
                'q.Lista
                'sk.Lista
                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o2 = getValue(sk.Kap)
                    End If
                End If
                'q.Lista
                'sk.Lista
                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o1 = getValue(sk.Kap)
                    End If
                End If
                'q.Lista
                'sk.Lista
                'Debug.Print "j········j " & (o1.number)
                'Debug.Print "j········j " & (o2.number)
                'Debug.Print "j········j " & getValue(sk.Kap2(sk.inptr)).number
                'Debug.Print "j········j " & (o1.number + o2.number)
                'If o1.number = 4 Then
                'End If

                e.Init TKN_NUMBER, (o1.number + o2.number), ""
                sk.Tesz e
                'q.Lista
                'Debug.Print "j········j45 " & getValue(sk.Kap2(sk.inptr - 1)).number
                'Debug.Print "j········j45 " & getValue(sk.Kap2(sk.inptr)).number
                'sk.Lista
                'Dim dd As New Token
                'Set dd = getValue(sk.Kap2(sk.inptr))
                'If sk.inptr = 1 Then
                '    sk.Tesz New Token
                'End If
            ElseIf t.str = "-" Then
                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o2 = getValue(sk.Kap)
                    End If
                End If

                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o1 = getValue(sk.Kap)
                    End If
                End If

                'Debug.Print "j········j2"
                e.Init TKN_NUMBER, (o1.number - o2.number), ""
                sk.Tesz e
            ElseIf t.str = "*" Then
                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o2 = getValue(sk.Kap)
                    End If
                End If

                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o1 = getValue(sk.Kap)
                    End If
                End If

                'Debug.Print "j········j2"
                e.Init TKN_NUMBER, (o1.number * o2.number), ""
                sk.Tesz e
            ElseIf t.str = "/" Then
                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o2 = getValue(sk.Kap)
                    End If
                End If

                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o1 = getValue(sk.Kap)
                    End If
                End If

                Debug.Print "j········j2"
                e.Init TKN_NUMBER, (o1.number / o2.number), ""
                sk.Tesz e
            ElseIf t.str = "^" Or t.str = "**" Then
                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o2 = getValue(sk.Kap)
                    End If
                End If

                If Not sk.Ures Then
                    If sk.Keresgel.tkntype = TKN_NUMBER Or sk.Keresgel.tkntype = TKN_NAME Then
                        Set o1 = getValue(sk.Kap)
                    End If
                End If

                'Debug.Print "j········j2"
                e.Init TKN_NUMBER, (o1.number ^ o2.number), ""
                sk.Tesz e
            End If
        End If
    Loop

    Dim d As Double

    'Debug.Print "sk.Ures: " & sk.Ures
    If Not sk.Ures Then
        d = getValue(sk.Kap).number
        'Debug.Print "VÈgeredmÈny2: " & getValue(sk.Kap).number
    Else
        d = 0#
    End If

    Debug.Print "VÈgeredmÈny: " & d
    'Debug.Print "VÈgeredmÈny2: " & getValue(sk.Kap).number
End Sub
