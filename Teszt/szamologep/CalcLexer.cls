VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcLexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim str As String
Dim pos As Integer
Dim length As Integer
Const swspace = " \n\r\t"
Const snumchr1 = "0123456789"
Const snumchr2 = "0123456789,_hxrob"
Const snumbers = "_0123456789,"
Const s1cops = "+-/*%^()"
Const s2cops = " ** "
Const upletters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ[]{}."
Const downletters = "abcdefghijklmnopqrstuvwxyz()(),"

Public Sub Init(mystr As String)
    str = strlower(mystr)
    pos = 0
    Debug.Print str
    Debug.Print Len(str)
    length = Len(str)
End Sub

Public Function strlower(s As String) As String
    Dim i As Integer
    Dim sa() As String
    sa = str2Array(s)
    'Debug.Print UBound(sa) - LBound(sa) + 1 & " vs "; Len(s)

    For i = LBound(sa) To UBound(sa)
        'Debug.Print sa(i)

        If InStr(upletters, sa(i)) Then
            Dim ds() As String
            ds = str2Array(downletters)
            'Debug.Print "ggggg: " & InStr(upletters, sa(i)) & " " & sa(i)
            sa(i) = ds(InStr(upletters, sa(i)) - 1)
        End If
        'Debug.Print sa(i)
    Next i

    For i = LBound(sa) To UBound(sa)
        strlower = strlower & sa(i)
    Next i
End Function

Private Function c_in_s(c As String, s As String) As Boolean
    c_in_s = InStr(s, c)
End Function

Private Function s_in_s(kw As String, s As String) As Boolean
    s_in_s = InStr(s, " " & kw & " ")
End Function

Public Function getFloat() As String
    Dim s As String
    Dim strarray() As String
    s = ""

    If pos < length Then
        strarray = str2Array(str)

        Do While pos < length
            If Not c_in_s(strarray(pos), snumbers) Then
                Exit Do
            End If

            If Not strarray(pos) = "_" Then
                s = s & strarray(pos)
            End If
        
            pos = pos + 1
        Loop
    End If

    pos = pos - 1
    getFloat = s
    'Debug.Print "GetFloat: " & s
End Function

Public Function getNextToken() As Token
    Dim tkn As New Token
    Dim strarray() As String

    If pos < length Then
        strarray = str2Array(str)

        Do While pos < length
            If Not c_in_s(strarray(pos), swspace) Then
                Exit Do
            End If

            pos = pos + 1
        Loop
    End If

    If pos < length Then
        Dim tmps As String
        strarray = str2Array(str)
        'Debug.Print "dádá: " & strarray(pos)

        '2 ch long ops
        If pos + 1 < length Then
            tmps = ""
            tmps = tmps & strarray(pos)
            tmps = tmps & strarray(pos + 1)

            If s_in_s(tmps, s2cops) Then
                tkn.tkntype = TKN_OPERATOR
                tkn.str = tmps
                pos = pos + 1 ' because its 2 char
            Else
GoTo asd
            End If
        Else
asd:
            ' 1 ch long ops
            If c_in_s(strarray(pos), s1cops) Then
                If strarray(pos) = "(" Then
                    tkn.tkntype = TKN_OPEN_PAR
                    tkn.str = "("
                ElseIf strarray(pos) = ")" Then
                    tkn.tkntype = TKN_CLOSE_PAR
                    tkn.str = ")"
                Else
                    tkn.tkntype = TKN_OPERATOR
                    tkn.str = strarray(pos)
                End If
            ' numbers
            ElseIf (InStr(snumchr1, strarray(pos))) Or strarray(pos) = "," Then
                tkn.tkntype = TKN_NUMBER
    
                If ((pos + 1) < length) Then
                    If c_in_s(strarray(pos + 1), snumchr2) Then
                        If c_in_s(strarray(pos + 1), snumbers) Then
                            tkn.number = CDbl(getFloat)
                        End If
                    Else
                        tkn.number = CDbl(strarray(pos))
                    End If
                Else
                    tkn.number = CDbl(strarray(pos))
                End If
            End If
        End If

        pos = pos + 1
    Else
        tkn.tkntype = TKN_TERMINATION
    End If

    Set getNextToken = tkn
End Function
