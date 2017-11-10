


Imports System.Net.Mail
Imports System.Net.Sockets
Imports System.Net


Module G9


    Dim ALPHA As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-.@:/#_\!*()<>{}[],?=+%$;abcdefghijklmnopqrstuvwxyz'`^|&~ "
   


    '-FOR MESSAGE HASHING--------------------------------------------------------------------

    Function Hash4D7(ByVal f As String) As Integer()
        Dim Plaintext As String = f
        Dim SBOX() As Integer
        Dim x As Integer
        Dim kChar As String, Fchar As String
        Dim MAP() As Integer
        Dim _0xFF As Integer = 255
        Dim Op As Integer

        ReDim MAP(Plaintext.Length)
        ReDim SBOX(Alpha.Length)

        Dim Lent As Integer = (Plaintext.Length)

        For i As Integer = 1 To UBound(SBOX)
            x = x + 32
            SBOX(i) = x Xor _0xFF
        Next i

        For t As Integer = 1 To Plaintext.Length
            kChar = Strings.Mid(Plaintext, t, 1)
            For y As Integer = 1 To ALPHA.Length
                Fchar = Mid(ALPHA, y, 1)
                If kChar = Fchar Then
                    MAP(t) = MAP(t) + ((SBOX(y) Xor (_0xFF) + (t * 32)) * Lent) >> t
                    Do While Not Op = MAP.GetUpperBound(0)
                        Op += 1
                        MAP(Op) = MAP(Op) + (SBOX(y))
                    Loop
                    Op = 0
                End If
            Next y
        Next t
        Return MAP
    End Function


    '-FOR MESSAGE ENCRYPTION-----------------------------------------------------------------

    Public Function Encrypt4D7(ByVal f As String, ByVal Master_key As Integer) As String


        Const _B4 As Integer = 2
        Dim DEG As String = """"

        Dim Plaintext As String = f
        Dim kChar As String = ""
        Dim Fchar As String = ""
        Dim Cipher As String = ""

        Dim SBOX(,) As UInt32    'STATE BOX
        Dim XBOX(,) As UInt32    'SUBSTITUTION BOX
        Dim KEYBOX(,) As UInt32  'SECURITY KEY BOX
        Dim KEY As UInt32 = 0
        Dim IV As UInt32 = 0
        Dim iX As Integer = 0
        Dim dX As Integer = 0
        Dim P As Integer
        Dim L As Integer
        Dim oCC As Integer
        Dim AlpaLent As Integer = Len(ALPA)
        Dim ANS(AlpaLent) As String
        Dim UBX As Integer = 20
        Try
            For iX = 1 To AlpaLent          '/* function for randomizing alpha character */
                P = Int(Rnd() * AlpaLent) + 1
                ANS(iX) = P

                For dX = 1 To AlpaLent
                    L = L + 1
                    If ANS(dX) = P Then
                        oCC += 1
                    End If

                    If oCC = 2 Then
                        oCC = 0
                        iX = iX - 1
                        Exit For
                    End If
                Next dX
                oCC = 0
                L = 0
            Next iX

            'Dim k As Func(Of String, String) = AddressOf Asc
            'Dim k1 As Func(Of String, String) = Function(s) s.ToString

            Dim MASTERKEY() As Integer = Array.CreateInstance(GetType(Integer), (8) + 1)

            For Y As Integer = 1 To Len(CStr(Master_key))
                MASTERKEY(Y) = CInt(Mid(Master_key, Y, 1) * 10)
            Next Y

            Dim JKEY_INTERCEPT As Array = Array.CreateInstance(GetType(Int16), 20)

            ReDim SBOX(5, UBX)
            ReDim KEYBOX(5, UBX)
            ReDim XBOX(5, UBX)

            Dim KEYLOCK(8) As Integer

            For _key1 As UInt32 = 1 To 5
                IV = IV + _B4
                For _key2 As UInt32 = 1 To UBX
                    KEY = KEY + IV
                    KEYBOX(_key1, _key2) = KEY
                Next _key2
                KEY = 0
            Next _key1

            KEY = 1028

            For _1S As Integer = 1 To UBound(SBOX, 1)
                For _2S As Integer = 1 To UBound(SBOX, 2)
                    KEY = KEY + 4
                    SBOX(_1S, _2S) = (KEY)
                Next _2S
            Next _1S

            Dim Counter As Integer = 0

            'XOR KEYBOX TO SBOX

            For _1s As Integer = 1 To 5
                For _2s As Integer = 1 To UBX
                    XBOX(_1s, _2s) = KEYBOX(_1s, _2s) Xor (SBOX(_1s, _2s) * AddKeyArray(MASTERKEY))
                Next _2s
            Next _1s

            Dim C2 As String


            For T As Integer = 1 To Len(f)

                Fchar = Mid(f, T, 1)

                For P = 1 To ALPHA.Length
                    kChar = Mid(ALPHA, P, 1)
                    If Fchar = kChar Then

                        For _1S As Integer = 1 To 5

                            For _2s As Integer = 1 To UBX
                                Counter = Counter + 1
                                If Counter = P Then
                                    C2 = Hex((XBOX(_1S, _2s)))
                                    If Len(C2) < 5 Then

                                    End If
                                    Cipher = Cipher & Hex((XBOX(_1S, _2s)))
                                End If
                            Next _2s
                        Next _1S
                        Counter = 0
                    End If
                Next P
            Next T

            Return (Cipher)

        Catch e As Exception
            Return e.Message
        End Try

    End Function


    Public Function Decrypt4D7(ByVal CipherText As String, ByVal Master_key As Integer) As String


        Const _B4 As Integer = 2

        Dim DEG As String = """"
        Dim kChar As String = ""
        Dim Fchar As String = ""
        Dim Cipher As String = ""

        Dim SBOX(,) As UInt32    'STATE BOX
        Dim XBOX(,) As UInt32 'SUBSTITUTION BOX
        Dim KEYBOX(,) As UInt32  'SECURITY KEY BOX
        Dim KEY As UInt32 = 0
        Dim IV As UInt32 = 0
        Dim iX As Integer = 0
        Dim dX As Integer = 0
        Dim UBX As Integer = 20

        Dim MASTERKEY() As Integer = Array.CreateInstance(GetType(Integer), (8) + 1)

        For Y As Integer = 1 To Len(CStr(Master_key))
            MASTERKEY(Y) = CInt(Mid(Master_key, Y, 1) * 10)
        Next Y

        ReDim SBOX(5, UBX)
        ReDim KEYBOX(5, UBX)
        ReDim XBOX(5, UBX)

        Dim KEYLOCK(8) As Integer

        For _key1 As UInt32 = 1 To 5
            IV = IV + _B4
            For _key2 As UInt32 = 1 To UBX
                KEY = KEY + IV
                KEYBOX(_key1, _key2) = KEY
            Next _key2
            KEY = 0
        Next _key1

        KEY = 1028

        For _1S As Integer = 1 To UBound(SBOX, 1)
            For _2S As Integer = 1 To UBound(SBOX, 2)
                KEY = KEY + 4
                SBOX(_1S, _2S) = (KEY)
            Next _2S
        Next _1S

        Dim Counter As Integer = 0

        ' XOR KEYBOX TO SBOX

        For _1s As Integer = 1 To 5
            For _2s As Integer = 1 To UBX
                XBOX(_1s, _2s) = KEYBOX(_1s, _2s) Xor (SBOX(_1s, _2s) * AddKeyArray(MASTERKEY))
            Next _2s
        Next _1s

        Dim CipherTXT As String = CipherText
        Dim C1 As String
        Dim C2 As Long
        Dim Lent As Integer = 5

        Dim TPos As Integer = Len(CipherTXT) / 5
        Dim Pos As Integer = 1
        Dim PosX As Integer
        Dim CharC As String
        Dim NL As String = "`"

        For i As Integer = 1 To TPos

            C1 = Mid(CipherTXT, Pos, Lent)
            Pos = Pos + 5

            C2 = "&H" & C1

            For _1x As Integer = 1 To 5
                For _2x As Integer = 1 To UBX
                    Counter += 1
                    If XBOX(_1x, _2x) = C2 Then
                        PosX = Counter
                        CharC = CharC & Mid(ALPHA, PosX, 1)
                        If CharC = NL Then
                            CharC = CharC & vbCr
                        End If
                    End If
                Next _2x
            Next _1x

            Counter = 0

        Next i
        Return CharC
    End Function

    Function AddKeyArray(ByVal ARR() As Integer) As Integer

        Dim Sum As Integer

        For i As Integer = 1 To UBound(ARR)
            Sum = Sum + CInt(ARR(i)) Xor i
        Next i
        Sum = Sum Xor 8
        Return Sum

    End Function

    Function IsEqual(ByVal a As Object, ByVal b As Object) As Boolean

        If a = b Then
            Return True
        Else
            Return False
        End If

    End Function

    '-FOR MESSAGE ENCRYPTION-----------------------------------------------------------------
End Module
