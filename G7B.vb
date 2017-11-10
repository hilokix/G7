'Imports krypt = System.Security.Cryptography
Imports System.Text
Imports System.Net
Imports SYSD = System.Diagnostics
Imports System.Threading
Imports System.Net.Mail


Public Module MailStorm

    Sub MailSample()

        Dim mailer As New SmtpClient
        Dim [sender] As New MailAddress("me@yahoo.com")
        Dim [reciever] As New MailAddress("him@yahoo.com")

        Dim Messages As MailMessage

    End Sub




End Module



Public Module G7B
    Public Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

    Public ErrorTracer As String
    Public D7Key As Integer = 0

    Public Pwd As String = ""
    Public C_Path As String = ""

    Public kBi() As Integer = {55, 66, 77, 84, 58}
    Dim Master_Key_On As Boolean
    Dim KEYMASTER As ULong
    Public C2Matrix(,) As Long
    Dim k1 As Integer : Dim k2 As Integer : Dim k3 As Integer
    Dim k4 As Integer : Dim k5 As Integer : Dim k6 As Integer
    Dim k7, k8, k9 As Integer
    Dim k10, k11, k12 As Integer
    Const DEG As String = """"
    Const lf As String = vbCr
    Public ALPA As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890!@#$%^&*()_-+={}[]|\<>,.?/`;:'~ "

    Dim ROWBOUND As Integer = Len(ALPA)
    Dim MATRIX(ROWBOUND, 1) As String
    Dim Separator As String = vbCr
    Dim NLF As String = "|"

    Const SUCCESS = 0
    Public KEYTEMP As String = ""
    Public CPT As String
    Private Achar As String
    Private Row1 As String
    Private Row2 As String
    Private Row3 As String
    Private k As Integer
    Private FactorL(3000) As Long
    Private FactorR(3000) As Long
    Private fP(3000) As Long
    Private msgs As String
    Private Lenght_list As String
    Private optlen As Integer
    Private x1Array() As Object
    Private vara As Integer : Private varb As Integer : Private varc As Integer
    Private PNum As Integer
    Private PrevHash As String
    Private xhex, r1, r2, r3
    Private n2(32) As Integer
    Private n0(32) As Integer
    Private ftemp As String
    Private FinalProd(32) As Long
    Private lent As Integer
    Private XBOX(1024, 1024) As String
    Private keylent As String
    Private ArrayPos(9000) As Long
    Private Template As String
    Private templateb As String
    Private arr(16) As String
    Private intDEC(3000) As String
    Private xy As Long
    Private yx As Long
    Private CTEM As Boolean
    Private KillBox As Integer
    Private KEYTEMPLATE As String
    Dim xP As Long
    Dim t As Long
    Private kENC As String
    Dim rq(1000) As String
    Dim lowBound As Integer
    Dim hiBound As Integer

    Dim output As String
    Dim LSK As Boolean = False
    Dim GenKeyOutput() As Long
    Dim CHOSENKEYR As Integer
    Dim CHOSENKEYL As Integer
    Dim CHOSENKEYM As Integer
    Public GLOBALTEXT As String

    Public BAlpha(,) As String = {{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
                                "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", _
                                "U", "V", "W", "X", "Y", "Z", "1", "2", "3", "4", "5"}, _
                                {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
                                 "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", _
                                 "u", "v", "w", "x", "y", "z", "6", "7", "8", "9", "0"} _
                                 }
    Function IntLen(ByVal n As Integer) As Integer

        Dim j As Integer
        Dim temp As String
        temp = CStr(n)
        For j = 1 To Len(temp) - 1 : Next
        Return j

    End Function


    Public Function GENERATEPWD() As Long()

        Dim _1Byte As Integer = 8

        On Error Resume Next
        Dim PadLock_key(32) As Long

        Dim ii As Integer
        Dim pwASC(32) As Long
        Dim pwL As Integer
        Dim Al As String

        Dim bits As Byte()
        Dim pwdShiftedAsc(32) As Long
        Dim evCount As Integer
        Dim oddCount As Integer
        Dim charPos As Integer
        Dim kG As String
        Dim kR As Integer
        Dim kT As Integer

        Dim CAlpha(,) As String = {{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
                                  "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", _
                                  "U", "V", "W", "X", "Y", "Z", "1", "2", "3", "4", "5"}, _
                                  {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
                                   "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", _
                                   "u", "v", "w", "x", "y", "z", "6", "7", "8", "9", "0"} _
                                   }
        '/------------------------------------------/
        Pwd = "TERMINATOR" ' THIS IS THE MASTERKEY STRING
        '/------------------------------------------/

        Dim CMatrix(1, 31) As Long
        Dim ix As Integer = 2 '256
        Dim ip As Integer = 32 '602
        Dim IncV As Integer = 8

        For ii = 0 To CMatrix.GetUpperBound(1) - 1
            ix += IncV
            ip += IncV
            CMatrix(0, ii) = ix
            CMatrix(1, ii) = ip
        Next

        If Pwd.Length <> 0 Then
            Master_Key_On = True
            Console.WriteLine("MASTER-KEY : " & Master_Key_On)
            Console.WriteLine("")
        End If

        If Pwd.Length = 0 Then
            Master_Key_On = False
            Console.WriteLine("MASTER-KEY : " & Master_Key_On)
            Console.WriteLine("")
        End If

        C2Matrix = CMatrix
        pwL = Pwd.Length

        For ii = 1 To pwL
            Al = Mid(Pwd, ii, 1)
            For ix = 0 To CAlpha.GetUpperBound(0)
                For ip = 0 To CAlpha.GetUpperBound(1)
                    If Al = CAlpha(ix, ip) Then
                        pwASC(ii) = CMatrix(ix, ip)
                        kR = pwASC(ii)

                        kT = IntLen(kR)

                        If (kT) = 1 Then
                            kG = Format(kR, "00#")
                        ElseIf (kT) = 2 Then
                            kG = Format(kR, "0##")
                        Else
                            kG = pwASC(ii)
                        End If

                    End If
                Next ip
            Next ix
        Next ii

        ii = 0

        oddCount = -1
        evCount = 0

        For ii = 1 To pwL '/ 2
            evCount += 1
            oddCount += 1
            ' pwdShiftedAsc(ii + oddCount) = pwASC(ii + evCount)
            ' pwdShiftedAsc(ii + evCount) = pwASC(ii + oddCount)
            pwdShiftedAsc(ii) = pwASC(ii)
        Next ii

        For ii = 1 To pwL
            PadLock_key(ii) = (pwdShiftedAsc(ii)) + (ii Xor _1Byte)
        Next

        k7 = PadLock_key(1) + PadLock_key(2) + PadLock_key(3) + PadLock_key(4)
        k8 = +PadLock_key(5) + PadLock_key(6) + PadLock_key(7) + PadLock_key(8)
        k9 = PadLock_key(7) + PadLock_key(8) + PadLock_key(9) + PadLock_key(10)
        k10 = +PadLock_key(11) + PadLock_key(12) + PadLock_key(13) + PadLock_key(14)
        k11 = +PadLock_key(15) + PadLock_key(16)
        k12 = (k7) + (k8) + (k9) + (k10) + (k11)

        Dim Tmp1, Tmp2, Tmp3, xR, gARRAY(15) As Integer

        'm.d.a.s calculation trick to get key value

        If pwL <= 5 Then

            For ii = 1 To 8

                Tmp1 = PadLock_key(ii)

                If ii = 2 Then
                    Tmp2 = Tmp3 * PadLock_key(ii)
                    Tmp3 = Tmp2
                    xR += 1
                    gARRAY(xR) = Tmp2
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                ElseIf ii = 3 Then
                    Tmp3 = Tmp2 / PadLock_key(ii)
                    Tmp2 = Tmp3
                    xR += 1
                    gARRAY(xR) = Tmp3
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                ElseIf ii = 4 Then
                    Tmp2 = Tmp3 + PadLock_key(ii)
                    Tmp3 = Tmp2
                    xR += 1
                    gARRAY(xR) = Tmp2
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                ElseIf ii = 5 Then
                    Tmp3 = Tmp2 - PadLock_key(ii)
                    Tmp2 = Tmp3
                    xR += 1
                    gARRAY(xR) = Tmp3
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)

                Else
                    Tmp3 = Tmp1

                End If

            Next ii

        End If



        If pwL >= 6 And pwL <= 12 Then

            For ii = 6 To 10

                Tmp1 = PadLock_key(ii)

                If ii = 7 Then
                    Tmp2 = Tmp3 * PadLock_key(ii)
                    Tmp3 = Tmp2
                    xR += 1
                    gARRAY(xR) = Tmp2
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                ElseIf ii = 8 Then
                    Tmp3 = Tmp2 / PadLock_key(ii)
                    Tmp2 = Tmp3
                    xR += 1
                    gARRAY(xR) = Tmp3
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                ElseIf ii = 9 Then
                    Tmp2 = Tmp3 + PadLock_key(ii)
                    Tmp3 = Tmp2
                    xR += 1
                    gARRAY(xR) = Tmp2
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                ElseIf ii = 10 Then
                    Tmp3 = Tmp2 - PadLock_key(ii)
                    Tmp2 = Tmp3
                    xR += 1
                    gARRAY(xR) = Tmp3
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                Else
                    Tmp3 = Tmp1

                End If

            Next ii

        End If

        If pwL >= 11 And pwL <= 16 Then

            For ii = 11 To 16

                Tmp1 = PadLock_key(ii)

                If ii = 12 Then
                    Tmp2 = Tmp3 * PadLock_key(ii)
                    Tmp3 = Tmp2
                    xR += 1
                    gARRAY(xR) = Tmp2
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                ElseIf ii = 13 Then
                    Tmp3 = Tmp2 / PadLock_key(ii)
                    Tmp2 = Tmp3
                    xR += 1
                    gARRAY(xR) = Tmp3
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                ElseIf ii = 14 Then
                    Tmp2 = Tmp3 + PadLock_key(ii)
                    Tmp3 = Tmp2
                    xR += 1
                    gARRAY(xR) = Tmp2
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                ElseIf ii = 15 Then
                    Tmp3 = Tmp2 - PadLock_key(ii)
                    Tmp2 = Tmp3
                    xR += 1
                    gARRAY(xR) = Tmp3
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)

                ElseIf ii = 16 Then
                    Tmp2 = Tmp3 - PadLock_key(ii)
                    Tmp3 = Tmp2
                    xR += 1
                    gARRAY(xR) = Tmp2
                    KEYMASTER = gARRAY(xR) Xor PadLock_key(ii)
                Else
                    Tmp3 = Tmp1
                End If

            Next ii

        End If

        KEYMASTER = KEYMASTER + KEYMASTER

        Return PadLock_key

    End Function


    Public Function DECRYPT(ByVal KEYRFILE As String, Optional ByVal TEXT As String = "", Optional ByVal PATH As String = "") As String

        Dim ciphertext As String = ""
        Dim ctr As Long
        Dim Chars As String = ""
        Dim pos As Long
        Dim Lent2 As String
        Dim CharA As String
        Dim Character As String = ""
        Dim lNum As Integer
        Dim Alpha1 As String
        Dim CH(Len(ALPA)) As Integer
        Dim K As Integer
        Dim rLent As String = ""
        Dim Xq As Long
        Dim Tq(120000) As Long
        Dim Kl As String
        Dim lK As Integer
        Dim rXQ As String
        Dim x As Integer
        Dim y As Integer
        Dim Sum() As Integer
        Dim II As Integer
        Dim c As Integer
        Dim PlText As String = ""
        Dim k7, k8, k9 As Integer
        Dim CounterMarker As Integer
        Dim PWD_Output_Array(3) As Integer
        Dim TmpA As String = ""
        Call GENERATEPWD()

10:     k7 = 0
        k8 = 0
        k9 = 0

        If Dir("KeyR.cpt") = "" Or Dir("template.key") = "" Then
            Console.WriteLine("FILE ERROR")
            Exit Function
        End If

        Try
            Template = ""
            If TEXT <> "" Then
                ciphertext = TEXT
            Else
                FileOpen(1, KEYRFILE, OpenMode.Input, OpenAccess.Default)
                Do While Not EOF(1)
                    ciphertext = LineInput(1)
                    ctr = ctr + 1
                    If ctr = 2 Then Exit Do
                Loop
                FileClose(1)
            End If

            ctr = 0

            Try
                FileOpen(1, C_Path & "Template.key", OpenMode.Input, OpenAccess.Read)
                Do While Not EOF(1)
                    TmpA = InputString(1, LOF(1))
                Loop
                FileClose(1)
            Catch e As Exception
            End Try

            FileOpen(1, C_Path & "Template.key", OpenMode.Input, OpenAccess.Default)
            Input(1, Character)
            FileClose(1)

            If D7Key <> 0 Then
                Character = G9.Decrypt4D7(Character, D7Key)
            Else
                Character = G9.Decrypt4D7(Character, 65445698)
            End If

            FileOpen(1, C_Path & "Template.key", OpenMode.Input, OpenAccess.Default)
            Do While Not EOF(1)
                rLent = LineInput(1)
                ctr = ctr + 1
                If ctr = 2 Then Exit Do
            Loop
            FileClose(1)

            ctr = 0

            FileOpen(1, C_Path & "Template.key", OpenMode.Input, OpenAccess.Default)
            Do While Not EOF(1)
                rXQ = LineInput(1)
                ctr = ctr + 1
                If ctr = 3 Then Exit Do
            Loop
            FileClose(1)

            pos = InStr(ciphertext, ":")
            ciphertext = Mid(ciphertext, pos + 2)


            ctr = 0
            Lent2 = 2



16:         For ctr = 1 To Character.Length - 2
                K = K + 1
                CharA = Mid(Character, ctr, Lent2)
                lNum = "&H" & CharA
                lNum = lNum Xor Right(Character, 2)
                ctr = ctr + 1
                CH(K) = lNum
                Alpha1 = Alpha1 & lNum & "|"
            Next

            ctr = 0

17:         For ctr = 1 To UBound(CH, 1)
                Template = Template + Mid(ALPA, CH(ctr), 1)
            Next ctr

            ErrorTracer = "Error"

            ctr = 0
            Erase CH

            For ctr = 1 To Len(ciphertext)
                Xq += 1
                If CounterMarker > 3 Then CounterMarker = 0
                CounterMarker += 1
                K = CInt(Mid(rLent, Xq, 1))
                CharA = Mid(ciphertext, ctr, K)
                Tq(Xq) = Hex2Dec(CharA) '    contains the cipher text on split form.
                ctr = K + ctr
                ctr -= 1
            Next

            ctr = 0
            LoadMatrix(Xq / 3)
            ReDim CH(Xq)

            ' 3RD SET OF KEYS
            For ctr = 1 To Len(rXQ)
                Kl = Mid(rXQ, ctr, 1)
                For x = 1 To Len(Template) '- 1
                    Alpha1 = Mid(Template, x, 1)
                    If Alpha1 = Kl Then
                        CH(ctr) = x
                    End If
                Next x
            Next ctr

            ctr = 0
            x = 0

            Dim plent As Integer = (Xq / 3 * 2)
            ReDim Sum(Xq / 3)
            Dim posZ(plent) As Integer
            Dim a As Integer
            Dim b As Integer

            r1 = Right(Character, 2)
            r2 = r1


            For II = 1 To Xq
                x = x + 1
                If x = 3 Then
                    y = y + 1
                    If Master_Key_On = True Then
                        Sum(y) = (Tq(II) Xor CH(II)) Xor KEYMASTER
                    Else
                        Sum(y) = (Tq(II) Xor CH(II)) Xor r2
                    End If

                    x = 0
                Else
                    lK += 1
                    If Master_Key_On = True Then
                        posZ(lK) = (Tq(II) Xor CH(II)) Xor KEYMASTER
                    Else
                        posZ(lK) = (Tq(II) Xor CH(II)) Xor r2
                    End If

                End If
            Next

            II = 0
            x = 0
            y = 0

            Erase CH
            Erase Tq

            II = 0


            For II = 1 To UBound(posZ, 1)

                a = posZ(II)
                b = posZ(II + 1)
                c = a + b

                If a < b Then
                    For x = 1 To UBound(MATRIX, 1)
                        If x = c Then
                            PlText = PlText + (MATRIX(x, 0))
                        End If
                    Next x
                End If

                If a > b Then
                    For y = 1 To UBound(MATRIX, 1)

                        If y = c Then
                            PlText = PlText + (MATRIX(y, 1))
                        End If
                    Next y
                End If

                II = II + 1

            Next II

            Dim KP As Integer
            Dim IOtext As String
            Dim ReText As String = ""


            For KP = 1 To Len(PlText) + 1

                IOtext = Mid(PlText, KP, 1)
                ReText = ReText & IOtext

                If IOtext = NLF Then
                    ReText = Replace(ReText, "|", " ")
                    ReText = ReText & vbCrLf
                End If

            Next KP

            PlText = ReText

            If PlText.Length = 0 Then
            Else
            End If

            FileOpen(1, My.Computer.FileSystem.CurrentDirectory & "\plaintext.txt", OpenMode.Output, OpenAccess.Default, OpenShare.Default)
            Print(1, PlText & vbCr)
            FileClose(1)

            Return PlText

        Catch E As Exception

            FileOpen(1, My.Computer.FileSystem.CurrentDirectory & "\plaintext.txt", OpenMode.Output, OpenAccess.Default, OpenShare.Default)
            Print(1, E.Message & " : " & ErrorTracer & D7Key & Template)
            FileClose(1)
            Return "SYSTEM ERROR"
        End Try

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

    Public Function Hex2Dec(ByVal arg) As Long
        Hex2Dec = "&H" & arg
    End Function

    Public Function ENCRYPT(ByVal PLText As String, Optional ByVal PlTextFile As String = "", Optional ByVal BinaryFile As String = "") As String

        Dim PLent As Long
        Dim i As Integer
        Dim x As Long
        Dim y As Long

        Dim B As String = "   "

        Dim T1 As Long
        Dim T2 As Long
        Dim T3 As Long
        Dim kE As String = ""

        Dim xA As String = ""
        Dim xB As String = ""
        Dim xC As String = ""
        Dim KEY() As Integer
        Dim Chart2 As String = ""

        Dim Lent2 As String = ""
        Dim Rbound As Long = 0
        Dim op As Integer
        Dim rE As String
        Dim r2 As Integer
        Dim fLenght As Integer
        Dim PTEXT As String = ""
        ' Try

        If PlTextFile.Length <> 0 Then
            If Right(PlTextFile, 3) <> "txt" Then
                Return ""
                Exit Function
            End If
        End If

        If PLText.Length = 0 Then
            If PlTextFile.Length = 0 Then
                Return ""
                Exit Function
            End If

            FileOpen(1, PlTextFile, OpenMode.Input, OpenAccess.Read)
            PLent = LOF(1)
            While Not EOF(1)
                PTEXT = InputString(1, LOF(1))
            End While
            FileClose(1)
            fLenght = PLent
        Else
            PTEXT = Trim(PLText)
            PLent = Strings.Len(PTEXT)
            fLenght = PLent
        End If

        If PLent < 8 Then
            ' Do Until PLent = 8
            ' Randomize()
            ' PTEXT = PTEXT & CStr(Int(Rnd() * 9))
            ' PLent = Len(PTEXT)
            ' Loop
        End If

        ReDim KEY(Strings.Len(PTEXT) * 3)

        PTEXT = Trim(PTEXT)

        For tX As Integer = 1 To Len(PTEXT)
            rE = Mid(PTEXT, tX, 1)
            If rE = Separator Then
                PTEXT = Replace(PTEXT, vbCr, "|")
            End If
        Next tX


        Call RNDALPHA(False)


        Call LoadMatrix(PTEXT.Length)


        GLOBALTEXT = Left(PTEXT, fLenght)
        Rbound = 0

        For i = 1 To Len(PTEXT)

            Dim charT1 As String = Mid(PTEXT, i, 1)
            For x = 0 To UBound(MATRIX, 2)

                For y = 1 To UBound(MATRIX, 1)

                    If charT1 = MATRIX(y, x) Then
                        Chart2 = Chart2 & charT1

                        If x = vara Then
                            Do Until k3 = y And k1 < k2 '-> PROBLEM WHEN Y = 2 CONDITION CANNOT BE EVALUATED PROPERLY STATUS:FIXED

                                If y = 2 Then
                                    k1 = 0
                                    k2 = 2
                                    k3 = k1 + k2
                                Else
                                    Randomize()
                                    k1 = Int(Rnd() * y)
                                    Randomize()
                                    k2 = Int(Rnd() * y)
                                    k3 = k1 + k2
                                End If

                                If y = 1 Then
                                    Randomize()
                                    k1 = Int(Rnd() * 2) ' -> solution
                                    Randomize()
                                    k2 = Int(Rnd() * 2)
                                    k3 = k1 + k2
                                End If
                            Loop
                        End If

                        If x = varb Then
                            Do Until k3 = y And k1 > k2

                                If y = 2 Then
                                    k1 = 2
                                    k2 = 0
                                    k3 = k1 + k2
                                Else
                                    Randomize()
                                    k1 = Int(Rnd() * y)
                                    Randomize()
                                    k2 = Int(Rnd() * y)
                                    k3 = k1 + k2
                                End If

                                If y = 1 Then
                                    Randomize()
                                    k1 = Int(Rnd() * 2)
                                    Randomize()
                                    k2 = Int(Rnd() * 2)
                                    k3 = k1 + k2
                                End If
                            Loop
                        End If


                        REM 2nd loop

                        Do Until op = 1
                            Randomize()
                            k4 = Int(Rnd() * ALPA.Length) + 1
                            If k4 = k3 Then op = 1
                            Randomize()
                            k5 = Int(Rnd() * ALPA.Length) + 1
                            If k5 = k3 Then op = 1
                            Randomize()
                            k6 = Int(Rnd() * ALPA.Length) + 1
                            If k6 = k3 Then op = 1
                        Loop

                        Do Until Rbound = Len(Chart2) * 3
                            KEY(Rbound + 1) = k4
                            KEY(Rbound + 2) = k5
                            KEY(Rbound + 3) = k6
                            op = Rbound + 3
                            Rbound = op
                        Loop

                        r2 = r1

                        If Master_Key_On = True Then
                            T1 = k1 Xor k4 Xor KEYMASTER
                            T2 = k2 Xor k5 Xor KEYMASTER
                            T3 = k3 Xor k6 Xor KEYMASTER
                        Else
                            T1 = k1 Xor k4 Xor r2
                            T2 = k2 Xor k5 Xor r2
                            T3 = k3 Xor k6 Xor r2
                        End If

                        kE = kE & Hex(T1) & Hex(T2) & Hex(T3)
                        Lent2 = Lent2 & Len(Hex(T1)) & Len(Hex(T2)) & Len(Hex(T3))

                    End If
                Next y
            Next x
        Next i

        k1 = 0
        k2 = 0
        k3 = 0

        Dim allText(3) As String
        Dim J2 As String

        FileOpen(11, C_Path & "KeyR.cpt", OpenMode.Output, OpenAccess.Default, OpenShare.Default)

        J2 = G9.Encrypt4D7(kE, 65445698)
        Print(11, "[G7B BEGIN]" & vbCr)
        Print(11, "ENCRYPTED-TEXT : " & kE & vbCrLf)
        FileClose(11)


        Dim CipherOut As String
        CipherOut = kE
        kE = ""

        Dim DBG As Integer = 0

        If DBG = 1 Then
            For n As Long = 1 To KEY.GetUpperBound(0)
                If KEY(n) > Len(Template) Then
                    Template = Template & Template
                End If
                If KEY(n) = 0 Then Exit For
                kE = kE & Mid(Template, KEY(n), 1)
            Next
        End If

        Static xw As Integer

        xw += 1

        For n As Long = 1 To KEY.GetUpperBound(0)
            If KEY(n) = 0 Then

                Exit For
            End If
            kE = kE & Mid(Template, KEY(n), 1)
        Next

        Dim B1 As String
        Dim B2 As String

        FileOpen(11, C_Path & "Template.key", OpenMode.Append, OpenAccess.Default)
        PrintLine(11, Lent2)
        PrintLine(11, kE)
        FileClose(11)

        '    Call ExcelTableCreate(CipherOut)

        Return CipherOut

        '  Catch e As Exception
        ' Return e.Message & " Error Found"
        ' End Try

    End Function

    Function BLOCKIMAGE(ByVal IMGPATH As String, Optional ByVal LOC As Integer = 333) As Boolean


        'offset 700 - lightness of pix

        Dim COR_STR As String = Chr(66)
        Dim ORHEXVAL As String


        Try

            Dim XBYTE(0 To 9) As Byte
            FileOpen(1, IMGPATH, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.Shared)
            FileGet(1, XBYTE, LOC)
            ORHEXVAL = (Hex(XBYTE(0)) & " : " & Chr(XBYTE(0)))
            MsgBox(ORHEXVAL)
            FilePut(1, COR_STR, LOC)
            FileClose(1)

        Catch E As Exception

        End Try

    End Function




    Function Sbox_Lame(ByVal KeyLock() As Long) As Long

        Dim S_Box(16, 16) As Integer
        Dim i As Integer
        Dim inc As Long = 2

        For i = 1 To S_Box.GetUpperBound(0)
            inc += 2
            For g = 1 To S_Box.GetUpperBound(1)
                S_Box(i, g) = ((KeyLock(g) / inc))
            Next g
        Next i

    End Function

    Public Sub RNDALPHA(Optional ByVal DBG As Boolean = False)

        Dim ALPA5 As String

        Dim MAP(Len(ALPA)) As Integer
        Dim TMP(Len(ALPA)) As Integer

        Dim S_NUM(,) As Integer = { _
                                   {54, 67, 87, 32, 55, 64, 73, 11, 24, 65}, _
                                   {34, 52, 54, 93, 13, 62, 57, 78, 15, 66}, _
                                   {75, 53, 37, 99, 18, 69, 76, 78, 90, 60}, _
                                   {43, 44, 38, 19, 45, 65, 28, 89, 97, 41}, _
                                   {11, 23, 34, 29, 56, 30, 70, 10, 51, 88} _
                                   }

        Dim S_NUM2(,) As Integer = {{256 Xor 32}, {512 Xor 64}, {1024 Xor 128}}

        Dim ii As Integer
        Dim Rn As Integer
        Dim x As Integer
        Dim Rn2 As Integer
        Dim hasDup As Integer
        Dim pos As Integer
        Dim GENKEYOUTPUT() As Long = GENERATEPWD()
        Erase TMP
        Erase MAP
        ftemp = ""

        '1 2 3 4 5 6 7 8 9 0 1 2 3
        'B D F H J L N P R T V X Z
        '1 2 3 4 5 6 7 8 9 0 - . @

        ' Rw = (3 * Rnd()) + 1
        ' Cl = (8 * Rnd()) + 1
        'S_NUM(1, 1) = {54, 67, 87, 90, 55, 64, 73, 11, 24, 65}
        'S_NUM(1, 2) = array(34, 53, 54, 99, 13, 69, 55, 78, 90, 66)
        'S_NUM(1, 3) = Array(43, 44, 38, 19, 45, 65, 28, 89, 97, 41)
        'S_NUM(1, 4) = Array(11, 23, 34, 23, 56, 30, 70, 10, 51, 88)
        'Randomize
        ' formula  x  = n * (hv - lv) + 1)
        'f:
        'y = (Int(Rnd * 10))
        's = (Int(Rnd * 4) + 1)
        'MsgBox S_NUM(1, s)(y) & ":::" & y & ":::" & s
        'GoTo f
        Template = ""
        templateb = ""
        ReDim TMP(Len(ALPA))
        ReDim MAP(Len(ALPA))

        If DBG = False Then
            For ii = 1 To Len(ALPA)
                Randomize()
                Rn = Int(Rnd() * Len(ALPA)) + 1
                pos = ii

                Do Until x = Len(ALPA)
                    x = x + 1
                    Rn2 = TMP(x)
                    If Rn2 = Rn Then hasDup = 1
                    If hasDup = 1 Then Exit Do
                Loop

                x = 0

                If hasDup = 0 Then
                    TMP(ii) = Rn
                Else
                    ii = pos - 1
                    hasDup = 0
                End If

            Next ii
        End If

        ii = 0

        If DBG = True Then
            For ii = 1 To Len(ALPA)
                TMP(ii) = ii
            Next ii
        End If

        ii = 0

        For ii = 1 To Len(ALPA)
            ALPA5 = Mid$(ALPA, TMP(ii), 1)
            Template = Template & ALPA5
        Next ii

        '  FileOpen(1, C_Path & "alpabhet.INI", OpenMode.Append, OpenAccess.Write)
        '  PrintLine(1, Template)
        ' FileClose(1)

        templateb = StrReverse(Template)

        ii = 0

        Randomize()
        '  r1 = S_NUM(1, Int(Rnd * 4) + 1)(Int(Rnd * 10))
        Randomize()
        Dim y1 As Integer = Int(Rnd() * 4)
        Randomize()
        Dim y2 As Integer = Int(Rnd() * 9)
        r1 = S_NUM.GetValue(y1, y2)

        PNum = r1

        For ii = 1 To UBound(TMP)
            r2 = r2 + 1
            xhex = TMP(ii) Xor PNum
            MAP(ii) = TMP(ii)
            If Len(Hex(xhex)) = 1 Then
                ftemp = ftemp & "0" & Hex(xhex)
            Else
                ftemp = ftemp & Hex(xhex)
            End If
        Next ii
        Dim TempC As String

        For ii = 1 To Len(ftemp)
            TempC = Mid(ftemp, ii, ii + 1)
        Next ii

        Dim B1$, D3Gen$

        D3Gen = ftemp & r1

        If D7Key <> 0 Then
            B1 = G9.Encrypt4D7(D3Gen, D7Key)
        ElseIf D7Key = 0 Then
            B1 = G9.Encrypt4D7(D3Gen, 65445698)
        End If

        FileOpen(1, C_Path + "Template.key", OpenMode.Output, OpenAccess.Write)
        PrintLine(1, B1)
        FileClose(1)


    End Sub

    Public Function LoadSameKey(ByVal keyK As String) As Boolean

        Dim Character As String = ""
        Dim CharA As String
        Dim lnum As Integer
        Dim alpha1 As String
        Dim Lent2 As Integer = 2
        Dim io As Integer
        Dim KeyTP As String
        Dim kl(48) As Integer

        Template = ""
        alpha1 = ""

        If keyK <> "" Then

            FileOpen(1, keyK, OpenMode.Input, OpenAccess.Default)
            KeyTP = LineInput(1)
            FileClose(1)
        End If
        k = 0
        For io = 1 To KeyTP.Length - 2
            k = k + 1
            CharA = Mid(KeyTP, io, Lent2)
            lnum = "&H" & CharA
            lnum = lnum Xor Right(KeyTP, 2)
            io = io + 1
            kl(k) = lnum
            alpha1 = alpha1 & lnum & "|"
        Next io

        For ctr = 1 To UBound(kl, 1)
            Template = Template + Mid(ALPA, kl(ctr), 1)
        Next ctr

        Return False

        If Template <> "" Then
            Return True
            LSK = True
        End If


    End Function


    Public Sub LoadMatrix(ByVal M2Lent As Integer)

        REM | active matrix loader |

        Dim i As Integer
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim kc As Integer
        Row1 = ""
        Row2 = ""
        Row3 = ""
        x = 1
        y = 2
        z = (ROWBOUND) * (2) + (1)

        hiBound = r1
        lowBound = hiBound - Len(ALPA) / 2 + 1

        Erase MATRIX
        ReDim MATRIX(ROWBOUND, 1)

        ' For i = lowBound To hiBound
        For i = 1 To Len(ALPA) / 2

            ' c = column , r = rows
            If M2Lent >= 4 Then
                'If M2Lent <= 4 Or M2Lent <= 32 Then ' M2Lent Is express In bytes
                vara = 0 : varb = 1
                MATRIX(i, vara) = Mid(Template, x, 1) ' CHANGE ALPA TO TEMPLATE VARIABLE
                MATRIX(i, varb) = Mid(Template, y, 1)
            End If
            x = x + 2
            y = y + 2
        Next i

        templateb = StrReverse(Template)

        For kc = 1 To UBound(MATRIX, 1)
            Row1 = Row1 + MATRIX(kc, vara)
            Row2 = Row2 + MATRIX(kc, varb)
        Next kc

        '  Console.WriteLine(Row1 & vbCrLf)
        ' Console.WriteLine(Row2 & vbCrLf)
        'Console.WriteLine("----------------------" & vbCrLf)
    End Sub

End Module
