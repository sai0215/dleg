Attribute VB_Name = "ModAccountsUD"
Option Explicit
Public lstNo As Long
Public ws As Workspace
Public db As Database
Public rs As Recordset
Public logString
Public LoginSucceeded As Boolean
Public msgtemp As String
Public tempStr As String
Public MasTemp As String
Public tempBln As Boolean
Public tempNum As Integer
Public LSTAC As Integer
Public LSTSUP As Integer
Public LSTAGN As Integer
Public cur As String
Public cod As String
Public con As Currency
Public Z As Integer
Public txtGFColor As ColorConstants
Public txtLFColor As ColorConstants
Public Sqlqry As String
Public Sqlqry1 As String
Public rs1 As Recordset
Dim padChar As String * 1
Dim TrimChar As String * 1
Public totdhs As Currency
Public totusd As Currency
Public Rep As Integer

'Public convertion As Currency
Public Function convertion()
convertion = 3.68
End Function

Public Function DaysinMonth(MonthNum As Integer, InYear As Integer) As Integer
    If MonthNum = 1 Or MonthNum = 3 Or MonthNum = 5 Or MonthNum = 7 Or MonthNum = 8 Or MonthNum = 10 Or MonthNum = 12 Then
        DaysinMonth = 31
    ElseIf MonthNum = 2 And InYear Mod 4 = 0 Then
        DaysinMonth = 29
    ElseIf MonthNum = 2 And InYear Mod 4 <> 0 Then
        DaysinMonth = 28
    ElseIf MonthNum = 4 Or MonthNum = 6 Or MonthNum = 9 Or MonthNum = 11 Then
        DaysinMonth = 30
    End If
End Function

Public Function InWords2(number As Long) As String
    Dim a(0 To 90) As String
    Dim B(3) As String
    Dim C(4) As String
    Dim d(6) As String
    Dim e(7) As String
    Dim X, i, j, k, m, n, p, Q, Y, Z As Variant
    Dim SarTemp As String
    Dim TempMill As String
    Dim MasTemp As String
    Dim tempStr As String
            X = number
            X = Val(X)
    a(0) = "Zero"
    a(1) = "One"
    a(2) = "Two"
    a(3) = "Three"
    a(4) = "Four"
    a(5) = "Five"
    a(6) = "Six"
    a(7) = "Seven"
    a(8) = "Eight"
    a(9) = "Nine"
    a(10) = "Ten"
    a(11) = "Eleven"
    a(12) = "Twelve"
    a(13) = "Thirteen"
    a(14) = "Fourteen"
    a(15) = "Fifteen"
    a(16) = "Sixteen"
    a(17) = "Seventeen"
    a(18) = "Eighteen"
    a(19) = "Nineteen"
    a(20) = "Twenty"
    a(30) = "Thirty"
    a(40) = "Fourty"
    a(50) = "Fifty"
    a(60) = "Sixty"
    a(70) = "Seventy"
    a(80) = "Eighty"
    a(90) = "Ninty"
        tempStr = ""
        MasTemp = ""
        Z = Len(X)
   
    If Z = 1 Then InWords2 = a(X)
        
    If Z = 2 Then
        If Val(X) > 20 And Val(X) < 100 And Right(X, 1) <> 0 Then
            i = Right(X, 1)
            m = Val(X) - i
        Else
            InWords2 = a(X)
            Exit Function
        End If
    End If
    
Hundred:
    If Z = 3 Then
        B(Z) = "Hundred"
        If Val(X) = 100 Then
            InWords2 = MasTemp & " " & a(1) & " " & B(Z)
            Exit Function
        ElseIf X > 100 And Val(X) < 1000 And (Mid(X, 3) = 0 And Mid(X, 2, 1) = 0) Then
            i = Right(X, 2)
            j = Left(X, 1)
            InWords2 = MasTemp & " " & a(j) & " " & B(Z)
            Exit Function
        ElseIf X > 100 And Val(X) < 1000 And (Mid(X, 3) = 0 Or Mid(X, 2, 1) = 0) Then
            i = Right(X, 2)
            j = Left(X, 1)
            InWords2 = MasTemp & " " & a(j) & " " & B(Z) & " and " & a(i)
            Exit Function
        ElseIf Val(X) > 100 And Val(X) < 1000 And Right(X, 2) < 20 Then
            i = Right(X, 2)
            k = Right(i, 1)
            j = Left(X, 1)
            InWords2 = MasTemp & " " & a(j) & " " & B(Z) & " and " & a(i)
            Exit Function
        ElseIf Val(X) > 100 And Val(X) < 1000 Then
            i = Right(X, 2)
            k = Right(i, 1)
            j = Left(X, 1)
            m = Val(i) - k
        End If
    End If

Thousand:
    If Z = 4 Then
        C(Z) = "Thousand"
        If Val(X) = 1000 Then
            InWords2 = MasTemp & " " & a(1) & " " & C(Z)
            Exit Function
        ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4) = 0 Then
            j = Left(X, 1)
            If j = 1 Then
            InWords2 = MasTemp & " " & a(j) & " " & C(Z)
            Else
            InWords2 = MasTemp & " " & a(j) & " " & C(Z) & "s"
            End If
            Exit Function
        ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Right(X, 2)
            j = Left(X, 1)
            If j = 1 Then
            InWords2 = MasTemp & " " & a(j) & " " & C(Z) & " and " & a(i)
            Else
            InWords2 = MasTemp & " " & a(j) & " " & C(Z) & "s" & " and " & a(i)
            End If
            Exit Function
         ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Right(X, 2)
            j = Left(X, 1)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            InWords2 = MasTemp & " " & a(j) & " " & C(Z) & " and " & a(m) & " " & a(k)
            Exit Function
         ElseIf X > 1000 And Val(X) < 10000 Then
            j = Left(X, 1)
            If j = 1 Then
            MasTemp = MasTemp & " " & a(j) & " " & C(Z)
            Else
            MasTemp = MasTemp & " " & a(j) & " " & C(Z) & "s"
            End If
            X = Mid(X, 2)
            Z = Len(X)
            GoTo Hundred
        End If
    End If

TenThousand:
    If Z = 5 Then
        C(Y) = "Thousand"
        If Val(X) = 10000 Then '1
            InWords2 = tempStr & " " & a(10) & " " & C(Y)
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            InWords2 = tempStr & " " & a(i) & " " & C(Y)
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            j = Right(X, 2)
            InWords2 = tempStr & " " & a(i) & " " & C(Y) & " and " & a(j)
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            InWords2 = tempStr & " " & a(i) & " " & C(Y) & " and " & a(m) & " " & a(k)
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) > 20 And Mid(Left(X, 2), 2, 1) <> 0) And (Right(X, 2) > 20 And Mid(Right(X, 2), 2, 1) <> 0) Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            InWords2 = tempStr & " " & a(Q) & " " & a(p) & " " & C(Y) & " and " & a(m) & " " & a(k)
            Exit Function
         ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            InWords2 = tempStr & " " & a(Q) & " " & a(p) & " " & C(Y) & " and " & a(j)
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Left(X, 2) > 20 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5) = 0 Then
            i = Left(X, 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            InWords2 = tempStr & " " & a(m) & " " & a(k) & " " & C(Y)
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Right(X, 2) < 20 And Left(X, 2) < 20 Then
            i = Left(X, 2)
            MasTemp = tempStr & " " & a(i) & " " & C(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 10000 And Val(X) < 100000 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            MasTemp = tempStr & " " & a(i) & " " & C(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 10000 And Val(X) < 100000 Then '17
            i = Left(X, 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            MasTemp = tempStr & " " & a(m) & " " & a(k) & " " & C(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Hundred
        End If
    End If
    
    If Z = 6 Then
        d(Y) = "Lakh"
        If Val(X) = 100000 Then '1
            InWords2 = a(1) & " " & d(Y)
            Exit Function
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6) = 0 Then
            i = Left(X, 1)
            InWords2 = a(i) & " " & d(Y) & "s"
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 1)
            j = Right(X, 2)
            InWords2 = a(i) & " " & d(Y) & " and " & a(j)
            Exit Function
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Left(X, 1)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            InWords2 = a(i) & " " & d(Y) & " and " & a(m) & " " & a(k)
            Exit Function
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 Then
            i = Left(X, 1)
            If i = 1 Then
                MasTemp = a(i) & " " & d(Y)
            Else
                MasTemp = a(i) & " " & d(Y) & "s"
            End If
            X = Mid(X, 4)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 Then
            i = Left(X, 1)
            If i = 1 Then
            MasTemp = a(i) & " " & d(Y)
            Else
            MasTemp = a(i) & " " & d(Y) & "s"
            End If
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Thousand
        ElseIf Val(X) > 100000 And Val(X) < 1000000 Then
            i = Left(X, 1)
            If i = 1 Then
                tempStr = a(i) & " " & d(Y)
            Else
                tempStr = a(i) & " " & d(Y) & "s"
            End If
            X = Mid(X, 2)
            Z = Len(X)
            GoTo TenThousand
        End If
    End If
    
    If Z = 7 Then
        d(Y) = "Lakhs"
        If Val(X) = 1000000 Then
            InWords2 = a(10) & " " & d(Y)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            InWords2 = a(i) & " " & d(Y)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 Then
            i = Left(X, 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            InWords2 = a(m) & " " & a(k) & " " & d(Y)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            j = Right(X, 2)
            InWords2 = a(i) & " " & d(Y) & " and " & a(j)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            InWords2 = a(i) & " " & d(Y) & " and " & a(m) & " " & a(k)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            InWords2 = a(Q) & " " & a(p) & " " & d(Y) & " and " & a(j)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And (Left(X, 2) > 20 And Mid(Left(X, 2), 2, 1) <> 0) And (Right(X, 2) > 20 And Mid(Right(X, 2), 2, 1) <> 0) Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            InWords2 = a(Q) & " " & a(p) & " " & d(Y) & " and " & a(m) & " " & a(k)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 Then
            i = Left(X, 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            InWords2 = a(m) & " " & a(k) & " " & d(Y)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            MasTemp = a(i) & " " & d(Y)
            X = Mid(X, 5)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 Then
            i = Left(X, 2)
            k = Right(X, 1)
            m = Val(i) - Val(k)
            MasTemp = a(m) & " " & a(k) & " " & d(Y)
            X = Mid(X, 5)
            Z = Len(X)
            GoTo Hundred
       ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            MasTemp = a(i) & " " & d(Y)
            X = Mid(X, 4)
            Z = Len(X)
            GoTo Thousand
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 Then
            i = Left(X, 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            MasTemp = a(m) & " " & a(k) & " " & d(Y)
            X = Mid(X, 4)
            Z = Len(X)
            GoTo Thousand
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            tempStr = a(i) & " " & d(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo TenThousand
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 Then
            i = Left(X, 2)
            k = Right(i, 1)
            j = Val(i) - Val(k)
            tempStr = a(j) & " " & a(k) & " " & d(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo TenThousand
        End If
    End If
    If Z = 8 Then
        e(Y) = "Crore"
        If Val(X) = 10000000 Then
            InWords2 = a(1) & " " & e(Y)
            Exit Function
        End If
    End If
    
    If Z = 2 And Len(i) = 1 Then
        InWords2 = a(m) & " " & a(i)
    ElseIf Z = 3 And Len(i) = 2 Then
        InWords2 = MasTemp & " " & a(j) & " " & B(Z) & " and " & a(m) & " " & a(k)
    End If
    
End Function

Public Function AEInWords(number As Long) As String
Dim tempStr As String
Dim MasTemp As String
Dim TempMill As String

    Dim a(0 To 90) As String
    Dim B(3) As String
    Dim C(4) As String
    Dim d(6) As String
    Dim e(7) As String
    Dim X, F3, i, j, k, m, n, p, Q, Y, Z, m1 As Variant
    Dim SarTemp As String
            X = number
            X = Val(X)
    a(0) = "Zero"
    a(1) = "Vaahad"
    a(2) = "Itnaan"
    a(3) = "Talata"
    a(4) = "Arba"
    a(5) = "Khamsa"
    a(6) = "Sittae"
    a(7) = "Sabha"
    a(8) = "Tamania"
    a(9) = "Theesa"
    a(10) = "Asra"
    a(11) = "Hidash"
    a(12) = "Itnash"
    a(13) = "Talatash"
    a(14) = "Arvathash"
    a(15) = "Khamstash"
    a(16) = "Sittash"
    a(17) = "Sabhatash"
    a(18) = "Tamantash"
    a(19) = "Teesatash"
    a(20) = "Asreen"
    a(30) = "Talateen"
    a(40) = "Arbaeen"
    a(50) = "Khamseen"
    a(60) = "Sitteen"
    a(70) = "Sabhaeen"
    a(80) = "Tamaeen"
    a(90) = "Teesaeen"
        tempStr = ""
        MasTemp = ""
        Z = Len(X)
   
    If Z = 1 Then AEInWords = a(X)
        
    If Z = 2 Then
        If Val(X) > 20 And Val(X) < 100 And Right(X, 1) <> 0 Then
            i = Right(X, 1)
            m = Val(X) - i
        Else
            AEInWords = a(X)
            Exit Function
        End If
    End If
    
Hundred:
    If Z = 3 Then
        B(Z) = "Miya"
        X = Val(X)
        m1 = Len(X)
        If m1 = 1 And X = 0 Then
            AEInWords = MasTemp
            Exit Function
        ElseIf m1 = 1 Then
            AEInWords = MasTemp & " " & a(X)
            Exit Function
        ElseIf m1 = 2 Then
            If Val(X) > 20 And Val(X) < 100 And Right(X, 1) <> 0 Then
                i = Right(Val(X), 1)
                m = Val(X) - i
                AEInWords = MasTemp & " " & a(i) & " O " & a(m)
                Exit Function
            Else
                AEInWords = MasTemp & " " & a(X)
                Exit Function
            End If
        End If
        If X = 0 Then
            AEInWords = MasTemp
            Exit Function
        End If
        If Val(X) = 100 Then
            AEInWords = MasTemp & " " & B(Z)
            Exit Function
        ElseIf Val(X) = 200 Then
            AEInWords = MasTemp & " " & "Methane"
            Exit Function
        ElseIf Val(X) = 300 Then
            AEInWords = MasTemp & " " & "Talat Miya"
            Exit Function
        ElseIf Val(X) = 400 Then
            AEInWords = MasTemp & " " & "Arba Miya"
            Exit Function
        ElseIf Val(X) = 500 Then
            AEInWords = MasTemp & " " & "Khams Miya"
            Exit Function
        ElseIf Val(X) = 600 Then
            AEInWords = MasTemp & " " & "Sitt Miya"
            Exit Function
        ElseIf Val(X) = 700 Then
            AEInWords = MasTemp & " " & "Sabha Miya"
            Exit Function
        ElseIf Val(X) = 800 Then
            AEInWords = MasTemp & " " & "Taman Miya"
            Exit Function
        ElseIf Val(X) = 900 Then
            AEInWords = MasTemp & " " & "Tees Miya"
            Exit Function
         ElseIf X > 200 And Val(X) < 300 And Mid(X, 3) = 0 Then
            i = Right(Val(X), 2)
            AEInWords = MasTemp & " " & "Methane" & " O " & a(i)  '& A(i) & " "
            Exit Function
        ElseIf X > 200 And Val(X) < 300 And Right(X, 2) < 20 Then
            i = Right(Val(X), 2)
            AEInWords = MasTemp & " " & "Methane" & " " & a(i)  '& " O " & A(i)
            Exit Function
        ElseIf Val(X) > 200 And Val(X) < 300 Then
            i = Right(Val(X), 2)
            k = Right(i, 1)
            m = Val(i) - k
            AEInWords = MasTemp & " " & "Methane" & " " & a(k) & " O " & a(m)
            Exit Function
        ElseIf X > 100 And Val(X) < 1000 And (Mid(X, 3) = 0 And Mid(X, 2, 1) = 0) Then
            i = Right(Val(X), 2)
            j = Left(Val(X), 1)
            AEInWords = MasTemp & " " & a(j) & " " & B(Z)
            Exit Function
       ElseIf X > 100 And Val(X) < 1000 And (Mid(X, 3) = 0 Or Mid(X, 2, 1) = 0) Then
            i = Right(Val(X), 2)
            j = Left(Val(X), 1)
            AEInWords = MasTemp & " " & B(Z) & " O " & a(i)  '& A(i) & " "
            Exit Function
        ElseIf Val(X) > 100 And Val(X) < 1000 And Right(X, 2) < 20 Then
            i = Right(Val(X), 2)
            k = Right(i, 1)
            j = Left(Val(X), 1)
            If j = 1 Then
                AEInWords = MasTemp & " " & B(Z) & " O " & a(i)
            Else
                AEInWords = MasTemp & " " & a(j) & " " & B(Z) & " O " & a(i)
            End If
            Exit Function
        ElseIf Val(X) > 100 And Val(X) < 1000 Then
            i = Right(Val(X), 2)
            k = Right(i, 1)
            j = Left(Val(X), 1)
            m = Val(i) - k
        End If
    End If

Thousand:
    If Z = 4 Then
        C(Z) = "Alf"
        If Val(X) = 1000 Then
            AEInWords = MasTemp & " " & C(Z)
            Exit Function
        ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4) = 0 Then
            j = Left(Val(X), 1)
            If j = 1 Then
            AEInWords = MasTemp & " " & a(j) & " " & C(Z)
            Else
            AEInWords = MasTemp & " " & a(j) & " " & C(Z)
            End If
            Exit Function
        ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Right(Val(X), 2)
            j = Left(Val(X), 1)
            If j = 1 Then
                AEInWords = MasTemp & C(Z) & " O " & a(i) & " " '" " & A(j) & " " &
            ElseIf j = 0 Then
                AEInWords = MasTemp & " " & C(Z) & " O " & a(i)
            Else
                AEInWords = MasTemp & " " & a(j) & " " & C(Z) & " O " & a(i)
            End If
            Exit Function
         ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Right(Val(X), 2)
            j = Left(Val(X), 1)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            AEInWords = MasTemp & " " & a(j) & " " & C(Z) & " O " & a(m) & " " & a(k)
            Exit Function
         ElseIf X > 1000 And Val(X) < 10000 Then
            j = Left(Val(X), 1)
            If j = 1 Then
            MasTemp = MasTemp & " " & a(j) & " " & C(Z)
            Else
            MasTemp = MasTemp & " " & a(j) & " " & C(Z)
            End If
            X = Mid(Val(X), 2)
            Z = Len(X)
            GoTo Hundred
        End If
    End If

TenThousand:
    If Z = 5 Then
        C(Y) = "Alf"
        If Val(X) = 10000 Then '1
            AEInWords = tempStr & " " & a(10) & " " & C(Y)
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            AEInWords = tempStr & " " & a(i) & " " & C(Y)
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(Val(X), 2)
            j = Right(Val(X), 2)
            AEInWords = tempStr & " " & a(i) & " " & C(Y) & " O " & a(j)
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Left(Val(X), 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(Val(X), 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            AEInWords = tempStr & " " & a(i) & " " & C(Y) & " " & a(k) & " O " & a(m)
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) > 20 And Mid(Left(X, 2), 2, 1) <> 0) And (Right(X, 2) > 20 And Mid(Right(X, 2), 2, 1) <> 0) Then
            i = Left(Val(X), 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(Val(X), 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            AEInWords = tempStr & " " & a(p) & " O " & a(Q) & " " & C(Y) & " " & a(k) & " O " & a(m)
            Exit Function
         ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(Val(X), 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(Val(X), 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If j = 0 Then
                AEInWords = tempStr & " " & a(p) & " O " & a(Q) & " " & C(Y) ' & " O " & A(j)
            Else
                AEInWords = tempStr & " " & a(p) & " O " & a(Q) & " " & C(Y) & " O " & a(j)
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Left(X, 2) > 20 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5) = 0 Then
            i = Left(Val(X), 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            AEInWords = tempStr & " " & a(m) & " " & a(k) & " " & C(Y)
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Right(X, 2) < 20 And Left(X, 2) < 20 Then
            i = Left(Val(X), 2)
            MasTemp = tempStr & " " & a(i) & " " & C(Y)
            X = Mid(Val(X), 3)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 10000 And Val(X) < 100000 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(Val(X), 2)
            MasTemp = tempStr & " " & a(i) & " " & C(Y)
            X = Mid(Val(X), 3)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 10000 And Val(X) < 100000 Then '17
            i = Left(Val(X), 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            MasTemp = tempStr & " " & a(k) & " O " & a(m) & " " & C(Y)
            X = Mid(Val(X), 3)
            Z = Len(X)
            GoTo Hundred
        End If
    End If
   
HundredThousand:

    If Z = 6 Then
        
        X = Val(X)
        m1 = Len(X)
        If m1 = 1 And X = 0 Then
            AEInWords = MasTemp
            Exit Function
        ElseIf m1 = 1 Then
            AEInWords = MasTemp & " " & a(X)
            Exit Function
        ElseIf m1 = 2 Then
            If Val(X) > 20 And Val(X) < 100 And Right(X, 1) <> 0 Then
                i = Right(X, 1)
                m = Val(X) - i
                AEInWords = MasTemp & " " & a(i) & " O " & a(m)
                Exit Function
            Else
                AEInWords = MasTemp & " " & a(X)
                Exit Function
            End If
        End If
         
        d(Y) = "Alf"
        B(Y) = "Miya"
        If Val(X) = 100000 Then '1
            AEInWords = TempMill & " " & B(Y) & " O " & d(Y)
            Exit Function
         ElseIf Val(X) = 200000 Then
            AEInWords = TempMill & " " & "Methane" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 300000 Then
            AEInWords = TempMill & " " & "Talat Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 400000 Then
            AEInWords = TempMill & " " & "Arba Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 500000 Then
            AEInWords = TempMill & " " & "Khams Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 600000 Then
            AEInWords = TempMill & " " & "Sitt Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 700000 Then
            AEInWords = TempMill & " " & "Sabha Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 800000 Then
            AEInWords = TempMill & " " & "Taman Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 900000 Then
            AEInWords = TempMill & " " & "Tees Miya" & " O " & d(Y)
            Exit Function
        End If
        F3 = Mid(Val(X), 1, 3)
        Z = Len(Mid(Val(X), 4))
        If Val(F3) = 100 Then '1
            MasTemp = TempMill & " " & B(Y) & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) = 200 Then
            MasTemp = TempMill & " " & "Methane" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) = 300 Then
            MasTemp = TempMill & " " & "Talat Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) = 400 Then
            MasTemp = TempMill & " " & "Arba Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) = 500 Then
            MasTemp = TempMill & " " & "Khams Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) = 600 Then
            MasTemp = TempMill & " " & "Sitt Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) = 700 Then
            MasTemp = TempMill & " " & "Sabha Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) = 800 Then
            MasTemp = TempMill & " " & "Taman Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) = 900 Then
            MasTemp = TempMill & " " & "Tees Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) > 200 And Val(F3) < 300 And Mid(F3, 3) = 0 Then
            i = Right(F3, 2)
            MasTemp = TempMill & " " & "Methane" & " O " & a(i) & " O " & d(Y) '& A(i) & " "
            X = Mid(Val(X), 4)
            GoTo Hundred
         ElseIf F3 > 200 And Val(F3) < 300 And Right(F3, 2) < 20 Then
            i = Right(F3, 2)
            MasTemp = TempMill & " " & "Methane" & " " & a(i) & " O " & d(Y) '& " O " & A(i)
            GoTo Hundred
        ElseIf Val(F3) > 200 And Val(F3) < 300 Then
            i = Right(F3, 2)
            k = Right(i, 1)
            m = Val(i) - k
            MasTemp = TempMill & " " & "Methane" & " " & a(k) & " O " & a(m) & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf F3 > 100 And Val(F3) < 1000 And (Mid(F3, 3) = 0 And Mid(F3, 2, 1) = 0) Then
            i = Right(F3, 2)
            j = Left(F3, 1)
            MasTemp = TempMill & " " & a(j) & " " & B(Y) & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
       ElseIf F3 > 100 And Val(F3) < 1000 And (Mid(F3, 3) = 0 Or Mid(F3, 2, 1) = 0) Then
            i = Right(F3, 2)
            j = Left(F3, 1)
            MasTemp = TempMill & " " & B(Y) & " " & a(j) & " O " & d(Y)  '& A(i) & " "
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) > 100 And Val(F3) < 1000 And Right(F3, 2) < 20 Then
            i = Right(F3, 2)
            k = Right(i, 1)
            j = Left(F3, 1)
            MasTemp = TempMill & " " & a(j) & " " & B(Y) & " O " & a(i) & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo Hundred
        ElseIf Val(F3) > 100 And Val(F3) < 1000 Then
            i = Right(F3, 2)
            k = Right(i, 1)
            j = Left(F3, 1)
            m = Val(i) - k
            If j = 1 Then
                MasTemp = TempMill & " " & B(Y) & " " & a(k) & " O " & a(m) & " O " & d(Y)
            Else
                MasTemp = TempMill & " " & a(j) & " " & B(Y) & " " & a(k) & " O " & a(m) & " O " & d(Y)
            End If
        X = Mid(Val(X), 4)
        GoTo Hundred
        End If
    End If
     
    If Z = 7 Then
        d(Y) = "Million"
        If Val(X) = 1000000 Then
            AEInWords = tempStr & " " & d(Y)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 1)
            AEInWords = a(i) & " " & d(Y)
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 Then
            i = Mid(Val(X), 1, 1)
            X = Mid(Val(X), 2)
            X = Val(X)
            Z = Len(X)
            TempMill = a(i) & " " & d(Y)
            If Z = 3 Then
                MasTemp = TempMill
                GoTo Hundred
            ElseIf Z = 4 Then
                MasTemp = TempMill
                GoTo Thousand
            ElseIf Z = 5 Then
                tempStr = TempMill
                GoTo TenThousand
            Else
                GoTo HundredThousand
            End If
        End If
    End If
    If Z = 8 Then
        d(Y) = "Million"
        If Val(X) >= 10000000 And Val(X) < 100000000 And Mid(X, 3, 1) = 0 And Mid(X, 1, 4) = 0 Or Mid(X, 1, 5) = 0 And (Mid(X, 1, 2) < 20 Or Mid(X, 2, 1) = 0) Then
            i = Left(X, 2)
            AEInWords = a(i) & " " & d(Y)
            Exit Function
        ElseIf Val(X) > 10000000 And Val(X) < 100000000 And Mid(X, 3, 1) = 0 And Mid(X, 1, 2) > 20 And Mid(X, 2, 1) <> 0 Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            AEInWords = tempStr & " " & a(p) & " O " & a(Q) & " " & d(Y) '& " O " & A(m) & " " & A(k)
            Exit Function
        ElseIf Val(X) > 10000000 And Val(X) < 100000000 Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If i < 20 And p = 0 Then
                TempMill = tempStr & " " & a(i) & " " & d(Y) '& " O " & A(m) & " " & A(k)
            Else
                TempMill = tempStr & " " & a(p) & " O " & a(Q) & " " & d(Y) '& " O " & A(m) & " " & A(k)
            End If
            X = Mid(Val(X), 3)
            X = Val(X)
            Z = Len(X)
            If Z = 3 Then
                MasTemp = TempMill
                GoTo Hundred
            ElseIf Z = 4 Then
                MasTemp = TempMill
                GoTo Thousand
            ElseIf Z = 5 Then
                tempStr = TempMill
                GoTo TenThousand
            Else
                GoTo HundredThousand
            End If
        End If
     End If
    
    If Z = 9 Then
        d(Y) = "Million"
        B(Y) = "Miya"
        If Val(X) = 100000000 Then '1
            AEInWords = TempMill & " " & B(Y) & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 200000000 Then
            AEInWords = TempMill & " " & "Methane" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 300000000 Then
            AEInWords = TempMill & " " & "Talat Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 400000000 Then
            AEInWords = TempMill & " " & "Arba Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 500000000 Then
            AEInWords = TempMill & " " & "Khams Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 600000000 Then
            AEInWords = TempMill & " " & "Sitt Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 700000000 Then
            AEInWords = TempMill & " " & "Sabha Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 800000000 Then
            AEInWords = TempMill & " " & "Taman Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) = 900000000 Then
            AEInWords = TempMill & " " & "Tees Miya" & " O " & d(Y)
            Exit Function
        ElseIf Val(X) >= 100000000 And Val(X) < 1000000000 And Mid(X, 3, 1) = 0 And Mid(X, 1, 4) = 0 Or Mid(X, 1, 5) = 0 And (Mid(X, 1, 2) < 20 Or Mid(X, 2, 1) = 0) Then
            i = Left(X, 2)
            AEInWords = a(i) & " " & B(Y) & " " & d(Y)
            Exit Function
        ElseIf Val(X) > 100000000 And Val(X) < 1000000000 And Mid(X, 3, 1) = 0 And Mid(X, 1, 2) > 20 And Mid(X, 2, 1) <> 0 Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            AEInWords = tempStr & " " & a(p) & " O " & a(Q) & " " & B(Y) & " " & d(Y) '& " O " & A(m) & " " & A(k)
            Exit Function

        End If
        F3 = Mid(Val(X), 1, 3)
        Z = Len(Mid(Val(X), 4))
        If Val(F3) = 100 Then '1
            TempMill = B(Y) & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) = 200 Then
            TempMill = "Methane" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) = 300 Then
            TempMill = "Talat Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) = 400 Then
            TempMill = "Arba Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) = 500 Then
            TempMill = "Khams Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) = 600 Then
            TempMill = "Sitt Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) = 700 Then
            TempMill = "Sabha Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) = 800 Then
            TempMill = "Taman Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) = 900 Then
            TempMill = "Tees Miya" & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        End If
        If Val(F3) > 200 And Val(F3) < 300 And Mid(F3, 3) = 0 Then
            i = Right(F3, 2)
            TempMill = "Methane" & " O " & a(i) & " O " & d(Y)  '& A(i) & " "
            X = Mid(Val(X), 4)
            GoTo HundredThousand
         ElseIf F3 > 200 And Val(F3) < 300 And Right(F3, 2) < 20 Then
            i = Right(F3, 2)
            TempMill = "Methane" & " " & a(i) & " O " & d(Y)  '& " O " & A(i)
            GoTo HundredThousand
        ElseIf Val(F3) > 200 And Val(F3) < 300 Then
            i = Right(F3, 2)
            k = Right(i, 1)
            m = Val(i) - k
            TempMill = "Methane" & " " & a(k) & " O " & a(m) & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf F3 > 100 And Val(F3) < 1000 And (Mid(F3, 3) = 0 And Mid(F3, 2, 1) = 0) Then
            i = Right(F3, 2)
            j = Left(F3, 1)
            TempMill = a(j) & " " & B(Y) & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf F3 > 100 And Val(F3) < 1000 And (Mid(F3, 3) = 0 Or Mid(F3, 2, 1) = 0) Then
            i = Right(F3, 2)
            j = Left(F3, 1)
            TempMill = B(Y) & " " & a(j) & " O " & d(Y)  '& A(i) & " "
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) > 100 And Val(F3) < 1000 And Right(F3, 2) < 20 Then
            i = Right(F3, 2)
            k = Right(i, 1)
            j = Left(F3, 1)
            TempMill = a(j) & " " & B(Y) & " O " & a(i) & " O " & d(Y)
            X = Mid(Val(X), 4)
            GoTo HundredThousand
        ElseIf Val(F3) > 100 And Val(F3) < 1000 Then
            i = Right(F3, 2)
            k = Right(i, 1)
            j = Left(F3, 1)
            m = Val(i) - k
            If j = 1 Then
                TempMill = B(Y) & " " & a(k) & " O " & a(m) & " O " & d(Y)
            Else
                TempMill = a(j) & " " & B(Y) & " " & a(k) & " O " & a(m) & " O " & d(Y)
            End If
        X = Mid(Val(X), 4)
        p = Val(X)
        Z = Len(p)
        If Len(p) = 6 Then
            GoTo HundredThousand
        ElseIf Len(p) = 5 Then
            tempStr = TempMill
            GoTo TenThousand
        ElseIf Len(p) = 4 Then
            MasTemp = TempMill
            GoTo Thousand
        End If
        End If
    End If
    
        
    If Z = 2 And Len(i) = 1 Then
        AEInWords = a(i) & " O " & a(m)
    ElseIf Z = 3 And Len(i) = 2 Then
        If j = 1 Then
            AEInWords = MasTemp & " " & B(Z) & " " & a(k) & " O " & a(m)
        Else
            AEInWords = MasTemp & " " & a(j) & " " & B(Z) & " " & a(k) & " O " & a(m)
        End If
    End If
    
End Function
Public Function inwordsusd(number As Currency) As String
    Dim a(0 To 99) As String
    Dim B(3) As String
    Dim C(4) As String
    Dim d(6) As String
    Dim e(7) As String
    Dim U, V
    Dim X, i, j, k, m, n, p, Q, Y, Z As Variant
    Dim SarTemp As String
    Dim TempMill As String
    Dim tempStr As String
    Dim MasTemp As String
            X = number
            V = Int(number)
            U = X - V
            U = Format(U, "0.00")
            U = Val(Mid(U, 3, 2))
           
           X = Int(X)
    a(0) = "Zero"
    a(1) = "One"
    a(2) = "Two"
    a(3) = "Three"
    a(4) = "Four"
    a(5) = "Five"
    a(6) = "Six"
    a(7) = "Seven"
    a(8) = "Eight"
    a(9) = "Nine"
    a(10) = "Ten"
    a(11) = "Eleven"
    a(12) = "Twelve"
    a(13) = "Thirteen"
    a(14) = "Fourteen"
    a(15) = "Fifteen"
    a(16) = "Sixteen"
    a(17) = "Seventeen"
    a(18) = "Eighteen"
    a(19) = "Nineteen"
    a(20) = "Twenty"
    a(21) = "Twenty One"
    a(22) = "Twenty Two"
    a(23) = "Twenty Three"
    a(24) = "Twenty Four"
    a(25) = "Twenty Five"
    a(26) = "Twenty Six"
    a(27) = "Twenty Seven"
    a(28) = "Twenty Eight"
    a(29) = "Twenty Nine"
    a(30) = "Thirty"
    a(31) = "Thirty One"
    a(32) = "Thirty Two"
    a(33) = "Thirty Three"
    a(34) = "Thirty Four"
    a(35) = "Thirty Five"
    a(36) = "Thirty Six"
    a(37) = "Thirty Seven"
    a(38) = "Thirty Eight"
    a(39) = "Thirty Nine"
    a(40) = "Fourty"
    a(41) = "Fourty One"
    a(42) = "Fourty Two"
    a(43) = "Fourty Three"
    a(44) = "Fourty Four"
    a(45) = "Fourty Five"
    a(46) = "Fourty Six"
    a(47) = "Fourty Seven"
    a(48) = "Fourty Eight"
    a(49) = "Fourty Nine"
    a(50) = "Fifty"
    a(51) = "Fifty One"
    a(52) = "Fifty Two"
    a(53) = "Fifty Three"
    a(54) = "Fifty Four"
    a(55) = "Fifty Five"
    a(56) = "Fifty Six"
    a(57) = "Fifty Seven"
    a(58) = "Fifty Eight"
    a(59) = "Fifty Nine"
    a(60) = "Sixty"
    a(61) = "Sixty One"
    a(62) = "Sixty Two"
    a(63) = "Sixty Three"
    a(64) = "Sixty Four"
    a(65) = "Sixty Five"
    a(66) = "Sixty Six"
    a(67) = "Sixty Seven"
    a(68) = "Sixty Eight"
    a(69) = "Sixty Nine"
    a(70) = "Seventy"
    a(71) = "Seventy One"
    a(72) = "Seventy Two"
    a(73) = "Seventy Three"
    a(74) = "Seventy Four"
    a(75) = "Seventy Five"
    a(76) = "Seventy Six"
    a(77) = "Seventy Seven"
    a(78) = "Seventy Eight"
    a(79) = "Seventy Nine"
    a(80) = "Eighty"
    a(81) = "Eighty One"
    a(82) = "Eighty Two"
    a(83) = "Eighty Three"
    a(84) = "Eighty Four"
    a(85) = "Eighty Five"
    a(86) = "Eighty Six"
    a(87) = "Eighty Seven"
    a(88) = "Eighty Eight"
    a(89) = "Eighty Nine"
    a(90) = "Ninty"
    a(91) = "Ninty One"
    a(92) = "Ninty Two"
    a(93) = "Ninty Three"
    a(94) = "Ninty Four"
    a(95) = "Ninty Five"
    a(96) = "Ninty Six"
    a(97) = "Ninty Seven"
    a(98) = "Ninty Eight"
    a(99) = "Ninty Nine"
    
        tempStr = ""
        MasTemp = ""
        Z = Len(X)
   
    If Z = 1 Then
     If U = 0 Then
       inwordsusd = a(X)
     Else
       inwordsusd = a(X) & " and Cents " & a(U)
       'Inwordsusd = a(X) & " and  " & a(U)
     End If
    End If
    If Z = 2 Then
        If Val(X) > 20 And Val(X) < 100 And Right(X, 1) <> 0 Then
            i = Right(X, 1)
            m = Val(X) - i
        Else
            If U = 0 Then
             inwordsusd = a(X)
            Else
              inwordsusd = a(X) & " and Cents " & a(U)
             ' Inwordsusd = a(X) & " and  " & a(U)
            End If
            Exit Function
        End If
    End If
    
Hundred:
    If Z = 3 Then
        B(Z) = "Hundred"
        If Val(X) = 100 Then
            If U <> 0 Then
             inwordsusd = MasTemp & " " & a(1) & " " & B(Z) & " and  Cents " & a(U)
             'Inwordsusd = MasTemp & " " & a(1) & " " & B(Z) & " and " & a(U)
            Else
             inwordsusd = MasTemp & " " & a(1) & " " & B(Z)
            End If
            Exit Function
        ElseIf X > 100 And Val(X) < 1000 And (Mid(X, 3) = 0 And Mid(X, 2, 1) = 0) Then
            i = Right(X, 2)
            j = Left(X, 1)
            If U = 0 Then
             inwordsusd = MasTemp & " " & a(j) & " " & B(Z)
            Else
             inwordsusd = MasTemp & " " & a(j) & " " & B(Z) & " and  Cents " & a(U)
             'Inwordsusd = MasTemp & " " & a(j) & " " & B(Z) & " and  " & a(U)
            End If
            Exit Function
        ElseIf X > 100 And Val(X) < 1000 And (Mid(X, 3) = 0 Or Mid(X, 2, 1) = 0) Then
            i = Right(X, 2)
            j = Left(X, 1)
            If U = 0 Then
             inwordsusd = MasTemp & " " & a(j) & " " & B(Z) & " and " & a(i)
            Else
             inwordsusd = MasTemp & " " & a(j) & " " & B(Z) & " " & a(i) & " and  Cents " & a(U)
             'Inwordsusd = MasTemp & " " & a(j) & " " & B(Z) & " " & a(i) & " and  " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100 And Val(X) < 1000 And Right(X, 2) < 20 Then
            i = Right(X, 2)
            k = Right(i, 1)
            j = Left(X, 1)
            If U = 0 Then
             inwordsusd = MasTemp & " " & a(j) & " " & B(Z) & " and " & a(i)
            Else
             inwordsusd = MasTemp & " " & a(j) & " " & B(Z) & " " & a(i) & " and  Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100 And Val(X) < 1000 Then
            i = Right(X, 2)
            k = Right(i, 1)
            j = Left(X, 1)
            m = Val(i) - k
        End If
    End If

Thousand:
    If Z = 4 Then
        C(Z) = "Thousand"
        If Val(X) = 1000 Then
            inwordsusd = MasTemp & " " & a(1) & " " & C(Z)
            Exit Function
        ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4) = 0 Then
            j = Left(X, 1)
            If U = 0 Then
            inwordsusd = MasTemp & " " & a(j) & " " & C(Z)
            Else
            inwordsusd = MasTemp & " " & a(j) & " " & C(Z) & " and  Cents " & a(U)
            End If
            Exit Function
        ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Right(X, 2)
            j = Left(X, 1)
            If U = 0 Then
            inwordsusd = MasTemp & " " & a(j) & " " & C(Z) & " and " & a(i)
            Else
            inwordsusd = MasTemp & " " & a(j) & " " & C(Z) & " " & a(i) & " and  Cents " & a(U)
            End If
            Exit Function
         ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Right(X, 2)
            j = Left(X, 1)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            If U = 0 Then
             inwordsusd = MasTemp & " " & a(j) & " " & C(Z) & " and " & a(m) & " " & a(k)
            Else
             inwordsusd = MasTemp & " " & a(j) & " " & C(Z) & " " & a(m) & " " & a(k) & " and  Cents " & a(U)
            End If
            Exit Function
         ElseIf X > 1000 And Val(X) < 10000 Then
            j = Left(X, 1)
            If j = 1 Then
            MasTemp = MasTemp & " " & a(j) & " " & C(Z)
            Else
            MasTemp = MasTemp & " " & a(j) & " " & C(Z)
            End If
            X = Mid(X, 2)
            Z = Len(X)
            GoTo Hundred
        End If
    End If

TenThousand:
    If Z = 5 Then
        C(Y) = "Thousand"
        If Val(X) = 10000 Then '1
            If U = 0 Then
             inwordsusd = MasTemp & " " & a(10) & " " & C(Y)
            Else
             inwordsusd = MasTemp & " " & a(10) & " " & C(Y) & " and  Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            If U = 0 Then
             inwordsusd = MasTemp & " " & a(i) & " " & C(Y)
            Else
             inwordsusd = MasTemp & " " & a(i) & " " & C(Y) & " and  Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            j = Right(X, 2)
            If U = 0 Then
              inwordsusd = MasTemp & " " & a(i) & " " & C(Y) & " and " & a(j)
            Else
              inwordsusd = MasTemp & " " & a(i) & " " & C(Y) & " " & a(j) & " and  Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If U = 0 Then
             inwordsusd = MasTemp & " " & a(i) & " " & C(Y) & " and " & a(m) & " " & a(k)
            Else
             inwordsusd = MasTemp & " " & a(i) & " " & C(Y) & " " & a(m) & " " & a(k) & " and  Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) > 20 And Mid(Left(X, 2), 2, 1) <> 0) And (Right(X, 2) > 20 And Mid(Right(X, 2), 2, 1) <> 0) Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If U = 0 Then
             inwordsusd = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " and " & a(m) & " " & a(k)
            Else
             inwordsusd = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " " & a(m) & " " & a(k) & " and  Cents " & a(U)
            End If
            Exit Function
         ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If j <> 0 Then
              If U = 0 Then
               inwordsusd = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " and " & a(j)
              Else
               inwordsusd = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " " & a(j) & " and  Cents " & a(U)
              End If
              
            Else
              If U = 0 Then
               inwordsusd = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y)
              Else
               inwordsusd = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " and  Cents " & a(U)
              End If
              
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Left(X, 2) > 20 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5) = 0 Then
            i = Left(X, 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            If U = 0 Then
              inwordsusd = MasTemp & " " & a(m) & " " & a(k) & " " & C(Y)
            Else
              inwordsusd = MasTemp & " " & a(m) & " " & a(k) & " " & C(Y) & " and  Cents " & a(U)
            End If
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Right(X, 2) < 20 And Left(X, 2) < 20 Then
            i = Left(X, 2)
            MasTemp = MasTemp & " " & a(i) & " " & C(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 10000 And Val(X) < 100000 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            MasTemp = MasTemp & " " & a(i) & " " & C(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 10000 And Val(X) < 100000 Then '17
            i = Left(X, 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            MasTemp = MasTemp & " " & a(m) & " " & a(k) & " " & C(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Hundred
        End If
    End If
    
HundredThousand:
    If Z = 6 Then
        If Mid(Val(X), 2, 1) = 0 And Mid(Val(X), 3, 1) = 0 Then
           d(Y) = "Hundred Thousand"
         Else
           d(Y) = "Hundred"
        End If
        If Val(X) = 100000 Then '1
            If U = 0 Then
             inwordsusd = a(1) & " " & d(Y)
            Else
             inwordsusd = a(1) & " " & d(Y) & " and  Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6) = 0 Then
            i = Left(X, 1)
            'Inwordsusd = a(i) & " " & d(Y) & "s"
            If U = 0 Then
             inwordsusd = a(i) & " " & d(Y)
            Else
             inwordsusd = a(i) & " " & d(Y) & " and  Cents " & a(U)
            End If
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 1)
            j = Right(X, 2)
            If U = 0 Then
             inwordsusd = a(i) & " " & d(Y) & " and " & a(j)
            Else
              inwordsusd = a(i) & " " & d(Y) & " " & a(j) & " and  Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Left(X, 1)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If U = 0 Then
             inwordsusd = a(i) & " " & d(Y) & " and " & a(m) & " " & a(k)
            Else
             inwordsusd = a(i) & " " & d(Y) & " " & a(m) & " " & a(k) & " and  Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 Then
            i = Left(X, 1)
            If i = 1 Then
                MasTemp = a(i) & " " & d(Y)
            Else
               'MasTemp = a(i) & " " & d(Y) & "s"
                MasTemp = a(i) & " " & d(Y)
            End If
            X = Mid(X, 4)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 Then
            i = Left(X, 1)
            If i = 1 Then
            MasTemp = a(i) & " " & d(Y)
            Else
            'MasTemp = a(i) & " " & d(Y) & "s"
            MasTemp = a(i) & " " & d(Y)
            End If
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Thousand
        ElseIf Val(X) > 100000 And Val(X) < 1000000 Then
            i = Left(X, 1)
            If i = 1 Then
                MasTemp = MasTemp & " " & a(i) & " " & d(Y)
            Else
                'tempStr = a(i) & " " & d(Y) & "s"
                MasTemp = MasTemp & " " & a(i) & " " & d(Y)
            End If
            X = Mid(X, 2)
            Z = Len(X)
            GoTo TenThousand
        End If
    End If
    
Million:
    If Z = 7 Then
        d(Y) = "Million"
        If Val(X) = 1000000 Then
            'Inwordsusd = a(10) & " " & d(Y)
            If U = 0 Then
             inwordsusd = a(1) & " " & d(Y)
            Else
             inwordsusd = a(1) & " " & d(Y) & " and Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 1)
            If U = 0 Then
             inwordsusd = a(i) & " " & d(Y)
            Else
             inwordsusd = a(i) & " " & d(Y) & " and Cents " & a(U)
            End If
              
            Exit Function
        'ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 Then
         ElseIf Val(X) > 1000000 And Val(X) < 10000000 Then
            i = Mid(Val(X), 1, 1)
            X = Mid(Val(X), 2)
            X = Val(X)
            Z = Len(X)
            TempMill = a(i) & " " & d(Y)
            If Z = 3 Then
              MasTemp = TempMill
              GoTo Hundred
            ElseIf Z = 4 Then
              MasTemp = TempMill
              GoTo Thousand
            ElseIf Z = 5 Then
              MasTemp = TempMill
              GoTo TenThousand
            Else
              MasTemp = TempMill
              GoTo HundredThousand
            End If
         End If
       End If
       
           
    If Z = 8 Then
        e(Y) = "Million"
        If Val(X) = 10000000 Then
            If U = 0 Then
             inwordsusd = a(10) & " " & e(Y)
            Else
             inwordsusd = a(10) & " " & e(Y) & " and Cents " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000000 And Val(X) < 100000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And Mid(X, 8, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            If U = 0 Then
             inwordsusd = a(i) & " " & e(Y)
            Else
             inwordsusd = a(i) & " " & e(Y) & " and Cents " & a(U)
            End If
              
            Exit Function
        'ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 Then
         ElseIf Val(X) > 10000000 And Val(X) < 100000000 Then
            i = Mid(Val(X), 1, 2)
            X = Mid(Val(X), 3)
            X = Val(X)
            Z = Len(X)
            TempMill = a(i) & " " & e(Y)
            If Z = 3 Then
              MasTemp = TempMill
              GoTo Hundred
            ElseIf Z = 4 Then
              MasTemp = TempMill
              GoTo Thousand
            ElseIf Z = 5 Then
              MasTemp = TempMill
              GoTo TenThousand
            ElseIf Z = 6 Then
              MasTemp = TempMill
              GoTo HundredThousand
            Else
              MasTemp = TempMill
              GoTo Million
            End If
         End If
         End If
    
    
    If Z = 2 And Len(i) = 1 Then
        If U = 0 Then
         inwordsusd = a(m) & " " & a(i)
        Else
         inwordsusd = a(m) & " " & a(i) & " and Cents " & a(U)
        End If
    ElseIf Z = 3 And Len(i) = 2 Then
        If U = 0 Then
         inwordsusd = MasTemp & " " & a(j) & " " & B(Z) & " and " & a(m) & " " & a(k)
        Else
         inwordsusd = MasTemp & " " & a(j) & " " & B(Z) & " " & a(m) & " " & a(k) & " and Cents " & a(U)
        End If
    End If


End Function
Public Function Inwords(number As Currency) As String

    Dim a(0 To 99) As String
    Dim B(3) As String
    Dim C(4) As String
    Dim d(6) As String
    Dim e(7) As String
    Dim U, V
    Dim X, i, j, k, m, n, p, Q, Y, Z As Variant
    Dim SarTemp As String
    Dim TempMill As String
    Dim tempStr As String
    Dim MasTemp As String
            X = number
            V = Int(number)
            U = X - V
            U = Format(U, "0.00")
            U = Val(Mid(U, 3, 2))
           
           X = Int(X)
    a(0) = "Zero"
    a(1) = "One"
    a(2) = "Two"
    a(3) = "Three"
    a(4) = "Four"
    a(5) = "Five"
    a(6) = "Six"
    a(7) = "Seven"
    a(8) = "Eight"
    a(9) = "Nine"
    a(10) = "Ten"
    a(11) = "Eleven"
    a(12) = "Twelve"
    a(13) = "Thirteen"
    a(14) = "Fourteen"
    a(15) = "Fifteen"
    a(16) = "Sixteen"
    a(17) = "Seventeen"
    a(18) = "Eighteen"
    a(19) = "Nineteen"
    a(20) = "Twenty"
    a(21) = "Twenty One"
    a(22) = "Twenty Two"
    a(23) = "Twenty Three"
    a(24) = "Twenty Four"
    a(25) = "Twenty Five"
    a(26) = "Twenty Six"
    a(27) = "Twenty Seven"
    a(28) = "Twenty Eight"
    a(29) = "Twenty Nine"
    a(30) = "Thirty"
    a(31) = "Thirty One"
    a(32) = "Thirty Two"
    a(33) = "Thirty Three"
    a(34) = "Thirty Four"
    a(35) = "Thirty Five"
    a(36) = "Thirty Six"
    a(37) = "Thirty Seven"
    a(38) = "Thirty Eight"
    a(39) = "Thirty Nine"
    a(40) = "Fourty"
    a(41) = "Fourty One"
    a(42) = "Fourty Two"
    a(43) = "Fourty Three"
    a(44) = "Fourty Four"
    a(45) = "Fourty Five"
    a(46) = "Fourty Six"
    a(47) = "Fourty Seven"
    a(48) = "Fourty Eight"
    a(49) = "Fourty Nine"
    a(50) = "Fifty"
    a(51) = "Fifty One"
    a(52) = "Fifty Two"
    a(53) = "Fifty Three"
    a(54) = "Fifty Four"
    a(55) = "Fifty Five"
    a(56) = "Fifty Six"
    a(57) = "Fifty Seven"
    a(58) = "Fifty Eight"
    a(59) = "Fifty Nine"
    a(60) = "Sixty"
    a(61) = "Sixty One"
    a(62) = "Sixty Two"
    a(63) = "Sixty Three"
    a(64) = "Sixty Four"
    a(65) = "Sixty Five"
    a(66) = "Sixty Six"
    a(67) = "Sixty Seven"
    a(68) = "Sixty Eight"
    a(69) = "Sixty Nine"
    a(70) = "Seventy"
    a(71) = "Seventy One"
    a(72) = "Seventy Two"
    a(73) = "Seventy Three"
    a(74) = "Seventy Four"
    a(75) = "Seventy Five"
    a(76) = "Seventy Six"
    a(77) = "Seventy Seven"
    a(78) = "Seventy Eight"
    a(79) = "Seventy Nine"
    a(80) = "Eighty"
    a(81) = "Eighty One"
    a(82) = "Eighty Two"
    a(83) = "Eighty Three"
    a(84) = "Eighty Four"
    a(85) = "Eighty Five"
    a(86) = "Eighty Six"
    a(87) = "Eighty Seven"
    a(88) = "Eighty Eight"
    a(89) = "Eighty Nine"
    a(90) = "Ninty"
    a(91) = "Ninty One"
    a(92) = "Ninty Two"
    a(93) = "Ninty Three"
    a(94) = "Ninty Four"
    a(95) = "Ninty Five"
    a(96) = "Ninty Six"
    a(97) = "Ninty Seven"
    a(98) = "Ninty Eight"
    a(99) = "Ninty Nine"
    
        tempStr = ""
        MasTemp = ""
        Z = Len(X)
   
    If Z = 1 Then
     If U = 0 Then
       Inwords = a(X)
     Else
       Inwords = a(X) & " and Fills " & a(U)
       'Inwords = a(X) & " and  " & a(U)
     End If
    End If
    If Z = 2 Then
        If Val(X) > 20 And Val(X) < 100 And Right(X, 1) <> 0 Then
            i = Right(X, 1)
            m = Val(X) - i
        Else
            If U = 0 Then
             Inwords = a(X)
            Else
              Inwords = a(X) & " and Fills " & a(U)
             ' Inwords = a(X) & " and  " & a(U)
            End If
            Exit Function
        End If
    End If
    
Hundred:
    If Z = 3 Then
        B(Z) = "Hundred"
        If Val(X) = 100 Then
            If U <> 0 Then
             Inwords = MasTemp & " " & a(1) & " " & B(Z) & " and  Fills " & a(U)
             'Inwords = MasTemp & " " & a(1) & " " & B(Z) & " and " & a(U)
            Else
             Inwords = MasTemp & " " & a(1) & " " & B(Z)
            End If
            Exit Function
        ElseIf X > 100 And Val(X) < 1000 And (Mid(X, 3) = 0 And Mid(X, 2, 1) = 0) Then
            i = Right(X, 2)
            j = Left(X, 1)
            If U = 0 Then
             Inwords = MasTemp & " " & a(j) & " " & B(Z)
            Else
             Inwords = MasTemp & " " & a(j) & " " & B(Z) & " and  Fills " & a(U)
             'Inwords = MasTemp & " " & a(j) & " " & B(Z) & " and  " & a(U)
            End If
            Exit Function
        ElseIf X > 100 And Val(X) < 1000 And (Mid(X, 3) = 0 Or Mid(X, 2, 1) = 0) Then
            i = Right(X, 2)
            j = Left(X, 1)
            If U = 0 Then
             Inwords = MasTemp & " " & a(j) & " " & B(Z) & " and " & a(i)
            Else
             Inwords = MasTemp & " " & a(j) & " " & B(Z) & " " & a(i) & " and  Fills " & a(U)
             'Inwords = MasTemp & " " & a(j) & " " & B(Z) & " " & a(i) & " and  " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100 And Val(X) < 1000 And Right(X, 2) < 20 Then
            i = Right(X, 2)
            k = Right(i, 1)
            j = Left(X, 1)
            If U = 0 Then
             Inwords = MasTemp & " " & a(j) & " " & B(Z) & " and " & a(i)
            Else
             Inwords = MasTemp & " " & a(j) & " " & B(Z) & " " & a(i) & " and  Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100 And Val(X) < 1000 Then
            i = Right(X, 2)
            k = Right(i, 1)
            j = Left(X, 1)
            m = Val(i) - k
        End If
    End If

Thousand:
    If Z = 4 Then
        C(Z) = "Thousand"
        If Val(X) = 1000 Then
            Inwords = MasTemp & " " & a(1) & " " & C(Z)
            Exit Function
        ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4) = 0 Then
            j = Left(X, 1)
            If U = 0 Then
            Inwords = MasTemp & " " & a(j) & " " & C(Z)
            Else
            Inwords = MasTemp & " " & a(j) & " " & C(Z) & " and  Fills " & a(U)
            End If
            Exit Function
        ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Right(X, 2)
            j = Left(X, 1)
            If U = 0 Then
            Inwords = MasTemp & " " & a(j) & " " & C(Z) & " and " & a(i)
            Else
            Inwords = MasTemp & " " & a(j) & " " & C(Z) & " " & a(i) & " and  Fills " & a(U)
            End If
            Exit Function
         ElseIf X > 1000 And Val(X) < 10000 And Mid(X, 2, 1) = 0 And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Right(X, 2)
            j = Left(X, 1)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            If U = 0 Then
             Inwords = MasTemp & " " & a(j) & " " & C(Z) & " and " & a(m) & " " & a(k)
            Else
             Inwords = MasTemp & " " & a(j) & " " & C(Z) & " " & a(m) & " " & a(k) & " and  Fills " & a(U)
            End If
            Exit Function
         ElseIf X > 1000 And Val(X) < 10000 Then
            j = Left(X, 1)
            If j = 1 Then
            MasTemp = MasTemp & " " & a(j) & " " & C(Z)
            Else
            MasTemp = MasTemp & " " & a(j) & " " & C(Z)
            End If
            X = Mid(X, 2)
            Z = Len(X)
            GoTo Hundred
        End If
    End If

TenThousand:
    If Z = 5 Then
        C(Y) = "Thousand"
        If Val(X) = 10000 Then '1
            If U = 0 Then
             Inwords = MasTemp & " " & a(10) & " " & C(Y)
            Else
             Inwords = MasTemp & " " & a(10) & " " & C(Y) & " and  Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            If U = 0 Then
             Inwords = MasTemp & " " & a(i) & " " & C(Y)
            Else
             Inwords = MasTemp & " " & a(i) & " " & C(Y) & " and  Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            j = Right(X, 2)
            If U = 0 Then
              Inwords = MasTemp & " " & a(i) & " " & C(Y) & " and " & a(j)
            Else
              Inwords = MasTemp & " " & a(i) & " " & C(Y) & " " & a(j) & " and  Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If U = 0 Then
             Inwords = MasTemp & " " & a(i) & " " & C(Y) & " and " & a(m) & " " & a(k)
            Else
             Inwords = MasTemp & " " & a(i) & " " & C(Y) & " " & a(m) & " " & a(k) & " and  Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And (Left(X, 2) > 20 And Mid(Left(X, 2), 2, 1) <> 0) And (Right(X, 2) > 20 And Mid(Right(X, 2), 2, 1) <> 0) Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If U = 0 Then
             Inwords = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " and " & a(m) & " " & a(k)
            Else
             Inwords = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " " & a(m) & " " & a(k) & " and  Fills " & a(U)
            End If
            Exit Function
         ElseIf Val(X) > 10000 And Val(X) < 100000 And Mid(X, 3, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            p = Right(i, 1)
            Q = Val(i) - Val(p)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If j <> 0 Then
              If U = 0 Then
               Inwords = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " and " & a(j)
              Else
               Inwords = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " " & a(j) & " and  Fills " & a(U)
              End If
              
            Else
              If U = 0 Then
               Inwords = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y)
              Else
               Inwords = MasTemp & " " & a(Q) & " " & a(p) & " " & C(Y) & " and  Fills " & a(U)
              End If
              
            End If
            Exit Function
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Left(X, 2) > 20 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5) = 0 Then
            i = Left(X, 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            If U = 0 Then
              Inwords = MasTemp & " " & a(m) & " " & a(k) & " " & C(Y)
            Else
              Inwords = MasTemp & " " & a(m) & " " & a(k) & " " & C(Y) & " and  Fills " & a(U)
            End If
        ElseIf Val(X) > 10000 And Val(X) < 100000 And Right(X, 2) < 20 And Left(X, 2) < 20 Then
            i = Left(X, 2)
            MasTemp = MasTemp & " " & a(i) & " " & C(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 10000 And Val(X) < 100000 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            MasTemp = MasTemp & " " & a(i) & " " & C(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 10000 And Val(X) < 100000 Then '17
            i = Left(X, 2)
            k = Right(i, 1)
            m = Val(i) - Val(k)
            MasTemp = MasTemp & " " & a(m) & " " & a(k) & " " & C(Y)
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Hundred
        End If
    End If
    
HundredThousand:
    If Z = 6 Then
        If Mid(Val(X), 2, 1) = 0 And Mid(Val(X), 3, 1) = 0 Then
           d(Y) = "Hundred Thousand"
         Else
           d(Y) = "Hundred"
        End If
        If Val(X) = 100000 Then '1
            If U = 0 Then
             Inwords = a(1) & " " & d(Y)
            Else
             Inwords = a(1) & " " & d(Y) & " and  Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6) = 0 Then
            i = Left(X, 1)
            'Inwords = a(i) & " " & d(Y) & "s"
            If U = 0 Then
             Inwords = a(i) & " " & d(Y)
            Else
             Inwords = a(i) & " " & d(Y) & " and  Fills " & a(U)
            End If
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And (Right(X, 2) < 20 Or Mid(Right(X, 2), 2, 1) = 0) Then
            i = Left(X, 1)
            j = Right(X, 2)
            If U = 0 Then
             Inwords = a(i) & " " & d(Y) & " and " & a(j)
            Else
              Inwords = a(i) & " " & d(Y) & " " & a(j) & " and  Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(Right(X, 2), 2, 1) <> 0 Then
            i = Left(X, 1)
            j = Right(X, 2)
            k = Right(j, 1)
            m = Val(j) - Val(k)
            If U = 0 Then
             Inwords = a(i) & " " & d(Y) & " and " & a(m) & " " & a(k)
            Else
             Inwords = a(i) & " " & d(Y) & " " & a(m) & " " & a(k) & " and  Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 And Mid(X, 3, 1) = 0 Then
            i = Left(X, 1)
            If i = 1 Then
                MasTemp = a(i) & " " & d(Y)
            Else
               'MasTemp = a(i) & " " & d(Y) & "s"
                MasTemp = a(i) & " " & d(Y)
            End If
            X = Mid(X, 4)
            Z = Len(X)
            GoTo Hundred
        ElseIf Val(X) > 100000 And Val(X) < 1000000 And Mid(X, 2, 1) = 0 Then
            i = Left(X, 1)
            If i = 1 Then
            MasTemp = a(i) & " " & d(Y)
            Else
            'MasTemp = a(i) & " " & d(Y) & "s"
            MasTemp = a(i) & " " & d(Y)
            End If
            X = Mid(X, 3)
            Z = Len(X)
            GoTo Thousand
        ElseIf Val(X) > 100000 And Val(X) < 1000000 Then
            i = Left(X, 1)
            If i = 1 Then
                MasTemp = MasTemp & " " & a(i) & " " & d(Y)
            Else
                'tempStr = a(i) & " " & d(Y) & "s"
                MasTemp = MasTemp & " " & a(i) & " " & d(Y)
            End If
            X = Mid(X, 2)
            Z = Len(X)
            GoTo TenThousand
        End If
    End If
    
Million:
    If Z = 7 Then
        d(Y) = "Million"
        If Val(X) = 1000000 Then
            'Inwords = a(10) & " " & d(Y)
            If U = 0 Then
             Inwords = a(1) & " " & d(Y)
            Else
             Inwords = a(1) & " " & d(Y) & " and Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 1)
            If U = 0 Then
             Inwords = a(i) & " " & d(Y)
            Else
             Inwords = a(i) & " " & d(Y) & " and Fills " & a(U)
            End If
              
            Exit Function
        'ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 Then
         ElseIf Val(X) > 1000000 And Val(X) < 10000000 Then
            i = Mid(Val(X), 1, 1)
            X = Mid(Val(X), 2)
            X = Val(X)
            Z = Len(X)
            TempMill = a(i) & " " & d(Y)
            If Z = 3 Then
              MasTemp = TempMill
              GoTo Hundred
            ElseIf Z = 4 Then
              MasTemp = TempMill
              GoTo Thousand
            ElseIf Z = 5 Then
              MasTemp = TempMill
              GoTo TenThousand
            Else
              MasTemp = TempMill
              GoTo HundredThousand
            End If
         End If
       End If
       
           
    If Z = 8 Then
        e(Y) = "Million"
        If Val(X) = 10000000 Then
            If U = 0 Then
             Inwords = a(10) & " " & e(Y)
            Else
             Inwords = a(10) & " " & e(Y) & " and Fills " & a(U)
            End If
            Exit Function
        ElseIf Val(X) > 10000000 And Val(X) < 100000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And Mid(X, 8, 1) = 0 And (Left(X, 2) < 20 Or Mid(Left(X, 2), 2, 1) = 0) Then
            i = Left(X, 2)
            If U = 0 Then
             Inwords = a(i) & " " & e(Y)
            Else
             Inwords = a(i) & " " & e(Y) & " and Fills " & a(U)
            End If
              
            Exit Function
        'ElseIf Val(X) > 1000000 And Val(X) < 10000000 And Mid(X, 3, 1) = 0 And Mid(X, 4, 1) = 0 And Mid(X, 5, 1) = 0 And Mid(X, 6, 1) = 0 And Mid(X, 7, 1) = 0 And Mid(Left(X, 2), 2, 1) <> 0 Then
         ElseIf Val(X) > 10000000 And Val(X) < 100000000 Then
            i = Mid(Val(X), 1, 2)
            X = Mid(Val(X), 3)
            X = Val(X)
            Z = Len(X)
            TempMill = a(i) & " " & e(Y)
            If Z = 3 Then
              MasTemp = TempMill
              GoTo Hundred
            ElseIf Z = 4 Then
              MasTemp = TempMill
              GoTo Thousand
            ElseIf Z = 5 Then
              MasTemp = TempMill
              GoTo TenThousand
            ElseIf Z = 6 Then
              MasTemp = TempMill
              GoTo HundredThousand
            Else
              MasTemp = TempMill
              GoTo Million
            End If
         End If
         End If
    
    
    If Z = 2 And Len(i) = 1 Then
        If U = 0 Then
         Inwords = a(m) & " " & a(i)
        Else
         Inwords = a(m) & " " & a(i) & " and Fills " & a(U)
        End If
    ElseIf Z = 3 And Len(i) = 2 Then
        If U = 0 Then
         Inwords = MasTemp & " " & a(j) & " " & B(Z) & " and " & a(m) & " " & a(k)
        Else
         Inwords = MasTemp & " " & a(j) & " " & B(Z) & " " & a(m) & " " & a(k) & " and Fills " & a(U)
        End If
    End If
    
End Function


Public Function AcadYear() As Variant

    If Month(Now) >= 9 And Month(Now) <= 12 Then
        AcadYear = Year(Now) + 1
    Else
        AcadYear = Year(Now)
    End If
    
End Function

Public Function Comma(Name As String) As String

    Dim Cnt As Integer
    Dim tempChar As String * 1
 
    Name = Trim(Name)
    tempStr = ""
    For Cnt = 1 To Len(Name) - 1
        tempChar = Mid(Name, Cnt, 1)
        tempStr = tempStr & tempChar
            
            If tempChar = "," Or Cnt = Len(Name) Or tempChar = Chr(13) Then
            If Len(tempStr) > 1 Then
                tempStr = UCase(Left(tempStr, 1)) & LCase(Mid(tempStr, 2)) & "," & Chr(13)
                Comma = Comma & tempStr
            End If
            
            tempStr = ""
        End If
    Next

End Function

Public Function NewLine(Name As String) As String

    Dim Cnt As Integer
    Dim tempChar As String * 1
 
    Name = Trim(Name)
    tempStr = ""
    For Cnt = 1 To Len(Name)
        tempChar = Mid(Name, Cnt, 1)
        tempStr = tempStr & tempChar

            If tempChar = "," Or Cnt = Len(Name) Then
            If Len(tempStr) > 1 Then
                tempStr = UCase(Left(tempStr, 1)) & LCase(Mid(tempStr, 2)) & Chr(13)
                NewLine = NewLine & tempStr
            End If
            tempStr = ""
        End If
    Next

End Function

Public Sub Main()

  '  txtGFColor = vbRed
  '  txtLFColor = vbBlack
  '  Load MainForm
  '  MainForm.Show
    End Sub

Public Function LPad(padString, padLen As Long, Optional padChar = " ") As Variant
    
    If padString = "" Or padChar = "" Then
        LPad = padString
        Exit Function
    ElseIf padLen = 0 Then
        LPad = ""
        Exit Function
    ElseIf padLen < 0 Then
        LPad = Mid(padString, Abs(padLen) + 1)
        Exit Function
    End If
    
    If Len(Trim(padString)) >= padLen Then
        tempStr = Left(padString, padLen)
    ElseIf Len(Trim(padString)) < padLen Then
        tempStr = String(padLen - Len(Trim(padString)), padChar) & padString
    End If
    
    LPad = tempStr

End Function

Public Function RPad(padString, padLen As Long, Optional padChar = " ") As Variant
       
    If padString = "" Or padChar = "" Then
        RPad = padString
        Exit Function
    ElseIf padLen = 0 Then
        RPad = ""
        Exit Function
    ElseIf padLen < 0 Then
        RPad = Left(padString, MaxOf(Len(padString) - Abs(padLen), 0))
        Exit Function
    End If
    
    If Len(Trim(padString)) >= padLen Then
        tempStr = Right(padString, padLen)
    ElseIf Len(Trim(padString)) < padLen Then
        tempStr = padString & String(padLen - Len(Trim(padString)), padChar)
    End If
    
    RPad = tempStr

End Function

Public Function CharTrim(charString As String, TrimChar) As String

    If charString = "" Or TrimChar = "" Then
        CharTrim = charString
        Exit Function
    ElseIf Len(TrimChar) = 1 Then
        TrimChar = Left(TrimChar, 1)
        tempStr = ""
        For tempNum = 1 To Len(charString)
            If Not Mid(charString, tempNum, Len(TrimChar)) = TrimChar Then _
                tempStr = tempStr + Mid(charString, tempNum, Len(TrimChar))
        Next
        CharTrim = tempStr
    ElseIf Len(TrimChar) > 1 Then
        TrimChar = Left(TrimChar, 1)
        tempStr = ""
        For tempNum = 1 To Len(charString)
            If Not Mid(charString, tempNum, Len(TrimChar)) = TrimChar Then _
                tempStr = tempStr + Mid(charString, tempNum, Len(TrimChar))
        Next
        CharTrim = tempStr
    End If
    
End Function

Public Function MsgError(errText As String, errTitle As String) As String

    MsgBox "The Following Error Occured :" & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
            & "Error : " & Err.Description & Chr(vbKeyReturn) _
            & "Error Code : " & Err.number & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
            & errText, vbCritical, errTitle

End Function

Public Function MinOf(Val1 As Variant, Val2 As Variant) As Variant
    
    MinOf = IIf(Val1 < Val2, Val1, Val2)

End Function

Public Function MaxOf(Val1 As Variant, Val2 As Variant) As Variant
    
    MaxOf = IIf(Val1 > Val2, Val1, Val2)

End Function

Public Function NameCase(Name As String) As String

    Dim Cnt As Integer
    Dim tempChar As String * 1
 
    Name = Trim(Name)
    tempStr = ""
    For Cnt = 1 To Len(Name)
        tempChar = Mid(Name, Cnt, 1)
        tempStr = tempStr & tempChar
        If tempChar = " " Or tempChar = "." Or _
            tempChar = Chr(vbKeyReturn) Or tempChar = Chr(10) _
            Or tempChar = "," Or Cnt = Len(Name) Then
            If Len(tempStr) > 0 Then
                tempStr = UCase(Left(tempStr, 1)) & LCase(Mid(tempStr, 2))
                NameCase = NameCase & tempStr
            End If
            tempStr = ""
        End If
    Next

End Function

Public Function Comma1(Name As String) As String

    Dim Cnt As Integer
    Dim tempChar As String * 1
 
    Name = Trim(Name)
    tempStr = ""
    For Cnt = 1 To Len(Name)
        tempChar = Mid(Name, Cnt, 1)
        tempStr = tempStr & tempChar
        'If tempChar = " " Or tempChar = "." Or _
            tempChar = Chr(vbKeyReturn) Or tempChar = Chr(10) _
            Or tempChar = "," Or cnt = Len(Name) Then
            If tempChar = "." Or Cnt = Len(Name) Or tempChar = Chr(13) Then
            If Len(tempStr) > 0 Then
                tempStr = UCase(Left(tempStr, 1)) & LCase(Mid(tempStr, 2)) & ":"
                Comma1 = Comma1 & tempStr
            End If
            tempStr = ""
        End If
    Next

End Function

Sub EnumFields(rsc As Recordset, intFldLen As Integer)

    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    Dim strTitle As String, strTemp As String

    ' Set the lngRecords variable to the number of
    ' records in the Recordset.
    lngRecords = rsc.RecordCount
    ' Set the lngFields variable to the number of
    ' fields in the Recordset.
    lngFields = rsc.Fields.Count
    
    Debug.Print "There are " & lngRecords _
        & " records containing " & lngFields _
        & " fields in the recordset."
    Debug.Print
    
    ' Form a string to print the column heading.
    strTitle = "Record  "
    For lngFldCount = 0 To lngFields - 1
        strTitle = strTitle _
        & Left(rsc.Fields(lngFldCount).Name _
        & Space(intFldLen), intFldLen)
    Next lngFldCount
    
    ' Print the column heading.
    Debug.Print strTitle
    Debug.Print
    
    ' Loop through the Recordset; print the record
    ' number and field values.
    rsc.MoveFirst
    For lngRecCount = 0 To lngRecords - 1

Debug.Print Right(Space(6) & _
            Str(lngRecCount), 6) & "  ";
        For lngFldCount = 0 To lngFields - 1
            ' Check for Null values.
            If IsNull(rsc.Fields(lngFldCount)) Then
                strTemp = "<null>"
            Else
                ' Set strTemp to the field contents.
                Select Case _
                    rsc.Fields(lngFldCount).Type
                    Case 11
                        strTemp = ""
                    Case dbText, dbMemo
                        strTemp = _
                            rsc.Fields(lngFldCount)

Case Else
                        strTemp = _
                            Str(rsc.Fields(lngFldCount))
                End Select
            End If
            Debug.Print Left(strTemp _
                & Space(intFldLen), intFldLen);
        Next lngFldCount
        Debug.Print
        rsc.MoveNext
    Next lngRecCount
End Sub
Function replacestr(Textin, ByVal searchstr As String, _
                    ByVal Replacement As String, _
                    ByVal CompMode As Integer)

  Dim Worktext As String, Pointer As Integer
   If IsNull(Textin) Then
    replacestr = Null
   Else
    Worktext = Textin
    Pointer = InStr(1, Worktext, searchstr, CompMode)
     Do While Pointer > 0
      Worktext = Left(Worktext, Pointer - 1) & Replacement & _
                 Mid(Worktext, Pointer + Len(searchstr))
                 
      Pointer = InStr(Pointer + Len(Replacement), Worktext, _
                 searchstr, CompMode)
                 
    Loop
    
    replacestr = Worktext
    
  
   End If
End Function

Function sqlfixup(Textin)
 sqlfixup = replacestr(Textin, "'", "''", 0)
End Function
Function jetsqlfixup(Textin)
 Dim Temp
  Temp = replacestr(Textin, "'", "''", 0)
  jetsqlfixup = replacestr(Temp, "|", "' & Chr(124) & '", 0)
End Function
 
Function findfirstfixup(Textin)
 Dim Temp
  Temp = replacestr(Textin, "'", "' & Chr(39) & '", 0)
  findfirstfixup = replacestr(Temp, "|", "' & Chr(124) & '", 0)

End Function



