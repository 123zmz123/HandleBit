Public Class HandleBit

    ' Set a bit to 1. input value length is 8.
    Public Shared Function SetOne(ByRef val As Byte, ByVal pos As Integer)
        Dim realPos As Integer
        realPos = pos
        If pos > (8 - 1) Then
            realPos = (pos Mod 8)
        End If
        val = val Or (1 << realPos)
    End Function

    Public Shared Function SetOne(ByRef val As UInt16, ByVal pos As Integer)
        Dim realPos As Integer
        realPos = pos
        If pos > (16 - 1) Then
            realPos = (pos Mod 16)
        End If
        val = val Or (1 << realPos)
    End Function

    Public Shared Function SetOne(ByRef val As UInt32, ByVal pos As Integer)
        Dim realPos As Integer
        realPos = pos
        If pos > (32 - 1) Then
            realPos = (pos Mod 32)
        End If
        val = val Or (1 << realPos)
    End Function
    Public Shared Function SetOne(ByRef val As UInt64, ByVal pos As Integer)
        Dim realPos As Integer
        realPos = pos
        If pos > (64 - 1) Then
            realPos = (pos Mod 64)
        End If
        val = val Or (1 << realPos)
    End Function

    Public Shared Function SetZero(ByRef val As Byte, ByVal pos As Integer)
        Dim realPos As Integer = pos
        Dim mask = 1
        If pos > (8 - 1) Then
            realPos = pos Mod 8
        End If
        mask = mask << realPos
        mask = Not mask
        val = val And mask
    End Function

    Public Shared Function SetZero(ByRef val As UInt16, ByVal pos As Integer)
        Dim realPos As Integer = pos
        Dim mask = 1
        If pos > (16 - 1) Then
            realPos = pos Mod 16
        End If
        mask = mask << realPos
        mask = Not mask
        val = val And mask

    End Function

    Public Shared Function SetZero(ByRef val As UInt32, ByVal pos As Integer)
        Dim realPos As Integer = pos
        Dim mask = 1
        If pos > (32 - 1) Then
            realPos = pos Mod 32
        End If
        mask = mask << realPos
        mask = Not mask
        val = val And mask

    End Function

    Public Shared Function SetZero(ByRef val As UInt64, ByVal pos As Integer)
        Dim realPos As Integer = pos
        Dim mask = 1
        If pos > (64 - 1) Then
            realPos = pos Mod 64
        End If
        mask = mask << realPos
        mask = Not mask
        val = val And mask
    End Function

    Public Shared Function DecToBin(ByVal s As Byte) As String
        Try
            Dim res As String
            Do
                res = (Str(s Mod 2) & res)
                s = s \ 2 ' divide exactly ,整除 , 割り切れる
            Loop Until s = 0
            Return res
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function DecToBin(ByVal s As UInt16) As String
        Try
            Dim res As String
            Do
                res = (Str(s Mod 2) & res)
                s = s \ 2 ' divide exactly ,整除 , 割り切れる
            Loop Until s = 0
            Return res
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function DecToBin(ByVal s As UInt32) As String
        Try
            Dim res As String
            Do
                res = (Str(s Mod 2) & res)
                s = s \ 2 ' divide exactly ,整除 , 割り切れる
            Loop Until s = 0
            Return res
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function DecToBin(ByVal s As UInt64) As String
        Try
            Dim res As String
            Do
                res = (Str(s Mod 2) & res)
                s = s \ 2 ' divide exactly ,整除 , 割り切れる
            Loop Until s = 0
            Return res
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function HexToBin(ByVal s As Byte) As String
        Try
            Dim res As String
            Do
                res = (Str(s Mod 2) & res)
                s = s \ 2 ' divide exactly ,整除 , 割り切れる
            Loop Until s = 0
            Return res
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function HexToBin(ByVal s As UInt16) As String
        Try
            Dim res As String
            Do
                res = (Str(s Mod 2) & res)
                s = s \ 2 ' divide exactly ,整除 , 割り切れる
            Loop Until s = 0
            Return res
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function HexToBin(ByVal s As UInt32) As String
        Try
            Dim res As String
            Do
                res = (Str(s Mod 2) & res)
                s = s \ 2 ' divide exactly ,整除 , 割り切れる
            Loop Until s = 0
            Return res
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function HexToBin(ByVal s As UInt64) As String
        Try
            Dim res As String
            Do
                res = (Str(s Mod 2) & res)
                s = s \ 2 ' divide exactly ,整除 , 割り切れる
            Loop Until s = 0
            Return res
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function BinToDec(ByVal b As String) As Integer
        Try
            Dim sorted = b.ToList()

            sorted.Sort()

            If sorted.First <> "0" And sorted.First <> "1" Then
                Throw New Exception("the binary string need only contain 0 or ")
            End If
            If sorted.Last <> "0" And sorted.Last <> "1" Then
                Throw New Exception("the binary string need only contain 0 or 1")
            End If


            Dim index = 0
            Dim res = 0
            For Each item In b.Reverse()
                res += Val(item) * (2 ^ index)
                index += 1

            Next
            Return res

        Catch ex As Exception
            Throw ex
        End Try
    End Function



    Public Shared Function SetBitsInRange8(ByRef d As Byte, ByVal val As Byte, ByVal pos As Integer, ByVal len As Integer)
        Try
            Dim init1 = 0, init2 = 0
            Dim Boundry = 8
            pos = pos Mod 8
            If len = 0 Or len > Boundry Then
                Return False
            End If


            init1 = Convert.ToByte(2 ^ len - 1)
            init2 = Not (init1 << pos)

            val = val And init1
            val <<= pos

            d = d And init2
            d = d Or val
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function SetBitsInRange16(ByRef d As UInt16, ByVal val As UInt16, ByVal pos As Integer, ByVal len As Integer)
        Try
            Dim init1 = 0, init2 = 0

            Dim Boundry = 16
            If len = 0 Or len > Boundry Then
                Return False
            End If

            pos = pos Mod 16

            init1 = Convert.ToUInt16(2 ^ len - 1)
            init2 = Not (init1 << pos)

            val = val And init1
            val <<= pos

            d = d And init2
            d = d Or val
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function SetBitsInRange32(ByRef d As UInt32, ByVal val As Integer, ByVal pos As Integer, ByVal len As Integer)
        Try
            Dim init1 = 0, init2 = 0
            Dim Boundry = 32

            If len = 0 Or len > Boundry Then
                Return False
            End If

            pos = pos Mod 16

            init1 = Convert.ToUInt32(2 ^ len - 1)
            init2 = Not (init1 << pos)
            val = val And init1
            val <<= pos
            d = d And init2
            d = d Or val
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Class
