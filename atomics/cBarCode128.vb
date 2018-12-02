Option Explicit On
Option Strict On

'''<summary>Class for Barcode 128b encoding.</summary>
'''<remarks>See https://en.wikipedia.org/wiki/Code_128 for details.</remarks>
Public Class cBarCode128

    Private Shared ReadOnly BitPat As String() = {"11011001100", "11001101100", "11001100110", "10010011000", "10010001100"}

    Public Shared Sub CharToBitPattern(ByVal Character As String, ByRef Binary As String, ByRef CharVal As Integer)

        Select Case Character
            Case = " "
                CharVal = 0 : Binary &= BitPat(CharVal)
            Case = "!"
                CharVal = 1 : Binary &= BitPat(CharVal)
            Case = """"
                CharVal = 2 : Binary &= BitPat(CharVal)
            Case = "#"
                CharVal = 3 : Binary &= BitPat(CharVal)
            Case = "$"
                CharVal = 4 : Binary &= BitPat(CharVal)
            Case "%"
                Binary &= "10001001100" : CharVal = 5
            Case "&"
                Binary &= "10011001000" : CharVal = 6
            Case "'"
                Binary &= "10011000100" : CharVal = 7
            Case "("
                Binary &= "10001100100" : CharVal = 8
            Case ")"
                Binary &= "11001001000" : CharVal = 9
            Case "*"
                Binary &= "11001000100" : CharVal = 10
            Case "+"
                Binary &= "11000100100" : CharVal = 11
            Case ","
                Binary &= "10110011100" : CharVal = 12
            Case "-"
                Binary &= "10011011100" : CharVal = 13
            Case "."
                Binary &= "10011001110" : CharVal = 14
            Case "/"
                Binary &= "10111001100" : CharVal = 15
            Case "0"
                Binary &= "10011101100" : CharVal = 16
            Case "1"
                Binary &= "10011100110" : CharVal = 17
            Case "2"
                Binary &= "11001110010" : CharVal = 18
            Case "3"
                Binary &= "11001011100" : CharVal = 19
            Case "4"
                Binary &= "11001001110" : CharVal = 20
            Case "5"
                Binary &= "11011100100" : CharVal = 21
            Case "6"
                Binary &= "11001110100" : CharVal = 22
            Case "7"
                Binary &= "11101101110" : CharVal = 23
            Case "8"
                Binary &= "11101001100"
                CharVal = 24
            Case "9"
                Binary &= "11100101100"
                CharVal = 25
            Case ":"
                Binary &= "11100100110"
                CharVal = 26
            Case " ;"
                Binary &= "11101100100"
                CharVal = 27
            Case "<"
                Binary &= "11100110100"
                CharVal = 28
            Case "="
                Binary &= "11100110010"
                CharVal = 29
            Case ">"
                Binary &= "11011011000"
                CharVal = 30
            Case " ?"
                Binary &= "11011000110"
                CharVal = 31
            Case "@"
                Binary &= "11000110110"
                CharVal = 32
            Case "A"
                Binary &= "10100011000"
                CharVal = 33
            Case "B"
                Binary &= "10001011000"
                CharVal = 34
            Case "C"
                Binary &= "10001000110"
                CharVal = 35
            Case "D"
                Binary &= "10110001000"
                CharVal = 36
            Case "E"
                Binary &= "10001101000"
                CharVal = 37
            Case "F"
                Binary &= "10001100010"
                CharVal = 38
            Case "G"
                Binary &= "11010001000"
                CharVal = 39
            Case "H"
                Binary &= "11000101000"
                CharVal = 40
            Case "I"
                Binary &= "11000100010"
                CharVal = 41
            Case "J"
                Binary &= "10110111000"
                CharVal = 42
            Case "K"
                Binary &= "10110001110"
                CharVal = 43
            Case "L"
                Binary &= "10001101110"
                CharVal = 44
            Case "M"
                Binary &= "10111011000"
                CharVal = 45
            Case "N"
                Binary &= "10111000110"
                CharVal = 46
            Case "O"
                Binary &= "10001110110"
                CharVal = 47
            Case "P"
                Binary &= "11101110110"
                CharVal = 48
            Case "Q"
                Binary &= "11010001110"
                CharVal = 49
            Case "R"
                Binary &= "11000101110"
                CharVal = 50
            Case "S"
                Binary &= "11011101000"
                CharVal = 51
            Case "T"
                Binary &= "11011100010"
                CharVal = 52
            Case "U"
                Binary &= "11011101110"
                CharVal = 53
            Case "V"
                Binary &= "11101011000"
                CharVal = 54
            Case "W"
                Binary &= "11101000110"
                CharVal = 55
            Case "X"
                Binary &= "11100010110"
                CharVal = 56
            Case "Y"
                Binary &= "11101101000"
                CharVal = 57
            Case "Z"
                Binary &= "11101100010"
                CharVal = 58
            Case "["
                Binary &= "11100011010"
                CharVal = 59
            Case "\"
                Binary &= "11101111010"
                CharVal = 60
            Case "]"
                Binary &= "11001000010"
                CharVal = 61
            Case "^"
                Binary &= "11110001010"
                CharVal = 62
            Case "_"
                Binary &= "10100110000"
                CharVal = 63
            Case "`"
                Binary &= "10100001100"
                CharVal = 64
            Case "a"
                Binary &= "10010110000"
                CharVal = 65
            Case "b"
                Binary &= "10010000110"
                CharVal = 66
            Case "c"
                Binary &= "10000101100"
                CharVal = 67
            Case "d"
                Binary &= "10000100110"
                CharVal = 68
            Case "e"
                Binary &= "10110010000"
                CharVal = 69
            Case "f"
                Binary &= "10110000100"
                CharVal = 70
            Case "g"
                Binary &= "10011010000"
                CharVal = 71
            Case "h"
                Binary &= "10011000010"
                CharVal = 72
            Case "i"
                Binary &= "10000110100"
                CharVal = 73
            Case "j"
                Binary &= "10000110010"
                CharVal = 74
            Case "k"
                Binary &= "11000010010"
                CharVal = 75
            Case "l"
                Binary &= "11001010000"
                CharVal = 76
            Case "m"
                Binary &= "11110111010"
                CharVal = 77
            Case "n"
                Binary &= "11000010100"
                CharVal = 78
            Case "o"
                Binary &= "10001111010"
                CharVal = 79
            Case "p"
                Binary &= "10100111100"
                CharVal = 80
            Case "q"
                Binary &= "10010111100"
                CharVal = 81
            Case "r"
                Binary &= "10010011110"
                CharVal = 82
            Case "s"
                Binary &= "10111100100"
                CharVal = 83
            Case "t"
                Binary &= "10011110100"
                CharVal = 84
            Case "u"
                Binary &= "10011110010"
                CharVal = 85
            Case "v"
                Binary &= "11110100100"
                CharVal = 86
            Case "w"
                Binary &= "11110010100"
                CharVal = 87
            Case "x"
                Binary &= "11110010010"
                CharVal = 88
            Case "y"
                Binary &= "11011011110"
                CharVal = 89
            Case "z"
                Binary &= "11011110110"
                CharVal = 90
            Case "{"
                Binary &= "11110110110"
                CharVal = 91
            Case "|"
                Binary &= "10101111000"
                CharVal = 92
            Case "}"
                Binary &= "10100011110"
                CharVal = 93
            Case "~"
                Binary &= "10001011110"
                CharVal = 94
        End Select

    End Sub

    Public Shared Function GenerateCheckSum(ByVal Dig As Integer) As String

        Select Case Dig
            Case 0, 1, 2
                Return BitPat(Dig)
            Case 3
                Return "10010011000"
            Case 4
                Return "10010001100"
            Case 5
                Return "10001001100"
            Case 6
                Return "10011001000"
            Case 7
                Return "10011000100"
            Case 8
                Return "10001100100"
            Case 9
                Return "11001001000"
            Case 10
                Return "11001000100"
            Case 11
                Return "11000100100"
            Case 12
                Return "10110011100"
            Case 13
                Return "10011011100"
            Case 14
                Return "10011001110"
            Case 15
                Return "10111001100"
            Case 16
                Return "10011101100"
            Case 17
                Return "10011100110"
            Case 18
                Return "11001110010"
            Case 19
                Return "11001011100"
            Case 20
                Return "11001001110"
            Case 21
                Return "11011100100"
            Case 22
                Return "11001110100"
            Case 23
                Return "11101101110"
            Case 24
                Return "11101001100"
            Case 25
                Return "11100101100"
            Case 26
                Return "11100100110"
            Case 27
                Return "11101100100"
            Case 28
                Return "11100110100"
            Case 29
                Return "11100110010"
            Case 30
                Return "11011011000"
            Case 31
                Return "11011000110"
            Case 32
                Return "11000110110"
            Case 33
                Return "10100011000"
            Case 34
                Return "10001011000"
            Case 35
                Return "10001000110"
            Case 36
                Return "10110001000"
            Case 37
                Return "10001101000"
            Case 38
                Return "10001100010"
            Case 39
                Return "11010001000"
            Case 40
                Return "11000101000"
            Case 41
                Return "11000100010"
            Case 42
                Return "10110111000"
            Case 43
                Return "10110001110"
            Case 44
                Return "10001101110"
            Case 45
                Return "10111011000"
            Case 46
                Return "10111000110"
            Case 47
                Return "10001110110"
            Case 48
                Return "11101110110"
            Case 49
                Return "11010001110"
            Case 50
                Return "11000101110"
            Case 51
                Return "11011101000"
            Case 52
                Return "11011100010"
            Case 53
                Return "11011101110"
            Case 54
                Return "11101011000"
            Case 55
                Return "11101000110"
            Case 56
                Return "11100010110"
            Case 57
                Return "11101101000"
            Case 58
                Return "11101100010"
            Case 59
                Return "11100011010"
            Case 60
                Return "11101111010"
            Case 61
                Return "11001000010"
            Case 62
                Return "11110001010"
            Case 63
                Return "10100110000"
            Case 64
                Return "10100001100"
            Case 65
                Return "10010110000"
            Case 66
                Return "10010000110"
            Case 67
                Return "10000101100"
            Case 68
                Return "10000100110"
            Case 69
                Return "10110010000"
            Case 70
                Return "10110000100"
            Case 71
                Return "10011010000"
            Case 72
                Return "10011000010"
            Case 73
                Return "10000110100"
            Case 74
                Return "10000110010"
            Case 75
                Return "11000010010"
            Case 76
                Return "11001010000"
            Case 77
                Return "11110111010"
            Case 78
                Return "11000010100"
            Case 79
                Return "10001111010"
            Case 80
                Return "10100111100"
            Case 81
                Return "10010111100"
            Case 82
                Return "10010011110"
            Case 83
                Return "10111100100"
            Case 84
                Return "10011110100"
            Case 85
                Return "10011110010"
            Case 86
                Return "11110100100"
            Case 87
                Return "11110010100"
            Case 88
                Return "11110010010"
            Case 89
                Return "11011011110"
            Case 90
                Return "11011110110"
            Case 91
                Return "11110110110"
            Case 92
                Return "10101111000"
            Case 93
                Return "10100011110"
            Case 94
                Return "10001011110"

            Case Else
                Return "00000000000"

        End Select

    End Function

End Class