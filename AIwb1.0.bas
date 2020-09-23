Attribute VB_Name = "AI"


Public ChessBoard(1 To 8, 1 To 8) As Byte
Public WhiteMobility, BlackMobility As Integer
Public MoveToPlay As String

Private Declare Function GetInputState Lib "user32" () As Long
Sub EngineGo() ' in this version it is not used
Call Alpha_Beta_Prune(1, 0, MINVALUE, MAXVALUE)
End Sub

'this  is initialization function for all needed data
Function Initialize_Board()

For i = 0 To 63
WCenter_Points(i) = 0
BCenter_Points(i) = 0

Next i

WCenter_Points(33) = 5
WCenter_Points(35) = 10
WCenter_Points(36) = 9
WCenter_Points(37) = 6
BCenter_Points(26) = 5
BCenter_Points(27) = 10
BCenter_Points(28) = 9
BCenter_Points(29) = 6

sign1 = 1
 RWRockMoved = False
 LWRockMoved = False
 RBRockMoved = False
 LBRockMoved = False
 WKingMoved = False
 BKingMoved = False
 WKO_O = False
 WKO_O_O = False
 BKO_O = False
 BKO_O_O = False

For i = 0 To 20
MB(i) = -1
MB(i + 99) = -1
Next i
MB(30) = -1
MB(40) = -1
MB(50) = -1
MB(60) = -1
MB(70) = -1
MB(80) = -1
MB(90) = -1

MB(39) = -1
MB(49) = -1
MB(59) = -1
MB(69) = -1
MB(79) = -1
MB(89) = -1
MB(99) = -1


MB(21) = 0
MB(22) = 1
MB(23) = 2
MB(24) = 3
MB(25) = 4
MB(26) = 5
MB(27) = 6
MB(28) = 7
MB(31) = 8
MB(32) = 9
MB(33) = 10
MB(34) = 11
MB(35) = 12
MB(36) = 13
MB(37) = 14
MB(38) = 15
MB(41) = 16
MB(42) = 17
MB(43) = 18
MB(44) = 19
MB(45) = 20
MB(46) = 21
MB(47) = 22
MB(48) = 23
MB(51) = 24
MB(52) = 25
MB(53) = 26
MB(54) = 27
MB(55) = 28
MB(56) = 29
MB(57) = 30
MB(58) = 31
MB(61) = 32
MB(62) = 33
MB(63) = 34
MB(64) = 35
MB(65) = 36
MB(66) = 37
MB(67) = 38
MB(68) = 39
MB(71) = 40
MB(72) = 41
MB(73) = 42
MB(74) = 43
MB(75) = 44
MB(76) = 45
MB(77) = 46
MB(78) = 47
MB(81) = 48
MB(82) = 49
MB(83) = 50
MB(84) = 51
MB(85) = 52
MB(86) = 53
MB(87) = 54
MB(88) = 55
MB(91) = 56
MB(92) = 57
MB(93) = 58
MB(94) = 59
MB(95) = 60
MB(96) = 61
MB(97) = 62
MB(98) = 63

LookUp(0) = 21
LookUp(1) = 22
LookUp(2) = 23
LookUp(3) = 24
LookUp(4) = 25
LookUp(5) = 26
LookUp(6) = 27
LookUp(7) = 28
LookUp(8) = 31
LookUp(9) = 32
LookUp(10) = 33
LookUp(11) = 34
LookUp(12) = 35
LookUp(13) = 36
LookUp(14) = 37
LookUp(15) = 38
LookUp(16) = 41
LookUp(17) = 42
LookUp(18) = 43
LookUp(19) = 44
LookUp(20) = 45
LookUp(21) = 46
LookUp(22) = 47
LookUp(23) = 48
LookUp(24) = 51
LookUp(25) = 52
LookUp(26) = 53
LookUp(27) = 54
LookUp(28) = 55
LookUp(29) = 56
LookUp(30) = 57
LookUp(31) = 58
LookUp(32) = 61
LookUp(33) = 62
LookUp(34) = 63
LookUp(35) = 64
LookUp(36) = 65
LookUp(37) = 66
LookUp(38) = 67
LookUp(39) = 68
LookUp(40) = 71
LookUp(41) = 72
LookUp(42) = 73
LookUp(43) = 74
LookUp(44) = 75
LookUp(45) = 76
LookUp(46) = 77
LookUp(47) = 78
LookUp(48) = 81
LookUp(49) = 82
LookUp(50) = 83
LookUp(51) = 84
LookUp(52) = 85
LookUp(53) = 86
LookUp(54) = 87
LookUp(55) = 88
LookUp(56) = 91
LookUp(57) = 92
LookUp(58) = 93
LookUp(59) = 94
LookUp(60) = 95
LookUp(61) = 96
LookUp(62) = 97
LookUp(63) = 98



MT(0, 0) = -11
MT(0, 1) = -10
MT(0, 2) = -9
MT(0, 3) = 0
MT(0, 4) = 0
MT(0, 5) = 0
MT(0, 6) = 0
MT(0, 7) = 0
MT(1, 0) = 9
MT(1, 1) = 10
MT(1, 2) = 11
MT(1, 3) = 0
MT(1, 4) = 0
MT(1, 5) = 0
MT(1, 6) = 0
MT(1, 7) = 0
MT(2, 0) = -21
MT(2, 1) = -19
MT(2, 2) = -12
MT(2, 3) = -8
MT(2, 4) = 8
MT(2, 5) = 12
MT(2, 6) = 19
MT(2, 7) = 21
MT(3, 0) = -11
MT(3, 1) = -9
MT(3, 2) = 9
MT(3, 3) = 0
MT(3, 4) = 0
MT(3, 5) = 0
MT(3, 6) = 0
MT(3, 7) = 0
MT(4, 0) = -10
MT(4, 1) = -1
MT(4, 2) = -1
MT(4, 3) = 10
MT(4, 4) = 0
MT(4, 5) = 0
MT(4, 6) = 0
MT(4, 7) = 0
MT(5, 0) = -11
MT(5, 1) = -10
MT(5, 2) = -9
MT(5, 3) = -1
MT(5, 4) = 1
MT(5, 5) = 9
MT(5, 6) = 10
MT(5, 7) = 11
MT(6, 0) = -11
MT(6, 1) = -10
MT(6, 2) = -9
MT(6, 3) = -1
MT(6, 4) = 1
MT(6, 5) = 9
MT(6, 6) = 10
MT(6, 7) = 11
RankLookUp(0, 0) = 0
RankLookUp(0, 1) = 7
RankLookUp(1, 0) = 0
RankLookUp(1, 1) = 7
RankLookUp(2, 0) = 0
RankLookUp(2, 1) = 7
RankLookUp(3, 0) = 0
RankLookUp(3, 1) = 7
RankLookUp(4, 0) = 0
RankLookUp(4, 1) = 7
RankLookUp(5, 0) = 0
RankLookUp(5, 1) = 7
RankLookUp(6, 0) = 0
RankLookUp(6, 1) = 7
RankLookUp(7, 0) = 0
RankLookUp(7, 1) = 7
RankLookUp(8, 0) = 8
RankLookUp(8, 1) = 15
RankLookUp(9, 0) = 8
RankLookUp(9, 1) = 15
RankLookUp(10, 0) = 8
RankLookUp(10, 1) = 15
RankLookUp(11, 0) = 8
RankLookUp(11, 1) = 15
RankLookUp(12, 0) = 8
RankLookUp(12, 1) = 15
RankLookUp(13, 0) = 8
RankLookUp(13, 1) = 15
RankLookUp(14, 0) = 8
RankLookUp(14, 1) = 15
RankLookUp(15, 0) = 8
RankLookUp(15, 1) = 15
RankLookUp(16, 0) = 16
RankLookUp(16, 1) = 23
RankLookUp(17, 0) = 16
RankLookUp(17, 1) = 23
RankLookUp(18, 0) = 16
RankLookUp(18, 1) = 23
RankLookUp(19, 0) = 16
RankLookUp(19, 1) = 23
RankLookUp(20, 0) = 16
RankLookUp(20, 1) = 23
RankLookUp(21, 0) = 16
RankLookUp(21, 1) = 23
RankLookUp(22, 0) = 16
RankLookUp(22, 1) = 23
RankLookUp(23, 0) = 16
RankLookUp(23, 1) = 23
RankLookUp(24, 0) = 24
RankLookUp(24, 1) = 31
RankLookUp(25, 0) = 24
RankLookUp(25, 1) = 31
RankLookUp(26, 0) = 24
RankLookUp(26, 1) = 31
RankLookUp(27, 0) = 24
RankLookUp(27, 1) = 31
RankLookUp(28, 0) = 24
RankLookUp(28, 1) = 31
RankLookUp(29, 0) = 24
RankLookUp(29, 1) = 31
RankLookUp(30, 0) = 24
RankLookUp(30, 1) = 31
RankLookUp(31, 0) = 24
RankLookUp(31, 1) = 31
RankLookUp(32, 0) = 32
RankLookUp(32, 1) = 39
RankLookUp(33, 0) = 32
RankLookUp(33, 1) = 39
RankLookUp(34, 0) = 32
RankLookUp(34, 1) = 39
RankLookUp(35, 0) = 32
RankLookUp(35, 1) = 39
RankLookUp(36, 0) = 32
RankLookUp(36, 1) = 39
RankLookUp(37, 0) = 32
RankLookUp(37, 1) = 39
RankLookUp(38, 0) = 32
RankLookUp(38, 1) = 39
RankLookUp(39, 0) = 32
RankLookUp(39, 1) = 39
RankLookUp(40, 0) = 40
RankLookUp(40, 1) = 47
RankLookUp(41, 0) = 40
RankLookUp(41, 1) = 47
RankLookUp(42, 0) = 40
RankLookUp(42, 1) = 47
RankLookUp(43, 0) = 40
RankLookUp(43, 1) = 47
RankLookUp(44, 0) = 40
RankLookUp(44, 1) = 47
RankLookUp(45, 0) = 40
RankLookUp(45, 1) = 47
RankLookUp(46, 0) = 40
RankLookUp(46, 1) = 47
RankLookUp(47, 0) = 40
RankLookUp(47, 1) = 47
RankLookUp(48, 0) = 48
RankLookUp(48, 1) = 55
RankLookUp(49, 0) = 48
RankLookUp(49, 1) = 55
RankLookUp(50, 0) = 48
RankLookUp(50, 1) = 55
RankLookUp(51, 0) = 48
RankLookUp(51, 1) = 55
RankLookUp(52, 0) = 48
RankLookUp(52, 1) = 55
RankLookUp(53, 0) = 48
RankLookUp(53, 1) = 55
RankLookUp(54, 0) = 48
RankLookUp(54, 1) = 55
RankLookUp(55, 0) = 48
RankLookUp(55, 1) = 55
RankLookUp(56, 0) = 56
RankLookUp(56, 1) = 63
RankLookUp(57, 0) = 56
RankLookUp(57, 1) = 63
RankLookUp(58, 0) = 56
RankLookUp(58, 1) = 63
RankLookUp(59, 0) = 56
RankLookUp(59, 1) = 63
RankLookUp(60, 0) = 56
RankLookUp(60, 1) = 63
RankLookUp(61, 0) = 56
RankLookUp(61, 1) = 63
RankLookUp(62, 0) = 56
RankLookUp(62, 1) = 63
RankLookUp(63, 0) = 56
RankLookUp(63, 1) = 63



FileLookUp(0, 0) = 0
FileLookUp(0, 1) = 56
FileLookUp(8, 0) = 0
FileLookUp(8, 1) = 56
FileLookUp(16, 0) = 0
FileLookUp(16, 1) = 56
FileLookUp(24, 0) = 0
FileLookUp(24, 1) = 56
FileLookUp(32, 0) = 0
FileLookUp(32, 1) = 56
FileLookUp(40, 0) = 0
FileLookUp(40, 1) = 56
FileLookUp(48, 0) = 0
FileLookUp(48, 1) = 56
FileLookUp(56, 0) = 0
FileLookUp(56, 1) = 56



FileLookUp(1, 0) = 1
FileLookUp(1, 1) = 57
FileLookUp(2, 0) = 2
FileLookUp(2, 1) = 58
FileLookUp(3, 0) = 3
FileLookUp(3, 1) = 59
FileLookUp(4, 0) = 4
FileLookUp(4, 1) = 60
FileLookUp(5, 0) = 5
FileLookUp(5, 1) = 61
FileLookUp(6, 0) = 6
FileLookUp(6, 1) = 62
FileLookUp(7, 0) = 7
FileLookUp(7, 1) = 63

FileLookUp(9, 0) = 1
FileLookUp(9, 1) = 57
FileLookUp(10, 0) = 2
FileLookUp(10, 1) = 58
FileLookUp(11, 0) = 3
FileLookUp(11, 1) = 59
FileLookUp(12, 0) = 4
FileLookUp(12, 1) = 60
FileLookUp(13, 0) = 5
FileLookUp(13, 1) = 61
FileLookUp(14, 0) = 6
FileLookUp(14, 1) = 62
FileLookUp(15, 0) = 7
FileLookUp(15, 1) = 63

FileLookUp(17, 0) = 1
FileLookUp(17, 1) = 57
FileLookUp(18, 0) = 2
FileLookUp(18, 1) = 58
FileLookUp(19, 0) = 3
FileLookUp(19, 1) = 59
FileLookUp(20, 0) = 4
FileLookUp(20, 1) = 60
FileLookUp(21, 0) = 5
FileLookUp(21, 1) = 61
FileLookUp(22, 0) = 6
FileLookUp(22, 1) = 62
FileLookUp(23, 0) = 7
FileLookUp(23, 1) = 63

FileLookUp(25, 0) = 1
FileLookUp(25, 1) = 57
FileLookUp(26, 0) = 2
FileLookUp(26, 1) = 58
FileLookUp(27, 0) = 3
FileLookUp(27, 1) = 59
FileLookUp(28, 0) = 4
FileLookUp(28, 1) = 60
FileLookUp(29, 0) = 5
FileLookUp(29, 1) = 61
FileLookUp(30, 0) = 6
FileLookUp(30, 1) = 62
FileLookUp(31, 0) = 7
FileLookUp(31, 1) = 63

FileLookUp(33, 0) = 1
FileLookUp(33, 1) = 57
FileLookUp(34, 0) = 2
FileLookUp(34, 1) = 58
FileLookUp(35, 0) = 3
FileLookUp(35, 1) = 59
FileLookUp(36, 0) = 4
FileLookUp(36, 1) = 60
FileLookUp(37, 0) = 5
FileLookUp(37, 1) = 61
FileLookUp(38, 0) = 6
FileLookUp(38, 1) = 62
FileLookUp(39, 0) = 7
FileLookUp(39, 1) = 63

FileLookUp(41, 0) = 1
FileLookUp(41, 1) = 57
FileLookUp(42, 0) = 2
FileLookUp(42, 1) = 58
FileLookUp(43, 0) = 3
FileLookUp(43, 1) = 59
FileLookUp(44, 0) = 4
FileLookUp(44, 1) = 60
FileLookUp(45, 0) = 5
FileLookUp(45, 1) = 61
FileLookUp(46, 0) = 6
FileLookUp(46, 1) = 62
FileLookUp(47, 0) = 7
FileLookUp(47, 1) = 63

FileLookUp(49, 0) = 1
FileLookUp(49, 1) = 57
FileLookUp(50, 0) = 2
FileLookUp(50, 1) = 58
FileLookUp(51, 0) = 3
FileLookUp(51, 1) = 59
FileLookUp(52, 0) = 4
FileLookUp(52, 1) = 60
FileLookUp(53, 0) = 5
FileLookUp(53, 1) = 61
FileLookUp(54, 0) = 6
FileLookUp(54, 1) = 61
FileLookUp(55, 0) = 7
FileLookUp(55, 1) = 63

FileLookUp(57, 0) = 1
FileLookUp(57, 1) = 57
FileLookUp(58, 0) = 2
FileLookUp(58, 1) = 58
FileLookUp(59, 0) = 3
FileLookUp(59, 1) = 59
FileLookUp(60, 0) = 4
FileLookUp(60, 1) = 60
FileLookUp(61, 0) = 5
FileLookUp(61, 1) = 61
FileLookUp(62, 0) = 6
FileLookUp(62, 1) = 62
FileLookUp(63, 0) = 7
FileLookUp(63, 1) = 63

'Call initiate_Masks
For i = 0 To 63
WhitePawn_enPassent(i) = 0
Blackpawn_enPassent(i) = 0
 Whitepawn_Position(i) = 0
 WhiteKing_Position(i) = 0
 WhiteQueen_Position(i) = 0
 WhiteKnight_Position(i) = 0
 WhiteBishop_Position(i) = 0
 WhiteRock_Position(i) = 0
 WhitePieces_Position(i) = 0
 Blackpawn_Position(i) = 0
 BlackKing_Position(i) = 0
 BlackQueen_Position(i) = 0
 BlackKnight_Position(i) = 0
 BlackBishop_Position(i) = 0
 BlackRock_Position(i) = 0
 BlackPieces_Position(i) = 0

Next i


For i = 8 To 15
Whitepawn_Position(i + 40) = 1
Blackpawn_Position(i) = 1
Next i

WhiteKing_Position(60) = 1
BlackKing_Position(4) = 1
BlackRock_Position(0) = 1
BlackRock_Position(7) = 1
WhiteRock_Position(56) = 1
WhiteRock_Position(63) = 1
BlackQueen_Position(3) = 1
WhiteQueen_Position(59) = 1
WhiteKnight_Position(62) = 1
WhiteBishop_Position(61) = 1
WhiteKnight_Position(57) = 1
WhiteBishop_Position(58) = 1
BlackKnight_Position(1) = 1
BlackBishop_Position(2) = 1
BlackKnight_Position(6) = 1
BlackBishop_Position(5) = 1



'a8h1

Diagonal1LookUp(0, 0) = 0
Diagonal1LookUp(9, 0) = 0
Diagonal1LookUp(18, 0) = 0
Diagonal1LookUp(27, 0) = 0
Diagonal1LookUp(36, 0) = 0
Diagonal1LookUp(45, 0) = 0
Diagonal1LookUp(54, 0) = 0
Diagonal1LookUp(63, 0) = 0
Diagonal1LookUp(0, 1) = 63
Diagonal1LookUp(9, 1) = 63
Diagonal1LookUp(18, 1) = 63
Diagonal1LookUp(27, 1) = 63
Diagonal1LookUp(36, 1) = 63
Diagonal1LookUp(45, 1) = 63
Diagonal1LookUp(54, 1) = 63
Diagonal1LookUp(63, 1) = 63

'a7g1
Diagonal1LookUp(8, 0) = 8
Diagonal1LookUp(17, 0) = 8
Diagonal1LookUp(26, 0) = 8
Diagonal1LookUp(35, 0) = 8
Diagonal1LookUp(44, 0) = 8
Diagonal1LookUp(53, 0) = 8
Diagonal1LookUp(62, 0) = 8
Diagonal1LookUp(8, 1) = 62
Diagonal1LookUp(17, 1) = 62
Diagonal1LookUp(26, 1) = 62
Diagonal1LookUp(35, 1) = 62
Diagonal1LookUp(44, 1) = 62
Diagonal1LookUp(53, 1) = 62
Diagonal1LookUp(62, 1) = 62

'a6f1
Diagonal1LookUp(16, 0) = 16
Diagonal1LookUp(25, 0) = 16
Diagonal1LookUp(34, 0) = 16
Diagonal1LookUp(43, 0) = 16
Diagonal1LookUp(52, 0) = 16
Diagonal1LookUp(61, 0) = 16

Diagonal1LookUp(16, 1) = 61
Diagonal1LookUp(25, 1) = 61
Diagonal1LookUp(34, 1) = 61
Diagonal1LookUp(43, 1) = 61
Diagonal1LookUp(52, 1) = 61
Diagonal1LookUp(61, 1) = 61

'a5e1
Diagonal1LookUp(24, 0) = 24
Diagonal1LookUp(33, 0) = 24
Diagonal1LookUp(42, 0) = 24
Diagonal1LookUp(51, 0) = 24
Diagonal1LookUp(60, 0) = 24

Diagonal1LookUp(24, 1) = 60
Diagonal1LookUp(33, 1) = 60
Diagonal1LookUp(42, 1) = 60
Diagonal1LookUp(51, 1) = 60
Diagonal1LookUp(60, 1) = 60

'a4d1
Diagonal1LookUp(32, 0) = 32
Diagonal1LookUp(41, 0) = 32
Diagonal1LookUp(50, 0) = 32
Diagonal1LookUp(59, 0) = 32

Diagonal1LookUp(32, 1) = 59
Diagonal1LookUp(41, 1) = 59
Diagonal1LookUp(50, 1) = 59
Diagonal1LookUp(59, 1) = 59


'a3c1
Diagonal1LookUp(40, 0) = 40
Diagonal1LookUp(49, 0) = 40
Diagonal1LookUp(58, 0) = 40

Diagonal1LookUp(40, 1) = 58
Diagonal1LookUp(49, 1) = 58
Diagonal1LookUp(58, 1) = 58

'a2b1
Diagonal1LookUp(48, 0) = 48
Diagonal1LookUp(57, 0) = 48

Diagonal1LookUp(48, 1) = 57
Diagonal1LookUp(57, 1) = 57
'a1a1
Diagonal1LookUp(56, 0) = 56
Diagonal1LookUp(56, 1) = 56


'b8h2
Diagonal1LookUp(1, 0) = 1
Diagonal1LookUp(10, 0) = 1
Diagonal1LookUp(19, 0) = 1
Diagonal1LookUp(28, 0) = 1
Diagonal1LookUp(37, 0) = 1
Diagonal1LookUp(46, 0) = 1
Diagonal1LookUp(55, 0) = 1

Diagonal1LookUp(1, 1) = 55
Diagonal1LookUp(10, 1) = 55
Diagonal1LookUp(19, 1) = 55
Diagonal1LookUp(28, 1) = 55
Diagonal1LookUp(37, 1) = 55
Diagonal1LookUp(46, 1) = 55
Diagonal1LookUp(55, 1) = 55

'c8h3
Diagonal1LookUp(2, 0) = 2
Diagonal1LookUp(11, 0) = 2
Diagonal1LookUp(20, 0) = 2
Diagonal1LookUp(29, 0) = 2
Diagonal1LookUp(38, 0) = 2
Diagonal1LookUp(47, 0) = 2

Diagonal1LookUp(2, 1) = 47
Diagonal1LookUp(11, 1) = 47
Diagonal1LookUp(20, 1) = 47
Diagonal1LookUp(29, 1) = 47
Diagonal1LookUp(38, 1) = 47
Diagonal1LookUp(47, 1) = 47

'd8h4
Diagonal1LookUp(3, 0) = 3
Diagonal1LookUp(12, 0) = 3
Diagonal1LookUp(21, 0) = 3
Diagonal1LookUp(30, 0) = 3
Diagonal1LookUp(39, 0) = 3

Diagonal1LookUp(3, 1) = 39
Diagonal1LookUp(12, 1) = 39
Diagonal1LookUp(21, 1) = 39
Diagonal1LookUp(30, 1) = 39
Diagonal1LookUp(39, 1) = 39

'e8h5
Diagonal1LookUp(4, 0) = 4
Diagonal1LookUp(13, 0) = 4
Diagonal1LookUp(22, 0) = 4
Diagonal1LookUp(31, 0) = 4

Diagonal1LookUp(4, 1) = 31
Diagonal1LookUp(13, 1) = 31
Diagonal1LookUp(22, 1) = 31
Diagonal1LookUp(31, 1) = 31

'f8h6
Diagonal1LookUp(5, 0) = 5
Diagonal1LookUp(14, 0) = 5
Diagonal1LookUp(23, 0) = 5

Diagonal1LookUp(5, 1) = 23
Diagonal1LookUp(14, 1) = 23
Diagonal1LookUp(23, 1) = 23

'g8h7
Diagonal1LookUp(6, 0) = 6
Diagonal1LookUp(15, 0) = 6

Diagonal1LookUp(6, 1) = 15
Diagonal1LookUp(15, 1) = 15

'h8h8
Diagonal1LookUp(7, 0) = 7
Diagonal1LookUp(7, 1) = 7



'a1h8

Diagonal2LookUp(7, 0) = 7
Diagonal2LookUp(14, 0) = 7
Diagonal2LookUp(21, 0) = 7
Diagonal2LookUp(28, 0) = 7
Diagonal2LookUp(35, 0) = 7
Diagonal2LookUp(42, 0) = 7
Diagonal2LookUp(49, 0) = 7
Diagonal2LookUp(56, 0) = 7

Diagonal2LookUp(7, 1) = 56
Diagonal2LookUp(14, 1) = 56
Diagonal2LookUp(21, 1) = 56
Diagonal2LookUp(28, 1) = 56
Diagonal2LookUp(35, 1) = 56
Diagonal2LookUp(42, 1) = 56
Diagonal2LookUp(49, 1) = 56
Diagonal2LookUp(56, 1) = 56

'g8a2
Diagonal2LookUp(6, 0) = 6
Diagonal2LookUp(13, 0) = 6
Diagonal2LookUp(20, 0) = 6
Diagonal2LookUp(27, 0) = 6
Diagonal2LookUp(34, 0) = 6
Diagonal2LookUp(41, 0) = 6
Diagonal2LookUp(48, 0) = 6

Diagonal2LookUp(6, 1) = 48
Diagonal2LookUp(13, 1) = 48
Diagonal2LookUp(20, 1) = 48
Diagonal2LookUp(27, 1) = 48
Diagonal2LookUp(41, 1) = 48
Diagonal2LookUp(48, 1) = 48

'f8a3
Diagonal2LookUp(5, 0) = 5
Diagonal2LookUp(12, 0) = 5
Diagonal2LookUp(19, 0) = 5
Diagonal2LookUp(26, 0) = 5
Diagonal2LookUp(33, 0) = 5
Diagonal2LookUp(40, 0) = 5

Diagonal2LookUp(5, 1) = 40
Diagonal2LookUp(12, 1) = 40
Diagonal2LookUp(19, 1) = 40
Diagonal2LookUp(26, 1) = 40
Diagonal2LookUp(33, 1) = 40
Diagonal2LookUp(40, 1) = 40

'e8a4
Diagonal2LookUp(4, 0) = 4
Diagonal2LookUp(11, 0) = 4
Diagonal2LookUp(18, 0) = 4
Diagonal2LookUp(25, 0) = 4
Diagonal2LookUp(32, 0) = 4

Diagonal2LookUp(4, 1) = 32
Diagonal2LookUp(11, 1) = 32
Diagonal2LookUp(18, 1) = 32
Diagonal2LookUp(25, 1) = 32
Diagonal2LookUp(32, 1) = 32

'd8a5
Diagonal2LookUp(3, 0) = 3
Diagonal2LookUp(10, 0) = 3
Diagonal2LookUp(17, 0) = 3
Diagonal2LookUp(24, 0) = 3

Diagonal2LookUp(3, 1) = 24
Diagonal2LookUp(10, 1) = 24
Diagonal2LookUp(17, 1) = 24
Diagonal2LookUp(24, 1) = 24

''c8a6
Diagonal2LookUp(2, 0) = 2
Diagonal2LookUp(9, 0) = 2
Diagonal2LookUp(16, 0) = 2

Diagonal2LookUp(2, 1) = 16
Diagonal2LookUp(9, 1) = 16
Diagonal2LookUp(16, 1) = 16

'b8a7
Diagonal2LookUp(1, 0) = 1
Diagonal2LookUp(8, 0) = 1

Diagonal2LookUp(1, 1) = 8
Diagonal2LookUp(8, 1) = 8

'a8a8
Diagonal2LookUp(0, 0) = 0
Diagonal2LookUp(0, 0) = 0

'h7b1
Diagonal2LookUp(15, 0) = 15
Diagonal2LookUp(22, 0) = 15
Diagonal2LookUp(29, 0) = 15
Diagonal2LookUp(36, 0) = 15
Diagonal2LookUp(43, 0) = 15
Diagonal2LookUp(50, 0) = 15
Diagonal2LookUp(57, 0) = 15

Diagonal2LookUp(15, 1) = 57
Diagonal2LookUp(22, 1) = 57
Diagonal2LookUp(29, 1) = 57
Diagonal2LookUp(36, 1) = 57
Diagonal2LookUp(43, 1) = 57
Diagonal2LookUp(50, 1) = 57
Diagonal2LookUp(57, 1) = 57

'h6c1
Diagonal2LookUp(23, 0) = 23
Diagonal2LookUp(30, 0) = 23
Diagonal2LookUp(37, 0) = 23
Diagonal2LookUp(44, 0) = 23
Diagonal2LookUp(51, 0) = 23
Diagonal2LookUp(58, 0) = 23

Diagonal2LookUp(23, 1) = 58
Diagonal2LookUp(30, 1) = 58
Diagonal2LookUp(37, 1) = 58
Diagonal2LookUp(44, 1) = 58
Diagonal2LookUp(51, 1) = 58
Diagonal2LookUp(58, 1) = 58

'h5d1
Diagonal2LookUp(31, 0) = 31
Diagonal2LookUp(38, 0) = 31
Diagonal2LookUp(45, 0) = 31
Diagonal2LookUp(52, 0) = 31
Diagonal2LookUp(59, 0) = 31

Diagonal2LookUp(31, 1) = 59
Diagonal2LookUp(38, 1) = 59
Diagonal2LookUp(45, 1) = 59
Diagonal2LookUp(52, 1) = 59
Diagonal2LookUp(59, 1) = 59

'h4e1
Diagonal2LookUp(39, 0) = 39
Diagonal2LookUp(46, 0) = 39
Diagonal2LookUp(53, 0) = 39
Diagonal2LookUp(60, 0) = 39

Diagonal2LookUp(39, 1) = 60
Diagonal2LookUp(46, 1) = 60
Diagonal2LookUp(53, 1) = 60
Diagonal2LookUp(60, 1) = 60

'h3f1
Diagonal2LookUp(47, 0) = 47
Diagonal2LookUp(54, 0) = 47
Diagonal2LookUp(61, 0) = 47

Diagonal2LookUp(47, 1) = 61
Diagonal2LookUp(54, 1) = 61
Diagonal2LookUp(61, 1) = 61

'h2g1
Diagonal2LookUp(55, 0) = 55
Diagonal2LookUp(62, 0) = 55

Diagonal2LookUp(55, 1) = 62
Diagonal2LookUp(62, 1) = 62

'h1h1
Diagonal2LookUp(63, 0) = 63
Diagonal2LookUp(63, 1) = 63
For i = 0 To 63
BlackPieces_Position(i) = BlackRock_Position(i) Or BlackKnight_Position(i) Or BlackBishop_Position(i) Or BlackRock_Position(i) Or BlackKing_Position(i) Or BlackQueen_Position(i) Or Blackpawn_Position(i)
WhitePieces_Position(i) = WhiteRock_Position(i) Or WhiteKnight_Position(i) Or WhiteBishop_Position(i) Or WhiteRock_Position(i) Or WhiteKing_Position(i) Or WhiteQueen_Position(i) Or Whitepawn_Position(i)
Chess_Board(i) = WhitePieces_Position(i) Or BlackPieces_Position(i)
Next i

'WPawnBonus(0) = 20
'WPawnBonus(1) = 21
'WPawnBonus(2) = 22
'WPawnBonus(3) = 23
'WPawnBonus(4) = 23
'WPawnBonus(5) = 22
'WPawnBonus(6) = 21
'WPawnBonus(7) = 20
'WPawnBonus(8) = 20
'WPawnBonus(9) = 21
''WPawnBonus(10) = 22
'WPawnBonus(11) = 23
''WPawnBonus(12) = 23
'WPawnBonus(13) = 22
'WPawnBonus(14) = 21
'WPawnBonus(15) = 20
'WPawnBonus(16) = 15
'WPawnBonus(17) = 16
'WPawnBonus(18) = 17
''WPawnBonus(19) = 18
'WPawnBonus(20) = 18
''WPawnBonus(21) = 17
'WPawnBonus(22) = 16
'WPawnBonus(23) = 15
'WPawnBonus(24) = 10
'WPawnBonus(25) = 11
'WPawnBonus(26) = 12
'WPawnBonus(27) = 13
''WPawnBonus(28) = 13
'WPawnBonus(29) = 12
''WPawnBonus(30) = 11
'WPawnBonus(31) = 10
'WPawnBonus(32) = 5
'WPawnBonus(33) = 62
'WPawnBonus(34) = 63
'WPawnBonus(35) = 64
''WPawnBonus(36) = 65
''WPawnBonus(37) = 66
'WPawnBonus(38) = 67
'WPawnBonus(39) = 68
'WPawnBonus(40) = 71
'WPawnBonus(41) = 72
'WPawnBonus(42) = 73
'WPawnBonus(43) = 74
'WPawnBonus(44) = 75
''WPawnBonus(45) = 76
''WPawnBonus(46) = 77
'WPawnBonus(47) = 78
'WPawnBonus(48) = 81
'WPawnBonus(49) = 82
'WPawnBonus(50) = 83
'WPawnBonus(51) = 84
'''WPawnBonus(52) = 85
''WPawnBonus(53) = 86
''WPawnBonus(54) = 87
'WPawnBonus(55) = 88
'WPawnBonus(56) = 91
''WPawnBonus(57) = 92
'WPawnBonus(58) = 93
'WPawnBonus(59) = 94
'WPawnBonus(60) = 95
'WPawnBonus(61) = 96
'WPawnBonus(62) = 97
'WPawnBonus(63) = 98
'
'BPawnBonus(0) = 21
'BPawnBonus(1) = 22
'BPawnBonus(2) = 23
'BPawnBonus(3) = 24
''BPawnBonus(4) = 25
''BPawnBonus(5) = 26
''BPawnBonus(6) = 27
''BPawnBonus(7) = 28
'BPawnBonus(8) = 31
'BPawnBonus(9) = 32
'BPawnBonus(10) = 33
''BPawnBonus(11) = 34
''BPawnBonus(12) = 35
''BPawnBonus(13) = 36
''BPawnBonus(14) = 37
'BPawnBonus(15) = 38
'BPawnBonus(16) = 41
'BPawnBonus(17) = 42
'BPawnBonus(18) = 43
'BPawnBonus(19) = 44
'BPawnBonus(20) = 45
'BPawnBonus(21) = 46
''BPawnBonus(22) = 47
''BPawnBonus(23) = 48
'BPawnBonus(24) = 51
'BPawnBonus(25) = 52
'BPawnBonus(26) = 53
'BPawnBonus(27) = 54
''BPawnBonus(28) = 55
''BPawnBonus(29) = 56
'BPawnBonus(30) = 57
'BPawnBonus(31) = 58
'BPawnBonus(32) = 61
'BPawnBonus(33) = 62
'BPawnBonus(34) = 63
''BPawnBonus(35) = 64
''BPawnBonus(36) = 65
'BPawnBonus(37) = 66
'BPawnBonus(38) = 67
'BPawnBonus(39) = 68
'BPawnBonus(40) = 71
'BPawnBonus(41) = 72
''BPawnBonus(42) = 73
'BPawnBonus(43) = 74
''BPawnBonus(44) = 75
'BPawnBonus(45) = 76
'BPawnBonus(46) = 77
'BPawnBonus(47) = 78
'BPawnBonus(48) = 81
'BPawnBonus(49) = 82
'BPawnBonus(50) = 83
''BPawnBonus(51) = 84
'BPawnBonus(52) = 85
'BPawnBonus(53) = 86
'BPawnBonus(54) = 87
'BPawnBonus(55) = 88
'BPawnBonus(56) = 91
''BPawnBonus(57) = 92
''BPawnBonus(58) = 93
'BPawnBonus(59) = 94
'BPawnBonus(60) = 95
'BPawnBonus(61) = 96
'BPawnBonus(62) = 97
'BPawnBonus(63) = 98
'
'
Empty_Mask(0) = 0
Empty_Mask(1) = 0
Empty_Mask(2) = 0
Empty_Mask(3) = 0
Empty_Mask(4) = 0
Empty_Mask(5) = 0
Empty_Mask(6) = 0
Empty_Mask(7) = 0
Empty_Mask(8) = 0
Empty_Mask(9) = 0
Empty_Mask(10) = 0
Empty_Mask(11) = 0
Empty_Mask(12) = 0
Empty_Mask(13) = 0
Empty_Mask(14) = 0
Empty_Mask(15) = 0
Empty_Mask(16) = 0
Empty_Mask(17) = 0
Empty_Mask(18) = 0
Empty_Mask(19) = 0
Empty_Mask(20) = 0
Empty_Mask(21) = 0
Empty_Mask(22) = 0
Empty_Mask(23) = 0
Empty_Mask(24) = 0
Empty_Mask(25) = 0
Empty_Mask(26) = 0
Empty_Mask(27) = 0
Empty_Mask(28) = 0
Empty_Mask(29) = 0
Empty_Mask(30) = 0
Empty_Mask(31) = 0
Empty_Mask(32) = 0
Empty_Mask(33) = 0
Empty_Mask(34) = 0
Empty_Mask(35) = 0
Empty_Mask(36) = 0
Empty_Mask(37) = 0
Empty_Mask(38) = 0
Empty_Mask(39) = 0
Empty_Mask(40) = 0
Empty_Mask(41) = 0
Empty_Mask(42) = 0
Empty_Mask(43) = 0
Empty_Mask(44) = 0
Empty_Mask(45) = 0
Empty_Mask(46) = 0
Empty_Mask(47) = 0
Empty_Mask(48) = 0
Empty_Mask(49) = 0
Empty_Mask(50) = 0
Empty_Mask(51) = 0
Empty_Mask(52) = 0
Empty_Mask(53) = 0
Empty_Mask(54) = 0
Empty_Mask(55) = 0
Empty_Mask(56) = 0
Empty_Mask(57) = 0
Empty_Mask(58) = 0
Empty_Mask(59) = 0
Empty_Mask(60) = 0
Empty_Mask(61) = 0
Empty_Mask(62) = 0
Empty_Mask(63) = 0
End Function



