Attribute VB_Name = "Module1"
' this moudle for all need data functions for move genertation validation evaluation and board updating


'pieces wieght for evaluation
Public Const PawnWieght = 100
Public Const KnightWieght = 300
Public Const BishopWieght = 300
Public Const RockWieght = 500
Public Const QueenWieght = 900
Public Const KingWieght = 9000
Public Const PassedWieght = 30
Public Const BackwardWieght = 10
Public Const ConectedWieght = 25
Public Const SaftyWieght = 100




Public MoveNotation As String ' move notation to be displayed or checked for validating the move


Public PostValue As Long
 
Type Mask    ' mask for attacking moves
Attack(0 To 63) As Byte
End Type

Public WhitePawn_enPassent(0 To 63) As Byte   ' to store enpassent squeres white ,black
Public Blackpawn_enPassent(0 To 63) As Byte
' be needed for  castle implementaion
Public RWRockMoved As Boolean
Public LWRockMoved As Boolean
Public RBRockMoved As Boolean
Public LBRockMoved As Boolean
Public WKingMoved As Boolean
Public BKingMoved As Boolean
Public WKO_O As Boolean
Public WKO_O_O As Boolean
Public BKO_O As Boolean
Public BKO_O_O As Boolean

' Chess board containin all pieces and Buffers
Public Chess_Board(0 To 63) As Byte
Public Chess_BoardBuffer(0 To 63) As Byte


Public MT(0 To 6, 0 To 7) As Integer 'MT "movment table" i used only for king and knight number of squeres to move
Public MB(0 To 119) As Integer ' used to check if the to squere generated is out of the board or not""
Public LookUp(0 To 63) As Integer ' it is used to calculate the squere value inside  MB if it is -1 then this squere out of the board.
'these maskes for attacking squeres
Public WKing_Mask(0 To 63) As Mask
Public Wknight_Mask(0 To 63) As Mask
Public WQueen_Mask(0 To 63) As Mask
Public WRock_Mask(0 To 63) As Mask
Public WBishop_Mask(0 To 63) As Mask
Public WPawn_Mask(0 To 63) As Mask

Public BKing_Mask(0 To 63) As Mask
Public Bknight_Mask(0 To 63) As Mask
Public BQueen_Mask(0 To 63) As Mask
Public BRock_Mask(0 To 63) As Mask
Public BBishop_Mask(0 To 63) As Mask
Public BPawn_Mask(0 To 63) As Mask

Public RankLookUp(0 To 63, 0 To 1) As Byte   'the rank begin and end squeres for each squere in the board
Public FileLookUp(0 To 63, 0 To 1) As Byte   'the file bein and end squeres  for each squere in the board
' the begin end for diagonals for each squere for the a1-h8 and h1-a8 directions
Public Diagonal1LookUp(0 To 63, 0 To 1) As Byte
Public Diagonal2LookUp(0 To 63, 0 To 1) As Byte


' maskes used for each type of piece for each color potions
Public Whitepawn_Position(0 To 63) As Byte
Public WhiteKing_Position(0 To 63) As Byte
Public WhiteQueen_Position(0 To 63) As Byte
Public WhiteKnight_Position(0 To 63) As Byte
Public WhiteBishop_Position(0 To 63) As Byte
Public WhiteRock_Position(0 To 63) As Byte
Public WhitePieces_Position(0 To 63) As Byte



Public Blackpawn_Position(0 To 63) As Byte
Public BlackKing_Position(0 To 63) As Byte
Public BlackQueen_Position(0 To 63) As Byte
Public BlackKnight_Position(0 To 63) As Byte
Public BlackBishop_Position(0 To 63) As Byte
Public BlackRock_Position(0 To 63) As Byte
Public BlackPieces_Position(0 To 63) As Byte



Public Whitepawn_PositionBuffer(0 To 63) As Byte
Public WhiteKing_PositionBuffer(0 To 63) As Byte
Public WhiteQueen_PositionBuffer(0 To 63) As Byte
Public WhiteKnight_PositionBuffer(0 To 63) As Byte
Public WhiteBishop_PositionBuffer(0 To 63) As Byte
Public WhiteRock_PositionBuffer(0 To 63) As Byte
Public WhitePieces_PositionBuffer(0 To 63) As Byte


Public Blackpawn_PositionBuffer(0 To 63) As Byte
Public BlackKing_PositionBuffer(0 To 63) As Byte
Public BlackQueen_PositionBuffer(0 To 63) As Byte
Public BlackKnight_PositionBuffer(0 To 63) As Byte
Public BlackBishop_PositionBuffer(0 To 63) As Byte
Public BlackRock_PositionBuffer(0 To 63) As Byte
Public BlackPieces_PositionBuffer(0 To 63) As Byte
'to clear any board or any 64 array
Public Empty_Mask(0 To 63) As Byte

' this type  used to collect all needed information about a single move
Public Type MoveList
From  As Byte
ToMov As Byte
PieceF As Byte
PieceT As Byte
Location As String
SideF As Integer
SideT As Integer
Attackvalue As Integer
DefenseValue As Integer
End Type


Public CurnTurn As Integer ' to identify whose  turn
Const WhiteTurn = 1
Const BlackTurn = 0

Const CheckMate = 1000000100  'high value for check mate

' exta points to encourage pawns advancaging in the center

Public WCenter_Points(0 To 63) As Integer
Public BCenter_Points(0 To 63) As Integer

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (anyDestination As Any, anySource As Any, ByVal lngLength As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Sub main_1()

End Sub
'this is where the computer evaluating all posations by using the parameters are defined before
Function EVALUATION() As Long
For i = 0 To 63
If Whitepawn_Position(i) = 1 Then
wscore = wscore + 100 + WCenter_Points(i)

ElseIf WhiteKnight_Position(i) = 1 Then
wscore = wscore + 250
ElseIf WhiteBishop_Position(i) = 1 Then
wscore = wscore + 300
ElseIf WhiteRock_Position(i) = 1 Then
wscore = wscore + 700

ElseIf WhiteQueen_Position(i) = 1 Then
wscore = wscore + 9000
ElseIf WhiteKing_Position(i) = 1 Then
wscore = wscore + 15000
ElseIf Blackpawn_Position(i) = 1 Then
bscore = bscore - 100 + BCenter_Points(i)
ElseIf BlackKnight_Position(i) = 1 Then
bscore = bscore - 250
ElseIf BlackBishop_Position(i) = 1 Then
bscore = bscore - 300
ElseIf BlackRock_Position(i) = 1 Then

bscore = bscore - 700

ElseIf BlackQueen_Position(i) = 1 Then
bscore = bscore - 9000
ElseIf BlackKing_Position(i) = 1 Then
bscore = bscore - 15000
End If







Next i


score = wscore + bscore

EVALUATION = score

End Function

Sub KingMove()
' generating king moves
 
For i = 0 To 63
If WhiteKing_Position(i) = 1 Then
CopyMemory WKing_Mask(i).Attack(0), Empty_Mask(0), 64

For B = 0 To 7
c = MB(LookUp(i) + MT(6, B))
If (c <> -1) Then
If WhitePieces_Position(c) <> 1 Then
WKing_Mask(i).Attack(c) = 1
End If

End If
Next B

ElseIf BlackKing_Position(i) = 1 Then
CopyMemory BKing_Mask(i).Attack(0), Empty_Mask(0), 64

For B = 0 To 7
c = MB(LookUp(i) + MT(6, B))
If (c <> -1) Then
If BlackPieces_Position(c) <> 1 Then
BKing_Mask(i).Attack(c) = 1
End If

End If
Next B
End If
Next i

End Sub
Sub RockMove()
' generating rook moves
Dim rockCount, rockMovment, A As Integer
rockCount = 0
rockMovment = 0
For A = 0 To 63
CopyMemory WRock_Mask(A).Attack(0), Empty_Mask(0), 64

    If rockCount < 11 Then
          If WhiteRock_Position(A) = 1 Then
             rockCount = rockCount + 1
             rockMovment = 0
             For i = A + 1 To RankLookUp(A, 1) Step 1
                 If rockMovment < 28 Then
                   If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
                        WRock_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'Brock_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
                        rockMovment = rockMovment + 1
                        Exit For

                    End If
                   Else
                    Exit For
           End If
        WRock_Mask(A).Attack(i) = 1

Next i

   For i = A - 1 To RankLookUp(A, 0) Step -1
       If rockMovment < 28 Then
          If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
             WRock_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'Brock_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
             rockMovment = rockMovment + 1
              Exit For

             End If
       Else
         Exit For
    End If
        WRock_Mask(A).Attack(i) = 1
 
   Next i

   rockMovment = 0
For i = A + 8 To FileLookUp(A, 1) Step 8
If rockMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
WRock_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'Brock_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
rockMovment = rockMovment + 1
Exit For

End If
Else
Exit For
End If
WRock_Mask(A).Attack(i) = 1

Next i

 For i = A - 8 To FileLookUp(A, 0) Step -8
If rockMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
WRock_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'Brock_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
rockMovment = rockMovment + 1
Exit For

End If
Else
Exit For
End If
WRock_Mask(A).Attack(i) = 1
Next i









End If
Else
Exit For
End If
WRock_Mask(A).Attack(A) = 0
Next A



'//////////////////
rockCount = 0
rockMovment = 0
For A = 0 To 63
CopyMemory BRock_Mask(A).Attack(0), Empty_Mask(0), 64

If rockCount < 11 Then
       If BlackRock_Position(A) = 1 Then
          rockCount = rockCount + 1
          rockMovment = 0
          For i = A + 1 To RankLookUp(A, 1) Step 1
If rockMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'Wrock_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BRock_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
rockMovment = rockMovment + 1
Exit For

End If
Else
Exit For
End If
BRock_Mask(A).Attack(i) = 1

Next i

 For i = A - 1 To RankLookUp(A, 0) Step -1
If rockMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'Wrock_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BRock_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
rockMovment = rockMovment + 1
Exit For

End If
Else
Exit For
End If

BRock_Mask(A).Attack(i) = 1
Next i





          rockMovment = 0
          For i = A + 8 To FileLookUp(A, 1) Step 8
If rockMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'Wrock_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BRock_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
rockMovment = rockMovment + 1
Exit For

End If
Else
Exit For
End If
BRock_Mask(A).Attack(i) = 1

Next i

 For i = A - 8 To FileLookUp(A, 0) Step -8
If rockMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'Brock_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BRock_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
rockMovment = rockMovment + 1
Exit For

End If
Else
Exit For
End If

BRock_Mask(A).Attack(i) = 1
Next i

End If

Else
Exit For
End If
BRock_Mask(A).Attack(A) = 0

Next A

'//////////////////







End Sub
Sub BishopMove()
'generating bishop moves
For A = 0 To 63
If WhiteBishop_Position(A) = 1 Then
CopyMemory WBishop_Mask(A).Attack(0), Empty_Mask(0), 64

For i = A - 7 To Diagonal2LookUp(A, 0) Step -7
If WhitePieces_Position(i) <> 1 Then
If BlackPieces_Position(i) <> 1 Then
WBishop_Mask(A).Attack(i) = 1
Else
WBishop_Mask(A).Attack(i) = 1
Exit For
End If
Else
Exit For
End If

Next i

For i = A + 7 To Diagonal2LookUp(A, 1) Step 7
If WhitePieces_Position(i) <> 1 Then
If BlackPieces_Position(i) <> 1 Then
WBishop_Mask(A).Attack(i) = 1
Else
WBishop_Mask(A).Attack(i) = 1
Exit For
End If
Else
Exit For
End If

Next i
For i = A - 9 To Diagonal1LookUp(A, 0) Step -9
If WhitePieces_Position(i) <> 1 Then
If BlackPieces_Position(i) <> 1 Then
WBishop_Mask(A).Attack(i) = 1
Else
WBishop_Mask(A).Attack(i) = 1
Exit For
End If
Else
Exit For
End If

Next i
For i = A + 9 To Diagonal1LookUp(A, 1) Step 9
If WhitePieces_Position(i) <> 1 Then
If BlackPieces_Position(i) <> 1 Then
WBishop_Mask(A).Attack(i) = 1
Else
WBishop_Mask(A).Attack(i) = 1
Exit For
End If
Else
Exit For
End If

Next i
WBishop_Mask(A).Attack(A) = 0
End If

Next A


For A = 0 To 63
If BlackBishop_Position(A) = 1 Then
CopyMemory BBishop_Mask(A).Attack(0), Empty_Mask(0), 64

For i = A - 7 To Diagonal2LookUp(A, 0) Step -7
If BlackPieces_Position(i) <> 1 Then
If WhitePieces_Position(i) <> 1 Then
BBishop_Mask(A).Attack(i) = 1
Else
BBishop_Mask(A).Attack(i) = 1
Exit For
End If
Else
Exit For
End If


Next i

For i = A + 7 To Diagonal2LookUp(A, 1) Step 7
If BlackPieces_Position(i) <> 1 Then
If WhitePieces_Position(i) <> 1 Then
BBishop_Mask(A).Attack(i) = 1
Else
BBishop_Mask(A).Attack(i) = 1
Exit For
End If
Else
Exit For
End If

Next i
For i = A - 9 To Diagonal1LookUp(A, 0) Step -9
If BlackPieces_Position(i) <> 1 Then
If WhitePieces_Position(i) <> 1 Then
BBishop_Mask(A).Attack(i) = 1
Else
BBishop_Mask(A).Attack(i) = 1
Exit For
End If
Else
Exit For
End If

Next i
For i = A + 9 To Diagonal1LookUp(A, 1) Step 9
If BlackPieces_Position(i) <> 1 Then
If WhitePieces_Position(i) <> 1 Then
BBishop_Mask(A).Attack(i) = 1
Else
BBishop_Mask(A).Attack(i) = 1
Exit For
End If
Else
Exit For
End If

Next i
BBishop_Mask(A).Attack(A) = 0
End If

Next A
End Sub

Sub QueenMove()
' generating queen moves
Dim QueenCount, QueenMovment, A As Integer
QueenCount = 0
QueenMovment = 0
For A = 0 To 63
CopyMemory WQueen_Mask(A).Attack(0), Empty_Mask(0), 64

    If QueenCount < 11 Then
          If WhiteQueen_Position(A) = 1 Then
             QueenCount = QueenCount + 1
             QueenMovment = 0
             For i = A + 1 To RankLookUp(A, 1) Step 1
                 If QueenMovment < 28 Then
                   If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
                        WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
                        QueenMovment = QueenMovment + 1
                        Exit For

                    End If
                   Else
                    Exit For
           End If
        WQueen_Mask(A).Attack(i) = 1

Next i

   For i = A - 1 To RankLookUp(A, 0) Step -1
       If QueenMovment < 28 Then
          If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
             WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
             QueenMovment = QueenMovment + 1
              Exit For

             End If
       Else
         Exit For
    End If
        WQueen_Mask(A).Attack(i) = 1
 
   Next i

   QueenMovment = 0
For i = A + 8 To FileLookUp(A, 1) Step 8
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
WQueen_Mask(A).Attack(i) = 1

Next i

 For i = A - 8 To FileLookUp(A, 0) Step -8
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
WQueen_Mask(A).Attack(i) = 1
Next i

         For i = A + 9 To Diagonal1LookUp(A, 1) Step 9
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
WQueen_Mask(A).Attack(i) = 1

Next i

 For i = A - 9 To Diagonal1LookUp(A, 0) Step -9
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
WQueen_Mask(A).Attack(i) = 1

Next i





          For i = A + 7 To Diagonal2LookUp(A, 1) Step 7
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
WQueen_Mask(A).Attack(i) = 1

Next i

 For i = A - 7 To Diagonal2LookUp(A, 0) Step -7
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
'BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
WQueen_Mask(A).Attack(i) = 1

Next i

End If
Else
Exit For
End If
WQueen_Mask(A).Attack(A) = 0
Next A



'//////////////////
QueenCount = 0
QueenMovment = 0
For A = 0 To 63
CopyMemory BQueen_Mask(A).Attack(0), Empty_Mask(0), 64

If QueenCount < 11 Then
       If BlackQueen_Position(A) = 1 Then
          QueenCount = QueenCount + 1
          QueenMovment = 0
          For i = A + 1 To RankLookUp(A, 1) Step 1
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
BQueen_Mask(A).Attack(i) = 1

Next i

 For i = A - 1 To RankLookUp(A, 0) Step -1
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If

BQueen_Mask(A).Attack(i) = 1
Next i





          QueenMovment = 0
          For i = A + 8 To FileLookUp(A, 1) Step 8
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
BQueen_Mask(A).Attack(i) = 1

Next i

 For i = A - 8 To FileLookUp(A, 0) Step -8
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'BQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If

BQueen_Mask(A).Attack(i) = 1
Next i

          For i = A + 9 To Diagonal1LookUp(A, 1) Step 9
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
BQueen_Mask(A).Attack(i) = 1

Next i

 For i = A - 9 To Diagonal1LookUp(A, 0) Step -9
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If

BQueen_Mask(A).Attack(i) = 1
Next i





          QueenMovment = 0
          For i = A + 7 To Diagonal2LookUp(A, 1) Step 7
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'WQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If
BQueen_Mask(A).Attack(i) = 1

Next i

 For i = A - 7 To Diagonal2LookUp(A, 0) Step -7
If QueenMovment < 28 Then
If WhitePieces_Position(i) = 1 Or BlackPieces_Position(i) = 1 Then
'BQueen_Mask(A).Attack(i) = 255 Xor ((WhitePieces_Position(i) And Not BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
BQueen_Mask(A).Attack(i) = 255 Xor ((Not WhitePieces_Position(i) And BlackPieces_Position(i)) Xor (Not WhitePieces_Position(i) And Not BlackPieces_Position(i)))
QueenMovment = QueenMovment + 1
Exit For

End If
Else
Exit For
End If

BQueen_Mask(A).Attack(i) = 1
Next i

End If

Else
Exit For
End If
BQueen_Mask(A).Attack(A) = 0

Next A

'//////////////////










End Sub
Sub KnightMove()
' knight move generation
For a1 = 0 To 63
CopyMemory Wknight_Mask(a1).Attack(0), Empty_Mask(0), 64
CopyMemory Bknight_Mask(a1).Attack(0), Empty_Mask(0), 64

Next a1

For i = 0 To 63
If WhiteKnight_Position(i) = 1 Then
For B = 0 To 7
c = MB(LookUp(i) + MT(2, B))
If (c <> -1) Then
'If Not WhitePieces_Position(c) Then
 If WhitePieces_Position(c) <> 1 Then
Wknight_Mask(i).Attack(c) = 1

End If

End If
Next B
Wknight_Mask(i).Attack(i) = 0
End If
Next i
For i = 0 To 63
If BlackKnight_Position(i) = 1 Then
For B = 0 To 7

c = MB(LookUp(i) + MT(2, B))
If (c <> -1) Then

'If Not WhitePieces_Position(c) Then
If BlackPieces_Position(c) <> 1 Then
Bknight_Mask(i).Attack(c) = 1
'End If
End If
End If

Next B
Bknight_Mask(i).Attack(i) = 0
End If

Next i
End Sub

Function ValidMove(ByVal From As Integer, ByVal ToM As Integer) As Boolean
'used to validate moves
Dim Side As Integer ' side = 1 for white ...side =0 for black
Dim sidefrm, sideto As Integer

Dim validation As Boolean
validation = False
Dim PType  As Integer


Side = -1
sidefrm = -1
sideto = -1


If WhitePieces_Position(From) Then
Side = 1
sidefrm = 1


ElseIf BlackPieces_Position(From) Then
Side = 0
sidefrm = 0


End If

If WhitePieces_Position(ToM) Then
sideto = 1
ElseIf BlackPieces_Position(ToM) Then
sideto = 0
End If
If WhiteKing_Position(ToM) Then
ValidMove = False
Exit Function
End If
If BlackKing_Position(ToM) Then
ValidMove = False
Exit Function
End If

If Side Then


If Whitepawn_Position(From) Then
If WPawn_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 1
Else
validation = False
End If

End If

If WhiteKnight_Position(From) Then
If (Wknight_Mask(From).Attack(ToM) = 1) Then
validation = True
PType = 2
Else
validation = False
End If
End If

If WhiteBishop_Position(From) Then
If WBishop_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 3
Else
validation = False
End If
End If

If WhiteRock_Position(From) Then
If WRock_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 4
Else
validation = False
End If
End If

If WhiteQueen_Position(From) Then
If WQueen_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 5
Else
validation = False
End If

End If

If WhiteKing_Position(From) Then
If WKing_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 6
Else
validation = False
End If

End If




Else

If Blackpawn_Position(From) Then
If BPawn_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 1
Else
validation = False
End If

End If

If BlackKnight_Position(From) Then
If Bknight_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 2

Else
validation = False
End If

End If

If BlackBishop_Position(From) Then
If BBishop_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 3
Else
validation = False
End If

End If

If BlackRock_Position(From) Then
If BRock_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 4
Else
validation = False
End If

End If

If BlackQueen_Position(From) Then
If BQueen_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 5
Else
validation = False
End If

End If

If BlackKing_Position(From) Then
If BKing_Mask(From).Attack(ToM) = 1 Then
validation = True
PType = 6
Else
validation = False
End If
End If
End If

If sideto = 1 Then
 
 If Whitepawn_Position(ToM) Then
 PtoType = 1
 ElseIf WhiteKnight_Position(ToM) Then
 PtoType = 2
 ElseIf WhiteBishop_Position(ToM) Then
 PtoType = 3
 ElseIf WhiteRock_Position(ToM) Then
 PtoType = 4
 ElseIf WhiteQueen_Position(ToM) Then
 PtoType = 5
 ElseIf WhiteKing_Position(ToM) Then
 PtoType = 6
End If
 ElseIf sideto = 0 Then
 If Blackpawn_Position(ToM) Then
 PtoType = 1
 ElseIf BlackKnight_Position(ToM) Then
 PtoType = 2
 ElseIf BlackBishop_Position(ToM) Then
 PtoType = 3
 ElseIf BlackRock_Position(ToM) Then
 PtoType = 4
 ElseIf BlackQueen_Position(ToM) Then
 PtoType = 5
 ElseIf BlackKing_Position(ToM) Then
 PtoType = 6
End If
 Else
 PtoType = 0
 End If
 
 If validation Then
MovFrm = From
MovTo = ToM
MoveNotation = ""
MoveNotation = MovToStr(MovFrm, PType, ToM, PtoType)
 Call UpdateBoards(MovFrm, PType, sidefrm, ToM, PtoType, sideto)
 End If
 
 

ValidMove = validation

End Function
Sub PawnMove()
' generating pawn moves
For i = 0 To 63
CopyMemory WPawn_Mask(i).Attack(0), Empty_Mask(0), 64
CopyMemory BPawn_Mask(i).Attack(0), Empty_Mask(0), 64

Next i

Dim B

For i = 48 To 55
If Whitepawn_Position(i) = 1 Then
c = i - 8
B = MB(LookUp(i) - 11)
If B <> -1 Then
If (BlackPieces_Position(B) = 1) Then
WPawn_Mask(i).Attack(B) = 1
End If
End If
B = MB(LookUp(i) - 9)
If B <> -1 Then
If (BlackPieces_Position(B) = 1) Then
WPawn_Mask(i).Attack(B) = 1
End If
End If
If WhitePieces_Position(c) <> 1 And BlackPieces_Position(c) <> 1 Then
WPawn_Mask(i).Attack(c) = 1


c = c - 8
If WhitePieces_Position(c) <> 1 And BlackPieces_Position(c) <> 1 Then
WPawn_Mask(i).Attack(c) = 1
End If
End If
End If
Next i

For i = 8 To 47
If Whitepawn_Position(i) = 1 Then
B = MB(LookUp(i) - 11)
If B <> -1 Then
If (BlackPieces_Position(B) = 1) Then
WPawn_Mask(i).Attack(B) = 1
End If
End If
B = MB(LookUp(i) - 9)
If B <> -1 Then
If (BlackPieces_Position(B) = 1) Then
WPawn_Mask(i).Attack(B) = 1
End If
End If
c = i - 8

If WhitePieces_Position(c) <> 1 And BlackPieces_Position(c) <> 1 Then
WPawn_Mask(i).Attack(c) = 1

End If
End If

Next i



For i = 8 To 15
If Blackpawn_Position(i) = 1 Then
c = i + 8
B = MB(LookUp(i) + 11)
If B <> -1 Then
If (WhitePieces_Position(B) = 1) Then
BPawn_Mask(i).Attack(B) = 1
End If
End If
B = MB(LookUp(i) + 9)
If B <> -1 Then
If (WhitePieces_Position(B) = 1) Then
BPawn_Mask(i).Attack(B) = 1
End If
End If
If WhitePieces_Position(c) <> 1 And BlackPieces_Position(c) <> 1 Then
BPawn_Mask(i).Attack(c) = 1

c = c + 8
B = MB(LookUp(i) + 11)
If B <> -1 Then
If (WhitePieces_Position(B) = 1) Then
BPawn_Mask(i).Attack(B) = 1
End If
End If
B = MB(LookUp(i) + 9)
If B <> -1 Then
If (WhitePieces_Position(B) = 1) Then
BPawn_Mask(i).Attack(B) = 1
End If
End If
If WhitePieces_Position(c) <> 1 And BlackPieces_Position(c) <> 1 Then
BPawn_Mask(i).Attack(c) = 1

End If
End If
End If
Next i

For i = 8 To 55
If Blackpawn_Position(i) = 1 Then

c = i + 8
B = MB(LookUp(i) + 11)
If B <> -1 Then
If (WhitePieces_Position(B) = 1) Then
BPawn_Mask(i).Attack(B) = 1
End If
End If
B = MB(LookUp(i) + 9)
If B <> -1 Then
If (WhitePieces_Position(B) = 1) Then
BPawn_Mask(i).Attack(B) = 1
End If
End If
If WhitePieces_Position(c) <> 1 And BlackPieces_Position(c) <> 1 Then
BPawn_Mask(i).Attack(c) = 1

End If
End If
Next i








End Sub
Sub UpdateBoards(ByVal MovFrm As Integer, ByVal piecfrm As Integer, ByVal sidefrm As Integer, ByVal MovTo As Integer, ByVal piecto As Integer, ByVal sideto As Integer)
' updating the board with the move
Select Case sidefrm

Case 1:
Select Case piecfrm
Case 1: Whitepawn_Position(MovFrm) = 0
        Whitepawn_Position(MovTo) = 1
        
Case 2: WhiteKnight_Position(MovFrm) = 0
        WhiteKnight_Position(MovTo) = 1

Case 3: WhiteBishop_Position(MovFrm) = 0
        WhiteBishop_Position(MovTo) = 1
        
Case 4: WhiteRock_Position(MovFrm) = 0
        WhiteRock_Position(MovTo) = 1

Case 5: WhiteQueen_Position(MovFrm) = 0
        WhiteQueen_Position(MovTo) = 1
        
Case 6: WhiteKing_Position(MovFrm) = 0
        WhiteKing_Position(MovTo) = 1

End Select
WhitePieces_Position(MovFrm) = 0
WhitePieces_Position(MovTo) = 1
Chess_Board(MovFrm) = 0
Chess_Board(MovTo) = 1

Case 0:
Select Case piecfrm
Case 1: Blackpawn_Position(MovFrm) = 0
        Blackpawn_Position(MovTo) = 1
        
Case 2: BlackKnight_Position(MovFrm) = 0
        BlackKnight_Position(MovTo) = 1

Case 3: BlackBishop_Position(MovFrm) = 0
        BlackBishop_Position(MovTo) = 1
        
Case 4: BlackRock_Position(MovFrm) = 0
        BlackRock_Position(MovTo) = 1

Case 5: BlackQueen_Position(MovFrm) = 0
        BlackQueen_Position(MovTo) = 1
        
Case 6: BlackKing_Position(MovFrm) = 0
        BlackKing_Position(MovTo) = 1

End Select
BlackPieces_Position(MovFrm) = 0
BlackPieces_Position(MovTo) = 1
Chess_Board(MovFrm) = 0
Chess_Board(MovTo) = 1

Case -1:

End Select

Select Case sideto

Case 1:
Select Case piecto
Case 1: Whitepawn_Position(MovTo) = 0
        
Case 2: WhiteKnight_Position(MovTo) = 0

Case 3: WhiteBishop_Position(MovTo) = 0
        
Case 4: WhiteRock_Position(MovTo) = 0

Case 5: WhiteQueen_Position(MovTo) = 0
        
Case 6: WhiteKing_Position(MovTo) = 0

End Select
WhitePieces_Position(MovTo) = 0
Chess_Board(MovFrm) = 0
Chess_Board(MovTo) = 1
Case 0:
Select Case piecto
Case 1: Blackpawn_Position(MovTo) = 0
        
Case 2: BlackKnight_Position(MovTo) = 0

Case 3: BlackBishop_Position(MovTo) = 0
        
Case 4: BlackRock_Position(MovTo) = 0
        
Case 5: BlackQueen_Position(MovTo) = 0
        
Case 6: BlackKing_Position(MovTo) = 0

End Select

BlackPieces_Position(MovTo) = 0
Chess_Board(MovFrm) = 0
Chess_Board(MovTo) = 1

Case -1:
End Select
Dim Mobility As Integer
Side = 0


Dim move_list(100) As MoveList


End Sub
Sub Gen_Movment()
'this function used to start all move generation functions
Call BishopMove
Call QueenMove
Call RockMove
Call KingMove
Call KnightMove
Call PawnMove
End Sub
Sub Generate_Moves(ByVal Side As Integer, ByRef Mobility As Integer, ByRef move_list() As MoveList)
' generate move list for the side to move depending on the pieces postion at that time and return also the mobility for number of possible moves

Dim Mobile As Integer
Dim sf, st, frm, Tos, pf, pt As Integer
Call Gen_Movment

If Side = 1 Then
sf = 1
For A = 0 To 63
If Whitepawn_Position(A) = 1 Then
For i = 0 To 63
If WPawn_Mask(A).Attack(i) = 1 Then
          If BlackPieces_Position(i) = 1 Then
            
sideto = 0
Else
sideto = -1
End If
If sideto = 0 Then
 If Blackpawn_Position(i) = 1 Then
 pt = 1
 ElseIf BlackKnight_Position(i) = 1 Then
 pt = 2
 ElseIf BlackBishop_Position(i) = 1 Then
 pt = 3
 ElseIf BlackRock_Position(i) = 1 Then
 pt = 4
 ElseIf BlackQueen_Position(i) = 1 Then
 pt = 5
 ElseIf BlackKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 1, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 1
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 1
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100
End Select

End If

Next i
ElseIf WhiteKnight_Position(A) = 1 Then

For i = 0 To 63
If Wknight_Mask(A).Attack(i) = 1 Then
          If BlackPieces_Position(i) = 1 Then
            
sideto = 0
Else
sideto = -1
End If
If sideto = 0 Then
 If Blackpawn_Position(i) = 1 Then
 pt = 1
 ElseIf BlackKnight_Position(i) = 1 Then
 pt = 2
 ElseIf BlackBishop_Position(i) = 1 Then
 pt = 3
 ElseIf BlackRock_Position(i) = 1 Then
 pt = 4
 ElseIf BlackQueen_Position(i) = 1 Then
 pt = 5
 ElseIf BlackKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 2, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 2
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 1
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100
End Select
End If

Next i

ElseIf WhiteBishop_Position(A) = 1 Then
For i = 0 To 63
If WBishop_Mask(A).Attack(i) = 1 And WhitePieces_Position(i) <> 1 Then
          If BlackPieces_Position(i) = 1 Then
            
sideto = 0
Else
sideto = -1
End If
If sideto = 0 Then
 If Blackpawn_Position(i) = 1 Then
 pt = 1
 ElseIf BlackKnight_Position(i) = 1 Then
 pt = 2
 ElseIf BlackBishop_Position(i) = 1 Then
 pt = 3
 ElseIf BlackRock_Position(i) = 1 Then
 pt = 4
 ElseIf BlackQueen_Position(i) = 1 Then
 pt = 5
 ElseIf BlackKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 3, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 3
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 1
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100
End Select
End If

Next i

ElseIf WhiteRock_Position(A) = 1 Then
For i = 0 To 63
If WRock_Mask(A).Attack(i) = 1 Then
          If BlackPieces_Position(i) = 1 Then
            
sideto = 0
Else
sideto = -1
End If
If sideto = 0 Then
 If Blackpawn_Position(i) Then
 pt = 1
 ElseIf BlackKnight_Position(i) = 1 Then
 pt = 2
 ElseIf BlackBishop_Position(i) = 1 Then
 pt = 3
 ElseIf BlackRock_Position(i) = 1 Then
 pt = 4
 ElseIf BlackQueen_Position(i) = 1 Then
 pt = 5
 ElseIf BlackKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 4, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 4
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 1
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100
End Select
End If

Next i
ElseIf WhiteQueen_Position(A) = 1 Then
For i = 0 To 63


If WQueen_Mask(A).Attack(i) = 1 Then
          If BlackPieces_Position(i) = 1 Then
          
sideto = 0
Else
sideto = -1
End If
If sideto = 0 Then
 If Blackpawn_Position(i) = 1 Then
 pt = 1
 ElseIf BlackKnight_Position(i) = 1 Then
 pt = 2
 ElseIf BlackBishop_Position(i) = 1 Then
 pt = 3
 ElseIf BlackRock_Position(i) = 1 Then
 pt = 4
 ElseIf BlackQueen_Position(i) = 1 Then
 pt = 5
 ElseIf BlackKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 5, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 5
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 1
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100
End Select
End If

Next i

ElseIf WhiteKing_Position(A) = 1 Then
For i = 0 To 63
If WKing_Mask(A).Attack(i) = 1 Then
          If BlackPieces_Position(i) = 1 Then
            
sideto = 0
Else
sideto = -1
End If
If sideto = 0 Then
 If Blackpawn_Position(i) Then
 pt = 1
 ElseIf BlackKnight_Position(i) = 1 Then
 pt = 2
 ElseIf BlackBishop_Position(i) = 1 Then
 pt = 3
 ElseIf BlackRock_Position(i) = 1 Then
 pt = 4
 ElseIf BlackQueen_Position(i) = 1 Then
 pt = 5
 ElseIf BlackKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 6, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 6
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 1
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 110
Case 6: move_list(Mobile).Attackvalue = 0
End Select
End If

Next i
End If

Next A


ElseIf Side = 0 Then
sf = 0

For A = 0 To 63
If Blackpawn_Position(A) = 1 Then
For i = 0 To 63
If BPawn_Mask(A).Attack(i) = 1 Then
          If WhitePieces_Position(i) = 1 Then
            
sideto = 1
Else
sideto = -1
End If
If sideto = 1 Then
 If Whitepawn_Position(i) = 1 Then
 pt = 1
 ElseIf WhiteKnight_Position(i) = 1 Then
 pt = 2
 ElseIf WhiteBishop_Position(i) = 1 Then
 pt = 3
 ElseIf WhiteRock_Position(i) = 1 Then
 pt = 4
 ElseIf WhiteQueen_Position(i) = 1 Then
 pt = 5
 ElseIf WhiteKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 1, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 1
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 0
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100
End Select
End If

Next i
ElseIf BlackKnight_Position(A) = 1 Then

For i = 0 To 63
If Bknight_Mask(A).Attack(i) = 1 Then
          If WhitePieces_Position(i) = 1 Then
            
sideto = 1
Else
sideto = -1
End If
If sideto = 1 Then
 If Whitepawn_Position(i) = 1 Then
 pt = 1
 ElseIf WhiteKnight_Position(i) = 1 Then
 pt = 2
 ElseIf WhiteBishop_Position(i) = 1 Then
 pt = 3
 ElseIf WhiteRock_Position(i) = 1 Then
 pt = 4
 ElseIf WhiteQueen_Position(i) = 1 Then
 pt = 5
 ElseIf WhiteKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 2, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 2
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 0
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100
End Select
End If
Next i

ElseIf BlackBishop_Position(A) = 1 Then
For i = 0 To 63
If BBishop_Mask(A).Attack(i) = 1 Then
          If WhitePieces_Position(i) = 1 Then
            
sideto = 1
Else
sideto = -1
End If
If sideto = 1 Then
 If Whitepawn_Position(i) = 1 Then
 pt = 1
 ElseIf WhiteKnight_Position(i) = 1 Then
 pt = 2
 ElseIf WhiteBishop_Position(i) = 1 Then
 pt = 3
 ElseIf WhiteRock_Position(i) = 1 Then
 pt = 4
 ElseIf WhiteQueen_Position(i) = 1 Then
 pt = 5
 ElseIf WhiteKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 3, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 3
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 0
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100
End Select
End If
Next i

ElseIf BlackRock_Position(A) = 1 Then
For i = 0 To 63
If BRock_Mask(A).Attack(i) = 1 Then
          If WhitePieces_Position(i) = 1 Then
            
sideto = 1
Else
sideto = -1
End If
If sideto = 1 Then
 If Whitepawn_Position(i) = 1 Then
 pt = 1
 ElseIf WhiteKnight_Position(i) = 1 Then
 pt = 2
 ElseIf WhiteBishop_Position(i) = 1 Then
 pt = 3
 ElseIf WhiteRock_Position(i) = 1 Then
 pt = 4
 ElseIf WhiteQueen_Position(i) = 1 Then
 pt = 5
 ElseIf WhiteKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 4, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 4
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 0
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100
End Select
End If

Next i
ElseIf BlackQueen_Position(A) = 1 Then
For i = 0 To 63
If BQueen_Mask(A).Attack(i) = 1 Then
          If WhitePieces_Position(i) = 1 Then
            
sideto = 1
Else
sideto = -1
End If
If sideto = 1 Then
 If Whitepawn_Position(i) = 1 Then
 pt = 1
 ElseIf WhiteKnight_Position(i) = 1 Then
 pt = 2
 ElseIf WhiteBishop_Position(i) = 1 Then
 pt = 3
 ElseIf WhiteRock_Position(i) = 1 Then
 pt = 4
 ElseIf WhiteQueen_Position(i) = 1 Then
 pt = 5
 ElseIf WhiteKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 5, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 5
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 0
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 10
Case 6: move_list(Mobile).Attackvalue = 100


End Select
End If

Next i

ElseIf BlackKing_Position(A) = 1 Then
For i = 0 To 63
If BKing_Mask(A).Attack(i) = 1 Then
          If BlackPieces_Position(i) = 1 Then
            
sideto = 1
Else
sideto = -1
End If
If sideto = 1 Then
 If Whitepawn_Position(i) = 1 Then
 pt = 1
 ElseIf WhiteKnight_Position(i) = 1 Then
 pt = 2
 ElseIf WhiteBishop_Position(i) = 1 Then
 pt = 3
 ElseIf WhiteRock_Position(i) = 1 Then
 pt = 4
 ElseIf WhiteQueen_Position(i) = 1 Then
 pt = 5
 ElseIf WhiteKing_Position(i) = 1 Then
 pt = 6
End If
 Else
 pt = 0
 End If
Mobile = Mobile + 1
'
MovNot = MovToStr(A, 6, i, pt)
move_list(Mobile).From = A
move_list(Mobile).ToMov = i
move_list(Mobile).Location = MovNot
move_list(Mobile).PieceF = 6
move_list(Mobile).PieceT = pt
move_list(Mobile).SideF = 0
move_list(Mobile).SideT = sideto
Select Case pt
Case 0: move_list(Mobile).Attackvalue = 0
Case 1: move_list(Mobile).Attackvalue = 1
Case 2: move_list(Mobile).Attackvalue = 3
Case 3: move_list(Mobile).Attackvalue = 4
Case 4: move_list(Mobile).Attackvalue = 5
Case 5: move_list(Mobile).Attackvalue = 110
Case 6: move_list(Mobile).Attackvalue = 0
End Select
End If

Next i
End If

Next A

End If
Mobility = Mobile

Call SortList(move_list)


End Sub
Function MovToStr(ByVal From As Integer, ByVal piecfrm As Integer, ByVal Tos As Integer, ByVal piecto As Integer) As String
' used to convert the from and to squeres to notaions
Dim Fstr, Tstr, pfstr, ptstr, Movstr1 As String
pfstr = ""
ptstr = ""

Select Case piecfrm

Case 1: pfstr = "P"
        
Case 2: pfstr = "N"

Case 3: pfstr = "B"
        
Case 4: pfstr = "R"

Case 5: pfstr = "Q"
        
Case 6: pfstr = "K"
End Select

Select Case piecto
Case 1: ptstr = "P"
        
Case 2: ptstr = "N"

Case 3: ptstr = "B"
        
Case 4: ptstr = "R"

Case 5: ptstr = "Q"
        
Case 6: ptstr = "K"
End Select
Select Case From

Case 0: Fstr = "a8"
Case 1: Fstr = "b8"
Case 2: Fstr = "c8"
Case 3: Fstr = "d8"
Case 4: Fstr = "e8"
Case 5: Fstr = "f8"
Case 6: Fstr = "g8"
Case 7: Fstr = "h8"

Case 8: Fstr = "a7"
Case 9: Fstr = "b7"
Case 10: Fstr = "c7"
Case 11: Fstr = "d7"
Case 12: Fstr = "e7"
Case 13: Fstr = "f7"
Case 14: Fstr = "g7"
Case 15: Fstr = "h7"

Case 16: Fstr = "a6"
Case 17: Fstr = "b6"
Case 18: Fstr = "c6"
Case 19: Fstr = "d6"
Case 20: Fstr = "e6"
Case 21: Fstr = "f6"
Case 22: Fstr = "g6"
Case 23: Fstr = "h6"

Case 24: Fstr = "a5"
Case 25: Fstr = "b5"
Case 26: Fstr = "c5"
Case 27: Fstr = "d5"
Case 28: Fstr = "e5"
Case 29: Fstr = "f5"
Case 30: Fstr = "g5"
Case 31: Fstr = "h5"

Case 32: Fstr = "a4"
Case 33: Fstr = "b4"
Case 34: Fstr = "c4"
Case 35: Fstr = "d4"
Case 36: Fstr = "e4"
Case 37: Fstr = "f4"
Case 38: Fstr = "g4"
Case 39: Fstr = "h4"

Case 40: Fstr = "a3"
Case 41: Fstr = "b3"
Case 42: Fstr = "c3"
Case 43: Fstr = "d3"
Case 44: Fstr = "e3"
Case 45: Fstr = "f3"
Case 46: Fstr = "g3"
Case 47: Fstr = "h3"

Case 48: Fstr = "a2"
Case 49: Fstr = "b2"
Case 50: Fstr = "c2"
Case 51: Fstr = "d2"
Case 52: Fstr = "e2"
Case 53: Fstr = "f2"
Case 54: Fstr = "g2"
Case 55: Fstr = "h2"

Case 56: Fstr = "a1"
Case 57: Fstr = "b1"
Case 58: Fstr = "c1"
Case 59: Fstr = "d1"
Case 60: Fstr = "e1"
Case 61: Fstr = "f1"
Case 62: Fstr = "g1"
Case 63: Fstr = "h1"
End Select

Select Case Tos

Case 0: Tstr = "a8"
Case 1: Tstr = "b8"
Case 2: Tstr = "c8"
Case 3: Tstr = "d8"
Case 4: Tstr = "e8"
Case 5: Tstr = "f8"
Case 6: Tstr = "g8"
Case 7: Tstr = "h8"

Case 8: Tstr = "a7"
Case 9: Tstr = "b7"
Case 10: Tstr = "c7"
Case 11: Tstr = "d7"
Case 12: Tstr = "e7"
Case 13: Tstr = "f7"
Case 14: Tstr = "g7"
Case 15: Tstr = "h7"

Case 16: Tstr = "a6"
Case 17: Tstr = "b6"
Case 18: Tstr = "c6"
Case 19: Tstr = "d6"
Case 20: Tstr = "e6"
Case 21: Tstr = "f6"
Case 22: Tstr = "g6"
Case 23: Tstr = "h6"

Case 24: Tstr = "a5"
Case 25: Tstr = "b5"
Case 26: Tstr = "c5"
Case 27: Tstr = "d5"
Case 28: Tstr = "e5"
Case 29: Tstr = "f5"
Case 30: Tstr = "g5"
Case 31: Tstr = "h5"

Case 32: Tstr = "a4"
Case 33: Tstr = "b4"
Case 34: Tstr = "c4"
Case 35: Tstr = "d4"
Case 36: Tstr = "e4"
Case 37: Tstr = "f4"
Case 38: Tstr = "g4"
Case 39: Tstr = "h4"

Case 40: Tstr = "a3"
Case 41: Tstr = "b3"
Case 42: Tstr = "c3"
Case 43: Tstr = "d3"
Case 44: Tstr = "e3"
Case 45: Tstr = "f3"
Case 46: Tstr = "g3"
Case 47: Tstr = "h3"

Case 48: Tstr = "a2"
Case 49: Tstr = "b2"
Case 50: Tstr = "c2"
Case 51: Tstr = "d2"
Case 52: Tstr = "e2"
Case 53: Tstr = "f2"
Case 54: Tstr = "g2"
Case 55: Tstr = "h2"

Case 56: Tstr = "a1"
Case 57: Tstr = "b1"
Case 58: Tstr = "c1"
Case 59: Tstr = "d1"
Case 60: Tstr = "e1"
Case 61: Tstr = "f1"
Case 62: Tstr = "g1"
Case 63: Tstr = "h1"
End Select
Movstr1 = pfstr & Fstr & "-" & ptstr & Tstr
MovToStr = Movstr1
End Function

Sub do_Move(Mov As MoveList)
' used to make the move on the board to evaluated or to generate the oppenent moves depending on the chanes happen by this move

CopyMemory Whitepawn_PositionBuffer(0), Whitepawn_Position(0), 64
CopyMemory WhiteKing_PositionBuffer(0), WhiteKing_Position(0), 64
CopyMemory WhiteBishop_PositionBuffer(0), WhiteBishop_Position(0), 64
CopyMemory WhiteKnight_PositionBuffer(0), WhiteKnight_Position(0), 64
CopyMemory WhiteRock_PositionBuffer(0), WhiteRock_Position(0), 64
CopyMemory WhiteQueen_PositionBuffer(0), WhiteQueen_Position(0), 64
CopyMemory WhitePieces_PositionBuffer(0), WhitePieces_Position(0), 64
CopyMemory Blackpawn_PositionBuffer(0), Blackpawn_Position(0), 64
CopyMemory BlackKing_PositionBuffer(0), BlackKing_Position(0), 64
CopyMemory BlackBishop_PositionBuffer(0), BlackBishop_Position(0), 64
CopyMemory BlackKnight_PositionBuffer(0), BlackKnight_Position(0), 64
CopyMemory BlackRock_PositionBuffer(0), BlackRock_Position(0), 64
CopyMemory BlackQueen_PositionBuffer(0), BlackQueen_Position(0), 64
CopyMemory BlackPieces_PositionBuffer(0), BlackPieces_Position(0), 64
CopyMemory Chess_BoardBuffer(0), Chess_Board(0), 64


If Mov.SideF = 1 Then
If Mov.SideT = 0 Then
WhitePieces_Position(Mov.From) = 0
WhitePieces_Position(Mov.ToMov) = 1
BlackPieces_Position(Mov.ToMov) = 0

Select Case Mov.PieceT
Case 1: Blackpawn_Position(Mov.ToMov) = 0
        
Case 2: BlackKnight_Position(Mov.ToMov) = 0
        
Case 3: BlackBishop_Position(Mov.ToMov) = 0
        
Case 4: BlackRock_Position(Mov.ToMov) = 0
        
Case 5: BlackQueen_Position(Mov.ToMov) = 0
        
Case 6: BlackKing_Position(Mov.ToMov) = 0
        
End Select


Else
WhitePieces_Position(Mov.From) = 0
WhitePieces_Position(Mov.ToMov) = 1
End If
Select Case Mov.PieceF
Case 1: Whitepawn_Position(Mov.From) = 0
        Whitepawn_Position(Mov.ToMov) = 1
Case 2: WhiteKnight_Position(Mov.From) = 0
        WhiteKnight_Position(Mov.ToMov) = 1
Case 3: WhiteBishop_Position(Mov.From) = 0
        WhiteBishop_Position(Mov.ToMov) = 1
Case 4: WhiteRock_Position(Mov.From) = 0
        WhiteRock_Position(Mov.ToMov) = 1
Case 5: WhiteQueen_Position(Mov.From) = 0
        WhiteQueen_Position(Mov.ToMov) = 1
Case 6: WhiteKing_Position(Mov.From) = 0
        WhiteKing_Position(Mov.ToMov) = 1
End Select



ElseIf Mov.SideF = 0 Then

If Mov.SideT = 1 Then
BlackPieces_Position(Mov.From) = 0
BlackPieces_Position(Mov.ToMov) = 1
WhitePieces_Position(Mov.ToMov) = 0

Select Case Mov.PieceT
Case 1: Whitepawn_Position(Mov.ToMov) = 0
        
Case 2: WhiteKnight_Position(Mov.ToMov) = 0
        
Case 3: WhiteBishop_Position(Mov.ToMov) = 0
        
Case 4: WhiteRock_Position(Mov.ToMov) = 0
        
Case 5: WhiteQueen_Position(Mov.ToMov) = 0
        
Case 6: WhiteKing_Position(Mov.ToMov) = 0
        
End Select


Else
BlackPieces_Position(Mov.From) = 0
BlackPieces_Position(Mov.ToMov) = 1
End If
Select Case Mov.PieceF
Case 1: Blackpawn_Position(Mov.From) = 0
        Blackpawn_Position(Mov.ToMov) = 1
Case 2: BlackKnight_Position(Mov.From) = 0
        BlackKnight_Position(Mov.ToMov) = 1
Case 3: BlackBishop_Position(Mov.From) = 0
        BlackBishop_Position(Mov.ToMov) = 1
Case 4: BlackRock_Position(Mov.From) = 0
        BlackRock_Position(Mov.ToMov) = 1
Case 5: BlackQueen_Position(Mov.From) = 0
        BlackQueen_Position(Mov.ToMov) = 1
Case 6: BlackKing_Position(Mov.From) = 0
        BlackKing_Position(Mov.ToMov) = 1
End Select


End If

Call Gen_Movment
End Sub


Sub UnDo_move(Movment As MoveList)

' used to un make this move and take back the board to the postion before make that move
'CopyMemory Whitepawn_Position(0), Whitepawn_PositionBuffer(0), 64
'CopyMemory WhiteKing_Position(0), WhiteKing_PositionBuffer(0), 64
'CopyMemory WhiteBishop_Position(0), WhiteBishop_PositionBuffer(0), 64
'CopyMemory WhiteKnight_Position(0), WhiteKnight_PositionBuffer(0), 64
'CopyMemory WhiteRock_Position(0), WhiteRock_PositionBuffer(0), 64
'CopyMemory WhiteQueen_Position(0), WhiteQueen_PositionBuffer(0), 64
'CopyMemory WhitePieces_Position(0), WhitePieces_PositionBuffer(0), 64
'CopyMemory Blackpawn_Position(0), Blackpawn_PositionBuffer(0), 64
'CopyMemory BlackKing_Position(0), BlackKing_PositionBuffer(0), 64
'CopyMemory BlackBishop_Position(0), BlackBishop_PositionBuffer(0), 64
'CopyMemory BlackKnight_Position(0), BlackKnight_PositionBuffer(0), 64
'CopyMemory BlackRock_Position(0), BlackRock_PositionBuffer(0), 64
'CopyMemory BlackQueen_Position(0), BlackQueen_PositionBuffer(0), 64
'CopyMemory BlackPieces_Position(0), BlackPieces_PositionBuffer(0), 64
'CopyMemory Chess_Board(0), Chess_BoardBuffer(0), 64
'


If Movment.SideF = 1 Then
If Movment.SideT = 0 Then
WhitePieces_Position(Movment.From) = 1
WhitePieces_Position(Movment.ToMov) = 0
BlackPieces_Position(Movment.ToMov) = 1

Select Case Movment.PieceT
Case 1: Blackpawn_Position(Movment.ToMov) = 1
        
Case 2: BlackKnight_Position(Movment.ToMov) = 1
        
Case 3: BlackBishop_Position(Movment.ToMov) = 1
        
Case 4: BlackRock_Position(Movment.ToMov) = 1
        
Case 5: BlackQueen_Position(Movment.ToMov) = 1
        
Case 6: BlackKing_Position(Movment.ToMov) = 0
        
End Select


Else
WhitePieces_Position(Movment.From) = 1
WhitePieces_Position(Movment.ToMov) = 0
End If
Select Case Movment.PieceF
Case 1: Whitepawn_Position(Movment.From) = 1
        Whitepawn_Position(Movment.ToMov) = 0
Case 2: WhiteKnight_Position(Movment.From) = 1
        WhiteKnight_Position(Movment.ToMov) = 0
Case 3: WhiteBishop_Position(Movment.From) = 1
        WhiteBishop_Position(Movment.ToMov) = 0
Case 4: WhiteRock_Position(Movment.From) = 1
        WhiteRock_Position(Movment.ToMov) = 0
Case 5: WhiteQueen_Position(Movment.From) = 1
        WhiteQueen_Position(Movment.ToMov) = 0
Case 6: WhiteKing_Position(Movment.From) = 1
        WhiteKing_Position(Movment.ToMov) = 0
End Select



ElseIf Movment.SideF = 0 Then

If Movment.SideT = 1 Then
BlackPieces_Position(Movment.From) = 1
BlackPieces_Position(Movment.ToMov) = 0
WhitePieces_Position(Movment.ToMov) = 1

Select Case Movment.PieceT
Case 1: Whitepawn_Position(Movment.ToMov) = 1
        
Case 2: WhiteKnight_Position(Movment.ToMov) = 1
        
Case 3: WhiteBishop_Position(Movment.ToMov) = 1
        
Case 4: WhiteRock_Position(Movment.ToMov) = 1
        
Case 5: WhiteQueen_Position(Movment.ToMov) = 1
        
Case 6: WhiteKing_Position(Movment.ToMov) = 0
        
End Select


Else
BlackPieces_Position(Movment.From) = 1
BlackPieces_Position(Movment.ToMov) = 0
End If
Select Case Movment.PieceF
Case 1: Blackpawn_Position(Movment.From) = 1
        Blackpawn_Position(Movment.ToMov) = 0
Case 2: BlackKnight_Position(Movment.From) = 1
        BlackKnight_Position(Movment.ToMov) = 0
Case 3: BlackBishop_Position(Movment.From) = 1
        BlackBishop_Position(Movment.ToMov) = 0
Case 4: BlackRock_Position(Movment.From) = 1
        BlackRock_Position(Movment.ToMov) = 0
Case 5: BlackQueen_Position(Movment.From) = 1
        BlackQueen_Position(Movment.ToMov) = 0
Case 6: BlackKing_Position(Movment.From) = 1
        BlackKing_Position(Movment.ToMov) = 0
End Select


End If






Call Gen_Movment


End Sub


Sub StrToMov(StMov As String, Mov1 As MoveList)
' used to  convert move notation to "from , to "squeres
Dim pf1, pt1 As Integer

Dim x As Integer
Dim st, frmstr, tostr, pfstr, ptstr As String
endstr = Len(StMov)
For i = 1 To endstr
st = Mid(StMov, i, 1)
If st <> "-" Then
frmstr = frmstr & st

Else
x = i + 1
Exit For
End If
Next i
For j = x To endstr
st = Mid(StMov, j, 1)
tostr = tostr & st
Next j
pfstr = Mid(frmstr, 1, 1)

frmstr = Mid(frmstr, 2, 2)
If Len(tostr) < 3 Then
tostr = tostr
ptstr = "/"
sideto1 = -1
Else
ptstr = Mid(tostr, 1, 1)
tostr = Mid(tostr, 2, 2)
End If


Select Case frmstr

Case "a8": frm = 0
Case "b8": frm = 1
Case "c8": frm = 2
Case "d8": frm = 3
Case "e8": frm = 4
Case "f8": frm = 5
Case "g8": frm = 6
Case "h8": frm = 7

Case "a7": frm = 8
Case "b7": frm = 9
Case "c7": frm = 10
Case "d7": frm = 11
Case "e7": frm = 12
Case "f7": frm = 13
Case "g7": frm = 14
Case "h7": frm = 15

Case "a6": frm = 16
Case "b6": frm = 17
Case "c6": frm = 18
Case "d6": frm = 19
Case "e6": frm = 20
Case "f6": frm = 21
Case "g6": frm = 22
Case "h6": frm = 23

Case "a5": frm = 24
Case "b5": frm = 25
Case "c5": frm = 26
Case "d5": frm = 27
Case "e5": frm = 28
Case "f5": frm = 29
Case "g5": frm = 30
Case "h5": frm = 31

Case "a4": frm = 32
Case "b4": frm = 33
Case "c4": frm = 34
Case "d4": frm = 35
Case "e4": frm = 36
Case "f4": frm = 37
Case "g4": frm = 38
Case "h4": frm = 39

Case "a3": frm = 40
Case "b3": frm = 41
Case "c3": frm = 42
Case "d3": frm = 43
Case "e3": frm = 44
Case "f3": frm = 45
Case "g3": frm = 46
Case "h3": frm = 47

Case "a2": frm = 48
Case "b2": frm = 49
Case "c2": frm = 50
Case "d2": frm = 51
Case "e2": frm = 52
Case "f2": frm = 53
Case "g2": frm = 54
Case "h2": frm = 55

Case "a1": frm = 56
Case "b1": frm = 57
Case "c1": frm = 58
Case "d1": frm = 59
Case "e1": frm = 60
Case "f1": frm = 61
Case "g1": frm = 62
Case "h1": frm = 63
End Select





Select Case tostr

Case "a8": Tol = 0
Case "b8": Tol = 1
Case "c8": Tol = 2
Case "d8": Tol = 3
Case "e8": Tol = 4
Case "f8": Tol = 5
Case "g8": Tol = 6
Case "h8": Tol = 7

Case "a7": Tol = 8
Case "b7": Tol = 9
Case "c7": Tol = 10
Case "d7": Tol = 11
Case "e7": Tol = 12
Case "f7": Tol = 13
Case "g7": Tol = 14
Case "h7": Tol = 15

Case "a6": Tol = 16
Case "b6": Tol = 17
Case "c6": Tol = 18
Case "d6": Tol = 19
Case "e6": Tol = 20
Case "f6": Tol = 21
Case "g6": Tol = 22
Case "h6": Tol = 23

Case "a5": Tol = 24
Case "b5": Tol = 25
Case "c5": Tol = 26
Case "d5": Tol = 27
Case "e5": Tol = 28
Case "f5": Tol = 29
Case "g5": Tol = 30
Case "h5": Tol = 31

Case "a4": Tol = 32
Case "b4": Tol = 33
Case "c4": Tol = 34
Case "d4": Tol = 35
Case "e4": Tol = 36
Case "f4": Tol = 37
Case "g4": Tol = 38
Case "h4": Tol = 39

Case "a3": Tol = 40
Case "b3": Tol = 41
Case "c3": Tol = 42
Case "d3": Tol = 43
Case "e3": Tol = 44
Case "f3": Tol = 45
Case "g3": Tol = 46
Case "h3": Tol = 47

Case "a2": Tol = 48
Case "b2": Tol = 49
Case "c2": Tol = 50
Case "d2": Tol = 51
Case "e2": Tol = 52
Case "f2": Tol = 53
Case "g2": Tol = 54
Case "h2": Tol = 55

Case "a1": Tol = 56
Case "b1": Tol = 57
Case "c1": Tol = 58
Case "d1": Tol = 59
Case "e1": Tol = 60
Case "f1": Tol = 61
Case "g1": Tol = 62
Case "h1": Tol = 63
End Select

Select Case pfstr

Case "P": pf1 = 1
        
Case "N": pf1 = 2

Case "B": pf1 = 3
        
Case "R": pf1 = 4

Case "Q": pf1 = 5
        
Case "K": pf1 = 6
Case "/": pf1 = 0
End Select



Select Case ptstr

Case "P": pt1 = 1
        
Case "N": pt1 = 2

Case "B": pt1 = 3
        
Case "R": pt1 = 4

Case "Q": pt1 = 5
        
Case "K": pt1 = 6
Case "/": pt1 = 0
End Select

Mov1.From = frm
Mov1.ToMov = Tol
Mov1.PieceF = pf1
Mov1.PieceT = pt1
Mov1.SideT = sideto1

End Sub



Sub Computer_Move(CompMOve As MoveList)
' ot update the board and make the computer move
If CompMOve.SideF = 1 Then
If CompMOve.SideT = 0 Then
WhitePieces_Position(CompMOve.From) = 0
WhitePieces_Position(CompMOve.ToMov) = 1
BlackPieces_Position(CompMOve.ToMov) = 0

Select Case CompMOve.PieceT
Case 1: Blackpawn_Position(CompMOve.ToMov) = 0
        
Case 2: BlackKnight_Position(CompMOve.ToMov) = 0
        
Case 3: BlackBishop_Position(CompMOve.ToMov) = 0
        
Case 4: BlackRock_Position(CompMOve.ToMov) = 0
        
Case 5: BlackQueen_Position(CompMOve.ToMov) = 0
        
'Case 6: BlackKing_Position(compmove.tomov) = 0
        
End Select


Else
WhitePieces_Position(CompMOve.From) = 0
WhitePieces_Position(CompMOve.ToMov) = 1
End If
Select Case CompMOve.PieceF
Case 1: Whitepawn_Position(CompMOve.From) = 0
        Whitepawn_Position(CompMOve.ToMov) = 1
Case 2: WhiteKnight_Position(CompMOve.From) = 0
        WhiteKnight_Position(CompMOve.ToMov) = 1
Case 3: WhiteBishop_Position(CompMOve.From) = 0
        WhiteBishop_Position(CompMOve.ToMov) = 1
Case 4: WhiteRock_Position(CompMOve.From) = 0
        WhiteRock_Position(CompMOve.ToMov) = 1
Case 5: WhiteQueen_Position(CompMOve.From) = 0
        WhiteQueen_Position(CompMOve.ToMov) = 1
Case 6: WhiteKing_Position(CompMOve.From) = 0
        WhiteKing_Position(CompMOve.ToMov) = 1
End Select



ElseIf CompMOve.SideF = 0 Then

If CompMOve.SideT = 1 Then
BlackPieces_Position(CompMOve.From) = 0
BlackPieces_Position(CompMOve.ToMov) = 1
WhitePieces_Position(CompMOve.ToMov) = 0

Select Case CompMOve.PieceT
Case 1: Whitepawn_Position(CompMOve.ToMov) = 0
        
Case 2: WhiteKnight_Position(CompMOve.ToMov) = 0
        
Case 3: WhiteBishop_Position(CompMOve.ToMov) = 0
        
Case 4: WhiteRock_Position(CompMOve.ToMov) = 0
        
Case 5: WhiteQueen_Position(CompMOve.ToMov) = 0
        
'Case 6: whiteKing_Position(compmove.tomov) = 0
        
End Select


Else
BlackPieces_Position(CompMOve.From) = 0
BlackPieces_Position(CompMOve.ToMov) = 1
End If
Select Case CompMOve.PieceF
Case 1: Blackpawn_Position(CompMOve.From) = 0
        Blackpawn_Position(CompMOve.ToMov) = 1
Case 2: BlackKnight_Position(CompMOve.From) = 0
        BlackKnight_Position(CompMOve.ToMov) = 1
Case 3: BlackBishop_Position(CompMOve.From) = 0
        BlackBishop_Position(CompMOve.ToMov) = 1
Case 4: BlackRock_Position(CompMOve.From) = 0
        BlackRock_Position(CompMOve.ToMov) = 1
Case 5: BlackQueen_Position(CompMOve.From) = 0
        BlackQueen_Position(CompMOve.ToMov) = 1
Case 6: BlackKing_Position(CompMOve.From) = 0
        BlackKing_Position(CompMOve.ToMov) = 1
End Select


End If

End Sub
Sub SortList(movelst() As MoveList)

' this is needed to order the move to check them first
Dim temp As MoveList
For i = 1 To UBound(movelst)
For j = 1 To UBound(movelst)
If movelst(j).Attackvalue < movelst(i).Attackvalue Then









temp.From = movelst(i).From
temp.ToMov = movelst(i).ToMov
temp.Location = movelst(i).Location
temp.PieceF = movelst(i).PieceF
temp.PieceT = movelst(i).PieceT
temp.SideF = movelst(i).SideF
temp.SideT = movelst(i).SideT
temp.Attackvalue = movelst(i).Attackvalue

movelst(i).From = movelst(j).From
movelst(i).ToMov = movelst(j).ToMov
movelst(i).Location = movelst(j).Location
movelst(i).PieceF = movelst(j).PieceF
movelst(i).PieceT = movelst(j).PieceT
movelst(i).SideF = movelst(j).SideF
movelst(i).SideT = movelst(j).SideT
movelst(i).Attackvalue = movelst(j).Attackvalue

movelst(j).From = temp.From
movelst(j).ToMov = temp.ToMov
movelst(j).Location = temp.Location
movelst(j).PieceF = temp.PieceF
movelst(j).PieceT = temp.PieceT
movelst(j).SideF = temp.SideF
movelst(j).SideT = temp.SideT
movelst(j).Attackvalue = temp.Attackvalue



End If

Next j

Next i
End Sub
