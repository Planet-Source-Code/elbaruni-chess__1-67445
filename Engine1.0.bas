Attribute VB_Name = "Engine"
' mins-plus infinity and max-min values used in Alpha beta tree search
Private Const MinsInfinity = -1000000000
Private Const PlusInfinity = 1000000000
Public Const MAXVALUE = 1000000000 'greater than any possible board value
    Public Const MINVALUE = -1000000000
Public StrMove As String
Public MoveToDo As MoveList ' this the move will the computer select to move
Function Max1(a1, a2) As Long
If a1 > a2 Then
Max1 = a1
Else
Max1 = a2
End If

End Function


'ab2 alphabetasearch are min-max(alpha beta search tree)tree
Function AB2(ByVal depth As Integer, ByVal Side As Integer, ByVal Alpha As Long, ByVal Beta As Long, root1 As Integer) As Long


Dim BestScore, value, v As Long
Dim Move_l(100) As MoveList
Dim Mobility As Integer
Dim Oppenent As Integer



' side is the current player to move 1 for white 0 for black
'depth is number of plies to search half-move

If Side = 1 Then
Oppenent = 0

ElseIf Side = 0 Then
Oppenent = 1
End If
If depth = 0 Then
AB2 = EVALUATION
Else
Call Generate_Moves(Side, Mobility, Move_l)
For i = 1 To Mobility

Call do_Move(Move_l(i))
     value = -1 * AB2(depth - 1, Oppenent, -1 * Beta, -1 * Alpha, 0)
     Call UnDo_move(Move_l(i))
     If value >= Beta Then
     AB2 = Beta
    If root1 = 1 Then
     
         If Side = CurnTurn Then
          MoveStr = Move_l(i).Location
            
             MoveToDo.From = Move_l(i).From
             MoveToDo.ToMov = Move_l(i).ToMov
             MoveToDo.PieceF = Move_l(i).PieceF
             MoveToDo.PieceT = Move_l(i).PieceT
             MoveToDo.SideF = Move_l(i).SideF
             MoveToDo.SideT = Move_l(i).SideT
             StrMove = MoveStr
             End If
     End If
     End If
     
     If value > Alpha Then
     Alpha = value
     If root1 = 1 Then
     
         If Side = CurnTurn Then
          MoveStr = Move_l(i).Location
             
             MoveToDo.From = Move_l(i).From
             MoveToDo.ToMov = Move_l(i).ToMov
             MoveToDo.PieceF = Move_l(i).PieceF
             MoveToDo.PieceT = Move_l(i).PieceT
             MoveToDo.SideF = Move_l(i).SideF
             MoveToDo.SideT = Move_l(i).SideT
             StrMove = MoveStr
             End If
     End If
     End If
     
     If Alpha >= Beta Then Exit For
Next i
AB2 = Alpha
End If
End Function












 Public Function AlphaBetaSearch(ByVal depth As Integer, ByVal Side As Integer, ByVal Limit As Long, ByVal root1 As Integer) As Long
        'returns the value of the board
        'aborts if the branch it is searching is worse than the limit
        Dim BestScore, v, i, BoardValue As Long
        Dim Move_l(100) As MoveList
        Dim Mobility As Integer
        Dim Oppenent As Integer
        'If AllDone Then Exit Function

        If depth <> 0 Then 'not at bottom of recursive branch yet...

            If Side = 0 Then

                Oppenent = 1
                BestScore = MAXVALUE
            Else
                BestScore = MINVALUE

                Oppenent = 0
            End If
            'check all possible moves from this location...
            Call Generate_Moves(Side, Mobility, Move_l)
            'calculate the new boardvalue
            If Mobility = 0 Then
                AlphaBetaSearch = EVALUATION()
            End If
            For i = 1 To Mobility
               Call do_Move(Move_l(i))
                            
                    v = AlphaBetaSearch(depth - 1, Oppenent, BestScore, 0)  'and recursively follow branch
                
                Call UnDo_move(Move_l(i))
                If v = BestScore Then 'insert a bit of randomness
                    If Rnd() < 0.3 Then
                        If root1 = 1 Then
                           If Side = CurnTurn Then
          MoveStr = Move_l(i).Location
            
             MoveToDo.From = Move_l(i).From
             MoveToDo.ToMov = Move_l(i).ToMov
             MoveToDo.PieceF = Move_l(i).PieceF
             MoveToDo.PieceT = Move_l(i).PieceT
             MoveToDo.SideF = Move_l(i).SideF
             MoveToDo.SideT = Move_l(i).SideT
             StrMove = MoveStr
             End If
                        End If
                    End If
                End If
                If Side = 1 Then ' Pick largest
                    If v > BestScore Then
                        BestScore = v
                        If root1 = 1 Then
                            If Side = CurnTurn Then
          MoveStr = Move_l(i).Location
            
             MoveToDo.From = Move_l(i).From
             MoveToDo.ToMov = Move_l(i).ToMov
             MoveToDo.PieceF = Move_l(i).PieceF
             MoveToDo.PieceT = Move_l(i).PieceT
             MoveToDo.SideF = Move_l(i).SideF
             MoveToDo.SideT = Move_l(i).SideT
             StrMove = MoveStr
             End If
                        End If

                    End If
                    If v > Limit Then 'we are past the limit, meaning there is no point exploring this branch further
                        Exit For
                    End If
                Else 'Pick smallest
                    If v < BestScore Then
                        BestScore = v
                        If root1 = 1 Then
                           If Side = CurnTurn Then
          MoveStr = Move_l(i).Location
            
             MoveToDo.From = Move_l(i).From
             MoveToDo.ToMov = Move_l(i).ToMov
             MoveToDo.PieceF = Move_l(i).PieceF
             MoveToDo.PieceT = Move_l(i).PieceT
             MoveToDo.SideF = Move_l(i).SideF
             MoveToDo.SideT = Move_l(i).SideT
             StrMove = MoveStr
             End If
                        End If
                    End If
                    If v < Limit Then
                        Exit For
                    End If
                End If

            Next
            AlphaBetaSearch = BestScore
        Else  'we are at the bottom of the branch. Just return the board value

            AlphaBetaSearch = EVALUATION()
        End If
    End Function
    

    


