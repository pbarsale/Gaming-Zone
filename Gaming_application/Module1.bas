Attribute VB_Name = "Module1"
Option Explicit
Dim inform As String

Function ticobject() As String
inform = "     The tic tac toe game requires two players x and 0 , who take turns marking the spaces in a 3×3 grid. The objective of the game is placing three respective marks in a horizontal,vertical, or diagonal row."
ticobject = inform
End Function
Function ticplaying() As String
inform = "      Playing this Game is a easy task. The two players are alloted the icons 'X' and '0' respectively . The player with 'X' icon is the first one to play. When first player marks,the turn goes to second player with '0'. This is how,the two players play till one of them wins. The player who succeeds in placing three respective marks in a horizontal, vertical, or diagonal row wins the game. "
ticplaying = inform
End Function

Function ticscoring() As String
inform = "      Scoring is an important part in any game. The winner of the game will get 20 points. These points are added to his previous points which are stored in the database  "
ticscoring = inform
End Function
Function puzobject() As String
inform = "      The puzzle game requires only one player. A sliding tile puzzle is a small handheld game, in which the objective is to slide a number of pieces into a certain order or configuration. Usually, each tile represents a number in a square, where the player must un-jumble the tiles to place them in numerical order."
puzobject = inform
End Function
Function puzplaying() As String
inform = "    The '15 Puzzle' consists of 15 squares numbered from 1 to 15 which are placed in a 4x4 box leaving one position out of the 16 empty. The goal is to reposition the squares from a given arbitrary starting arrangement by sliding them one at a time into the configuration as shown below To play you first scramble the tiles and then try to put them back in order. To move a tile you simply click on it. The only tiles you can move are those adjacent to the hole."
puzplaying = inform
End Function



