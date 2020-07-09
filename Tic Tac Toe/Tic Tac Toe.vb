Public Class Form1

    'define global variables
    Dim playerTurn As String = "Player1"
    Dim numberOfTurns As Integer = 0

    Private Sub square1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles square1.Click
        square1.Enabled = False
        If playerTurn = "Player1" Then
            square1.BackColor = Color.Red
            player1Turn()
        ElseIf playerTurn = "Player2" Then
            square1.BackColor = Color.Blue
            player2Turn()
        End If

    End Sub

    Private Sub square2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles square2.Click
        square2.Enabled = False
        If playerTurn = "Player1" Then
            square2.BackColor = Color.Red
            player1Turn()
        ElseIf playerTurn = "Player2" Then
            square2.BackColor = Color.Blue
            player2Turn()
        End If

    End Sub

    Private Sub square3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles square3.Click
        square3.Enabled = False
        If playerTurn = "Player1" Then
            square3.BackColor = Color.Red
            player1Turn()
        ElseIf playerTurn = "Player2" Then
            square3.BackColor = Color.Blue
            player2Turn()
        End If

    End Sub

    Private Sub square4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles square4.Click
        square4.Enabled = False
        If playerTurn = "Player1" Then
            square4.BackColor = Color.Red
            player1Turn()
        ElseIf playerTurn = "Player2" Then
            square4.BackColor = Color.Blue
            player2Turn()
        End If

    End Sub

    Private Sub square5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles square5.Click
        square5.Enabled = False
        If playerTurn = "Player1" Then
            square5.BackColor = Color.Red
            player1Turn()
        ElseIf playerTurn = "Player2" Then
            square5.BackColor = Color.Blue
            player2Turn()
        End If

    End Sub

    Private Sub square6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles square6.Click
        square6.Enabled = False
        If playerTurn = "Player1" Then
            square6.BackColor = Color.Red
            player1Turn()
        ElseIf playerTurn = "Player2" Then
            square6.BackColor = Color.Blue
            player2Turn()
        End If

    End Sub

    Private Sub square7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles square7.Click
        square7.Enabled = False
        If playerTurn = "Player1" Then
            square7.BackColor = Color.Red
            player1Turn()
        ElseIf playerTurn = "Player2" Then
            square7.BackColor = Color.Blue
            player2Turn()
        End If

    End Sub

    Private Sub square8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles square8.Click
        square8.Enabled = False
        If playerTurn = "Player1" Then
            square8.BackColor = Color.Red
            player1Turn()
        ElseIf playerTurn = "Player2" Then
            square8.BackColor = Color.Blue
            player2Turn()
        End If

    End Sub


    Private Sub square9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles square9.Click
        square9.Enabled = False
        If playerTurn = "Player1" Then
            square9.BackColor = Color.Red
            player1Turn()
        ElseIf playerTurn = "Player2" Then
            square9.BackColor = Color.Blue
            player2Turn()
        End If

    End Sub

    Private Function isGameOver() As Boolean
        'There are 8 ways to get three in a row in Tic Tac Toe:
        '3 in a row in any row
        '3 in a row in any column
        '3 in a row diagonally
        'if all squares have been selected, yet there is no three in a row, then the game is a draw

        'Scenario 1: if the first row has three in a row (squares 1, 2, 3)
        If square1.BackColor = Color.Red And square2.BackColor = Color.Red And square3.BackColor = Color.Red Then
            player1Wins()
            Return True
        ElseIf square1.BackColor = Color.Blue And square2.BackColor = Color.Blue And square3.BackColor = Color.Blue Then
            player2Wins()
            Return True
        End If

        'Scenario 2: if the second row has three in a row (squares 4, 5, 6)
        If square4.BackColor = Color.Red And square5.BackColor = Color.Red And square6.BackColor = Color.Red Then
            player1Wins()
            Return True
        ElseIf square4.BackColor = Color.Blue And square5.BackColor = Color.Blue And square6.BackColor = Color.Blue Then
            player2Wins()
            Return True
        End If

        'Scenario 3: if third row has three in a row (squares 7, 8, 9)
        If square7.BackColor = Color.Red And square8.BackColor = Color.Red And square9.BackColor = Color.Red Then
            player1Wins()
            Return True
        ElseIf square7.BackColor = Color.Blue And square8.BackColor = Color.Blue And square9.BackColor = Color.Blue Then
            player2Wins()
            Return True
        End If

        'Scenario 4: if the first column has three in a row (squares 1, 4, 7)
        If square1.BackColor = Color.Red And square4.BackColor = Color.Red And square7.BackColor = Color.Red Then
            player1Wins()
            Return True
        ElseIf square1.BackColor = Color.Blue And square4.BackColor = Color.Blue And square7.BackColor = Color.Blue Then
            player2Wins()
            Return True
        End If

        'Scenario 5: if the second column has three in a row (squares 2, 5, 8)
        If square2.BackColor = Color.Red And square5.BackColor = Color.Red And square8.BackColor = Color.Red Then
            player1Wins()
            Return True
        ElseIf square2.BackColor = Color.Blue And square5.BackColor = Color.Blue And square8.BackColor = Color.Blue Then
            player2Wins()
            Return True
        End If

        'Scenario 6: if the third column has three in a row (squares 3, 6, 9)
        If square3.BackColor = Color.Red And square6.BackColor = Color.Red And square9.BackColor = Color.Red Then
            player1Wins()
            Return True
        ElseIf square3.BackColor = Color.Blue And square6.BackColor = Color.Blue And square9.BackColor = Color.Blue Then
            player2Wins()
            Return True
        End If

        'Scenario 7: if diagonally three in a row (squares 1, 5, 9)
        If square1.BackColor = Color.Red And square5.BackColor = Color.Red And square9.BackColor = Color.Red Then
            player1Wins()
            Return True
        ElseIf square1.BackColor = Color.Blue And square5.BackColor = Color.Blue And square9.BackColor = Color.Blue Then
            player2Wins()
            Return True
        End If

        'Scenario 8: if diagonally three in a row (squares 7, 5 ,3)
        If square7.BackColor = Color.Red And square5.BackColor = Color.Red And square3.BackColor = Color.Red Then
            player1Wins()
            Return True
        ElseIf square7.BackColor = Color.Blue And square5.BackColor = Color.Blue And square3.BackColor = Color.Blue Then
            player2Wins()
            Return True
        End If

        'Draw scenario: if all squares have been seleccted, yet no winner then it is a draw
        If numberOfTurns = 9 Then
            gameIsDraw()
            Return True
        End If


        Return False

    End Function

    Private Sub player1Turn()
        'call this code at game start and every time
        'player 1 takes a turn

        'switch the labels
        Player1Label.Visible = False
        Player2Label.Visible = True

        'transfer control to player 2
        playerTurn = "Player2"

        'increment turn count
        numberOfTurns += 1

        'see if game over or winner
        isGameOver()

    End Sub
    Private Sub player2Turn()
        'call this code every time player 2 takes a turn

        'switch the labels
        Player2Label.Visible = False
        Player1Label.Visible = True

        'transfer control to player 1
        playerTurn = "Player1"

        'increment turn count
        numberOfTurns += 1

        'see if game over or winner
        isGameOver()

    End Sub
    Private Sub player1Wins()
        disableButtons()
        MessageBox.Show("Player 1 wins!")
        Player1Label.Visible = True
        Player2Label.Visible = False
        Player1Label.Text = "Winner!"
        btnReset.Text = "Play Again"
    End Sub

    Private Sub player2Wins()
        disableButtons()
        MessageBox.Show("Player 2 wins!")
        Player2Label.Visible = True
        Player1Label.Visible = False
        Player2Label.Text = "Winner!"
        btnReset.Text = "Play Again"
    End Sub

    Private Sub gameIsDraw()
        MessageBox.Show("Game is a draw!")
        Player1Label.Visible = False
        Player2Label.Visible = False
        drawLabel.Visible = True
        btnReset.Text = "Play Again"
    End Sub

    Private Sub btnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuit.Click
        'quit the game and close the form
        Me.Close()
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        'reset all aspects of the game

        'reset all the squares to original state and enable them
        square1.BackColor = Color.White
        square1.Enabled = True

        square2.BackColor = Color.White
        square2.Enabled = True

        square3.BackColor = Color.White
        square3.Enabled = True

        square4.BackColor = Color.White
        square4.Enabled = True

        square5.BackColor = Color.White
        square5.Enabled = True

        square6.BackColor = Color.White
        square6.Enabled = True

        square7.BackColor = Color.White
        square7.Enabled = True

        square8.BackColor = Color.White
        square8.Enabled = True

        square9.BackColor = Color.White
        square9.Enabled = True

        'reset the turn back to player 1 and reset labels
        playerTurn = "Player1"
        Player1Label.Text = "Player 1"
        Player2Label.Text = "Player 2"
        Player1Label.Visible = True
        Player2Label.Visible = False
        drawLabel.Visible = False

        'reset turn count
        numberOfTurns = 0


        'change reset button text
        btnReset.Text = "Reset Game"

    End Sub

    Private Sub disableButtons()
        'Disables all the squares to end game.  
        'Forces user to reset game.
        square1.Enabled = False
        square2.Enabled = False
        square3.Enabled = False
        square4.Enabled = False
        square5.Enabled = False
        square6.Enabled = False
        square7.Enabled = False
        square8.Enabled = False
        square9.Enabled = False
    End Sub
End Class
