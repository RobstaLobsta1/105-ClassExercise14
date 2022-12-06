Public Class Form1
    Dim gameNumber As Integer = 1
    Dim person As New Human
    Dim machine As New Computer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim humanChoice As String
        Dim computerChoice As String
        Dim result As String
        humanChoice = person.Choice()
        computerChoice = machine.Choice()

        Select Case humanChoice
            Case "Rock"
                Select Case computerChoice
                    Case "Paper"
                        machine.Score += 1
                        result = "COMPUTER WINS"
                    Case "Scissors"
                        person.Score += 1
                        result = "YOU WIN"
                    Case "Rock"
                        result = "TIE"
                End Select
            Case "Paper"
                Select Case computerChoice
                    Case "Rock"
                        person.Score += 1
                        result = "YOU WIN"
                    Case "Scissors"
                        machine.Score += 1
                        result = "COMPUTER WINS"
                    Case "Paper"
                        result = "TIE"
                End Select
            Case "Scissors"
                Select Case computerChoice
                    Case "Rock"
                        machine.Score += 1
                        result = "COMPUTER WINS"
                    Case "Paper"
                        person.Score += 1
                        result = "YOU WIN"
                    Case "Scissors"
                        result = "TIE"
                End Select
        End Select

        MessageBox.Show(person.name & "'s CHOICE : " & humanChoice & machine.name & "'s CHOICE : " & computerChoice & " " & result, "RESULT")
        TextBox1.Text = CStr(person.Score)
        TextBox2.Text = CStr(machine.Score)
        gameNumber += 1
        Button1.Text = "Play Game #" & (gameNumber)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        person.name = "Robby"
        machine.name = "Winston"
        Label1.Text = person.name & "'s Score :"
        Label2.Text = machine.name & "'s Score :"
    End Sub
End Class

MustInherit Class Contestant
    Public Property Score As Integer
    Public name As String
    MustOverride Function Choice() As String
End Class 'Contesant

Class Human
    Inherits Contestant
    Overrides Function Choice() As String
        Return InputBox("Enter your choice (Rock, Paper, or Scissors) : ")
    End Function
End Class 'Human

Class Computer
    Inherits Contestant
    Overrides Function Choice() As String
        Dim choices() As String = {"Rock", "Paper", "Scissors"}
        Dim randomNumber As New Random
        Return choices(randomNumber.Next(0, 3))
    End Function
End Class 'Computer
