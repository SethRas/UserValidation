Public Class UserValidation

    Sub validateInput()
        Dim UserAge As Integer
        'If AgeTextBox3.Text = "" Then
        '    accumulateMessage("AGE is required")
        '    AgeTextBox3.Focus()
        'End If

        Try
            UserAge = CInt(AgeTextBox3.Text)
        Catch ex As Exception
            accumulateMessage("AGE must be a valid entry")
            AgeTextBox3.Focus()
        End Try

        If LastNameTextBox2.Text = "" Then
            accumulateMessage("LAST name is required")
            LastNameTextBox2.Focus()
        End If

        If FirstNAmeTextBox1.Text = "" Then
            accumulateMessage("FIRST name is required")
            FirstNAmeTextBox1.Focus()
        End If

        If accumulateMessage() <> "" Then
            MsgBox(accumulateMessage())
            accumulateMessage(, True)
        End If

    End Sub

    Private Function accumulateMessage(Optional newMessage As String = "", Optional clear As Boolean = False) As String
        Static _message As String
        Select Case clear
            Case False
                If newMessage <> "" Then
                    _message &= newMessage & vbCrLf
                End If

            Case Else
                _message = ""
        End Select

        Return _message
    End Function

    Private Sub Reset()
        'reset All form controls to default
        FirstNAmeTextBox1.Text = ""
        LastNameTextBox2.Text = ""
        AgeTextBox3.Text = ""
        'Clear any stored messages
        accumulateMessage(, True)


    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles FirstNAmeLabel1.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles LastNameLabel2.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles AgeLabel3.Click

    End Sub

    Private Sub Label1_Click_1(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles LastNameTextBox2.TextChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        validateInput()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Submit_Click(sender As Object, e As EventArgs) Handles Submit.Click
        Reset()
    End Sub
End Class
