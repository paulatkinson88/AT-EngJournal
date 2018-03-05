
Public Class form_EmailSend
    Public proj As String = ""

    Private Sub button_record_Click(sender As Object, e As EventArgs) Handles button_record.Click
        proj = TextBox1.Text

        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'do not enable the record button unless the project number selected is 4 digits and numeric.
        If TextBox1.Text.Length = 4 And IsNumeric(TextBox1.Text) Then
            button_record.Enabled = True
        Else
            button_record.Enabled = False
        End If
    End Sub

    Private Sub button_skip_Click(sender As Object, e As EventArgs) Handles button_skip.Click
        Me.Close()
    End Sub
End Class