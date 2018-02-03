Public Class form_setDiscipline
    Private Sub form_setDiscipline_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim tempdisc = ASL_Tools.get_discipline()

        If Not (tempdisc = "") Then
            combobox_discipline.Text = tempdisc
        End If
    End Sub

    Private Sub combobox_discipline_SelectedIndexChanged(sender As Object, e As EventArgs) Handles combobox_discipline.SelectedIndexChanged
        If Not (combobox_discipline.Text = "") Then
            button_save.Enabled = True
        Else
            button_save.Enabled = False
        End If
    End Sub

    Private Sub button_cancel_Click(sender As Object, e As EventArgs) Handles button_cancel.Click
        Me.Close()
    End Sub

    Private Sub button_save_Click(sender As Object, e As EventArgs) Handles button_save.Click
        'store the discipline value in the registry
        If Not (combobox_discipline.Text = "") Then
            ASL_Tools.enable_discipline(combobox_discipline.Text)
        Else
            MsgBox("Please select a discipline.", vbCritical, "Error")
            Exit Sub
        End If

        Me.Close()
    End Sub
End Class