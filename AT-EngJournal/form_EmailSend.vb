Public Class form_EmailSend

    Public Item As Outlook.MailItem

    Private Sub button_record_Click(sender As Object, e As EventArgs) Handles button_record.Click
        'begin the structure for the send function
        Dim proj As String = TextBox1.Text

        'copy the message to the folder
        '   find the projectfolder in the aslstore
        Dim fld As Outlook.Folder = ASL_Tools.Get_ProjectFolder_In_ASLStoreInbox(proj)

        'if the project is not found then create it.
        If IsNothing(fld) Then
            'MsgBox("No Project Found", vbCritical, "Error")
            'create the project folder
            fld = ASL_Tools.Create_ProjectFolder_In_ASLStoreInbox(proj)
            If IsNothing(fld) Then
                MsgBox("Error getting or creating project", vbCritical, "Error")
                Exit Sub
            End If
        End If

        'with the project folder
        'get the sent folder
        Dim fldSent As Outlook.Folder = ASL_Tools.Get_ProjectFolderSent(fld)

        If IsNothing(fldSent) Then
            fldSent = ASL_Tools.Create_ProjectFolderSent(fld)
            If IsNothing(fldSent) Then
                MsgBox("Error creating sent folder", vbCritical, "Error")
                Exit Sub
            End If
        End If

        'if the project exists then store the message information to the server
        Dim itemCopy As Outlook.MailItem = Item.Copy
        Dim cD As Date = New Date.Now
        Dim uS As String = ASL_Tools.aslStore.DisplayName
        itemCopy.Subject = "(" & Format(cD, "yyyy-MM-dd:HHmmss") & " " & uS & ") " & itemCopy.Subject
        itemCopy.Move(fldSent)

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