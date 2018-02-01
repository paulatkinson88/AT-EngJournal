Public Class form_EmailRecord
    Public emList As List(Of Outlook.MailItem) = New List(Of Outlook.MailItem)

    Private Sub button_skip_Click(sender As Object, e As EventArgs) Handles button_skip.Click
        Me.Close()
    End Sub

    Private Sub button_record_Click(sender As Object, e As EventArgs) Handles button_record.Click
        'begin the structure for the send function
        Dim proj As String = TextBox1.Text

        'copy the message to the folder
        '   find the projectfolder in the aslstore
        Dim fld As Outlook.Folder = ASL_Tools.Get_ProjectFolder_In_ASLStoreInbox(proj)

        'if the project is not found then 
        'ask the user if they wish to create it or re enter a new
        'project number
        If IsNothing(fld) Then
            Dim rsp = MsgBox("Project folder not found in your email inbox." & vbLf & "Yes to Create Project folder" & vbLf & "No to Re-Enter Project Number", vbYesNo + vbCritical, "Error")
            If rsp = vbNo Then
                Exit Sub
            End If
            'MsgBox("No Project Found", vbCritical, "Error")
            'create the project folder
            fld = ASL_Tools.Create_ProjectFolder_In_ASLStoreInbox(proj)
            If IsNothing(fld) Then
                MsgBox("Error getting or creating project", vbCritical, "Error")
                Exit Sub
            End If
        End If

        'if the project exists then store the message information to the server
        'for each message move it to the project folder and create a copy on the server.

        For Each itm As Outlook.MailItem In emList
            Dim uS As String = ASL_Tools.aslStore.DisplayName

            Dim cD As Date = itm.ReceivedTime
            Dim st As String = "(" & proj & ")(" & Format(cD, "yyyy-MM-dd-HHmmss") & ")(re)"
            'get the email recieved date store this messageKeyValue in a user property in the message
            ASL_Tools.Set_StampProperty(itm, st)

            'check to see if the user is in the office.
            'if they are then save a copy of the email to the project folder
            'if not then flag the message with the category offline
            'offline messages can get copied to network at a later date.
            If ASL_Tools.networkReady = True Then
                'copy to network
                Dim di As System.IO.DirectoryInfo = ASL_Tools.Check_For_ProjectDirectoryEngJournal(proj)
                If IsNothing(di) Then
                    itm.Categories = "Offline"
                Else
                    'use the messageKeyValue as the message name when saving to the 
                    'network
                    itm.SaveAs(di.FullName & "\" & st & ".msg")
                End If
            Else
                itm.Categories = "Offline"
            End If

            itm.Move(fld)
        Next

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

    Private Sub form_EmailRecord_Shown(sender As Object, e As EventArgs) Handles Me.Shown

    End Sub

    Private Sub form_EmailRecord_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class