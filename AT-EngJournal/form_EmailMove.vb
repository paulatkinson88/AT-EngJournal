Public Class form_emailMove
    Public proj As String = ""

    Private Sub button_skip_Click(sender As Object, e As EventArgs) Handles button_skip.Click
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

    Private Sub form_emailMove_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub button_record_Click(sender As Object, e As EventArgs) Handles button_record.Click
        'begin the structure for the send function
        proj = TextBox1.Text
        'Dim fldSent As Outlook.Folder = Nothing

        'copy the message to the folder
        '   find the projectfolder in the aslstore
        'Dim fld As Outlook.Folder = ASL_Tools.Get_ProjectFolder_In_ASLStoreInbox(proj)

        'if the project is not found then 
        'ask the user if they wish to create it or re enter a new
        'project number
        'If IsNothing(fld) Then
        'Dim rsp = MsgBox("Project folder not found in your email inbox." & vbLf & "Yes to Create Project folder" & 'vbLf & "No to Re-Enter Project Number", vbYesNo + vbCritical, "Error")
        'If rsp = vbNo Then
        'Exit Sub
        'End If
        ''MsgBox("No Project Found", vbCritical, "Error")
        ''create the project folder
        'fld = ASL_Tools.Create_ProjectFolder_In_ASLStoreInbox(proj)
        'If IsNothing(fld) Then
        'MsgBox("Error getting or creating project", vbCritical, "Error")
        'Exit Sub
        'End If
        'End If

        'If msgProp.messagetype = "se" Then
        ''the message is a sent item so we need to check there is a sent items folder under
        ''   the project folder.
        ''with the project folder
        ''get the sent folder
        'fldSent = ASL_Tools.Get_ProjectFolderSent(fld)

        'If IsNothing(fldSent) Then
        'fldSent = ASL_Tools.Create_ProjectFolderSent(fld)
        'If IsNothing(fldSent) Then
        'MsgBox("Error creating sent folder", vbCritical, "Error")
        'Exit Sub
        'End If
        'End If
        'MsgBox("Error moving stored sent items.", vbCritical, "Error")
        'Exit Sub
        'End If

        'Dim st As String = "(" & proj & ")(" & msgProp.timestamp & ")(" & msgProp.messagetype & ")"
        ''get the email recieved date store this messageKeyValue in a user property in the message
        ''ASL_Tools.Set_StampProperty(em2, st)

        ''msgProp.Set_TimeStampProperty()
        ''check to see if the user is in the office.
        ''if they are then save a copy of the email to the project folder
        ''if not then flag the message with the category offline
        ''offline messages can get copied to network at a later date.

        'Dim username As String = ASL_Tools.aslStore.DisplayName

        'If ASL_Tools.networkReady = True Then
        ''copy to network
        ' Dim di As System.IO.DirectoryInfo = ASL_Tools.Check_For_ProjectDirectoryEngJournal(proj, username)
        'If IsNothing(di) Then
        ''em2.Categories = "Offline"
        'Else
        ''use the messageKeyValue as the message name when saving to the 
        ''network
        'Dim diOld As System.IO.DirectoryInfo = ASL_Tools.Check_For_ProjectDirectoryEngJournal(msgProp.proj, username)
        ''ASL_Tools.remove_ASLMessage(msgProp, diOld)
        ''Dim emSave As Outlook.MailItem = em2
        ''emSave.SaveAs(di.FullName & "\" & st & ".msg")
        ''em2.Categories = ""
        ' End If
        'Else
        ''em2.Categories = "Offline"
        'End If

        ''if the message is a recieved item then
        ''else the message is a sent item

        'If msgProp.messagetype = "re" Then
        ''em2.Move(fld)
        'ElseIf msgProp.messagetype = "se" Then
        ''Dim em3 As Outlook.MailItem = em2.Copy()
        ''em3.Move(fldSent)
        ''em2.Delete()
        'End If


        Me.Close()
    End Sub
End Class